"""Charge-month postpaid transport + scoped November import on the user's file.

Preserve the original RAW date/type and stable transaction IDs. A card's usage
summary is evidence, while its bank settlement is the fixed-cost recognition
event. The latter is a narrowly identified exception to the transfer exclusion;
ordinary card settlements remain excluded. Neither rule reads free memos.
"""
from collections import Counter,defaultdict
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json
import re

from lxml import etree as ET
import openpyxl
from openpyxl.utils import get_column_letter as col

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from raw_sort_id_20261101 import current,field,reorder,TABLE,END
from update_raw_20261101 import explicit_class,asset,HEADERS
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches

DEFERRED='후불교통 사용기록'
digest=lambda path:sha256(Path(path).read_bytes()).hexdigest()


def compact(value):return str(value or '').replace(' ','').replace('-','').upper()


def program_card(value):
    value=compact(value)
    return 'K패스' in value or '모두의카드' in value


def usage(item):
    title=compact(item['title'])
    return (item['kind']=='지출' and program_card(item['asset'])
            and ('후불교통' in title or '교통카드후불' in title))


def settlement(item):
    return (item['kind'] in ['지출','이체지출','이체출금']
            and compact(item['title'])=='카드정산결제' and program_card(item['category']))


def excel_card(value):
    normalized=f'UPPER(SUBSTITUTE(SUBSTITUTE({value}," ",""),"-",""))'
    return f'OR(ISNUMBER(SEARCH("K패스",{normalized})),ISNUMBER(SEARCH("모두의카드",{normalized})))'


def usage_condition():
    title=f'SUBSTITUTE({current("내용")}," ","")'
    return (f'AND({current("수입/지출")}="지출",{excel_card(current("자산"))},'
            f'OR(ISNUMBER(SEARCH("후불교통",{title})),ISNUMBER(SEARCH("교통카드후불",{title}))))')


def settlement_condition():
    kind=current('수입/지출')
    return (f'AND(OR({kind}="지출",{kind}="이체지출",{kind}="이체출금"),'
            f'SUBSTITUTE({current("내용")}," ","")="카드정산결제",{excel_card(current("분류"))})')


def classification_formula(original):
    return f'IF({usage_condition()},"{DEFERRED}",IF({settlement_condition()},"고정비",{original.lstrip("=")}))'


def slot_formula(original):
    return f'IF(AND({current("적용분류")}="고정비",{settlement_condition()}),11,{original.lstrip("=")})'


def fixed_amount_formula():
    return (f'IF(AND({current("적용분류")}="고정비",'
            f'OR({current("수입/지출")}="지출",{settlement_condition()}),'
            f'{current("적용결제")}="본인",{current("동일거래순번")}=1),{current("적용부담")},0)')


def read_source(path):
    w=openpyxl.load_workbook(path,data_only=True);s=w.active
    assert tuple(next(s.values))==HEADERS
    result=[]
    for sr,v in enumerate(s.iter_rows(min_row=2,values_only=True),2):
        if not v[0]:continue
        assert isinstance(v[0],datetime) and v[0].year==2026 and v[0].month in [9,10,11]
        assert v[9]=='KRW' and v[5]==v[8]
        result.append(dict(source_row=sr,date=v[0],asset=asset(v[1]),category=v[2],sub=v[3] or None,
                           title=v[4],amount=v[5],kind=v[6],memo=v[7] or None))
    w.close();return result


def key(item):return item['date'],asset(item['asset']),item['amount'],item['kind']


def make_plan(baseline,source):
    w=openpyxl.load_workbook(baseline,data_only=True);raw=w['RAW']
    ids={raw.cell(r,42).value:r for r in range(2,END+1)}
    assert set(ids)==set(range(2,END+1))
    existing=defaultdict(list)
    for ident,r in ids.items():
        dt=raw.cell(r,1).value
        if isinstance(dt,datetime):
            existing[(dt,asset(raw.cell(r,3).value),raw.cell(r,7).value,raw.cell(r,8).value)].append(ident)
    source_rows=read_source(source)
    bills=[v for v in source_rows if settlement(v) and v['date'].month==10 and v['amount']==136350]
    uses=[v for v in source_rows if usage(v) and v['date'].month==9 and v['amount']==136350]
    assert len(bills)==len(uses)==1,(bills,uses)
    bill=bills[0];assert len(existing[key(bill)])==1
    bill_id=existing[key(bill)][0]
    matched=[];added=[];deferred=[]
    for item in source_rows:
        hits=existing.get(key(item),[])
        assert len(hits)<=1,('ambiguous source identity',item['source_row'],hits)
        if hits:matched.append({'source_row':item['source_row'],'transaction_id':hits[0]});continue
        if item['date'].month==11 or item is uses[0]:added.append(dict(item))
        else:deferred.append(dict(item))
    assert len(added)==9 and len([v for v in added if v['date'].month==11])==8
    assert len(matched)==504 and len(deferred)==17
    empty=sorted(i for i,r in ids.items() if raw.cell(r,1).value in [None,''])
    added.sort(key=lambda v:(v['date'],v['source_row']))
    for ident,item in zip(empty,added):item['transaction_id']=ident;item['input_row']=ids[ident]
    assert len(empty)>=len(added)
    assert not any(v['asset']=='미래에셋_WRAP형CMA' or v['category']=='미래에셋_WRAP형CMA' for v in added)
    source_keys={key(v) for v in source_rows}
    unmatched_old_september=[i for i,r in ids.items() if isinstance(raw.cell(r,1).value,datetime)
                            and raw.cell(r,1).value.month==9
                            and (raw.cell(r,1).value,asset(raw.cell(r,3).value),raw.cell(r,7).value,raw.cell(r,8).value) not in source_keys]
    w.close()
    return {'new':added,'matched':matched,'deferred_unrelated_september':deferred,
            'existing_september_not_in_new_export_preserved':unmatched_old_september,
            'billing_transaction_id':bill_id,'source_usage':uses[0],
            'scope':'Only requested September transport evidence plus November transactions; unrelated September differences are not merged.'}


def remove_calc_chain(p):
    p.parts.pop('xl/calcChain.xml',None)
    rels=p.tree('xl/_rels/workbook.xml.rels')
    for node in list(rels):
        if node.get('Type','').endswith('/calcChain'):rels.remove(node)
    types=p.tree('[Content_Types].xml')
    for node in list(types):
        if node.get('PartName')=='/xl/calcChain.xml':types.remove(node)


def prepare(baseline,source,output):
    metadata=make_plan(baseline,source)
    w=openpyxl.load_workbook(baseline);p=ExtendedPackage(baseline)
    raw=w['RAW'];ids={raw.cell(r,42).value:r for r in range(2,END+1)}
    for item in metadata['new']:
        r=item['input_row']
        for c,k in [('A','date'),('C','asset'),('D','category'),('E','sub'),('F','title'),('G','amount'),('H','kind'),('I','memo')]:
            p.write('RAW',f'{c}{r}',item[k])
        classification=explicit_class(item)
        if classification:p.write('RAW',f'J{r}',classification)
    # Preserve manual AA overrides (two unrelated installment/subscription rows).
    changed=Counter()
    for r in range(2,END+1):
        for c,factory in [('AJ',classification_formula),('AA',slot_formula),('AB',lambda old:fixed_amount_formula())]:
            old=raw[f'{c}{r}'].value
            if not isinstance(old,str) or not old.startswith('='):
                assert c=='AA',('Unexpected manual helper',r,c,old)
                continue
            p.write('RAW',f'{c}{r}',factory(old),True);changed[c]+=1
    # Existing input J2438 is already an explicit fixed-cost classification.
    bill_row=ids[metadata['billing_transaction_id']]
    assert raw[f'J{bill_row}'].value=='고정비'
    old_note=raw[f'I{bill_row}'].value or ''
    use_id=next(v['transaction_id'] for v in metadata['new'] if v['date'].month==9)
    confirmation=(f'11/2 확인: 9/30 후불교통 사용기록 거래ID {use_id}와 동일한 136,350원. '
                  'K패스=모두의카드. 10/12 토스뱅크 결제 기준으로 10월 고정비에 한 번만 반영하며 9월에 재차 합산하지 않습니다.')
    p.write('RAW',f'I{bill_row}','\n'.join(x for x in [old_note,confirmation] if x))
    metadata['billing_note_before']=old_note;metadata['billing_note_confirmation']=confirmation
    table_part=next(n for n in p.parts if n.startswith('xl/tables/') and n.endswith('.xml')
                    and ET.fromstring(p.parts[n]).get('displayName')==TABLE)
    table=p.tree(table_part)
    for column in table.find(Q('tableColumns')):
        f=column.find(Q('calculatedColumnFormula'))
        if f is None:continue
        name=column.get('name')
        if name=='기본분류':f.text=classification_formula(f.text)
        if name=='고정항목':f.text=slot_formula(f.text)
        if name=='고정금액':f.text=fixed_amount_formula()
    category_last=max(r for r in range(50,100) if raw[f'AQ{r}'].value is not None)
    p.write('RAW',f'AQ{category_last+1}',DEFERRED)
    p.cell('RAW',f'AQ{category_last+1}').set('s',p.cell('RAW',f'AQ{category_last}').get('s','0'))
    for dv in p.sheet('RAW').find(Q('dataValidations')):
        if dv.get('sqref')=='J2:J4001':dv.find(Q('formula1')).text=f'$AQ$50:$AQ${category_last+1}'
    total=f'SUMIFS({field("고정금액")},{field("월")},$T$1,{field("고정항목")},11)'
    count=f'COUNTIFS({field("월")},$T$1,{field("고정항목")},11,{field("고정금액")},">0")'
    for m in [10,11,12]:
        sn=f'{m}월'
        p.write(sn,'L11','모두의카드');p.write(sn,'M11','토스뱅크');p.write(sn,'N11','전월 사용분')
        formula=total if m==10 else f'IF({count}>0,{total},\'{m-1}월\'!O11)'
        p.write(sn,'O11',formula,True)
        fallback='미입력' if m==10 else '예상'
        p.write(sn,'U11',f'IF({count}>0,"확정","{fallback}")',True)
    p.write('10월','A70','모두의카드(기존 K패스): 9월 사용 136,350원은 10/12 청구 기준 10월 고정비에 1회 반영합니다.')
    p.write('11월','A70','11/1~2 RAW 8건 · 변동 6건 / 고정비 1건 / 이자 1건. 교통비 예상액과 실제 확정액은 구분합니다.')
    p.write('사용안내','A26','거래 기준일 2026-11-02 · 11월 RAW 8건. 모두의카드=K패스, 후불교통은 사용기록과 청구월 고정비를 분리합니다.')
    remove_calc_chain(p)
    calc=p.tree('xl/workbook.xml').find(Q('calcPr'))
    calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    # Use the user's descending order; stable IDs carry every old relationship.
    p.save(output)
    fresh=openpyxl.load_workbook(output,data_only=True)
    metadata['sort']=reorder(p,fresh,'descending');fresh.close()
    p.save(output);w.close();gc.collect()
    metadata['helper_formula_counts_updated']=dict(changed)
    return metadata


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for a in ['baseline','source','output','metadata']:parser.add_argument(a,type=Path)
    a=parser.parse_args()
    result=prepare(a.baseline,a.source,a.output)
    a.metadata.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({'new_records':len(result['new']),'payment_id':result['billing_transaction_id'],
                      'deferred_unrelated_september':len(result['deferred_unrelated_september']),
                      'helper_formula_counts_updated':result['helper_formula_counts_updated']},ensure_ascii=False),flush=True)
