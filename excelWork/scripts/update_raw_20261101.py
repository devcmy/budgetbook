"""Append the complete October export to the memo-independent RAW workbook.

Preserve confirmed historical inputs and formulas. Recalculate disposable
macro-free copies and import caches into the original XLSM package only.
"""
from collections import Counter,defaultdict
from datetime import datetime
from hashlib import sha256
import argparse
import json
import math
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree as ET
import openpyxl
from openpyxl.utils import get_column_letter as col

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from update_cash_details_20261019 import import_changed_caches
from verify_reserve_redesign import recalculate
from verify_reserve_cash_alignment import verify_month,independent_balances
from balance_selected_ledger import check_view
from validate_excel_structure import validate_package
from verify_raw_auto_link_20261026 import errors
from raw_auto_link_20261026 import VARIABLE,BACK

WRAP='미래에셋_WRAP형CMA'
SHARED={'모임통장','가족통장','여행통장','20260606'}
HEADERS=('날짜','자산','분류','소분류','내용','금액(원)','수입/지출','메모','금액','화폐')
digest=lambda p:sha256(Path(p).read_bytes()).hexdigest()
asset=lambda name:'모임통장' if name=='20260606' else name
blank=lambda value:value in [None,'']


def read_source(path):
    w=openpyxl.load_workbook(path,data_only=True);s=w.active
    rows=list(s.values);assert tuple(rows[0])==HEADERS
    result=[]
    for sr,row in enumerate(rows[1:],2):
        if row[0] is None:continue
        dt,ac,category,sub,title,amount,kind,memo,original,currency=row
        assert isinstance(dt,datetime) and dt.year==2026 and dt.month==10,(sr,dt)
        assert currency=='KRW' and amount==original and isinstance(amount,(int,float)),sr
        result.append({'source_row':sr,'date':dt,'asset':asset(ac),'category':category,
                       'sub':sub or None,'title':title,'amount':amount,'kind':kind,'memo':memo or None})
    w.close();return result


def key(item):return item['date'],item['asset'],item['amount'],item['kind']


def plan(baseline,source):
    w=openpyxl.load_workbook(baseline,data_only=True);raw=w['RAW']
    lookup=defaultdict(list);all_rows=[]
    for r in range(2,4002):
        when=raw[f'A{r}'].value
        if not isinstance(when,datetime):continue
        all_rows.append(r)
        if when.year==2026 and when.month==10:
            lookup[when,asset(raw[f'C{r}'].value),raw[f'G{r}'].value,raw[f'H{r}'].value].append(r)
    matched=set();mapping=[];new=[];labels=[]
    for item in read_source(source):
        found=lookup.get(key(item),[])
        assert len(found)<=1,('Ambiguous existing identity',item['source_row'],found)
        if found:
            rr=found[0];assert rr not in matched,('Repeated source row',item['source_row'],rr)
            matched.add(rr);mapping.append({'source_row':item['source_row'],'row':rr,'status':'preserved'})
            if [raw[f'{c}{rr}'].value for c in ['D','E','F']]!=[item['category'],item['sub'],item['title']]:
                labels.append({'source_row':item['source_row'],'row':rr,'source_title':item['title'],'preserved_title':raw[f'F{rr}'].value})
        else:new.append(item)
    assert matched=={r for values in lookup.values() for r in values},'Existing October records missing from export'
    assert len(matched)==171 and len(new)==96,(len(matched),len(new))
    last=max(all_rows);assert last==2493 and len(all_rows)==2492
    new.sort(key=lambda v:(v['date'],v['source_row']))
    for r,item in enumerate(new,last+1):item['row']=r
    assert last+len(new)<=4001
    for item in new:mapping.append({'source_row':item['source_row'],'row':item['row'],'status':'added'})
    w.close();return {'new':new,'source_mapping':sorted(mapping,key=lambda v:v['source_row']),
                     'existing_label_differences_preserved':labels,'old_raw_count':len(all_rows),'old_october_count':len(matched)}


def explicit_class(item):
    if item['asset'] in SHARED:return '공금'
    if item['kind']!='지출':return None
    # Actual installment/regular bills remain fixed regardless of export category.
    if item['category']=='정기구독' or any(t in item['title'] for t in ['띵샵','애플케어']):return None
    if item['category']=='경조사/회비/선물/대인':return '경조·선물'
    if item['category']=='가전/전자/잡화':return '일회성전자'
    if item['category']=='여행/약속/공사' and item['sub'] in ['약속','만남']:return '약속·만남'
    return None


def prepare(baseline,source,candidate):
    metadata=plan(baseline,source);w=openpyxl.load_workbook(baseline,data_only=True)
    p=ExtendedPackage(baseline);ledger_row=250;wrap=[]
    for item in metadata['new']:
        r=item['row']
        assert w['RAW'][f'A{r}'].value is None
        values={'A':item['date'],'C':item['asset'],'D':item['category'],'E':item['sub'],
                'F':item['title'],'G':item['amount'],'H':item['kind'],'I':item['memo']}
        for c,value in values.items():p.write('RAW',f'{c}{r}',value)
        classification=explicit_class(item)
        if classification:p.write('RAW',f'J{r}',classification)
        if item['asset'] in SHARED:
            p.write('RAW',f'K{r}',0);p.write('RAW',f'L{r}','공동')
        # Generic vendors must not overwrite a differently identified old item:
        # the existing 6/36 AirPods installment stays separate from a new 1/36 bill.
        if item['title'] in ['구글플레이','롯데카드 띵샵(1/36)']:
            p.write('RAW',f'AA{r}',26)
            item['fixed_slot_override']=26
        incoming=item['category']==WRAP and item['asset']!=WRAP
        in_wrap=item['asset']==WRAP or incoming
        if not in_wrap:continue
        assert item['amount']>0
        assert w['잔액'][f'A{ledger_row}'].value is None
        transfer=item['kind'].startswith('이체')
        if incoming:
            p.write('RAW',f'S{r}','입금')  # Source-side transfer-out is WRAP's transfer-in.
            source_account,dest_account=item['asset'],WRAP
            source_category,dest_category='외부','생활비'
            title=f'{source_account} → 미래에셋 WRAP'
        else:
            source_account=WRAP;dest_account=asset(item['category']) if transfer else '외부'
            source_category,dest_category='생활비','외부' if transfer else '지출'
            title=f'미래에셋 WRAP → {dest_account}' if transfer else item['title']
        inputs={'A':item['date'],'B':'계좌이체' if transfer else '출금',
                'C':source_category,'D':dest_category,'E':item['amount'],'F':title,
                'J':source_account,'K':dest_account,'L':'WRAP계좌이체' if transfer else 'WRAP입출금','W':r}
        for c,value in inputs.items():p.write('잔액',f'{c}{ledger_row}',value)
        wrap.append({'raw_row':r,'ledger_row':ledger_row,'amount':item['amount'],'direction':'입금' if incoming else '출금',
                     'source_account':source_account,'destination_account':dest_account,'type':inputs['B']})
        ledger_row+=1
    assert len(wrap)==8
    # Preserve the pre-existing known/unknown date state as explicit metadata.
    # A free memo must not change the ledger date display via a text search.
    p.write(BACK,'AS1','원장 날짜 상태 · 메모와 분리')
    unknown=[]
    for r in range(64,501):
        state='날짜 미기록' if w['잔액'][f'U{r}'].value=='날짜 미기록' else None
        if state:unknown.append(r)
        p.write(BACK,f'AS{r}',state)
        p.write('잔액',f'U{r}',f'=IF(Q{r}="","",IF(\'{BACK}\'!AS{r}="날짜 미기록","날짜 미기록",A{r}))',True)
    p.sheet(BACK).find(Q('dimension')).set('ref','A1:AS31699')
    p.write('잔액','D3','2026-10-31 첨부 기준 · WRAP RAW 24건 연결. 월급의 국민은행→WRAP 이동은 수입에 재가산하지 않습니다.')
    p.write('10월','A70','10/1~31 RAW 267건 · 기존 171건 보존 + 신규 96건. 메모는 I열에만 입력합니다.')
    p.write('사용안내','A26','거래 기준일 2026-10-31 · 10월 RAW 267건. 월급·공동자금·투자 이체와 후속 매수를 분리해 집계합니다.')
    p.save(candidate);w.close()
    metadata.update(wrap_entries=wrap,unknown_ledger_date_rows=unknown,
                    date_status_independent_of_memo=True)
    return metadata


def verify(baseline,final,source,metadata):
    before=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(final,data_only=True)
    raw,month,calc,ledger=[w[t] for t in ['RAW','10월','10월계산기','잔액']]
    assert not errors(w),errors(w)[:20]
    assert len(w.sheetnames)==46 and '검증' not in w.sheetnames
    assert sum(isinstance(c.value,datetime) for c in raw['A'][1:])==2588
    source_rows=read_source(source)
    assert len(source_rows)==267
    for item in metadata['source_mapping']:
        r=item['row'];s=next(v for v in source_rows if v['source_row']==item['source_row'])
        assert (raw[f'A{r}'].value,raw[f'C{r}'].value,raw[f'G{r}'].value,raw[f'H{r}'].value)==key(s),(r,'source identity')
        if item['status']=='added':
            for c,name in [('D','category'),('E','sub'),('F','title'),('I','memo')]:assert raw[f'{c}{r}'].value==s[name],(r,c)
            assert raw[f'Q{r}'].value==1,(r,'new duplicate')
    # No historical input or memo is overwritten by the new export.
    for r in range(2,2494):
        for c in list(range(1,14)):
            assert before['RAW'].cell(r,c).value==raw.cell(r,c).value,(r,c,'old RAW')
    for r in range(64,250):
        for c in ['A','B','C','D','E','F','J','K','L','W','X']:
            assert before['잔액'][f'{c}{r}'].value==ledger[f'{c}{r}'].value,(r,c,'old ledger input')
        a,b=before['잔액'][f'U{r}'].value,ledger[f'U{r}'].value
        equal=math.isclose(a,b,rel_tol=0,abs_tol=1e-8) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else a==b
        assert equal,(r,'date display state',a,b)
    added=metadata['new'];new_variable=[v for v in added if raw[f'Y{v["row"]}'].value!=0]
    variable_delta=sum(v['amount'] for v in new_variable)
    rounded_delta=sum(math.ceil(v['amount']/1000)*1000 for v in new_variable)
    assert month['I33'].value==before['10월']['I33'].value+variable_delta
    assert calc['W12'].value==before['10월계산기']['W12'].value+rounded_delta
    assert month['I38'].value==calc['W12'].value
    for dr in [3,8,13,18,23,28]:
        for c in range(2,9):
            day=before['10월'].cell(dr,c).value
            if not isinstance(day,datetime) or day.month!=10:continue
            entries=[v for v in new_variable if v['date'].date()==day.date()]
            assert (month.cell(dr+2,c).value or 0)==(before['10월'].cell(dr+2,c).value or 0)+sum(v['amount'] for v in entries),(day,'daily actual')
            assert (month.cell(dr+1,c).value or 0)==(before['10월'].cell(dr+1,c).value or 0)+sum(math.ceil(v['amount']/1000)*1000 for v in entries),(day,'daily rounded')
    for m in [*range(1,10),11,12]:
        for addr in ['I33','H37','H38','H44','C41','E78','E79']:
            assert before[f'{m}월'][addr].value==w[f'{m}월'][addr].value,(m,addr)
        for addr in ['W12','W13']+[f'C{r}' for r in range(3,11)]:assert before[f'{m}월계산기'][addr].value==w[f'{m}월계산기'][addr].value,(m,addr)
    assert month['C37'].value==2776050
    assert month['N37'].value==150000 and month['N38'].value==20000 and month['N39'].value==120000
    assert month['N57'].value==250000 and month['N59'].value==20000
    assert month['C40'].value==before['10월']['C40'].value+94650
    assert month['H43'].value==before['10월']['H43'].value+2000
    assert month['N69'].value==239000
    expected_fixed=before['10월']['H37'].value+sum(raw[f'AB{v["row"]}'].value for v in added)
    assert month['E74'].value==month['H37'].value==expected_fixed
    assert month['O23'].value==10200  # Existing installment is not replaced by new headphones.
    assert month['O26'].value==5900+11004
    assert month['E79'].value==month['E78'].value==month['E74'].value+month['I33'].value+103600
    assert ledger['G7'].value==6284741 and ledger['G16'].value==8434741
    links=Counter(ledger[f'{c}{r}'].value for r in range(64,501) for c in ['W','X'] if ledger[f'{c}{r}'].value is not None)
    wrap_rows={r for r in range(2323,2590) if raw[f'R{r}'].value=='WRAP'}
    assert len(wrap_rows)==24 and all(links[r]==1 for r in wrap_rows)
    for item in metadata['wrap_entries']:
        rr,lr=item['raw_row'],item['ledger_row']
        assert ledger[f'I{lr}'].value=='OK' and ledger[f'E{lr}'].value==raw[f'G{rr}'].value
        assert raw[f'S{rr}'].value==item['direction']
        assert (ledger[f'G{lr}'].value or '')==(raw[f'I{rr}'].value or '')
    reserve=verify_month(w,10);view=check_view(w,10)
    assert w['예비비']['F13'].value==239000 and w['예비비']['I13'].value==2150000
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    for m in range(1,13):
        old=p.sheet(f'{m}월계산기').findall(Q('conditionalFormatting'))
        new=q.sheet(f'{m}월계산기').findall(Q('conditionalFormatting'))
        assert [ET.tostring(x) for x in old]==[ET.tostring(x) for x in new],(m,'calculator formatting rules')
    assert calc['AF3'].value==2649 and sum(calc[f'C{r}'].value for r in range(3,11))==2649
    for name in p.parts:
        if name=='xl/vbaProject.bin' or name.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/')) or name.endswith('.rels') or name=='xl/styles.xml':
            assert p.parts[name]==q.parts[name],name
    for r in range(64,501):
        assert 'SEARCH' not in (q.cell('잔액',f'U{r}').find(Q('f')).text or '')
    structure=validate_package(final);assert not structure['errors'],structure
    result={'date':'2026-11-01','as_of':'2026-10-31','latest_transaction':max(s['date'] for s in source_rows).isoformat(),
            'source_sha256':digest(source),'baseline_sha256':digest(baseline),'workbook_sha256':digest(final),
            'old_raw_records_preserved':2492,'old_october_records_preserved':171,'new_transactions':96,
            'october_records':267,'total_raw_records':2588,'new_personal_variable_count':len(new_variable),
            'new_personal_variable_actual':variable_delta,'new_personal_variable_tetris':rounded_delta,
            'calendar_actual':month['I33'].value,'calendar_tetris':calc['W12'].value,
            'salary_actual':month['C37'].value,'fixed_actual':month['E74'].value,'fixed_actual_and_planned':month['E75'].value,
            'actual_usage':month['E78'].value,'final_usage':month['E79'].value,
            'wrap_total':ledger['G16'].value,'living_balance':ledger['G7'].value,'wrap_raw_records':24,
            'new_wrap_inflows':2776050,'new_wrap_outflows':901000,
            'reserve_alignment':reserve,'selected_month_view':view,'structure':structure,
            'other_month_financial_values_preserved':True,'memo_single_source_preserved':True,
            'ledger_date_display_no_longer_reads_memo':True,'formula_errors':0,'vba_styles_controls_relationships_preserved':True,
            'calculator_conditional_rules_preserved':True,'calculator_additional_region_cells':49,
            'native_excel_ui_executed':False,'sharepoint_file_updated':False}
    before.close();w.close();return result


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for name in ['baseline','source','output','workdir','report']:parser.add_argument('--'+name,required=True)
    args=parser.parse_args();folder=Path(args.workdir);folder.mkdir(parents=True,exist_ok=True)
    candidate=folder/'candidate.xlsm';metadata=prepare(args.baseline,args.source,candidate)
    (folder/'import_metadata.json').write_text(json.dumps(metadata,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Prepared 96 RAW records and 8 WRAP ledger links.',flush=True)
    calculated=recalculate(candidate,folder/'calculation')
    import_changed_caches(candidate,calculated,args.output)
    result=verify(args.baseline,args.output,args.source,metadata)
    result['metadata']=metadata
    Path(args.report).write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:v for k,v in result.items() if k!='metadata'},ensure_ascii=False,indent=2))
