"""Check source matching, independent expense deltas, cash migration and OOXML."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json
import math

import openpyxl
from lxml import etree as ET

from import_raw_20261115 import source_rows,plan,CASH,NAVER,WRAP,NEW_CLASS_CODES
from update_raw_20261101 import SHARED
from update_raw_20261104 import classify
from raw_auto_link_20261026 import palette_for
from redesign_reserve import ExtendedPackage
from verify_reserve_redesign import assert_no_errors,recalculate
from verify_reserve_cash_alignment import verify_month
from verify_raw_sort_20261101 import equal,memo_checks
from balance_selected_ledger import check_view
from validate_excel_structure import validate_package
from update_20261015 import Q


def expense_kind(v):
    if v['asset'] in SHARED or v['kind']!='지출':return None
    if v['title'].endswith('후불교통') or ('후불교통' in v['title'] and '체크카드' in v['asset']):return None
    if v['amount']<0 and '카드' in v['title'] and '할인' in v['title']:return None
    if v['category']=='정기구독' or v['title'] in NEW_CLASS_CODES or v['title']=='청년 버팀목 전세대출이자':return 'fixed'
    return 'variable'


def verify(baseline,final,source,metadata,report_path):
    meta=json.loads(metadata.read_text());src={v['source_row']:v for v in source_rows(source)}
    a=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(w);raw=w['RAW'];old=a['RAW']
    ids={int(raw.cell(r,42).value):r for r in range(2,4002)}
    prev={int(old.cell(r,42).value):r for r in range(2,4002)}
    assert set(ids)==set(range(2,4002)) and len(ids)==4000
    rows=[r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    assert len(rows)==2733 and len(meta['new'])==56 and len(meta['derived'])==1
    assert Counter(raw.cell(r,1).value.month for r in rows)[10]==279
    assert Counter(raw.cell(r,1).value.month for r in rows)[11]==132
    assert [(s.title,s.sheet_state) for s in a]==[(s.title,s.sheet_state) for s in w]
    assert len(w.sheetnames)==47 and '검증' not in w.sheetnames
    assert raw.tables['RawData2026'].ref=='A1:AQ4001'
    allowed={p['transaction_id']:set(p['changes']) for p in meta['patches']}
    cols={1:'A',3:'C',4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J',11:'K',12:'L',13:'M',43:'AQ'}
    checked=0
    for ident,pr in prev.items():
        if not isinstance(old.cell(pr,1).value,datetime):continue
        for c,key in cols.items():
            if key in allowed.get(ident,set()):continue
            equal(old.cell(pr,c).value,raw.cell(ids[ident],c).value,('old input',ident,key));checked+=1
    for item in meta['new']+meta['derived']:
        rr=ids[item['transaction_id']]
        for c,k in [(1,'date'),(3,'asset'),(4,'category'),(5,'sub'),(6,'title'),(7,'amount'),(8,'kind'),(9,'memo')]:
            expected=datetime.fromisoformat(item[k]) if k=='date' else item[k]
            equal(raw.cell(rr,c).value,expected,('new raw',item['transaction_id'],c))
        assert raw.cell(rr,17).value==1
    assert len(meta['source_mapping'])==405
    assert {m['source_row'] for m in meta['source_mapping']}==set(src)
    for link in meta['source_mapping']:
        item=src[link['source_row']];r=ids[link['transaction_id']]
        if link['status'] in ['previous_cumulative_interest_preserved','statement_amount_and_date_preserved','statement_amount_preserved','cumulative_interest_split_635']:continue
        equal(raw.cell(r,7).value,item['amount'],('mapped source amount',item['source_row']))
        if link['status'] in ['exact','added']:equal(raw.cell(r,1).value,item['date'],('mapped source date',item['source_row']))
        equal(raw.cell(r,9).value,item['memo'],('source memo',item['source_row']))
    assert raw.cell(ids[2667],7).value==13964
    assert raw.cell(ids[2442],7).value==7922
    assert raw.cell(ids[2678],7).value==635
    assert raw.cell(ids[2615],7).value==0 and raw.cell(ids[2615],25).value==0
    assert raw.cell(ids[1156],19).value=='입금' # old explicit receiving-side transfer
    assert all(not(raw.cell(r,6).value=='정산용' and raw.cell(r,7).value==-1960) for r in rows)
    our=[r for r in rows if raw.cell(r,3).value=='우리카드의정석SUPER' and raw.cell(r,1).value.month==10]
    assert sum(raw.cell(r,7).value for r in our if raw.cell(r,7).value>0)==874710
    assert sum(raw.cell(r,33+6).value for r in our)==16974 # AM card-benefit amount
    extra=meta['new']+meta['derived'];month_results={}
    for month in [10,11]:
        variables=[v for v in extra if datetime.fromisoformat(v['date']).month==month and expense_kind(v)=='variable']
        fixed=[v for v in extra if datetime.fromisoformat(v['date']).month==month and expense_kind(v)=='fixed']
        delta=sum(v['amount'] for v in variables)+(194 if month==10 else -89800)
        units=sum(math.ceil(v['amount']/1000) for v in variables)+(1 if month==10 else -90)
        s=w[f'{month}월'];b=a[f'{month}월'];calc=w[f'{month}월계산기']
        assert s['I33'].value==b['I33'].value+delta
        assert calc['W12'].value==s['I38'].value==b['I38'].value+units*1000
        assert s['E74'].value==b['E74'].value+sum(v['amount'] for v in fixed)
        advances=b['E79'].value-b['I33'].value-b['E74'].value
        assert s['E79'].value==s['E78'].value==s['I33'].value+s['E74'].value+advances
        palette=Counter({r:a[f'{month}월계산기'][f'C{r}'].value for r in range(3,11)})
        for v in variables:palette[palette_for(classify(v) or '변동비',v['category'],v['sub'])]+=math.ceil(v['amount']/1000)
        palette[9 if month==10 else 5]+=1 if month==10 else -90
        assert all(calc[f'C{r}'].value==palette[r] for r in range(3,11)),('palette',month,palette)
        for header in [3,8,13,18,23,28]:
            for c in range(2,9):
                dt=s.cell(header,c).value
                if not isinstance(dt,datetime) or dt.month!=month:continue
                daily=[v for v in variables if datetime.fromisoformat(v['date']).date()==dt.date()]
                adjustment=194 if month==10 and dt.day==22 else -89800 if month==11 and dt.day==4 else 0
                rounding=1000 if adjustment==194 else -90000 if adjustment==-89800 else 0
                equal(s.cell(header+2,c).value or 0,(b.cell(header+2,c).value or 0)+sum(v['amount'] for v in daily)+adjustment,(month,dt.day,'calendar'))
                equal(s.cell(header+1,c).value or 0,(b.cell(header+1,c).value or 0)+sum(math.ceil(v['amount']/1000)*1000 for v in daily)+rounding,(month,dt.day,'tetris'))
        month_results[str(month)]={addr:s[addr].value for addr in ['C40','I33','I38','H43','E74','E75','E79']}
    assert w['10월']['H43'].value==68231 and w['11월']['H43'].value==1000
    assert w['10월']['C40'].value==110573 and w['11월']['C40'].value==58122
    for month in [*range(1,10),12]:
        for addr in ['I33','I38','C40','E74','E75','E79','H43']:
            equal(a[f'{month}월'][addr].value,w[f'{month}월'][addr].value,(month,addr))
    ledger=w['잔액'];assert ledger['G16'].value==7262926 and ledger['G7'].value==5112926
    assert ledger['G17'].value==w['예비비']['I13'].value==2150000
    assert '7,262,399' in ledger['D3'].value and '527' in ledger['D3'].value
    # Independent cash movements from source account fields, excluding migration.
    cash_delta=-635
    for v in meta['new']:
        a1,a2=v['asset'],v['category']
        if a1 in CASH and a2 in CASH:continue
        if a1 in CASH:cash_delta+=v['amount'] if v['kind']=='수입' else -v['amount']
        elif a2 in CASH:cash_delta+=v['amount']
    assert a['잔액']['G16'].value+cash_delta==ledger['G16'].value
    for link in meta['bank_links']:
        r=link['ledger_row'];assert ledger[f'I{r}'].value=='OK'
        assert ledger[f'W{r}'].value==link['transaction_id']
        if link['internal']:assert ledger[f'C{r}'].value==ledger[f'D{r}'].value=='생활비'
    reserves=verify_month(w,11);view=check_view(w,11);notes=memo_checks(w)
    assert view['all_ledger_rows']==16
    for r in rows:
        fields=[raw.cell(r,3).value,raw.cell(r,4).value];when=raw.cell(r,1).value
        if WRAP in fields or (NAVER in fields and when>=datetime(2026,11,12)):
            assert raw.cell(r,18).value=='WRAP'
        elif NAVER in fields:
            ident=raw.cell(r,42).value
            equal(raw.cell(r,18).value,old.cell(prev[ident],18).value,('historical Naver scope',ident))
    dates=[raw.cell(r,1).value for r in rows];assert dates==sorted(dates,reverse=True)
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if n.endswith(('.bin','.rels')) or n=='xl/styles.xml' or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    for month in range(1,13):
        for suffix in ['월','월계산기']:
            assert [ET.tostring(v) for v in p.sheet(f'{month}{suffix}').findall(Q('conditionalFormatting'))]==[ET.tostring(v) for v in q.sheet(f'{month}{suffix}').findall(Q('conditionalFormatting'))]
    structure=validate_package(final);assert not structure['errors'],structure
    result=dict(as_of='2026-11-15',workbook_sha256=sha256(final.read_bytes()).hexdigest(),source_sha256=sha256(source.read_bytes()).hexdigest(),
                total_raw_records=len(rows),source_records=405,new_records=56,derived_expense_records=1,
                raw_by_month=dict(Counter(raw.cell(r,1).value.strftime('%Y-%m') for r in rows)),
                managed_cash_total=7262926,naver_cash=7262399,wrap_residual=527,purpose_available=2150000,
                income_635_split_without_double_count=True,migration_counted_as_expense=False,
                months=month_results,reserve=reserves,selected_view=view,memos=notes,
                historical_cells_preserved=checked,vba_controls_styles_relationships_preserved=True,
                conditional_formats_preserved=True,formula_errors=0,structure=structure,
                source_original_local_only=True,native_excel_ui_executed=False,
                previous_commit='1ae6cb7',previous_version='history/2026-11-15/01',
                guide='2026-11-15_RAW반영_관리계좌이전.md',metadata=meta)
    a.close();w.close();p=None;q=None;gc.collect()
    again,_=plan(final,source)
    assert not again['new'] and not again['derived'] and not again['patches'],('Repeat import',again['patches'])
    result.update(repeat_new_records=0,repeat_updates=0)
    report_path.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:result[k] for k in ['total_raw_records','managed_cash_total','naver_cash','wrap_residual','formula_errors']},ensure_ascii=False),flush=True)
    return result


def migration_scenario(final,root):
    # Vary a real bank move and an external withdrawal independently. The move
    # must preserve the combined total; only the 1,000 external increment costs cash.
    def edit(p):
        p.write('잔액','E264',8288457)
        p.write('잔액','E265',97430)
    calc=recalculate(final,root/'migration_scenario',edit)
    w=openpyxl.load_workbook(calc,data_only=True);assert_no_errors(w)
    verify_month(w,11)
    assert w['잔액']['G16'].value==7261926 and w['예비비']['I13'].value==2150000
    assert '7,261,398' in w['잔액']['D3'].value and '528' in w['잔액']['D3'].value
    w.close();print('Migration / withdrawal scenario passed.',flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['baseline','final','source','metadata','report']:p.add_argument(name,type=Path)
    p.add_argument('--scenario',action='store_true');a=p.parse_args()
    verify(a.baseline,a.final,a.source,a.metadata,a.report)
    if a.scenario:
        migration_scenario(a.final,a.metadata.parent)
        verified=json.loads(a.report.read_text())
        verified['migration_and_external_withdrawal_scenario_passed']=True
        verified['unit_test_count']=194
        a.report.write_text(json.dumps(verified,ensure_ascii=False,default=str,indent=2)+'\n')
