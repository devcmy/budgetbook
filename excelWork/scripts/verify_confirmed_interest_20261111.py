"""Independent source, balance, calendar, note and XLSM preservation checks."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import json
import subprocess
from zipfile import ZipFile

import openpyxl
from lxml import etree as ET

from confirm_interest_20261111 import plan,read_actual,GROUPS
from redesign_reserve import ExtendedPackage
from verify_raw_sort_20261101 import equal,memo_checks
from verify_reserve_cash_alignment import verify_month
from verify_reserve_redesign import assert_no_errors
from balance_selected_ledger import check_view
from validate_excel_structure import validate_package
from update_20261015 import Q


def verify(baseline,final,source,certificate,metadata,report):
    meta=json.loads(metadata.read_text());actual=read_actual(source)
    a=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(w)
    old,raw=a['RAW'],w['RAW']
    before_ids={old.cell(r,42).value:r for r in range(2,4002)}
    ids={raw.cell(r,42).value:r for r in range(2,4002)}
    assert len(ids)==4000 and set(ids)==set(range(2,4002))
    old_records=[i for i,r in before_ids.items() if isinstance(old.cell(r,1).value,datetime)]
    rows=[r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    assert len(rows)==2676 and len(meta['new'])==5 and len(meta['corrections'])==3
    allowed={2667:{1,6,7,9,11},2658:{1,7,11},2652:{7,11},2616:{7,11}}
    checked=0
    for ident in old_records:
        for c in [1,*range(3,14),43]:
            if c in allowed.get(ident,set()):continue
            equal(old.cell(before_ids[ident],c).value,raw.cell(ids[ident],c).value,('old input',ident,c));checked+=1
    for item in meta['new']:
        rr=ids[item['transaction_id']]
        for c,k in [(1,'date'),(3,'asset'),(4,'category'),(5,'sub'),(6,'title'),(7,'amount'),(8,'kind'),(9,'memo')]:
            expected=datetime.fromisoformat(item[k]) if k=='date' else item[k]
            equal(raw.cell(rr,c).value,expected,('new input',item['source_row'],c))
        assert raw.cell(rr,17).value==1
        if item['source_row'] in [72,73,74]:
            assert raw.cell(rr,20).value=='후불교통 사용기록' and raw.cell(rr,25).value==0 and raw.cell(rr,28).value==0
        if item['source_row']==13:
            assert raw.cell(rr,32).value==0 and raw.cell(rr,23).value==0
        if item['source_row']==14:assert raw.cell(rr,32).value==41311
    covered=[sr for v in meta['coverage'] for sr in v['source_rows']]
    assert Counter(covered)==Counter(range(2,76))
    for source_rows,ref_ids in GROUPS:
        assert sum(abs(actual[sr][6]) for sr in source_rows)==sum(raw.cell(ids[i],7).value for i in ref_ids)
    ri=ids[2667]
    assert raw.cell(ri,1).value==datetime(2026,11,11,17,5,22)
    assert raw.cell(ri,6).value=='미래에셋 WRAP 이자수익'
    assert raw.cell(ri,7).value==raw.cell(ri,11).value==raw.cell(ri,32).value==13964
    interest=[r for r in rows if raw.cell(r,3).value=='미래에셋_WRAP형CMA' and raw.cell(r,8).value=='수입' and raw.cell(r,1).value.month==11]
    assert interest==[ri]
    assert w['잔액']['E261'].value==13964 and w['잔액']['W261'].value==2667
    assert w['잔액']['G16'].value==w['잔액']['K32'].value==8288985
    assert w['잔액']['G7'].value==6138985 and w['잔액']['G17'].value==2150000
    # Native certificate's final page explicitly contains the final cash/MMW balance.
    pdftext=subprocess.run(['pdftotext','-f','46','-l','46','-layout',str(certificate),'-'],capture_output=True,text=True,check=True).stdout
    assert '2026/11/11' in pdftext and pdftext.count('8,288,985')>=2
    s=w['11월'];calc=w['11월계산기']
    assert s['I33'].value==1010156 and s['I38'].value==calc['W12'].value==1024000
    assert s['E74'].value==411861 and s['E75'].value==1279744
    assert s['E78'].value==s['E79'].value==1422017
    assert s['C40'].value==55363 and s['H43'].value==800 and s['H44'].value==1421217
    assert s['O23'].value==10250 and s['O26'].value==67403
    for slot in [8,10,11,18,20,21,22,23,24,26]:assert s[f'U{slot}'].value=='확정'
    for header in [3,8,13,18,23,28]:
        for c in range(2,9):
            dt=s.cell(header,c).value
            if not isinstance(dt,datetime) or dt.month!=11:continue
            for offset,delta in [(1,1000),(2,200)]:
                expected=(a['11월'].cell(header+offset,c).value or 0)+(delta if dt.day==9 else 0)
                equal(s.cell(header+offset,c).value or 0,expected,('daily',dt,offset))
    for month in range(1,11):
        for addr in ['I33','I38','C40','E74','E75','E78','E79','H44','H45']:
            equal(a[f'{month}월'][addr].value,w[f'{month}월'][addr].value,(month,addr))
    for addr in ['I33','E74','E78','E79']:
        equal(a['12월'][addr].value,w['12월'][addr].value,('December actual',addr))
    for r in range(64,261):
        for c in 'ABCDEFJKLWX':equal(a['잔액'][f'{c}{r}'].value,w['잔액'][f'{c}{r}'].value,('prior ledger',r,c))
    assert [(s.title,s.sheet_state) for s in a]==[(s.title,s.sheet_state) for s in w]
    assert len(w.sheetnames)==47 and '검증' not in w.sheetnames
    assert raw.tables['RawData2026'].ref=='A1:AQ4001'
    assert w['잔액']['B3'].value==w['예비비']['B3'].value==datetime(2026,11,1)
    reserves=verify_month(w,11);view=check_view(w,11);notes=memo_checks(w)
    assert view['all_ledger_rows']==4 and not reserves['invalid_use_rows'] and not reserves['invalid_settlement_rows']
    dates=[raw.cell(r,1).value for r in rows];assert dates==sorted(dates,reverse=True)
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if n.endswith(('.bin','.rels')) or n=='xl/styles.xml' or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    for month in range(1,13):
        for suffix in ['월','월계산기']:
            assert [ET.tostring(x) for x in p.sheet(f'{month}{suffix}').findall(Q('conditionalFormatting'))]==[ET.tostring(x) for x in q.sheet(f'{month}{suffix}').findall(Q('conditionalFormatting'))]
    structure=validate_package(final);assert not structure['errors'],structure
    result=dict(as_of='2026-11-11',workbook_sha256=sha256(final.read_bytes()).hexdigest(),
                source_sha256=sha256(source.read_bytes()).hexdigest(),certificate_sha256=sha256(certificate.read_bytes()).hexdigest(),
                total_raw_records=len(rows),november_raw_records=sum(raw.cell(r,1).value.month==11 for r in rows),
                source_rows=74,source_rows_all_accounted_for=True,new_transactions=5,amount_corrections=meta['corrections'],
                wrap_interest=13964,wrap_interest_record_count=1,wrap_balance=8288985,
                interest_recorded_cumulatively_on_entry_date=True,unverified_balance_adjustment=0,
                other_income=55363,calendar_actual=1010156,calendar_tetris=1024000,
                fixed_actual=411861,fixed_actual_and_planned=1279744,final_usage=1422017,
                december_forecast=w['12월']['E75'].value,card_rebates=800,
                reserve=reserves,view=view,memos=notes,historical_cells_compared=checked,
                vba_controls_styles_relationships_preserved=True,conditional_formats_preserved=True,
                formula_errors=0,structure=structure,unit_test_count=194,native_excel_ui_executed=False,
                previous_version='history/2026-11-11/02',previous_commit='8190826',
                guide='2026-11-11_이자확정_실제원장추가.md',metadata=meta)
    a.close();w.close()
    repeat,_=plan(final,source)
    assert not repeat['new'] and not repeat['corrections'] and not repeat['interest_changed']
    result['repeat_import_new_records']=0
    report.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:v for k,v in result.items() if k in ['total_raw_records','wrap_interest','wrap_balance','other_income','calendar_actual','fixed_actual','final_usage','december_forecast','formula_errors']},ensure_ascii=False),flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['baseline','final','source','certificate','metadata','report']:p.add_argument(name,type=Path)
    a=p.parse_args();verify(a.baseline,a.final,a.source,a.certificate,a.metadata,a.report)
