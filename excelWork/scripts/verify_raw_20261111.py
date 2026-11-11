"""Check source preservation, independent monthly totals, and package integrity."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import json
import math

from lxml import etree as ET
import openpyxl

from redesign_reserve import ExtendedPackage
from update_raw_20261111 import read_source, plan
from update_raw_20261104 import classify
from update_raw_20261101 import SHARED
from raw_auto_link_20261026 import palette_for
from verify_raw_sort_20261101 import equal, memo_checks
from verify_reserve_cash_alignment import verify_month
from balance_selected_ledger import check_view
from verify_reserve_redesign import assert_no_errors
from validate_excel_structure import validate_package
from update_20261015 import Q


def fixed(item):
    title=item['title'].replace(' ','')
    return ((item['kind']=='지출' and (item['category']=='정기구독' or
            title.startswith('11번가') or '띵샵' in title or
            title in ['SKT','LG헬로비전','코나아이(MONA)','G마켓(1/6)'])) or
            (title=='카드정산결제' and item['category']=='토스뱅크K패스체크카드'))


def verify(baseline, final, source, metadata, report):
    meta=json.loads(metadata.read_text());items=read_source(source)
    a=openpyxl.load_workbook(baseline,data_only=True)
    w=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(w)
    old,raw=a['RAW'],w['RAW']
    ids={raw.cell(r,42).value:r for r in range(2,4002)}
    old_ids={old.cell(r,42).value:r for r in range(2,4002)}
    assert len(ids)==4000 and set(ids)==set(range(2,4002))
    count=sum(isinstance(raw.cell(r,1).value,datetime) for r in range(2,4002))
    assert count==2671 and len(items)==82 and len(meta['new'])==54
    assert [(s.title,s.sheet_state) for s in a]==[(s.title,s.sheet_state) for s in w]
    assert '검증' not in w.sheetnames and len(w.sheetnames)==47
    assert raw.tables['RawData2026'].ref=='A1:AQ4001' and not raw.protection.sheet
    checked=0
    for ident,ar in old_ids.items():
        if not isinstance(old.cell(ar,1).value,datetime):continue
        for c in [1,*range(3,14),43]:
            equal(old.cell(ar,c).value,raw.cell(ids[ident],c).value,('old input',ident,c));checked+=1
    source_map={v['source_row']:v for v in items}
    for link in meta['source_mapping']:
        item=source_map[link['source_row']];rr=ids[link['transaction_id']]
        fields=[(1,'date'),(3,'asset'),(7,'amount'),(8,'kind')]
        if link['status']=='added':fields += [(4,'category'),(5,'sub'),(6,'title'),(9,'memo')]
        for c,k in fields:equal(raw.cell(rr,c).value,item[k],('source',item['source_row'],c))
        assert raw.cell(rr,17).value==1
    personal=[v for v in items if v['asset'] not in SHARED]
    variable=[v for v in personal if v['kind']=='지출' and not fixed(v)]
    expected_fixed=sum(v['amount'] for v in personal if fixed(v))
    expected_variable=sum(v['amount'] for v in variable)
    rounded=sum(math.ceil(v['amount']/1000)*1000 for v in variable)
    s,calc=w['11월'],w['11월계산기']
    assert expected_fixed==s['E74'].value==411813
    assert expected_variable==s['I33'].value==1009956
    assert rounded==calc['W12'].value==s['I38'].value==1023000
    assert s['E78'].value==s['E79'].value==expected_fixed+expected_variable==1421769
    assert s['E75'].value==1279696
    assert s['C40'].value==12890 and s['H43'].value==800
    assert s['H44'].value==1420969
    checked_days=[]
    for header in [3,8,13,18,23,28]:
        for c in range(2,9):
            dt=s.cell(header,c).value
            if not isinstance(dt,datetime) or dt.month!=11:continue
            daily=[v for v in variable if v['date'].date()==dt.date()]
            equal(s.cell(header+2,c).value or 0,sum(v['amount'] for v in daily),(dt,'daily actual'))
            equal(s.cell(header+1,c).value or 0,sum(math.ceil(v['amount']/1000)*1000 for v in daily),(dt,'daily tetris'))
            checked_days.append(dt.day)
    assert sorted(checked_days)==list(range(1,31))
    palette=Counter()
    for item in variable:
        cl=classify(item) or '변동비'
        palette[palette_for(cl,item['category'],item['sub'])]+=math.ceil(item['amount']/1000)
    for r in range(3,11):assert calc[f'C{r}'].value==palette[r],(r,palette[r],calc[f'C{r}'].value)
    for row,value in [(8,53620),(10,27610),(11,137300),(18,13875),(20,303),(21,18000),(22,74700),(23,10200),(24,8800),(26,67405)]:
        assert s[f'O{row}'].value==value and s[f'U{row}'].value=='확정'
    for month in [*range(1,11),12]:
        for addr in ['C37','C40','C41','I33','E74','E75','E78','E79','H43','H44','H45']:
            if month==12 and addr=='E75':
                # Existing December forecast formula references November O11.
                assert w['12월'][addr].value==a['12월'][addr].value+950==1207010
                assert w['12월']['O11'].value==137300 and w['12월']['U11'].value=='예상'
                continue
            equal(a[f'{month}월'][addr].value,w[f'{month}월'][addr].value,(month,addr))
        for addr in ['W12','W13']+[f'C{r}' for r in range(3,11)]:
            equal(a[f'{month}월계산기'][addr].value,w[f'{month}월계산기'][addr].value,(month,addr))
    for r in range(64,258):
        for c in 'ABCDEFJKLWX':equal(a['잔액'][f'{c}{r}'].value,w['잔액'][f'{c}{r}'].value,('old ledger',r,c))
    reserve=verify_month(w,10)
    memos=memo_checks(w)
    dates=[raw.cell(r,1).value for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    assert dates==sorted(dates,reverse=True)
    assert [w['입력·연결'][f'H{r}'].value for r in range(8,18)]==[raw.cell(r,42).value for r in range(2,12)]
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if n.endswith(('.bin','.rels')) or n=='xl/styles.xml' or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    for month in range(1,13):
        for suffix in ['월','월계산기']:
            assert [ET.tostring(n) for n in p.sheet(f'{month}{suffix}').findall(Q('conditionalFormatting'))]==[ET.tostring(n) for n in q.sheet(f'{month}{suffix}').findall(Q('conditionalFormatting'))]
    structure=validate_package(final);assert not structure['errors'],structure
    a.close();w.close()
    assert not plan(final,source)['new']
    view=openpyxl.load_workbook(metadata.parent/'november_view/out/input.xlsx',data_only=True)
    november_reserve=verify_month(view,11);november_view=check_view(view,11)
    assert november_reserve['physical_wrap_total']==8287823 and november_view['all_ledger_rows']==4
    view.close()
    result=dict(as_of='2026-11-11',workbook_sha256=sha256(final.read_bytes()).hexdigest(),
                source_sha256=sha256(source.read_bytes()).hexdigest(),baseline_sha256=sha256(baseline.read_bytes()).hexdigest(),
                total_raw_records=count,source_november_records=82,new_transactions=54,
                calendar_actual=expected_variable,calendar_tetris=rounded,fixed_actual=expected_fixed,
                fixed_actual_and_planned=1279696,final_usage=1421769,card_rebates=800,
                other_income_as_submitted=12890,unverified_balance_adjustment_in_income=12802,
                other_income_excluding_adjustment=88,wrap_november_as_submitted=8287823,
                wrap_november_excluding_adjustment=8275021,reserve_selected_october=reserve,
                reserve_selected_november=november_reserve,november_selected_ledger=november_view,
                december_forecast_transport=137300,december_fixed_and_planned=1207010,
                historical_input_cells_compared=checked,memos=memos,
                repeated_import_new_records=0,all_30_calendar_days_checked=True,
                calculator_palette_checked=True,vba_buttons_styles_relationships_byte_preserved=True,
                conditional_formats_preserved=True,native_excel_ui_executed=False,
                unit_test_count=194,formula_errors=0,structure=structure,
                guide='2026-11-11_RAW추가_WRAP차액검토.md',
                previous_commit='30b65a5543e2e08782312241c52b26be5d156f9c',
                previous_version='history/2026-11-11/01',raw_attachment='raw/2026-11-11/'+source.name)
    report.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:result[k] for k in ['total_raw_records','new_transactions','calendar_actual','fixed_actual','final_usage','formula_errors']},ensure_ascii=False),flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['baseline','final','source','metadata','report']:p.add_argument(name,type=Path)
    a=p.parse_args();verify(a.baseline,a.final,a.source,a.metadata,a.report)
