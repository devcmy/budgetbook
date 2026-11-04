"""Independent November totals, date cells, categories, old data and ZIP checks."""
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
from update_postpaid_20261102 import read_source,compact
from raw_sort_id_20261101 import END,TABLE
from raw_auto_link_20261026 import palette_for
from verify_raw_sort_20261101 import equal,memo_checks
from verify_reserve_cash_alignment import verify_month
from verify_reserve_redesign import assert_no_errors
from validate_excel_structure import validate_package
from update_20261015 import Q


def fixed(item):
    title=compact(item['title'])
    return item['kind']=='지출' and (item['category']=='정기구독' or title.startswith('11번가') or '띵샵' in title)


def verify(baseline,final,source,metadata_path,report_path):
    metadata=json.loads(metadata_path.read_text());source_rows=read_source(source)
    a=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(w);raw=w['RAW'];old=a['RAW'];s=w['11월'];calc=w['11월계산기']
    ids={int(raw.cell(r,42).value):r for r in range(2,END+1)}
    before_ids={int(old.cell(r,42).value):r for r in range(2,END+1)}
    assert set(ids)==set(range(2,END+1)) and len(ids)==4000
    assert [(x.title,x.sheet_state) for x in a]==[(x.title,x.sheet_state) for x in w]
    assert len(w.sheetnames)==47 and '검증' not in w.sheetnames
    assert raw.tables[TABLE].ref=='A1:AQ4001' and not raw.protection.sheet
    count=sum(isinstance(raw.cell(r,1).value,datetime) for r in range(2,END+1))
    assert count==2617
    checked=0
    for ident,ar in before_ids.items():
        if not isinstance(old.cell(ar,1).value,datetime):continue
        br=ids[ident]
        for c in [1,*range(3,14),43]:
            if ident==2597 and c in [7,11]:
                assert old.cell(ar,c).value==5000 and raw.cell(br,c).value==4500
            else:equal(old.cell(ar,c).value,raw.cell(br,c).value,('historical input',ident,c));checked+=1
    source_map={item['source_row']:item for item in source_rows}
    for link in metadata['source_mapping']:
        item=source_map[link['source_row']];rr=ids[link['transaction_id']]
        for c,k in [(1,'date'),(3,'asset'),(7,'amount'),(8,'kind')]:
            equal(raw.cell(rr,c).value,item[k],('source identity',item['source_row'],c))
        if link['status']=='added':
            for c,k in [(4,'category'),(5,'sub'),(6,'title'),(9,'memo')]:
                equal(raw.cell(rr,c).value,item[k],('new original value',item['source_row'],c))
            assert raw.cell(rr,17).value==1
    assert len(metadata['new'])==20 and len(metadata['corrections'])==1
    assert metadata['corrections'][0]['transaction_id']==2597
    november=[item for item in source_rows if item['date'].month==11]
    variable=[item for item in november if item['kind']=='지출' and not fixed(item)]
    expected_fixed=sum(item['amount'] for item in november if fixed(item))
    expected_variable=sum(item['amount'] for item in variable)
    expected_units=sum(math.ceil(item['amount']/1000) for item in variable)
    assert len(november)==28 and len(variable)==17
    assert expected_fixed==s['E74'].value==125878
    assert expected_variable==s['I33'].value==438270
    assert expected_units*1000==s['I38'].value==calc['W12'].value==443000
    assert s['E78'].value==s['E79'].value==expected_fixed+expected_variable==564148
    assert s['E75'].value==1211341
    assert s['C40'].value==88 and s['H43'].value==300 and s['H44'].value==563848
    dates=[]
    for header in [3,8,13,18,23,28]:
        for c in range(2,9):
            dt=s.cell(header,c).value
            if not isinstance(dt,datetime) or dt.month!=11:continue
            daily=[item for item in variable if item['date'].date()==dt.date()]
            equal(s.cell(header+2,c).value or 0,sum(item['amount'] for item in daily),(dt,'calendar actual'))
            equal(s.cell(header+1,c).value or 0,sum(math.ceil(item['amount']/1000)*1000 for item in daily),(dt,'calendar rounded'))
            dates.append(dt.day)
    assert sorted(dates)==list(range(1,31))
    palette=Counter()
    for item in variable:
        cl='일회성전자' if item['category']=='가전/전자/잡화' else '변동비'
        palette[palette_for(cl,item['category'],item['sub'])]+=math.ceil(item['amount']/1000)
    for r in range(3,11):assert calc[f'C{r}'].value==palette[r],(r,calc[f'C{r}'].value,palette[r])
    assert sum(calc[f'C{r}'].value for r in range(3,11))*1000==443000
    for slot,amount in [(18,13875),(20,303),(21,18000),(22,74700),(23,10200),(24,8800)]:
        assert s[f'O{slot}'].value==amount and s[f'U{slot}'].value=='확정'
    assert s['O26'].value==0
    for ident,slot,amount in [(2616,23,10200),(2617,22,74700),(2618,24,8800)]:
        r=ids[ident];assert raw.cell(r,27).value==slot and raw.cell(r,28).value==amount
        assert raw.cell(r,25).value==0
    assert raw.cell(ids[2604],20).value=='일회성전자' and raw.cell(ids[2604],25).value==199000
    preserved_month_checks=0
    for month in [*range(1,11),12]:
        for addr in ['C37','C38','C39','C40','C41','I33','E74','E75','E78','E79','H43','H44','H45']:
            equal(a[f'{month}월'][addr].value,w[f'{month}월'][addr].value,(month,addr));preserved_month_checks+=1
        for addr in [f'C{r}' for r in range(3,11)]+['W12','W13']:
            equal(a[f'{month}월계산기'][addr].value,w[f'{month}월계산기'][addr].value,(month,'calculator',addr));preserved_month_checks+=1
    # No cash / reserve transaction was introduced by card purchases or Toss refunds.
    for name in ['잔액','예비비']:
        for cell in a[name]._cells.values():
            equal(cell.value,w[name][cell.coordinate].value,(name,cell.coordinate))
    reserve=verify_month(w,10)
    notes=memo_checks(w)
    dates_in_raw=[raw.cell(r,1).value for r in range(2,END+1) if isinstance(raw.cell(r,1).value,datetime)]
    assert dates_in_raw==sorted(dates_in_raw,reverse=True)
    preview_ids=[w['입력·연결'][f'H{r}'].value for r in range(8,18)]
    assert preview_ids==[raw.cell(r,42).value for r in range(2,12)]
    assert w['입력·연결']['C8'].value==raw.cell(ids[2618],6).value
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if n.endswith('.bin') or n.endswith('.rels') or n=='xl/styles.xml'
               or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    for month in range(1,13):
        for suffix in ['월','월계산기']:
            name=f'{month}{suffix}'
            assert [ET.tostring(x) for x in p.sheet(name).findall(Q('conditionalFormatting'))]==[ET.tostring(x) for x in q.sheet(name).findall(Q('conditionalFormatting'))]
    structure=validate_package(final);assert not structure['errors'],structure
    report=dict(date='2026-11-04',as_of='2026-11-04',
                source_sha256=sha256(source.read_bytes()).hexdigest(),
                baseline_sha256=sha256(baseline.read_bytes()).hexdigest(),workbook_sha256=sha256(final.read_bytes()).hexdigest(),
                source_records=len(source_rows),source_october_records=267,source_november_records=28,
                total_raw_records=count,new_transactions=20,amount_corrections=metadata['corrections'],
                future_invoices_confirmed_actual_by_user=metadata['future_invoices'],
                latest_booking_date=max(item['date'] for item in source_rows).isoformat(),
                latest_nonfuture_booking_date=max(item['date'] for item in source_rows if item['date'].date()<=datetime(2026,11,4).date()).isoformat(),
                calendar_actual=expected_variable,calendar_tetris=expected_units*1000,
                fixed_actual=expected_fixed,fixed_actual_and_planned=s['E75'].value,final_usage=s['E79'].value,
                card_rebates=300,other_income=88,usage_after_card_rebates=563848,
                wrap_total=w['잔액']['G16'].value,wrap_living_balance=w['잔액']['G7'].value,
                reserve=reserve,memos=notes,historical_input_cells_compared=checked,
                other_month_summary_cells_compared=preserved_month_checks,
                existing_ids_and_memos_preserved=True,all_30_calendar_days_checked=True,
                calculator_palette_independently_checked=True,new_raw_not_duplicated=True,
                new_wrap_or_reserve_links=0,search_preview_refreshed=True,
                vba_buttons_styles_relationships_byte_preserved=True,conditional_formats_preserved=True,
                native_excel_ui_executed=False,formula_errors=0,structure=structure,
                guide='2026-11-04_11월RAW추가_검증.md')
    report_path.write_text(json.dumps(report,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:report[k] for k in ['workbook_sha256','new_transactions','total_raw_records','calendar_actual','calendar_tetris','fixed_actual','final_usage','formula_errors']},ensure_ascii=False),flush=True)
    a.close();w.close()


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['baseline','final','source','metadata','report']:p.add_argument(name,type=Path)
    a=p.parse_args();verify(a.baseline,a.final,a.source,a.metadata,a.report)
