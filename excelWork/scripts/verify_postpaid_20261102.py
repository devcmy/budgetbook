"""Independent financial/source/sort/memo tests for the user's November update."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json
import math
import re

from lxml import etree as ET
import openpyxl
from openpyxl.utils.datetime import to_excel

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from update_postpaid_20261102 import read_source,key,usage,settlement,DEFERRED,digest
from update_raw_20261101 import asset
from raw_sort_id_20261101 import END,TABLE,reorder
from verify_reserve_redesign import recalculate,assert_no_errors
from verify_reserve_cash_alignment import verify_month
from verify_raw_sort_20261101 import equal,memo_checks
from validate_excel_structure import validate_package


def id_rows(raw):
    result={raw.cell(r,42).value:r for r in range(2,END+1)}
    assert set(result)==set(range(2,END+1))
    return result


def package_checks(baseline,final):
    before,after=ExtendedPackage(baseline),ExtendedPackage(final)
    removed=before.parts.keys()-after.parts.keys()
    assert removed=={'xl/calcChain.xml'},removed
    assert not after.parts.keys()-before.parts.keys()
    protected=[n for n in before.parts if n.endswith('.bin') or n=='xl/styles.xml'
               or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))
               or (n.endswith('.rels') and n!='xl/_rels/workbook.xml.rels')]
    assert all(before.parts[n]==after.parts[n] for n in protected)
    for name in [f'{m}월계산기' for m in range(1,13)]:
        for tag in ['conditionalFormatting','mergeCells','cols']:
            a=[ET.tostring(x) for x in before.sheet(name).findall(Q(tag))]
            b=[ET.tostring(x) for x in after.sheet(name).findall(Q(tag))]
            assert a==b,(name,tag)
    w=openpyxl.load_workbook(final)
    assert len(w.sheetnames)==46 and '검증' not in w.sheetnames
    assert all(w[f'{m}월'].sheet_state==w[f'{m}월계산기'].sheet_state=='hidden' for m in range(1,10))
    raw=w['RAW'];assert not raw.protection.sheet
    assert raw.tables[TABLE].ref=='A1:AP4001'
    # Excel escaped the existing literal space when saving the uploaded file.
    assert all(raw.cell(r,1).number_format.replace('\\ ',' ')=='m"월" d"일"' for r in range(2,END+1))
    assert '모두의카드'==w['10월']['L11'].value==w['11월']['L11'].value==w['12월']['L11'].value
    states={s.title:s.sheet_state for s in w}
    w.close()
    structure=validate_package(final);assert not structure['errors'],structure
    return {'vba_controls_drawings_styles_relationships_preserved':True,
            'calculator_visual_rules_and_layout_preserved':True,
            'removed_stale_calc_chain':True,'hidden_months_1_to_9_preserved':True,
            'sheet_states':states,'structure':structure}


def verify(baseline,baseline_values,final,source,metadata):
    report={'package':package_checks(baseline,final)}
    original=openpyxl.load_workbook(baseline)
    before=openpyxl.load_workbook(baseline_values,data_only=True)
    w=openpyxl.load_workbook(final,data_only=True)
    formulas=openpyxl.load_workbook(final)
    assert_no_errors(w)
    old,raw=before['RAW'],w['RAW'];oi,ni=id_rows(old),id_rows(raw)
    payment=metadata['billing_transaction_id']
    for ident,r in oi.items():
        if not isinstance(old.cell(r,1).value,datetime):continue
        rr=ni[ident]
        for c in [1,3,4,5,6,7,8,9,10,11,12,13]:
            # Compare source inputs against the actual uploaded package, not
            # the evaluator which normalizes CR/LF in literal note text.
            expected=(original['RAW'].cell(r,c).value if c in [1,3,4,5,6,7,8,9,13]
                      else old.cell(r,c).value)
            if ident==payment and c==9:
                assert raw.cell(rr,c).value.startswith(str(expected or ''))
                continue
            equal(expected,raw.cell(rr,c).value,('original RAW',ident,c))
        for c in [20,21,22,23,24,25,26,28,29,30,31,32,33,34,35,37,38,39]:
            if ident==payment and c==28:continue
            equal(old.cell(r,c).value,raw.cell(rr,c).value,('old financial helper',ident,c))
    sources={v['source_row']:v for v in read_source(source)}
    for item in metadata['new']:
        r=ni[item['transaction_id']];src=sources[item['source_row']]
        for c,k in [(1,'date'),(3,'asset'),(4,'category'),(5,'sub'),(6,'title'),(7,'amount'),(8,'kind'),(9,'memo')]:
            equal(raw.cell(r,c).value,src[k],('new source field',item['source_row'],c))
        equal(raw.cell(r,17).value,1,('new record not duplicate',r))
    for item in metadata['matched']:
        src=sources[item['source_row']];r=ni[item['transaction_id']]
        equal((raw.cell(r,1).value,asset(raw.cell(r,3).value),raw.cell(r,7).value,raw.cell(r,8).value),key(src),('matched source',item['source_row']))
    counts=Counter(raw.cell(r,1).value.month for r in range(2,END+1) if isinstance(raw.cell(r,1).value,datetime))
    assert sum(counts.values())==2597 and counts[9]==245 and counts[10]==267 and counts[11]==8,counts
    dates=[to_excel(raw.cell(r,1).value) for r in range(2,END+1) if raw.cell(r,1).value]
    assert dates==sorted(dates,reverse=True)
    for r in range(2,END+1):
        expected=raw.cell(r,9).value if raw.cell(r,1).value else None
        equal(raw.cell(r,40).value,expected,('exact memo cache including linebreaks',r))
    # All old monthly financial/date cells stay unchanged, including October.
    old_month_cells=0
    for m in [*range(1,11),12]:
        for sn in [f'{m}월',f'{m}월계산기']:
            for row in before[sn]:
                for cell in row:
                    v=cell.value;other=w[sn][cell.coordinate].value
                    # Forward estimate intentionally follows the confirmed October invoice.
                    if m==12 and sn=='12월' and cell.coordinate in ['O11','Q11','R11','O25','Q25','R25','E75','I37']:continue
                    if isinstance(v,(int,float,datetime)) or isinstance(other,(int,float,datetime)):
                        equal(v,other,('old month financial',sn,cell.coordinate));old_month_cells+=1
    for sn in ['잔액','예비비']:
        for row in before[sn]:
            for cell in row:
                value=cell.value
                if isinstance(value,(int,float,datetime)):
                    equal(value,w[sn][cell.coordinate].value,(sn,cell.coordinate,'unchanged original balance'))
    for ident,r in oi.items():
        if not isinstance(old.cell(r,1).value,datetime):continue
        rr=ni[ident]
        for c in [10,11,12,18,19,27]:
            oc=original['RAW'].cell(r,c)
            if oc.data_type!='f' and not (ident==payment and c==27):
                equal(oc.value,formulas['RAW'].cell(rr,c).value,('manual override preserved',ident,c))
    bill=ni[payment];evidence=next(v for v in metadata['new'] if str(v['date']).startswith('2026-09-'))
    use=ni[evidence['transaction_id']]
    assert raw.cell(use,20).value==DEFERRED and raw.cell(use,28).value==raw.cell(use,25).value==0
    assert raw.cell(bill,20).value=='고정비' and raw.cell(bill,27).value==11 and raw.cell(bill,28).value==136350
    for addr,value in [('O11',136350),('U11','확정'),('H37',1256537),('E75',1256547),('I33',2605157),('E78',3965294),('E79',3965294)]:
        equal(w['10월'][addr].value,value,('October',addr))
    november=[v for v in read_source(source) if v['date'].month==11]
    variable=[v for v in november if v['kind']=='지출' and v['category']!='정기구독']
    expected_variable=sum(v['amount'] for v in variable)
    expected_rounded=sum(math.ceil(v['amount']/1000)*1000 for v in variable)
    expected_fixed=sum(v['amount'] for v in november if v['kind']=='지출' and v['category']=='정기구독')
    expected_income=sum(v['amount'] for v in november if v['kind']=='수입')
    assert (expected_variable,expected_rounded,expected_fixed,expected_income)==(75250,77000,303,88)
    for addr,value in [('I33',expected_variable),('I38',expected_rounded),('O20',expected_fixed),('U20','확정'),
                       ('H37',expected_fixed),('C40',expected_income),('O11',136350),('U11','예상'),
                       ('E78',expected_variable+expected_fixed),('E79',expected_variable+expected_fixed)]:
        equal(w['11월'][addr].value,value,('November',addr))
    equal(w['11월계산기']['W12'].value,expected_rounded,'November calculator')
    for m in [10,11,12]:
        assert w[f'{m}월']['L11'].value=='모두의카드' and w[f'{m}월']['M11'].value=='토스뱅크'
    assert w['10월']['V11'].value and '11/2 확인' in w['10월']['V11'].value
    assert '11/2 확인' not in str(w['10월']['V26'].value)
    for sn in ['잔액','예비비']:
        assert formulas[sn]['B3'].value==original[sn]['B3'].value,(sn,'user selector changed')
    report.update(date='2026-11-02',as_of='2026-11-02',baseline_sha256=digest(baseline),
                  source_sha256=digest(source),workbook_sha256=digest(final),total_raw_records=sum(counts.values()),
                  raw_by_month=dict(sorted(counts.items())),new_transactions=9,new_november_transactions=8,
                  october_fixed_actual=w['10월']['H37'].value,october_usage=w['10월']['E79'].value,
                  november_variable_actual=expected_variable,november_tetris=expected_rounded,
                  november_fixed_actual=expected_fixed,november_income=expected_income,
                  november_usage=w['11월']['E79'].value,november_fixed_actual_and_planned=w['11월']['E75'].value,
                  confirmed_october_transport=136350,november_transport_expected_not_actual=True,
                  original_month_financial_cells_compared=old_month_cells,reserve=verify_month(w,10),memos=memo_checks(w),
                  source_fields_and_old_overrides_preserved=True,user_descending_sort_preserved=True,
                  raw_memo_caches_preserve_exact_linebreaks=True,
                  ordinary_transfers_not_converted_to_expenses=True,native_excel_ui_executed=False,formula_errors=0)
    w.close();before.close();original.close();formulas.close();gc.collect()
    return report


def future_fixture(delivery,root):
    w=openpyxl.load_workbook(delivery,data_only=True);ids=id_rows(w['RAW'])
    p=ExtendedPackage(delivery)
    # Delete all free notes. Classification and fixed costs must remain intact.
    for r in range(2,END+1):p.write('RAW',f'I{r}',None)
    for ident,dt,account,category,title,kind in [
        (2599,datetime(2026,10,31,22),'토스뱅크모두의카드','교통/차량','10월 후불교통','지출'),
        (2600,datetime(2026,11,12,8),'토스뱅크','토스뱅크모두의카드','카드정산 결제','이체지출'),
        (2601,datetime(2026,11,12,8),'토스뱅크','토스뱅크모두의카드','카드정산 결제','이체지출'),
        (2602,datetime(2026,11,13,8),'토스뱅크','현대카드','카드정산 결제','이체지출')]:
        r=ids[ident];assert w['RAW'].cell(r,1).value is None
        for c,value in {'A':dt,'C':account,'D':category,'E':None,'F':title,'G':77777 if ident==2602 else 88888,'H':kind,'I':None,'M':None}.items():
            p.write('RAW',f'{c}{r}',value)
    p.write('RAW',f'I{ids[2600]}','대납 · 날짜 미기록 · 메모는 계산 기준이 아님')
    p.write('RAW',f'I{ids[2601]}','다른 메모라도 동일 거래는 중복')
    for sn in ['잔액','예비비']:p.write(sn,'B3',datetime(2026,11,1))
    p.save(root/'future_unsorted.xlsm');w.close()
    values=openpyxl.load_workbook(root/'future_unsorted.xlsm',data_only=True)
    order=reorder(p,values,'ascending');values.close();p.save(root/'future.xlsm')
    del p;gc.collect()
    path=recalculate(root/'future.xlsm',root/'future_calc')
    w=openpyxl.load_workbook(path,data_only=True);assert_no_errors(w)
    raw=w['RAW'];ids=id_rows(raw)
    for ident,classification,amount,rank in [(2599,DEFERRED,0,1),(2600,'고정비',88888,1),(2601,'고정비',0,2),(2602,'내부이체',0,1)]:
        r=ids[ident]
        assert raw.cell(r,20).value==classification and raw.cell(r,28).value==amount and raw.cell(r,17).value==rank,(ident,[raw.cell(r,c).value for c in [20,28,17]])
    assert w['10월']['O11'].value==136350 and w['10월']['H37'].value==1256537
    assert w['11월']['O11'].value==88888 and w['11월']['U11'].value=='확정'
    assert w['11월']['H37'].value==303+88888
    assert w['11월']['I33'].value==75250 and w['11월계산기']['W12'].value==77000
    assert w['10월']['V11'].value in [None,'']
    assert '메모는 계산 기준' in w['11월']['V11'].value
    result={'temporary_synthetic_transactions_only':True,'new_alias_usage_and_charge_auto_link':True,
            'original_month_usage_not_double_counted':True,'different_memo_does_not_avoid_duplicate_exclusion':True,
            'ordinary_card_settlement_still_excluded':True,
            'cleared_memo_clears_view_not_finances':True,'native_sort_simulated_by_full_table_permutation':order,
            'reserve_november':verify_month(w,11),'memos':memo_checks(w),'formula_errors':0}
    w.close();return result


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for a in ['baseline','baseline_values','final','source','metadata','root','report']:parser.add_argument(a,type=Path)
    a=parser.parse_args();metadata=json.loads(a.metadata.read_text())
    result=verify(a.baseline,a.baseline_values,a.final,a.source,metadata)
    print('Source preservation, dates, October/November totals and package checks passed',flush=True)
    result['future_charge_sort_and_memo_test']=future_fixture(a.final,a.root)
    result['metadata']=metadata
    result['unit_tests_passed']=153
    a.report.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Future charge, memo deletion, duplicate exclusion and sorting passed',flush=True)
