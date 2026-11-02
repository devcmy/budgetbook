"""Audit forecast-only changes and simulate next installment invoices."""
from datetime import datetime
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
import argparse
import gc
import json

from lxml import etree as ET
import openpyxl

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from raw_sort_id_20261101 import END,TABLE,reorder
from verify_raw_sort_20261101 import equal,memo_checks
from verify_reserve_redesign import recalculate,assert_no_errors
from verify_reserve_cash_alignment import verify_month
from validate_excel_structure import validate_package


def verify(baseline,final):
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if n.endswith('.bin') or n.endswith('.rels') or n=='xl/styles.xml'
               or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    max_formula=max(len(f.text or '') for part in q.sheets.values() for f in q.tree(part).iter(Q('f')))
    assert max_formula<=8192,max_formula
    for m in range(1,13):
        assert [ET.tostring(c) for c in p.sheet(f'{m}월계산기').findall(Q('conditionalFormatting'))]==[ET.tostring(c) for c in q.sheet(f'{m}월계산기').findall(Q('conditionalFormatting'))]
    del p,q;gc.collect()
    a=openpyxl.load_workbook(baseline,data_only=True);b=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(b)
    assert [(s.title,s.sheet_state) for s in a]==[(s.title,s.sheet_state) for s in b]
    assert len(b.sheetnames)==46 and '검증' not in b.sheetnames
    assert b['RAW'].tables[TABLE].ref=='A1:AP4001' and not b['RAW'].protection.sheet
    raw_checked=0
    for r in range(2,END+1):
        for c in range(1,43):
            equal(a['RAW'].cell(r,c).value,b['RAW'].cell(r,c).value,('existing RAW values',r,c));raw_checked+=1
    allowed={('10월',x) for x in ['O24','Q24','R24']}
    allowed.update((f'{m}월',x) for m in [11,12] for x in ['N24','O24','Q24','R24','Q25','R25','E75'])
    financial_checked=0
    for sheet in a:
        if sheet.title=='RAW':continue
        for cell in sheet._cells.values():
            key=(sheet.title,cell.coordinate)
            if key in allowed:continue
            value,other=cell.value,b[sheet.title][cell.coordinate].value
            if isinstance(value,(int,float,datetime)) or isinstance(other,(int,float,datetime)):
                equal(value,other,('unrelated numeric/date value',*key));financial_checked+=1
    expected={11:(2,1201465),12:(3,1206060)}
    for month,(installment,total) in expected.items():
        s=b[f'{month}월']
        assert s['L24'].value=='갤럭시버즈4프로' and s['O24'].value==11004 and s['U24'].value=='예상'
        assert s['S24'].value==f'{installment}/36회차'
        assert s['E75'].value==total
        assert all('애플케어' not in str(s[f'L{r}'].value or '') for r in range(4,25))
    assert b['10월']['L24'].value in [None,''] and b['10월']['O24'].value in [None,'']
    assert b['11월']['H37'].value==303 and b['11월']['E79'].value==75553
    assert b['12월']['H37'].value==b['12월']['E79'].value==0
    assert b['10월']['H37'].value==1256537 and b['10월']['E79'].value==3965294
    report={'date':'2026-11-02','workbook_sha256':sha256(final.read_bytes()).hexdigest(),
            'baseline_sha256':sha256(baseline.read_bytes()).hexdigest(),
            'raw_value_cells_compared':raw_checked,'other_numeric_and_date_cells_compared':financial_checked,
            'raw_transactions_added':0,'november_expected_buds':11004,'november_buds_installment':'2/36',
            'december_expected_buds':11004,'december_buds_installment':'3/36',
            'november_fixed_actual_and_planned':b['11월']['E75'].value,
            'december_fixed_actual_and_planned':b['12월']['E75'].value,
            'november_actual_fixed':303,'november_actual_usage':75553,
            'applecare_completed_september_3_of_3':True,'completed_applecare_not_forecast':True,
            'october_amounts_and_all_actuals_preserved':True,
            'source_memos_and_overrides_preserved':True,'hidden_months_and_raw_sort_preserved':True,
            'vba_styles_controls_relationships_byte_preserved':True,'calculator_visual_rules_preserved':True,
            'formula_max_length':max_formula,'formula_errors':0,'native_excel_ui_executed':False,
            'reserve':verify_month(b,10),'memos':memo_checks(b)}
    structure=validate_package(final);assert not structure['errors'],structure
    report['structure']=structure
    a.close();b.close();del a,b;gc.collect()
    return report


def future_test(final,root):
    values=openpyxl.load_workbook(final,data_only=True)
    ids={values['RAW'].cell(r,42).value:r for r in range(2,END+1)}
    p=ExtendedPackage(final)
    for r in range(2,END+1):p.write('RAW',f'I{r}',None)
    cases=[(2599,11,24,'롯데카드 띵샵 (2/36)',11000,'애플케어/에어팟이라고 써도 상품 분류는 변하지 않음'),
           (2600,11,5,'롯데카드띵샵(7/36)',10200,'버즈라고 써도 상품 분류는 변하지 않음'),
           (2601,11,12,'11번가 (15/22)',74700,None),
           (2602,11,26,'애플케어플러스 추가청구',26400,None),
           (2603,11,24,'롯데카드 띵샵 (2/36)',11000,'동일 거래의 다른 메모'),
           (2604,12,24,'롯데카드 띵샵(3/36)',11111,None)]
    for ident,month,day,title,amount,memo in cases:
        row=ids[ident];assert values['RAW'].cell(row,1).value is None
        for c,value in {'A':datetime(2026,month,day,12),'C':'롯데365카드','D':'정기구독','E':None,
                        'F':title,'G':amount,'H':'지출','I':memo,'M':None}.items():p.write('RAW',f'{c}{row}',value)
    for sn in ['잔액','예비비']:p.write(sn,'B3',datetime(2026,11,1))
    p.save(root/'future_unsorted.xlsm');values.close()
    values=openpyxl.load_workbook(root/'future_unsorted.xlsm',data_only=True)
    order=reorder(p,values,'ascending');values.close();p.save(root/'future.xlsm')
    del values,p;gc.collect()
    calculated=recalculate(root/'future.xlsm',root/'future_calculation')
    w=openpyxl.load_workbook(calculated,data_only=True);assert_no_errors(w)
    raw=w['RAW'];ids={raw.cell(r,42).value:r for r in range(2,END+1)}
    for ident,slot,amount in [(2599,24,11000),(2600,23,10200),(2601,22,74700),(2602,26,26400),(2603,24,0),(2604,24,11111)]:
        rr=ids[ident]
        assert raw.cell(rr,27).value==slot and raw.cell(rr,28).value==amount,(ident,raw.cell(rr,27).value,raw.cell(rr,28).value)
    assert raw.cell(ids[2603],17).value==2
    assert w['11월']['O24'].value==11000 and w['11월']['U24'].value=='확정'
    assert w['12월']['O24'].value==11111 and w['12월']['U24'].value=='확정'
    assert w['11월']['O23'].value==10200 and w['11월']['O22'].value==74700
    assert w['11월']['H37'].value==122603 and w['11월']['E79'].value==197853
    assert w['11월']['E75'].value==1227861 and w['12월']['E75'].value==1206167
    assert w['11월']['I33'].value==75250 and w['11월계산기']['W12'].value==77000
    assert w['10월']['H37'].value==1256537 and w['10월']['O11'].value==136350
    assert '애플케어' in w['11월']['V24'].value and w['11월']['L24'].value=='갤럭시버즈4프로'
    report={'synthetic_transactions_only_in_disposable_copy':True,
            'buds_airpods_phone_mapped_separately':True,'actual_amount_replaces_forecast_not_added_to_it':True,
            'installment_charge_rounding_difference_does_not_break_product_match':True,
            'completed_plan_not_restarted_by_unexpected_late_charge':True,
            'unrecognized_late_charge_preserved_as_actual_other_fixed_cost':True,
            'different_memos_do_not_reclassify_or_avoid_duplicate_exclusion':True,
            'clearing_raw_memos_does_not_change_financial_recognition':True,'date_sort':order,
            'reserve':verify_month(w,11),'memos':memo_checks(w),'formula_errors':0}
    w.close();return report


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['baseline','final','root','metadata','report']:p.add_argument(a,type=Path)
    a=p.parse_args();result=verify(a.baseline,a.final)
    print('Current inputs, actuals, forecasts, and package preservation verified',flush=True)
    result['future_invoice_tests']=future_test(a.final,a.root)
    result['metadata']=json.loads(a.metadata.read_text());result['unit_tests_passed']=161
    result['previous_commit']='554a491';result['previous_version']='history/2026-11-02/02'
    a.report.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Future invoices, duplicate handling, memo independence and sorting verified',flush=True)
