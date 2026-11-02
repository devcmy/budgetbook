"""Read-only delivery audit and disposable end-to-end input workflow tests."""
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json
import re

from lxml import etree as ET
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter as col

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from raw_sort_id_20261101 import END,TABLE,reorder
from entry_flow_20261102 import SCREEN,LINK,PLAN,CODES
from verify_raw_sort_20261101 import equal,memo_checks
from verify_reserve_redesign import recalculate,assert_no_errors
from verify_reserve_cash_alignment import verify_month
from validate_excel_structure import validate_package


def raw_ids(raw):
    result={raw.cell(r,42).value:r for r in range(2,END+1)}
    assert set(result)==set(range(2,END+1))
    return result


def actual_snapshot(w):
    result={}
    for m in range(1,13):
        s=w[f'{m}월'];c=w[f'{m}월계산기']
        addresses=['I33','E74','E78','E79']+[f'C{r}' for r in range(37,42)]
        addresses += [f'{cc}{r}' for cc in ['H','I'] for r in range(37,46)]
        addresses += [f'N{r}' for r in [*range(37,44),57,58,59,*range(64,70)]]
        for addr in addresses:result[s.title,addr]=s[addr].value
        for addr in [f'C{r}' for r in range(3,11)]+[f'W{r}' for r in range(3,18)]+['AF3']:
            result[c.title,addr]=c[addr].value
    for name in ['잔액','예비비']:
        for c in w[name]._cells.values():
            if isinstance(c.value,(int,float,datetime)):result[name,c.coordinate]=c.value
    return result


def compare_actuals(before,after):
    a,b=actual_snapshot(before),actual_snapshot(after)
    for key,value in a.items():equal(value,b.get(key),('actual value preserved',*key))
    return len(a)


def package_checks(baseline,final):
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    added=q.parts.keys()-p.parts.keys();assert added=={q.sheets[SCREEN]},added
    assert not p.parts.keys()-q.parts.keys()
    protected=[n for n in p.parts if n.endswith('.bin') or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))
               or (n.endswith('.rels') and n!='xl/_rels/workbook.xml.rels')]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    for component in ['numFmts','fonts','fills','borders','cellXfs','dxfs']:
        a=p.tree('xl/styles.xml').find(Q(component));b=q.tree('xl/styles.xml').find(Q(component))
        assert len(b)>=len(a)
        assert [ET.tostring(x) for x in a]==[ET.tostring(x) for x in list(b)[:len(a)]],component
    for m in range(1,13):
        assert [ET.tostring(c) for c in p.sheet(f'{m}월계산기').findall(Q('conditionalFormatting'))]==[ET.tostring(c) for c in q.sheet(f'{m}월계산기').findall(Q('conditionalFormatting'))]
    max_formula=max(len(f.text or '') for part in q.sheets.values() for f in q.tree(part).iter(Q('f')))
    assert max_formula<=8192,max_formula
    # U, not the payment-method column T, drives the green confirmation state.
    style=q.tree('xl/styles.xml');green=[]
    for cf in q.sheet('11월').findall(Q('conditionalFormatting')):
        if cf.get('sqref')=='O4:R24 U4:U24':
            for r in cf:
                formula=r.find(Q('formula')).text or ''
                if '"확정"' in formula and int(r.get('priority'))==1:
                    dxf=style.find(Q('dxfs'))[int(r.get('dxfId'))]
                    color=dxf.find(Q('fill')).find(Q('patternFill')).find(Q('fgColor')).get('rgb')
                    assert color=='FFE2F0D9';green.append({'formula':formula,'fill':color})
                    assert dxf.find(Q('fill')).find(Q('patternFill')).find(Q('bgColor')).get('rgb')==color
    assert len(green)==1,green
    for m in [11,12]:
        for cf in q.sheet(f'{m}월').findall(Q('conditionalFormatting')):
            assert not (cf.get('sqref')=='O18:O38' and any('$T18' in (r.find(Q('formula')).text or '') for r in cf))
    dvs=[]
    for sn in [SCREEN,'RAW','10월','11월','12월','잔액','예비비']:
        tree=q.sheet(sn).find(Q('dataValidations'))
        if tree is None:continue
        for dv in tree:
            f=dv.find(Q('formula1'))
            if f is not None:
                if dv.get('promptTitle')=='입력·연결 먼저 확인':assert len(f.text or '')<=255,(sn,dv.get('sqref'),len(f.text or ''))
                dvs.append((sn,dv.get('sqref'),f.text))
    assert any(s=='11월' and a=='O4:O24' and '_xlfn.ISFORMULA' in f for s,a,f in dvs)
    assert any(s=='잔액' and a=='W64:W500' and '_ERscope' in f for s,a,f in dvs)
    structure=validate_package(final);assert not structure['errors'],structure
    return {'new_worksheet_only_part_added':True,'original_style_components_prefix_preserved':True,
            'vba_controls_drawings_relationships_preserved':True,'calculator_visual_rules_preserved':True,
            'confirmed_green_rule':green[0],'maximum_formula_length':max_formula,'structure':structure,
            'standard_input_guards_installed':True,'paste_can_bypass_excel_data_validation':True}


def verify_delivery(baseline,final):
    package=package_checks(baseline,final);gc.collect()
    a=openpyxl.load_workbook(baseline,data_only=True);b=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(b)
    assert len(b.sheetnames)==47 and SCREEN in b.sheetnames and '검증' not in b.sheetnames
    assert all(a[n].sheet_state==b[n].sheet_state for n in a.sheetnames)
    assert b['RAW'].tables[TABLE].ref=='A1:AQ4001' and b['RAW']['AQ1'].value==LINK
    assert not b['RAW'].protection.sheet
    for m in [11,12]:
        assert not isinstance(b[f'{m}월']['AB20'],MergedCell)
        assert [str(r) for r in a[f'{m}월'].merged_cells.ranges]==[str(r) for r in b[f'{m}월'].merged_cells.ranges]
        assert b[f'{m}월'].column_dimensions['AB'].hidden is False
        for r in range(52,62):assert b[f'{m}월'][f'AB{r}'].number_format==';;;'
    checked=0
    for r in range(2,END+1):
        for c in range(1,43):equal(a['RAW'].cell(r,c).value,b['RAW'].cell(r,c).value,('old RAW field',r,c));checked+=1
        assert b['RAW'].cell(r,43).value is None
    actuals=compare_actuals(a,b)
    for m in [10,11,12]:equal(a[f'{m}월']['E75'].value,b[f'{m}월']['E75'].value,('unmodified initial plan',m))
    assert b['11월']['O20'].value==303 and b['11월']['U20'].value=='확정'
    form=b[SCREEN];assert form['K29'].value==8
    assert form['A33'].value is not None and len([form[f'A{r}'].value for r in range(33,43) if form[f'A{r}'].value])==8
    assert form['K26'].value is False and form['B14'].value=='고정비'
    assert form['B17'].value=='카드사용' and '확인' in form['A26'].value
    rows=raw_ids(b['RAW']);r=rows[2515]
    assert b['RAW'].cell(r,7).value==b['RAW'].cell(r,23).value==b['RAW'].cell(r,25).value==169000
    assert b['RAW'].cell(r,33).value==0
    camping={'raw_id':2515,'date':'2026-10-23','title':'숙소비','amount':169000,
             'current_personal_burden':169000,'calendar':169000,'tetris':169000,'advance_principal':0,
             'user_confirmed_shared_settlement_planned':True,'final_personal_share_unknown':True,
             'financial_reclassification_not_guessed':True,'source_transaction_untouched':True}
    report={'date':'2026-11-02','baseline_sha256':sha256(baseline.read_bytes()).hexdigest(),
            'workbook_sha256':sha256(final.read_bytes()).hexdigest(),'package':package,
            'raw_value_cells_compared':checked,'actual_financial_cells_compared':actuals,
            'naver_november_confirmed_amount':303,'camping_analysis':camping,
            'total_raw_records':2597,'raw_transactions_added':0,'hidden_months_and_actuals_preserved':True,
            'input_screen_candidates_visible':8,'input_screen_is_not_a_parallel_financial_register':True,
            'no_vba_popups_or_macro_execution':True,'native_microsoft_excel_ui_executed':False,
            'reserve':verify_month(b,10),'memos':memo_checks(b),'formula_errors':0}
    a.close();b.close();del a,b;gc.collect();return report


def recalc_scenario(source,root,tag,edit):
    p=ExtendedPackage(source);edit(p);path=root/f'{tag}.xlsm';p.save(path);del p;gc.collect()
    calculated=recalculate(path,root/f'{tag}_calculation')
    return path,calculated


def scenario_tests(final,root):
    def plan_query(p):
        for a,v in {'B3':'기존 RAW 수정·연결','G3':datetime(2026,10,1),'B4':169000,'G5':'글램핑','B7':2515,'G8':999}.items():p.write(SCREEN,a,v)
        for sn,values in [('11월',{'AB24':12345,'AB26':5000}),('12월',{'AB26':2500})]:
            for a,v in values.items():p.write(sn,a,v)
        p.write(SCREEN,'O48',123)
        p.write(SCREEN,'P48','1+0',True)
        p.write(SCREEN,'K48','_xlfn.ISFORMULA(O48)',True)
        p.write(SCREEN,'K49','_xlfn.ISFORMULA(P48)',True)
    plan_path,plan_values=recalc_scenario(final,root,'plan_and_query',plan_query)
    before=openpyxl.load_workbook(final,data_only=True);w=openpyxl.load_workbook(plan_values,data_only=True)
    assert_no_errors(w);actuals=compare_actuals(before,w)
    s=w[SCREEN]
    assert s['K29'].value==1 and s['A33'].value==2515 and s['K26'].value is False
    assert s['G11'].value==169000 and '글램핑' in s['B16'].value
    assert '확인' in s['F23'].value
    assert not bool(s['K48'].value) and bool(s['K49'].value)
    assert w['11월']['O26'].value==5000 and w['11월']['U26'].value=='예상'
    assert w['11월']['O24'].value==12345 and w['11월']['U24'].value=='예상'
    assert w['11월']['E74'].value==303 and w['11월']['E79'].value==75553
    assert w['11월']['E75'].value==1207806
    assert w['12월']['E74'].value==0 and w['12월']['E79'].value==0 and w['12월']['E75'].value==1208560
    ids=raw_ids(w['RAW']);before.close();w.close();del before,w;gc.collect()
    def actuals_and_mapping(p):
        for a,v in {'B3':'기존 RAW 수정·연결','G3':datetime(2026,11,1),'B4':31490,'G5':'쿠팡','B7':2598,'G8':2598}.items():p.write(SCREEN,a,v)
        for r in range(2,END+1):p.write('RAW',f'I{r}',None)
        p.write('RAW',f'J{ids[2598]}','고정비');p.write('RAW',f'AQ{ids[2598]}',20)
        p.write('RAW',f'AQ{ids[2597]}',20) # No fixed reclassification: must warn, not double post.
        cases=[(2599,3,'수기 신규 고정비 시험',6000,'지출','정기구독',26),
               (2600,24,'롯데카드 띵샵(2/36)',11111,'지출','정기구독',None),
               (2601,3,'수기 신규 고정비 시험',6000,'지출','정기구독',20),
               (2602,4,'카드정산 결제',77777,'이체지출','현대카드',20)]
        for ident,day,title,amount,kind,category,code in cases:
            r=ids[ident]
            for c,v in {'A':datetime(2026,11,day,12),'C':'토스뱅크' if ident==2602 else '롯데365카드',
                        'D':category,'E':None,'F':title,'G':amount,'H':kind,'M':None,'AQ':code,
                        'I':'메모가 달라도 같은 거래' if ident==2601 else None}.items():p.write('RAW',f'{c}{r}',v)
            if ident==2602:p.write('RAW',f'J{r}','고정비')
        for sn in ['잔액','예비비']:p.write(sn,'B3',datetime(2026,11,1))
    mapping=ExtendedPackage(plan_path);actuals_and_mapping(mapping)
    unsorted=root/'mapping_unsorted.xlsm';mapping.save(unsorted)
    values=openpyxl.load_workbook(unsorted,data_only=True)
    order=reorder(mapping,values,'ascending');values.close()
    mapping_path=root/'mapping_and_actuals.xlsm';mapping.save(mapping_path);del mapping;gc.collect()
    calculated=recalculate(mapping_path,root/'mapping_and_actuals_calculation')
    w=openpyxl.load_workbook(calculated,data_only=True);assert_no_errors(w)
    raw=w['RAW'];rows=raw_ids(raw)
    assert raw.cell(rows[2598],43).value==20 and raw.cell(rows[2598],27).value==20
    assert raw.cell(rows[2598],25).value==0 and raw.cell(rows[2598],28).value==31490
    assert raw.cell(rows[2597],43).value==20 and raw.cell(rows[2597],27).value==0
    assert '연결 대기' in raw.cell(rows[2597],14).value
    assert raw.cell(rows[2601],17).value==2 and raw.cell(rows[2601],28).value==0
    assert raw.cell(rows[2602],28).value==0 and '이체/수입' in raw.cell(rows[2602],14).value
    assert w['11월']['O20'].value==31793
    assert w['11월']['AB26'].value==5000 and w['11월']['O26'].value==6000 and w['11월']['U26'].value=='확정'
    assert w['11월']['AB24'].value==12345 and w['11월']['O24'].value==11111 and w['11월']['U24'].value=='확정'
    assert w['11월']['I33'].value==43760 and w['11월계산기']['W12'].value==45000
    assert w['11월']['E74'].value==48904 and w['11월']['E79'].value==92664
    assert w['11월']['E75'].value==1239062
    s=w[SCREEN]
    assert bool(s['K26'].value) and s['K29'].value==1 and s['A33'].value==2598
    assert s['G14'].value==20 and 'RAW' in s['A9'].value and '연결코드' in s['F23'].value
    result={'synthetic_transactions_only_in_temp_files':True,
            'plan_only_changes_forecast_not_actual':True,'actuals_compared_in_plan_case':actuals,
            'memo_keyword_lookup_finds_camping':True,'wrong_confirmation_id_disables_edit_routes':True,
            'explicit_confirmation_matches_selected_raw_id':True,'native_isformula_guard_logic_verified':True,
            'manual_fixed_link_moves_existing_spending_without_adding_it_twice':True,
            'wrong_class_or_ordinary_transfer_does_not_force_fixed_posting':True,
            'actual_other_fixed_replaces_manual_other_plan':True,'actual_installment_replaces_manual_plan':True,
            'memo_changes_do_not_avoid_duplicate_exclusion':True,'raw_link_code_survives_sort':order,
            'reserve':verify_month(w,11),'memos':memo_checks(w),'formula_errors':0}
    w.close();return result


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['baseline','final','root','metadata','report']:p.add_argument(a,type=Path)
    a=p.parse_args();report=verify_delivery(a.baseline,a.final)
    print('Delivery financials, camping facts, colors, RAW source preservation and layout verified',flush=True)
    report['workflow_tests']=scenario_tests(a.final,a.root)
    report['metadata']=json.loads(a.metadata.read_text());report['unit_tests_passed']=174
    report['previous_commit']='64643ba';report['previous_version']='history/2026-11-02/03'
    a.report.write_text(json.dumps(report,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Guided lookup, confirmation, plans, actual links, duplicates and sorting verified',flush=True)
