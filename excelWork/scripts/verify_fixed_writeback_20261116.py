"""Inspect VBA event installation and exercise financial contracts on copies.

Microsoft Excel/VBA is unavailable here. These checks do not claim native
event execution: they check extracted source, protections, formula restoration,
and independent source-edit/recalculation fixtures for the promised behavior.
"""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json
import re

import openpyxl
from lxml import etree as ET

from fixed_writeback_20261116 import source_modules,candidate_ids,SLOTS,FIRST,META,MODULE
from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from verify_reserve_redesign import recalculate,assert_no_errors
from verify_raw_sort_20261101 import equal,memo_checks
from verify_reserve_cash_alignment import verify_month
from verify_chronological_blocks_20261116 import chronology
from validate_excel_structure import validate_package
from preserve_memo_caches_20261102 import restore


def vba_checks(baseline,final):
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    before=source_modules(p.parts['xl/vbaProject.bin']);after=source_modules(q.parts['xl/vbaProject.bin'])
    code=after['BBFixed'].decode('cp949');workbook=p.tree('xl/workbook.xml').find(Q('workbookPr')).get('codeName')
    assert set(after)==set(before)|{'BBFixed'}
    assert all(after[k]==v for k,v in before.items() if k!=workbook)
    assert after[workbook].startswith(before[workbook])
    events=after[workbook].decode('cp949')
    for event in ['Workbook_Open','Workbook_BeforeSave','Workbook_AfterSave','Workbook_SheetChange']:
        assert events.count('Sub '+event+'(')==1
    assert 'BBFixed.BB_FixedChange Sh, Target' in events
    assert code.replace('\r\n','\n')==MODULE.read_text().replace('\r\n','\n')
    for prohibited in ['Shell(', 'CreateObject(', 'VBProject', 'SaveAs(', 'Kill ', 'Declare ', 'Environ(']:
        assert prohibited not in code
    assert 'Application.Undo' not in code
    assert 'raw.Cells(rawRow, 43).Value2 = slot' in code
    assert 'portion = Number(raw.Cells(rawRow, 23).Value2) + desired - oldTotal' in code
    assert 'raw.Cells(rawRow, 11).Value2 = portion' in code
    assert 'raw.Cells(rawRow, 9).ClearContents' in code
    assert 'raw.Range("A" & rawRow & ":AQ" & rawRow).Formula2 = oldRaw' in code
    assert 'Application.EnableEvents = savedEvents' in code
    assert 'Application.Calculation = savedCalc' in code
    assert code.index('If portion < 0 Then')<code.index('changed = True')
    assert 'Private Const FIRST_META As Long = 32002' in code
    assert 'Cells(r, IIf(column = 15, 88, 89))' in code
    # Strings must close on every physical source line; VBA continuation limits.
    continued=0
    for line in code.splitlines():
        if line.lstrip().startswith("'"):continue
        assert len(line)<1024
        assert line.count('"')%2==0, line
        continued=continued+1 if line.rstrip().endswith(' _') else 0
        assert continued<=24
    return dict(existing_modules_preserved_except_workbook_event_additions=True,
                new_module='BBFixed',modules=len(after),event_bindings_verified=True,
                original_search_source_identical=before['BBSearch']==after['BBSearch'],
                rollback_and_application_state_guards_present=True,native_macro_executed=False)


def verify(baseline,final,fresh_calculation):
    a=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(final,data_only=True)
    fresh=openpyxl.load_workbook(fresh_calculation,data_only=True);f=openpyxl.load_workbook(final,data_only=False)
    assert_no_errors(w);assert_no_errors(fresh)
    assert [(s.title,s.sheet_state) for s in a]==[(s.title,s.sheet_state) for s in w]
    for r in range(2,4002):
        for c in range(1,44):equal(a['RAW'].cell(r,c).value,w['RAW'].cell(r,c).value,('RAW',r,c))
    for m in range(1,13):
        for slot in SLOTS:
            for c in ['O','P','Q','R','U','V']:
                equal(a[f'{m}월'][f'{c}{slot}'].value,w[f'{m}월'][f'{c}{slot}'].value,('fixed',m,c,slot))
                if c!='V':equal(a[f'{m}월'][f'{c}{slot}'].value,fresh[f'{m}월'][f'{c}{slot}'].value,('fresh fixed',m,c,slot))
            assert f[f'{m}월'].protection.sheet
            assert f[f'{m}월'][f'O{slot}'].protection.locked
            assert f[f'{m}월'][f'V{slot}'].protection.locked
            rr=FIRST+(m-1)*23+slot-4
            assert w[META].cell(rr,88).value==f[f'{m}월'][f'O{slot}'].value
            assert w[META].cell(rr,89).value==f[f'{m}월'][f'V{slot}'].value
            assert not w[META].cell(rr,88).data_type=='f'
            assert not w[META].cell(rr,89).data_type=='f'
        for address in ['I33','I38','C40','H37','E74','E75','E79','H43']:
            equal(a[f'{m}월'][address].value,w[f'{m}월'][address].value,(m,address))
            equal(a[f'{m}월'][address].value,fresh[f'{m}월'][address].value,('fresh',m,address))
        for address in [*[f'C{r}' for r in range(3,11)],'W12','W13']:
            equal(a[f'{m}월계산기'][address].value,w[f'{m}월계산기'][address].value,('calculator',m,address))
    assert not f['11월']['AB4'].protection.locked and not f['12월']['AB26'].protection.locked
    for s in ['잔액','예비비']:
        for address in (['B3','C9','C10','C11','G7','G9','G10','G11','G16','G17'] if s=='잔액' else ['B3','I13']):
            equal(a[s][address].value,w[s][address].value,(s,address))
    notes=memo_checks(w);memo_checks(fresh)
    chronology(w,11,stored=True);chronology(w,12,stored=True)
    reserve=verify_month(w,10)
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if (n.endswith(('.bin','.rels')) and n!='xl/vbaProject.bin') or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    for m in range(1,13):
        name=f'{m}월계산기';assert p.parts[p.sheets[name]]==q.parts[q.sheets[name]]
    old_styles=p.tree('xl/styles.xml');new_styles=q.tree('xl/styles.xml')
    for old in old_styles:
        new=new_styles.find(old.tag)
        if old.tag==Q('cellXfs'):
            assert len(new)>=len(old) and all(ET.tostring(x)==ET.tostring(new[i]) for i,x in enumerate(old))
        else:assert ET.tostring(old)==ET.tostring(new)
    structure=validate_package(final);assert not structure['errors'],structure
    result=dict(as_of='2026-11-16',workbook_sha256=sha256(final.read_bytes()).hexdigest(),
                initial_raw_and_financial_values_preserved=True,all_12_months_checked=True,
                freshly_recalculated_memo_routing_matches=True,delivery_memo_linebreaks_preserved=True,
                monthly_memos=notes,reserve=reserve,
                chronological_calculators_preserved=True,vba=vba_checks(baseline,final),
                saved_cells_protected_until_macros_enabled=True,forecast_inputs_editable=True,
                controls_relationships_preserved=True,original_style_ids_preserved=True,
                formula_errors=0,structure=structure,native_excel_ui_executed=False,
                validation_scope='OOXML/VBA-source inspection and independent source-patch/recalculation fixtures; no native VBA event execution.',
                previous_commit='1473050',previous_version='history/2026-11-16/03')
    for book in [a,w,fresh,f]:book.close()
    gc.collect();print('Initial finances, memos, chronology, protections, VBA and structure passed.',flush=True)
    return result


def scenarios(final,root,report):
    w=openpyxl.load_workbook(final,data_only=True);raw=w['RAW']
    ids={int(raw.cell(r,42).value):r for r in range(2,4002)}
    empty=next(r for r in range(2,4002) if not raw.cell(r,1).value)
    new_id=int(raw.cell(empty,42).value)
    # Literal patches are the contract expected from the VBA, not VBA execution.
    # One copy exercises singleton, aggregate, partial burden, link/create/plan,
    # memo roundtrip, and a real CMA withdrawal without altering delivery data.
    ledger_row=max(r for r in range(64,501) if w['잔액'][f'A{r}'].value)+1
    def edit(p):
        # Singleton 11월 에어팟: 10,250 -> 10,800 (full burden).
        for c in ['G','K']:p.write('RAW',f'{c}{ids[2616]}',10800)
        # Aggregate: one invoice +1,000; the other invoices remain independent.
        for c in ['G','K']:p.write('RAW',f'{c}{ids[2658]}',55983)
        # Partial burden: change the personal portion, retain original gross.
        p.write('RAW',f'K{ids[2720]}',100000)
        p.write('RAW',f'I{ids[2658]}','월별 메모에서 수정\n여러 줄 유지')
        p.write('RAW',f'I{ids[2617]}',None)
        # Reconnect an existing variable expense to an empty actual fixed slot.
        p.write('RAW',f'J{ids[2738]}','고정비');p.write('RAW',f'AQ{ids[2738]}',9)
        # New December rent in an existing free ID: create one source, one bank
        # movement. Forecast is replaced, not added a second time.
        for c,value in dict(A=datetime(2026,12,25),C='네이버미래에셋',D='고정비',F='월세',G=391000,H='지출',J='고정비',K=391000,L='본인',AQ=4).items():p.write('RAW',f'{c}{empty}',value)
        for c,value in dict(A=datetime(2026,12,25),B='출금',C='생활비',D='지출',E=391000,F='월세',J='네이버미래에셋',K='외부',L='WRAP입출금',W=new_id).items():p.write('잔액',f'{c}{ledger_row}',value)
        p.write('12월','AB5',180000)
        # Existing CMA rent corrected through a monthly actual: cash follows.
        for c in ['G','K']:p.write('RAW',f'{c}{ids[2539]}',395000)
        p.write('잔액','E256',395000)
        for name in ['잔액','예비비']:p.write(name,'B3',datetime(2026,12,1))
    w.close();w=None;raw=None;gc.collect()
    if (root/'edits/out/input.xlsx').is_file():
        expected=ExtendedPackage(final);edit(expected);expected.save(root/'edits/expected_input.xlsm')
        assert ExtendedPackage(root/'edits/expected_input.xlsm').parts==ExtendedPackage(root/'edits/input.xlsm').parts
        calc=root/'edits/out/input.xlsx'
    else:calc=recalculate(final,root/'edits',edit)
    restored=root/'edits/restored.xlsx';restore(final,calc,restored);calc=restored
    s=openpyxl.load_workbook(calc,data_only=True);assert_no_errors(s)
    assert s['11월']['O23'].value==10800
    assert s['11월']['O26'].value==183303 # 188637+1000-6334
    assert s['RAW'].cell(ids[2720],7).value==106334
    assert s['RAW'].cell(ids[2720],11).value==100000
    assert s['11월']['O9'].value==10887
    assert s['11월']['I33'].value==1252014
    assert s['11월']['E74'].value==626515+550+1000-6334+10887
    assert s['11월계산기']['W12'].value==1273000
    assert s['12월']['O4'].value==s['12월']['E74'].value==391000
    assert s['12월']['U4'].value=='확정' and s['12월']['O5'].value==180000
    assert s['12월']['E75'].value==1216660
    assert s['잔액']['G16'].value==4653217 and s['잔액']['G7'].value==2303217
    assert s['예비비']['I13'].value==2350000
    assert s['10월']['O4'].value==395000
    assert s['잔액'][f'I{ledger_row}'].value=='OK'
    memo_checks(s);verify_month(s,12);chronology(s,11)
    assert '월별 메모에서 수정\n여러 줄 유지' in s['11월']['V26'].value
    assert s['11월']['V22'].value in [None,'']
    report['source_edit_contract_scenario']=dict(singleton_amount=True,aggregate_selected_delta=True,
        partial_burden_preserves_gross=True,memo_change_and_clear=True,existing_raw_relinked=True,
        new_actual_replaces_forecast=True,forecast_only_does_not_create_raw=True,
        cma_existing_and_new_withdrawals_match=True,vba_event_itself_executed=False)
    s.close();s=None;gc.collect()
    # A second copy removes memo, cancels an actual cash charge, and ensures
    # the remaining aggregate memo/other invoices and chronology are intact.
    def remove(p):
        p.write('RAW',f'I{ids[2658]}',None)
        for c in ['G','K']:p.write('RAW',f'{c}{ids[2539]}',0)
        for c in ['A','B','C','D','E']:p.write('잔액',f'{c}256',None)
    calc2=recalculate(calc,root/'clears',remove)
    restored=root/'clears/restored.xlsx';restore(final,calc2,restored);calc2=restored
    s=openpyxl.load_workbook(calc2,data_only=True);assert_no_errors(s)
    assert s['10월']['O4'].value==0
    assert s['잔액']['G16'].value==5048217 and s['잔액']['G7'].value==2698217
    assert '월별 메모에서 수정' not in (s['11월']['V26'].value or '')
    assert '브랜든 리포트' in s['11월']['V26'].value and '모나' in s['11월']['V26'].value
    memo_checks(s);verify_month(s,12);s.close()
    report['clear_and_cash_cancellation_contract_scenario_passed']=True
    print('Independent edit, aggregate, memo, partial burden, relink, create, plan and cash scenarios passed.',flush=True)


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for name in ['baseline','final','fresh','root','report']:a.add_argument(name,type=Path)
    a.add_argument('--initial-only',action='store_true')
    a.add_argument('--scenarios-only',action='store_true')
    args=a.parse_args();args.root.mkdir(exist_ok=True)
    result=json.loads(args.report.read_text()) if args.scenarios_only else verify(args.baseline,args.final,args.fresh)
    if not args.initial_only:scenarios(args.final,args.root,result)
    args.report.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
