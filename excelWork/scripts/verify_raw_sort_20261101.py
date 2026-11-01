"""Independent permutation/recalculation tests for sortable RAW transaction IDs."""
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json
import math
import re

import openpyxl
from lxml import etree as ET
from openpyxl.utils.datetime import to_excel

from raw_sort_id_20261101 import END, TABLE, reorder
from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from verify_reserve_redesign import recalculate, assert_no_errors
from verify_reserve_cash_alignment import verify_month
from validate_excel_structure import validate_package


def normalized(value):
    if value in (None,''): return None
    if isinstance(value,datetime): return to_excel(value)
    return value


def equal(a,b,context):
    a,b = normalized(a),normalized(b)
    if isinstance(a,(int,float)) and isinstance(b,(int,float)):
        assert math.isclose(a,b,rel_tol=0,abs_tol=1e-7),(context,a,b)
    else:
        assert a == b,(context,a,b)


def compare_finances(before,after,stable_before=True):
    checked = 0
    for name in before.sheetnames:
        if name == 'RAW': continue
        a,b = before[name],after[name]
        for row in a:
            for c in row:
                value = c.value
                other = b[c.coordinate].value
                if isinstance(value,(int,float,datetime)) or isinstance(other,(int,float,datetime)):
                    equal(value,other,(name,c.coordinate));checked += 1
    a,b = before['RAW'],after['RAW']
    left = {int(a.cell(r,42).value) if stable_before else r:r for r in range(2,END+1)}
    right = {int(b.cell(r,42).value):r for r in range(2,END+1)}
    assert left.keys() == right.keys() == set(range(2,END+1))
    for ident,ar in left.items():
        br = right[ident]
        for col in range(1,42):
            value,other = a.cell(ar,col).value,b.cell(br,col).value
            # Text helpers depend on physical concatenation order only outside
            # RAW. RAW text can be checked exactly, including duplicate keeper.
            equal(value,other,('RAW ID',ident,col));checked += 1
    return checked


def memo_checks(w):
    raw = w['RAW'];records = []
    for r in range(2,END+1):
        when,memo = raw.cell(r,1).value,raw.cell(r,40).value
        if not isinstance(when,datetime) or when.year!=2026 or memo in (None,''): continue
        ident = int(raw.cell(r,42).value)
        records.append((ident,r,when.month,str(memo)))
    result = []
    for month in range(1,13):
        notes = sorted(v for v in records if v[2]==month)
        sheet,calc = w[f'{month}월'],w[f'{month}월계산기']
        assert sheet['AC40'].value == len(notes),(month,'memo count')
        page = max(1,min(int(sheet['W40'].value or 1),max(1,math.ceil(len(notes)/5))))
        for i in range(5):
            j = (page-1)*5+i
            text = notes[j][3] if j<len(notes) else None
            equal(sheet[f'W{35+i}'].value,text,(month,'memo page',i))
            equal(calc[f'B{158+i}'].value,text,(month,'calculator memo',i))
        for slot in [*range(4,25),26]:
            values = [str(raw.cell(r,40).value) for r in range(2,END+1)
                      if normalized(raw.cell(r,2).value)==to_excel(datetime(2026,month,1))
                      and raw.cell(r,27).value==slot and raw.cell(r,20).value=='고정비'
                      and raw.cell(r,17).value==1 and raw.cell(r,40).value not in (None,'')]
            equal(sheet[f'V{slot}'].value,'\n'.join(values),(month,'fixed memo',slot))
        result.append({'month':month,'memos':len(notes),'pages_and_fixed_memos_match':True})
    return result


def package_checks(source,candidate):
    before,after = ExtendedPackage(source),ExtendedPackage(candidate)
    assert before.parts.keys()==after.parts.keys()
    changed = [n for n in before.parts if before.parts[n]!=after.parts[n]]
    protected = [n for n in before.parts if n.endswith('.bin') or '/drawings/' in n
                 or '/ctrlProps/' in n or '/activeX/' in n or n=='xl/styles.xml']
    assert all(before.parts[n]==after.parts[n] for n in protected)
    b = openpyxl.load_workbook(source);w = openpyxl.load_workbook(candidate)
    raw = w['RAW']
    assert not raw.protection.sheet
    assert raw.tables[TABLE].ref=='A1:AP4001'
    assert raw.tables[TABLE].autoFilter.ref=='A1:AP4001'
    assert raw['AP1'].value=='거래ID'
    assert [raw.cell(r,42).value for r in range(2,END+1)]==list(range(2,END+1))
    assert raw.column_dimensions['AP'].hidden is False
    assert any(str(d.sqref)=='AP2:AP4001' for d in raw.data_validations.dataValidation)
    for r in range(2,END+1):
        assert raw.cell(r,1).number_format=='m"월" d"일"'
        for col in [1,3,4,5,6,7,8,9,13]:
            assert raw.cell(r,col).value==b['RAW'].cell(r,col).value,('input altered',r,col)
        for col in [10,11,12,18,19,27]:
            original=b['RAW'].cell(r,col)
            if original.data_type!='f':
                assert raw.cell(r,col).value==original.value,('override altered',r,col)
    unsafe = []
    for s in w:
        for row in s:
            for c in row:
                value = getattr(c.value,'text',c.value)
                if not isinstance(value,str) or not value.startswith('='): continue
                # The sole RAW physical navigation hyperlink is intentional.
                if s.title!='RAW' and re.search(r"(?:RAW|'RAW')!\$?[A-Z]+\$?\d",value):
                    unsafe.append((s.title,c.coordinate,value))
                if s.title=='RAW' and c.column<=41:
                    assert not re.search(r'\bROW\(\)',value),(c.coordinate,'physical identity')
    assert not unsafe,unsafe[:4]
    assert w.sheetnames==b.sheetnames and '검증' not in w.sheetnames
    structure=validate_package(candidate);assert not structure['errors'],structure
    w.close();b.close()
    return {'changed_parts':changed,'vba_controls_drawings_styles_byte_identical':True,
            'source_transactions_and_manual_overrides_preserved':True,
            'raw_native_sort_not_sheet_protected':True,'stable_ids_unique':4000,
            'unsafe_raw_physical_links':0,'structure':structure}


def main(source,candidate,work):
    report={'package':package_checks(source,candidate)}
    gc.collect()
    def calculate(path,folder):
        cached=folder/'out/input.xlsx'
        if cached.is_file():
            # A resumable diagnostic must never reuse an unrelated/stale cache.
            original=ExtendedPackage(path);actual=ExtendedPackage(folder/'input.xlsm')
            for package in [original,actual]:
                package.parts['xl/workbook.xml']=ET.tostring(package.tree('xl/workbook.xml'),encoding='UTF-8',xml_declaration=True)
            assert original.parts==actual.parts,('recalculation input does not match source',path,folder)
            return cached
        if folder.exists():
            # A failed disposable attempt remains available for diagnosis.
            folder=folder.with_name(folder.name+'_retry')
        return recalculate(path,folder)
    # Structured row references make the evaluator memory hungry. Running
    # these serially avoids two Office processes exhausting the cloud limit.
    paths={'baseline':calculate(source,work/'baseline')}
    paths['default']=calculate(candidate,work/'default')
    print('Default recalculations finished',flush=True)
    old=openpyxl.load_workbook(paths['baseline'],data_only=True)
    new=openpyxl.load_workbook(paths['default'],data_only=True)
    assert_no_errors(old);assert_no_errors(new)
    report['default']={'financial_values_checked':compare_finances(old,new,False),
                       'reserve':verify_month(new,10),'memos':memo_checks(new)}
    old.close();del old;gc.collect()
    cases={}
    for mode in ['ascending','descending']:
        p=ExtendedPackage(candidate)
        cases[mode]=reorder(p,new,mode)
        path=work/f'{mode}.xlsm';p.save(path);cases[mode]['path']=path
    del p;gc.collect()
    for mode,v in cases.items(): paths[mode]=calculate(v['path'],work/f'{mode}_calc')
    report['sort_cases']=[]
    for mode in cases:
        w=openpyxl.load_workbook(paths[mode],data_only=True)
        assert_no_errors(w)
        item={k:v for k,v in cases[mode].items() if k!='path'}
        dates=[to_excel(w['RAW'].cell(r,1).value) for r in range(2,END+1) if w['RAW'].cell(r,1).value]
        assert dates==sorted(dates,reverse=mode=='descending')
        item.update(financial_values_checked=compare_finances(new,w),reserve=verify_month(w,10),memos=memo_checks(w))
        report['sort_cases'].append(item);w.close()
        print(mode,'passed',flush=True)
    # Restore after descending, not just a fresh original-order fixture.
    desc=openpyxl.load_workbook(paths['descending'],data_only=True)
    p=ExtendedPackage(cases['descending']['path'])
    restored=reorder(p,desc,'id')
    p.save(work/'restored.xlsm')
    # In this baseline IDs were initially assigned in chronological order.
    # Also prove an actual date-ascending permutation from the descending file;
    # if its package is identical to the restored package, one recalculation
    # suffices for both operations (not an unchanged zero-move ascending test).
    ascending=ExtendedPackage(cases['descending']['path'])
    ascending_meta=reorder(ascending,desc,'ascending');desc.close()
    ascending.save(work/'ascending_from_descending.xlsm')
    same_package=(ExtendedPackage(work/'ascending_from_descending.xlsm').parts ==
                  ExtendedPackage(work/'restored.xlsm').parts)
    paths['restored']=calculate(work/'restored.xlsm',work/'restored_calc')
    w=openpyxl.load_workbook(paths['restored'],data_only=True);assert_no_errors(w)
    restored.update(financial_values_checked=compare_finances(new,w),reserve=verify_month(w,10))
    report['restored']=restored;w.close()
    assert same_package,'If a later baseline is not chronological by ID, recalculate ascending separately'
    ascending_meta.update(package_byte_identical_to_restored=True,
                          financial_values_checked=restored['financial_values_checked'])
    report['ascending_from_descending']=ascending_meta
    report['default_recalculated_path']=str(paths['default'])
    report['paths']={k:str(v) for k,v in paths.items()}
    report['native_microsoft_excel_ui_executed']=False
    report['scope']='ZIP/XML validation and full-record permutations with independent macro-free recalculation; no Excel UI automation'
    report['unchanged_limitations']=['Whole-RAW reset/reimport is not supported','New transactions belong in new blank rows; keep IDs and helper formulas','Data validation can be bypassed by paste; do not edit or delete transaction IDs']
    new.close()
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('source',type=Path);parser.add_argument('candidate',type=Path)
    parser.add_argument('work',type=Path);parser.add_argument('report',type=Path)
    args=parser.parse_args();args.work.mkdir(exist_ok=True)
    result=main(args.source,args.candidate,args.work)
    args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print('All sort invariants passed',flush=True)
