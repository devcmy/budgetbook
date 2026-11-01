"""Change RAW date display only and test memo linkage in all twelve months.

The delivery package receives only new number-format styles and RAW A-column
style references. Synthetic transactions and memo edits exist in temp copies.
"""
from collections import defaultdict
from copy import deepcopy
from datetime import datetime
from hashlib import sha256
import argparse
import json
import math
from pathlib import Path
import re
import subprocess
import sys

import openpyxl
from lxml import etree as ET

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from validate_excel_structure import validate_package
from verify_reserve_redesign import recalculate
from verify_raw_auto_link_20261026 import errors,append_row
from verify_memo_only_20261026 import clear_memos,check_cleared,same_finances

DATE_FORMAT='m"월" d"일"'
END=4001


def date_style(xf,numfmt):
    result=deepcopy(xf)
    result.set('numFmtId',str(numfmt));result.set('applyNumberFormat','1')
    return result


def prepare(source,output):
    p=ExtendedPackage(source);styles=p.tree('xl/styles.xml')
    formats=styles.find(Q('numFmts'));assert formats is not None
    numfmt=next((int(n.get('numFmtId')) for n in formats if n.get('formatCode')==DATE_FORMAT),None)
    if numfmt is None:
        numfmt=max([163]+[int(n.get('numFmtId')) for n in formats])+1
        ET.SubElement(formats,Q('numFmt'),numFmtId=str(numfmt),formatCode=DATE_FORMAT)
        formats.set('count',str(len(formats)))
    xfs=styles.find(Q('cellXfs'));mapping={}
    def clone(style_id):
        if style_id not in mapping:
            mapping[style_id]=len(xfs);xfs.append(date_style(xfs[style_id],numfmt))
        return mapping[style_id]
    for r in range(2,END+1):
        c=p.cell('RAW',f'A{r}');c.set('s',str(clone(int(c.get('s','0')))))
    # Column default remains locked outside the 4,000 prepared input rows.
    column=next(c for c in p.sheet('RAW').find(Q('cols')) if c.get('min')=='1' and c.get('max')=='1')
    column.set('style',str(clone(int(column.get('style','0')))))
    column.set('width','13');column.set('customWidth','1')
    xfs.set('count',str(len(xfs)))
    p.trees.pop('xl/workbook.xml',None)
    p.save(output)
    return {'format':DATE_FORMAT,'styled_input_cells':4000,'style_mapping':mapping,'raw_column_width':13}


def check_package(source,final):
    a,b=ExtendedPackage(source),ExtendedPackage(final)
    assert a.parts.keys()==b.parts.keys()
    changed=[n for n in a.parts if a.parts[n]!=b.parts[n]]
    assert set(changed)=={'xl/styles.xml',b.sheets['RAW']},changed
    old={c.get('r'):c for c in a.sheet('RAW').iter(Q('c'))}
    new={c.get('r'):c for c in b.sheet('RAW').iter(Q('c'))}
    assert old.keys()==new.keys()
    for address,c in old.items():
        other=deepcopy(new[address])
        if re.fullmatch(r'A([2-9]|[1-9][0-9]+)',address):
            if 's' in c.attrib:other.set('s',c.get('s'))
            else:other.attrib.pop('s',None)
        assert ET.tostring(c)==ET.tostring(other),('RAW content changed',address)
    w=openpyxl.load_workbook(final);before=openpyxl.load_workbook(source,data_only=True)
    for r in range(2,END+1):
        assert w['RAW'][f'A{r}'].number_format==DATE_FORMAT
        assert w['RAW'][f'A{r}'].value==before['RAW'][f'A{r}'].value,(r,'date/time altered')
        assert not w['RAW'][f'A{r}'].protection.locked
    assert w['RAW']['A1'].value=='일시'
    assert w['RAW']['B2'].number_format==before['RAW']['B2'].number_format=='yyyy-mm'
    assert w['RAW'].protection.sheet and not w['RAW'].protection.autoFilter
    assert w['RAW'].tables['RawData2026'].ref=='A1:AO4001'
    assert len(w.sheetnames)==46 and '검증' not in w.sheetnames
    w.close();before.close()
    structure=validate_package(final);assert not structure['errors'],structure
    return {'changed_package_parts':changed,'raw_values_formulas_caches_preserved':True,
            'actual_year_and_time_preserved':True,'all_other_package_parts_byte_identical':True,
            'month_filter_and_input_protection_preserved':True,'structure':structure}


def text(value):return '' if value in [None,''] else str(value)


def check_monthly_memos(path):
    """Independently reconstruct current memo pages and fixed/extra views."""
    w=openpyxl.load_workbook(path,data_only=True);f=openpyxl.load_workbook(path)
    assert not errors(w),errors(w)[:10]
    raw=w['RAW'];by_month=defaultdict(list);fixed=defaultdict(list);normalized_rows=[]
    for r in range(2,END+1):
        dt=raw[f'A{r}'].value;memo=raw[f'I{r}'].value
        if not isinstance(dt,datetime) or dt.year!=2026:continue
        if not text(memo):continue
        helper=text(raw[f'AN{r}'].value)
        original=text(memo)
        if helper!=original:
            # The disposable evaluator removes input-cell CR/LF before formula
            # evaluation, even for =A1. Compare the source bytes separately;
            # never rewrite user text to match the evaluator's limitation.
            assert helper==original.replace('\r','').replace('\n',''),(r,'memo text differs beyond line breaks')
            normalized_rows.append(r)
        by_month[dt.month].append((r,helper))
        if raw[f'T{r}'].value=='고정비' and raw[f'Q{r}'].value==1:
            fixed[dt.month,raw[f'AA{r}'].value].append(helper)
    report=[];extras=0
    for m in range(1,13):
        s=w[f'{m}월'];calc=w[f'{m}월계산기'];notes=by_month[m]
        try:page=max(1,min(int(s['W40'].value),max(1,math.ceil(len(notes)/5))))
        except (TypeError,ValueError):page=1
        expected=notes[(page-1)*5:page*5]
        assert s['AC40'].value==len(notes),(m,'memo count')
        for i in range(5):
            value=expected[i][1] if i<len(expected) else ''
            assert text(s[f'W{35+i}'].value)==value,(m,i,'monthly memo page')
            assert text(calc[f'B{158+i}'].value)==value,(m,i,'calculator memo page')
        for slot in [*range(4,25),26]:
            assert text(s[f'V{slot}'].value)=='\n'.join(fixed[m,slot]),(m,slot,'fixed memo')
        for r in range(52,62):
            formula=f[f'{m}월'][f'I{r}'].value
            if not isinstance(formula,str) or not formula.startswith('='):continue
            rows=sorted(set(int(x) for x in re.findall(r'RAW!AN(\d+)',formula)))
            if rows:
                value='\n'.join(text(raw[f'AN{rr}'].value) for rr in rows if raw[f'A{rr}'].value and text(raw[f'AN{rr}'].value))
                assert text(s[f'I{r}'].value)==value,(m,r,'extra memo')
                extras+=1
        report.append({'month':m,'raw_memos':len(notes),'page':page,'memo_browser_and_calculator_match':True,'fixed_memos_match':True})
    w.close();f.close()
    return {'months':report,'linked_extra_memo_cells_checked':extras,'formula_errors':0,
            'existing_linebreak_normalized_cache_rows':normalized_rows,
            'linebreak_scope':'Temporary evaluator normalizes input CR/LF; original RAW memo bytes and all existing caches are preserved, not rewritten.'}


def fixture(p,with_memos):
    clear_memos(p)
    for m in range(1,13):
        variable=2590+(m-1)*2;fixed=variable+1
        append_row(p,variable,m,26,1200+m,title=f'월별 메모 검증 {m}',memo=f'{m}월 변동 메모 수정 시험' if with_memos else '')
        append_row(p,fixed,m,27,2000+m,title='Apple Appstore',category='정기구독',asset='MG+S 하나카드',memo=f'{m}월 고정 메모 수정 시험' if with_memos else '')
        p.write(f'{m}월','W40',1)


def check_fixture(baseline,result,with_memos):
    check_monthly_memos(result)
    before=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(result,data_only=True)
    assert not errors(w),errors(w)[:10]
    for m in range(1,13):
        s=w[f'{m}월'];calc=w[f'{m}월계산기']
        expected_variable=f'{m}월 변동 메모 수정 시험' if with_memos else ''
        expected_fixed=f'{m}월 고정 메모 수정 시험' if with_memos else ''
        for name,addr,expected in [(f'{m}월','W35',expected_variable),(f'{m}월','W36',expected_fixed),
                                   (f'{m}월','V19',expected_fixed),(f'{m}월계산기','B158',expected_variable),
                                   (f'{m}월계산기','B159',expected_fixed)]:
            assert text(w[name][addr].value)==expected,(name,addr)
        assert s['I33'].value==before[f'{m}월']['I33'].value+1200+m,(m,'calendar')
        assert calc['W12'].value==before[f'{m}월계산기']['W12'].value+2000,(m,'calculator')
    before.close();w.close()
    return {'months_tested':list(range(1,13)),'monthly_browser_fixed_and_calculator_follow_raw':True,
            'memos_present':with_memos,'calendar_and_calculator_keep_transaction_amounts':True}


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for name in ['source','output','workdir','report']:parser.add_argument('--'+name,required=True)
    args=parser.parse_args();folder=Path(args.workdir);folder.mkdir(parents=True,exist_ok=True)
    report={'date':'2026-11-01','data_as_of':'2026-10-31','date_display':prepare(args.source,args.output)}
    report['package']=check_package(args.source,args.output)
    report['current_memo_views']=check_monthly_memos(args.output)
    # Metadata lives at excelWork, independent of snapshot nesting.
    root=Path(__file__).resolve().parent.parent
    metadata=json.loads((root/'memo_only_20261026_migration.json').read_text())
    cleared=recalculate(args.output,folder/'clear_existing_memos',clear_memos)
    report['clear_existing_memos']=check_cleared(args.output,cleared,metadata)
    print(json.dumps({'passed':'clear_existing_memos',**report['clear_existing_memos']},ensure_ascii=False),flush=True)
    edited=recalculate(args.output,folder/'all_months_edited',lambda p:fixture(p,True))
    report['all_months_edited']=check_fixture(args.output,edited,True)
    print(json.dumps({'passed':'all_months_edited',**report['all_months_edited']},ensure_ascii=False),flush=True)
    deleted=recalculate(args.output,folder/'all_months_deleted',lambda p:fixture(p,False))
    report['all_months_deleted']=check_fixture(args.output,deleted,False)
    report['delete_memo_financial_preservation']=same_finances(edited,deleted)
    print(json.dumps({'passed':'all_months_deleted',**report['delete_memo_financial_preservation']},ensure_ascii=False),flush=True)
    tests=subprocess.run([sys.executable,'-m','unittest','discover','-s',str(Path(__file__).parent),'-p','test_*.py'],capture_output=True,text=True,check=True)
    report['unit_tests_passed']=int(re.search(r'Ran (\d+) tests',tests.stderr)[1])
    report['source_sha256']=sha256(Path(args.source).read_bytes()).hexdigest()
    report['workbook_sha256']=sha256(Path(args.output).read_bytes()).hexdigest()
    report['native_excel_ui_executed']=False
    report['synthetic_transactions_only_in_temporary_copies']=True
    report['previous_version']='history/2026-11-01/02'
    report['guide']='2026-11-01_RAW월일표시_메모연동_검증.md'
    Path(args.report).write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'report':args.report,'workbook_sha256':report['workbook_sha256']},ensure_ascii=False))
