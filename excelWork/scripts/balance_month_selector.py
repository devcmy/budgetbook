"""Add and regression-test the balance query month without rewriting VBA.

Recalculation runs only on disposable macro-free copies. The canonical XLSM
keeps its original formulas, styles, controls and binary parts.
"""
import argparse
from datetime import datetime
from hashlib import sha256
import json
import math
import os
from pathlib import Path
import subprocess
from zipfile import ZipFile

from lxml import etree as ET
import openpyxl
from openpyxl.utils import get_column_letter

from update_20261015 import Package, Q, macro_free, import_caches


def prepare(source, output):
    p = Package(source)
    formulas = openpyxl.load_workbook(source, data_only=False)
    sheet = p.sheet('잔액')
    validations = sheet.find(Q('dataValidations'))
    assert not any('B3' in x.get('sqref', '').split() for x in validations)
    validation = ET.SubElement(validations, Q('dataValidation'), {
        'type': 'list', 'sqref': 'B3', 'allowBlank': '0',
        'showDropDown': '0', 'showInputMessage': '1',
        'showErrorMessage': '1', 'errorStyle': 'stop',
        'promptTitle': '조회월 선택',
        'prompt': '목록에서 3~12월을 선택하세요. 위쪽 계정별 잔액과 합계가 선택한 달 기준으로 바뀝니다.',
        'errorTitle': '목록에서 월을 선택하세요',
        'error': '조회월은 목록에 있는 2026년 3~12월 중에서 선택해주세요.',
    })
    ET.SubElement(validation, Q('formula1')).text = '$C$21:$L$21'
    validations.set('count', str(len(validations)))
    p.write('잔액', 'D3',
            '=TEXT($B$3,"yyyy-mm")&IF($B$3>MAX($A$64:$A$500),'
            '" · 해당 월 미입력 / 이월 잔액",IF(EOMONTH($B$3,0)<MAX($A$64:$A$500),'
            '" · 월말 기록 기준"," · 입력된 거래 기준 (월 진행 중)"))', True)
    p.write('잔액', 'A18',
            '조회월은 위쪽 잔액 기준, 내역 월은 아래 거래내역 기준입니다. 미입력 월은 이월 잔액이며 WRAP 밖 보관금은 제외합니다.')
    # Historical comparison and audit tables must not depend on a UI selection.
    # Keep actual recorded months visible; do not present future carry as actuals.
    for row in list(range(22, 31)) + [32]:
        for col in range(3, 13):
            addr = get_column_letter(col) + str(row)
            f = formulas['잔액'][addr].value
            assert '$B$3' in f
            p.write('잔액', addr, f.replace('$B$3', 'MAX($A$64:$A$500)'), True)
    p.write('검증', 'I168', '=잔액!I32', True)
    p.write('검증', 'I175', '=잔액!J32-I174', True)
    for row in range(72, 77):
        addr = 'L' + str(row)
        f = formulas['예비비'][addr].value
        assert '잔액!$B$3' in f
        p.write('예비비', addr, f.replace('잔액!$B$3', 'MAX(잔액!$A$64:$A$500)'), True)
    p.write('사용안내', 'A14',
            '잔액 탭은 WRAP 통장만 관리합니다. 조회월(B3)을 목록에서 고르면 해당 월의 전월말·입출금·계정별 잔액·합계가 바뀝니다.')
    calc = p.tree('xl/workbook.xml').find(Q('calcPr'))
    calc.set('calcMode', 'auto')
    calc.set('fullCalcOnLoad', '1')
    calc.set('forceFullCalc', '1')
    p.save(output)


def expected_summary(sheet, month):
    start = datetime(2026, month, 1)
    end = datetime(2027, 1, 1) if month == 12 else datetime(2026, month + 1, 1)
    expected = {}
    for row in range(7, 16):
        account = sheet[f'A{row}'].value
        before = income = expense = outgoing = incoming = 0
        for r in range(64, 501):
            date = sheet[f'A{r}'].value
            if not isinstance(date, datetime):
                continue
            amount = sheet[f'Y{r}'].value or 0
            sender, recipient = sheet[f'C{r}'].value, sheet[f'D{r}'].value
            kind = sheet[f'B{r}'].value
            if date < start:
                before += amount * ((recipient == account) - (sender == account))
            elif date < end:
                if recipient == account:
                    if kind in {'입금', '적립', '기초', '환입'}:
                        income += amount
                    else:
                        incoming += amount
                if sender == account:
                    if kind in {'사용', '출금'}:
                        expense += amount
                    else:
                        outgoing += amount
        for col, value in zip('BCDEFG', [before, income, expense, outgoing, incoming,
                                       before + income - expense - outgoing + incoming]):
            expected[f'{col}{row}'] = value
    for col in 'BCDEFG':
        expected[f'{col}16'] = sum(expected[f'{col}{r}'] for r in range(7, 16))
    expected['G17'] = sum(expected[f'G{r}'] for r in range(8, 13))
    return expected


def test_and_cache(baseline, candidate, output, workdir, report):
    root = Path(workdir)
    root.mkdir(parents=True, exist_ok=True)
    original = openpyxl.load_workbook(baseline, data_only=True)
    bools = [(s.title, c.coordinate) for s in original for row in s for c in row
             if isinstance(c.value, bool)]
    def test_month(month):
        folder = root / f'month_{month:02d}'
        folder.mkdir()
        (folder / 'out').mkdir()
        (folder / 'cache').mkdir()
        p = Package(candidate)
        p.write('잔액', 'B3', datetime(2026, month, 1))
        p.save(folder / 'input.xlsm')
        macro_free(folder / 'input.xlsm', folder / 'input.xlsx')
        result = subprocess.run([
            'soffice', '-env:UserInstallation=' + (folder / 'profile').as_uri(),
            '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML',
            '--outdir', str(folder / 'out'), str(folder / 'input.xlsx'),
        ], env={**os.environ, 'XDG_CACHE_HOME': str(folder / 'cache')},
            capture_output=True, text=True, timeout=120)
        assert result.returncode == 0, (month, result.returncode, result.stdout, result.stderr)
        recalculated = folder / 'out/input.xlsx'
        assert recalculated.is_file(), result.stdout + result.stderr
        w = openpyxl.load_workbook(recalculated, data_only=True)
        expected = expected_summary(original['잔액'], month)
        for addr, value in expected.items():
            assert w['잔액'][addr].value == value, (month, addr, w['잔액'][addr].value, value)
        for sheet, addr in bools:
            assert w[sheet][addr].value == original[sheet][addr].value, (month, sheet, addr)
        errors = [(s.title, c.coordinate, c.value) for s in w for row in s for c in row if c.data_type == 'e']
        assert not errors, (month, errors[:10])
        for row in range(72, 77):
            assert w['예비비'][f'L{row}'].value in [True, 1], (month, row)
        assert w['잔액']['B35'].value == original['잔액']['B35'].value
        # All other sheets keep their original cached values after recalculation.
        for name in original.sheetnames:
            if name == '잔액':
                continue
            for row in original[name]:
                for c in row:
                    if (name, c.coordinate) == ('사용안내', 'A14'):
                        continue  # Updated instructions, not financial data.
                    actual = w[name][c.coordinate].value
                    if c.value in [None, ''] and actual in [None, '']:
                        continue
                    if isinstance(c.value, (int, float)) and isinstance(actual, (int, float)):
                        assert math.isclose(c.value, actual, abs_tol=1e-6), (month, name, c.coordinate, c.value, actual)
                    else:
                        assert c.value == actual, (month, name, c.coordinate, c.value, actual)
        return {'month': month, 'total': expected['G16'], 'purpose_funds': expected['G17'],
                'header': w['잔액']['D3'].value, 'checked_summary_cells': len(expected),
                'formula_errors': len(errors), 'unchanged_boolean_checks': len(bools)}
    # Calc and openpyxl both use considerable memory for this workbook.
    results = [test_month(month) for month in range(3, 13)]
    import_caches(candidate, root / 'month_10/out/input.xlsx', output, baseline)
    f = openpyxl.load_workbook(output, data_only=False)
    b = openpyxl.load_workbook(baseline, data_only=False)
    assert f.sheetnames == b.sheetnames
    assert f['잔액']['B3'].value == datetime(2026, 10, 1)
    dv = [x for x in f['잔액'].data_validations.dataValidation if 'B3' in x.sqref]
    assert len(dv) == 1 and dv[0].formula1 == '$C$21:$L$21'
    assert dv[0].showDropDown is False and dv[0].showErrorMessage is True
    for sheet in b:
        assert sheet.sheet_state == f[sheet.title].sheet_state
        assert list(map(str, sheet.merged_cells.ranges)) == list(map(str, f[sheet.title].merged_cells.ranges))
        for row in sheet:
            for c in row:
                assert c._style == f[sheet.title][c.coordinate]._style, (sheet.title, c.coordinate, 'style')
    with ZipFile(baseline) as old, ZipFile(output) as new:
        assert old.namelist() == new.namelist()
        preserved = [n for n in old.namelist() if not n.startswith('xl/worksheets/sheet') and n != 'xl/workbook.xml']
        for name in preserved:
            assert old.read(name) == new.read(name), name
        assert new.testzip() is None
    audit = {'date': '2026-10-15', 'cell': '잔액!B3', 'source': '잔액!C21:L21',
             'default_month': '2026-10', 'month_tests': results,
             'other_sheet_financial_values_unchanged': True, 'styles_and_merges_unchanged': True,
             'vba_controls_and_binary_parts_unchanged': True,
             'preserved_nonworksheet_parts': len(preserved),
             'workbook_sha256': sha256(Path(output).read_bytes()).hexdigest(),
             'recalculation': 'LibreOffice macro-free temporary copies; original XLSM retains its package and formulas',
             'excel_ui_and_vba_executed': False}
    Path(report).write_text(json.dumps(audit, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(audit, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    cli = argparse.ArgumentParser()
    sub = cli.add_subparsers(dest='mode', required=True)
    p = sub.add_parser('prepare')
    p.add_argument('source'); p.add_argument('output')
    t = sub.add_parser('test')
    for name in ['baseline', 'candidate', 'output', 'workdir', 'report']:
        t.add_argument('--' + name, required=True)
    args = cli.parse_args()
    if args.mode == 'prepare':
        prepare(args.source, args.output)
    else:
        test_and_cache(args.baseline, args.candidate, args.output, args.workdir, args.report)
