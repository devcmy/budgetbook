"""Record a purpose-account transfer, not a new bank transaction or income."""
import argparse
from datetime import datetime
import gc
from hashlib import sha256
import json
from pathlib import Path
from zipfile import ZipFile
from zoneinfo import ZoneInfo

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from verify_reserve_redesign import (
    read, recalculate, assert_no_errors, verify_month, cache_original,
)


def prepare(source, output):
    w = read(source)
    assert w['잔액']['G15'].value == 61000
    assert w['잔액']['G7'].value == 4474466
    assert w['잔액']['A247'].value is None
    date = datetime.now(ZoneInfo('Asia/Seoul')).replace(tzinfo=None, microsecond=0)
    assert date.date() == datetime(2026,10,15).date()
    p = ExtendedPackage(source)
    values = {
        'A247': date, 'B247': '이체', 'C247': '임시보관', 'D247': '생활비',
        'E247': 61000, 'F247': '생신 잔금 임시보관 → 생활비',
        'G247': '사용자 요청으로 회수 원금 61,000원을 생활비로 분류 이동. 날짜는 분류 기록 시각이며 실제 은행 송금 아님. 원금 회수는 원장 246행에 별도 보존. 신규 수입·지출·적립 아님.',
        'J247': '미래에셋_WRAP형CMA', 'K247': '미래에셋_WRAP형CMA',
        'L247': 'WRAP내계정이동',
    }
    for addr, value in values.items():
        p.write('잔액', addr, value)
    p.write('잔액','G246',w['잔액']['G246'].value+' 회수 후 원장 247행에서 임시보관 → 생활비로 분류 이동.')
    p.write('검증','I44','=COUNTIF(잔액!$I$231:$I$247,"입력 확인")+COUNTIF(잔액!$I$231:$I$247,"WRAP 범위 확인")',True)
    calc=p.tree('xl/workbook.xml').find(Q('calcPr'))
    calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    p.save(output)
    return date.isoformat()


def verify(before_path, final_path):
    before, after = read(before_path), read(final_path)
    assert before.sheetnames == after.sheetnames
    assert [(s.title,s.sheet_state) for s in before] == [(s.title,s.sheet_state) for s in after]
    expected={'G7':4535466,'G15':0,'G16':6685466,'E16':2598839,'F16':748510,'Y247':61000}
    for addr, value in expected.items():
        assert after['잔액'][addr].value == value,(addr,after['잔액'][addr].value,value)
    assert after['잔액']['I247'].value == 'OK'
    assert after['잔액']['Q247'].value == '이체'
    assert after['잔액']['W247'].value is None and after['잔액']['X247'].value is None
    # Every financial value outside the balance sheet remains unchanged. The
    # hidden transfer-view date helper must pick up the new ledger entry.
    for s in before:
        if s.title == '잔액':continue
        for row in s:
            for c in row:
                if (s.title,c.coordinate) == ('_예비비계산','BB303'):
                    continue
                actual=after[s.title][c.coordinate].value
                if c.value in [None,''] and actual in [None,'']:continue
                assert c.value == actual,(s.title,c.coordinate,c.value,actual)
    # Existing balances for March through September cannot change.
    for row in range(22,33):
        for col in range(3,10):
            assert before['잔액'].cell(row,col).value == after['잔액'].cell(row,col).value
    assert after['예비비']['L13'].value == 239000
    assert after['예비비']['F13'].value == 0
    assert after['예비비']['I13'].value == 2150000
    assert_no_errors(after)
    for row in after['검증']:
        for c in row:
            if isinstance(c.value,bool):assert c.value is True,c.coordinate
    with ZipFile(before_path) as a,ZipFile(final_path) as b:
        assert set(a.namelist()) == set(b.namelist())
        assert b.testzip() is None
        for n in a.namelist():
            if n in ['xl/styles.xml','xl/vbaProject.bin'] or n.endswith('.rels') or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/')):
                assert a.read(n) == b.read(n),n
    return {'raw_monthly_and_non_balance_financial_values_unchanged':True,'past_month_balances_unchanged':True,
            'total_wrap_cash_unchanged':True,'no_new_income_expense_or_raw_transaction':True,
            'birthday_usage_and_settlement_unchanged':True,'vba_controls_and_styles_preserved':True,
            'all_validation_booleans_true':True}


def run(a):
    root=Path(a.workdir);root.mkdir(parents=True,exist_ok=True)
    candidate=root/'candidate.xlsm'
    timestamp=prepare(a.source,candidate)
    months=[];default=None
    for month in [9,10,11]:
        def query(p,m=month):
            p.write('잔액','B3',datetime(2026,m,1));p.write('예비비','B3',datetime(2026,m,1))
        path=recalculate(candidate,root/f'month_{month}',query)
        w=read(path);assert_no_errors(w);result=verify_month(w,month)
        expected=6388603 if month==9 else 4535466
        assert w['잔액']['G7'].value == expected
        assert w['잔액']['G15'].value == 0
        if month==10:default=path
        months.append(result)
        print(json.dumps({'month_passed':month,**result},ensure_ascii=False),flush=True)
        del w;gc.collect()
    cached=cache_original(candidate,default,a.output,a.source)
    checks=verify(a.source,a.output)
    report={'as_of':'2026-10-15','classification_recorded_at_kst':timestamp,'ledger_row':247,
            'from':'임시보관','to':'생활비','amount':61000,
            'living_balance':4535466,'temporary_balance':0,'wrap_total':6685466,
            'selected_month_usage':239000,'outstanding_reserve_settlement':0,
            'month_tests':months,'checks':checks,'formula_errors':0,'cached_formulas':cached,
            'baseline_sha256':sha256(Path(a.source).read_bytes()).hexdigest(),
            'workbook_sha256':sha256(Path(a.output).read_bytes()).hexdigest(),
            'recalculation':'Disposable macro-free XLSX in LibreOffice; original XLSM package retained.',
            'excel_ui_and_vba_executed':False}
    Path(a.report).write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'verified':True,'months':len(months),'hash':report['workbook_sha256']}),flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for name in ['source','output','workdir','report']:parser.add_argument('--'+name,required=True)
    run(parser.parse_args())
