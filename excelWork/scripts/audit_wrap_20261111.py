"""Reconstruct September netted records; reconcile October/November source links.

This proves consistency of supplied records, not the completeness of an external
bank statement. Historical month-boundary matches are evidence, not bank IDs.
"""
from collections import Counter
from datetime import datetime
from pathlib import Path
import argparse
import json

import openpyxl

from update_raw_20261101 import WRAP
from update_postpaid_20261102 import read_source, key, asset
from verify_reserve_cash_alignment import ACCOUNTS


def audit(path, older_source, output):
    w = openpyxl.load_workbook(path, data_only=True)
    raw = {int(r[41]): r for r in w['RAW'].iter_rows(min_row=2, max_row=4001, values_only=True)
           if isinstance(r[0], datetime)}
    ledger = w['잔액']
    wr = {i:r for i,r in raw.items() if r[0] >= datetime(2026,9,1) and WRAP in [r[2],r[3]]}
    def signed(r):
        return r[6] if r[3] == WRAP or r[7] == '수입' else -r[6]
    groups = [
        (192, [2087], '9/1 이자 4,467: 기존 8월 이자에 동일 금액'),
        (195, [2088], '9/1 카카오페이→WRAP 814,000: 기존 8월 여행 정산에 동일 금액'),
        (210, [2080], '9/1 샴푸 11,000: 기존 8월 샴푸에 동일 내용·금액'),
        (228, [2090], '신한→WRAP 165,100'),
        (218, [2105], 'WRAP→신한V2 7,030'),
        (216, [2157], '보증보험료 14,701'),
        (217, [2173], '전세대출이자 준비 이체 93,420'),
        (229, [2185,2186,2187,2188,2189], '5개 은행 송금 합계 4,848,153'),
        (215, [2190], '9/13 이자 8,803'),
        (222, [2203], '생일 50,000'),
        (223, [2292], '가족통장 200,000'),
        (214, [2293,2298,2305,2286,2299,2300,2303,2304], '월급 후 순정산 2,082,285'),
    ]
    bridge = []
    used = []
    for lr, ids, explanation in groups:
        expected = sum(signed(wr[i]) for i in ids)
        a,b,amount = [ledger[f'{c}{lr}'].value for c in ['C','D','E']]
        actual = amount * ((b in ACCOUNTS) - (a in ACCOUNTS))
        assert expected == actual, (lr, expected, actual)
        used.extend(ids)
        bridge.append(dict(ledger_row=lr, ledger_date=ledger[f'A{lr}'].value,
                           transaction_ids=ids, raw_signed_total=expected,
                           ledger_signed_total=actual, explanation=explanation))
    sep = {i:r for i,r in wr.items() if r[0].month == 9}
    assert set(used) == set(sep) and len(used) == len(set(used)) == 23
    sep_net = sum(signed(r) for r in sep.values())
    historical_net = 0
    for lr in range(64,501):
        date = ledger[f'A{lr}'].value
        if isinstance(date,datetime) and date.month == 9:
            a,b,amount = [ledger[f'{c}{lr}'].value for c in ['C','D','E']]
            historical_net += amount * ((b in ACCOUNTS) - (a in ACCOUNTS))
    boundary = 814000 + 4467 - 11000
    assert sep_net == -2149649 and historical_net == -2957116
    assert sep_net - historical_net == boundary == 807467
    assert sep_net - boundary - historical_net == 0
    october = {i:r for i,r in wr.items() if r[0].month==10}
    november = {i:r for i,r in wr.items() if r[0].month==11}
    links = Counter(); linked = []
    for lr in range(64,501):
        dt = ledger[f'A{lr}'].value
        if not isinstance(dt,datetime) or dt < datetime(2026,10,1):
            continue
        ids = [ledger[f'{c}{lr}'].value for c in ['W','X'] if ledger[f'{c}{lr}'].value]
        if not ids:
            # Existing 61,000 purpose transfer changes no bank balance.
            assert ledger[f'C{lr}'].value in ACCOUNTS and ledger[f'D{lr}'].value in ACCOUNTS
            continue
        for ident in ids:
            assert ident in wr; links[ident] += 1
        a,b,amount = [ledger[f'{c}{lr}'].value for c in ['C','D','E']]
        expected = sum(signed(wr[i]) for i in ids)
        assert expected == amount * ((b in ACCOUNTS) - (a in ACCOUNTS))
        assert all(wr[i][0] == dt for i in ids)
        linked.append(dict(ledger_row=lr, transaction_ids=ids, signed_amount=expected))
    assert set(links)==set(october)|set(november) and all(n==1 for n in links.values())
    assert len(october)==24 and len(november)==4
    october_net=sum(signed(r) for r in october.values())
    assert october_net==-103862
    assert ledger['I32'].value+october_net==ledger['J32'].value==8434741
    provisional=[i for i,r in november.items() if r[6]==12802 and r[7]=='수입']
    assert len(provisional)==1
    known_nov_net=sum(signed(r) for i,r in november.items() if i not in provisional)
    assert known_nov_net==-159720
    assert ledger['J32'].value+known_nov_net==8275021
    assert ledger['K32'].value==8275021+12802==8287823
    # The previously deferred September source differences contain no WRAP side.
    src=read_source(older_source)
    source_wr=[item for item in src if item['date'].month==9 and WRAP in [item['asset'],item['category']]]
    assert Counter(key(item) for item in source_wr)==Counter((r[0],asset(r[2]),r[6],r[7]) for r in sep.values())
    result=dict(as_of='2026-11-11', conclusion='12,802 is a user-entered balancing difference, not verified interest. No matching omitted WRAP transaction found in supplied September-onward records.',
                september_raw_count=23,september_bridge=bridge,
                september_raw_net=sep_net,september_legacy_net=historical_net,
                september_month_boundary_difference=boundary,september_bridge_residual=0,
                september_bridge_limit='Month labels and netted aggregate amounts reconstructed from existing records; no bank transaction IDs establish these historical matches.',
                september_older_export_wrap_records_identical=True,
                october_raw_count=24,october_net=october_net,october_missing_or_duplicate_links=0,
                november_wrap_count=4,november_known_outflows=159720,
                november_balance_without_unresolved_adjustment=8275021,
                november_balance_with_user_adjustment=8287823,
                user_adjustment_transaction_id=provisional[0],unresolved_amount=12802,
                links=linked,bank_statement_verified=False,
                historical_direction_helper_issue=[i for i,r in sep.items() if r[3]==WRAP and r[18]=='출금'],
                historical_direction_helper_issue_scope='Display helper error only; legacy balances use reconstructed monthly ledger, not this direction helper. Not changed in this diagnostic update.')
    output.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:result[k] for k in ['september_bridge_residual','october_missing_or_duplicate_links','unresolved_amount','november_balance_without_unresolved_adjustment','november_balance_with_user_adjustment']},ensure_ascii=False))
    w.close()
    return result


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['workbook','older_source','output']:p.add_argument(name,type=Path)
    a=p.parse_args();audit(a.workbook,a.older_source,a.output)
