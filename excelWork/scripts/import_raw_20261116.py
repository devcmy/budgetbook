"""Nov 16 source reconciliation and persisted, automatically refreshed blocks.

Use stable transaction IDs; bank-confirmed historical corrections take priority
over stale export values. No financial classification depends on a memo.
Patch the XLSM package, never save it through a spreadsheet application.
"""
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json

import openpyxl

from import_raw_20261115 import plan as previous_plan, source_rows, CASH, NEW_CLASS_CODES
from update_raw_20261104 import classify
from update_raw_20261101 import SHARED
from redesign_reserve import ExtendedPackage
from raw_sort_id_20261101 import reorder
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore


def plan(baseline, source):
    meta, ids = previous_plan(baseline, source)
    w = openpyxl.load_workbook(baseline, data_only=True)
    s = w['RAW']
    # The revised export replaces, rather than repeats, the earlier card credit.
    # Link this one explicitly reconciled event by ID, not by its editable memo.
    revised = [v for v in meta['new'] if v['title'] == 'NH농협카드 연회비'
               and v['asset'] == '토스뱅크' and v['amount'] == -6000
               and v['date'].date() == datetime(2026, 11, 16).date()]
    if revised:
        assert len(revised) == 1 and 2708 in meta['retained_old_ids']
        old_row = ids[2708]
        assert s.cell(old_row, 3).value == 'NH별다줄카드'
        assert s.cell(old_row, 6).value == 'NH농협카드 연회비'
        assert s.cell(old_row, 7).value == -6000
        v = revised[0]
        meta['new'].remove(v)
        meta['retained_old_ids'].remove(2708)
        meta['patches'].append(dict(transaction_id=2708, old_amount=-6000,
                                    changes=dict(A=v['date'], C=v['asset'], I=v['memo'])))
        link = next(x for x in meta['source_mapping'] if x['source_row'] == v['source_row'])
        link.update(transaction_id=2708, status='revised_refund_date_and_account')
        meta['memo_edits'].append(2708)
    # Settlement was revised by one won. Preserve the funding transfer and
    # update the existing separate fee component instead of adding another fee.
    assert s.cell(ids[2734], 6).value == '현대네이버카드 지연금'
    source_by_row = {v['source_row']: v for v in source_rows(source)}
    linked = {v['transaction_id']: source_by_row[v['source_row']] for v in meta['source_mapping']}
    fee = linked[2706]['amount'] - linked[2705]['amount']
    assert fee == 158
    meta['derived'] = []
    if s.cell(ids[2734], 7).value != fee:
        meta['patches'].append(dict(transaction_id=2734, old_amount=s.cell(ids[2734], 7).value,
                                    changes=dict(G=fee, K=fee)))
    meta.update(as_of='2026-11-16', refund_replaces_transaction_id=2708,
                revised_fee=fee, fee_parent_transaction_ids=[2706, 2705])
    w.close()
    return meta, ids


def prepare(baseline, source, output):
    meta, ids = plan(baseline, source)
    p = ExtendedPackage(baseline)
    w = openpyxl.load_workbook(baseline, data_only=True)
    ledger = w['잔액']
    links = {ledger[f'W{r}'].value: r for r in range(64, 501) if ledger[f'W{r}'].value}
    ledger_row = max(r for r in range(64, 501) if ledger[f'A{r}'].value) + 1
    for patch in meta['patches']:
        ident = patch['transaction_id']
        for c, value in patch['changes'].items():
            p.write('RAW', f'{c}{ids[ident]}', value)
        assert ident not in links, ('Linked ledger correction needs explicit reconciliation', ident)
    bank_links = []
    for v in meta['new']:
        r = v['input_row']
        for c, key in [('A', 'date'), ('C', 'asset'), ('D', 'category'), ('E', 'sub'),
                       ('F', 'title'), ('G', 'amount'), ('H', 'kind'), ('I', 'memo')]:
            p.write('RAW', f'{c}{r}', v[key])
        classification = classify(v)
        if v['title'] in NEW_CLASS_CODES:
            classification, code = NEW_CLASS_CODES[v['title']]
            p.write('RAW', f'AQ{r}', code)
        if classification:
            p.write('RAW', f'J{r}', classification)
        if v['asset'] in SHARED:
            p.write('RAW', f'K{r}', 0)
            p.write('RAW', f'L{r}', '공동')
        src, dst = v['asset'], v['category']
        if src not in CASH and dst not in CASH:
            continue
        assert src in CASH and dst not in CASH and v['kind'] == '이체지출'
        inputs = dict(A=v['date'], B='계좌이체', C='생활비', D='외부', E=v['amount'],
                      F=f'{src} → {dst}', J=src, K=dst, L='WRAP계좌이체', W=v['transaction_id'])
        assert ledger_row <= 500 and inputs['E'] > 0
        for c, value in inputs.items():
            p.write('잔액', f'{c}{ledger_row}', value)
        bank_links.append(dict(ledger_row=ledger_row, transaction_id=v['transaction_id'],
                               direction=-1, amount=v['amount']))
        ledger_row += 1
    p.write('11월', 'A70', '11/16 원장 반영 · 연회비 환급은 날짜·계좌를 정정해 한 번만 반영. 카드대금 송금은 소비에 재가산하지 않습니다.')
    p.write('사용안내', 'A26', '11/16 갱신 · RAW 추가·수정·삭제는 월별 달력과 계산기에 연동됩니다. 11·12월 계산기는 색상도 자동 갱신되며 RAW I열 메모는 분류에 쓰지 않습니다.')
    p.save(output)
    w.close()
    values = openpyxl.load_workbook(output, data_only=True)
    meta['sort'] = reorder(p, values, 'descending')
    values.close()
    p.save(output)
    values = openpyxl.load_workbook(output, data_only=True)
    for r, rr in enumerate(range(2, 12), 8):
        p.write('입력·연결', f'H{r}', values['RAW'].cell(rr, 42).value)
    values.close()
    p.save(output)
    meta['bank_links'] = bank_links
    gc.collect()
    return meta


def slots():
    """Same two-column / five-row blocks as the original 1,000-won grid."""
    for start, stop in [(15, 145), (165, 435)]:
        for r0 in range(start, stop, 5):
            for c0 in range(2, 22, 2):
                for c in range(c0, c0 + 2):
                    for r in range(r0, r0 + 5):
                        yield r, c


def install_blocks(candidate, output):
    """Keep the shared importer on the user\'s chronological layout."""
    from chronological_blocks_20261116 import install
    return install(candidate, output)

if __name__ == '__main__':
    a = argparse.ArgumentParser()
    for name in ['baseline', 'source', 'root', 'output', 'metadata']:
        a.add_argument(name, type=Path)
    args = a.parse_args()
    args.root.mkdir(exist_ok=True)
    meta = prepare(args.baseline, args.source, args.root / 'candidate.xlsm')
    print('Merge:', len(meta['new']), 'new,', len(meta['patches']), 'corrected', flush=True)
    calc = recalculate(args.root / 'candidate.xlsm', args.root / 'numbers')
    import_changed_caches(args.root / 'candidate.xlsm', calc, args.root / 'numbers.xlsm')
    meta['blocks'] = install_blocks(args.root / 'numbers.xlsm', args.root / 'blocks.xlsm')
    calc = recalculate(args.root / 'blocks.xlsm', args.root / 'calculation')
    import_changed_caches(args.root / 'blocks.xlsm', calc, args.root / 'cached.xlsm')
    meta['memo_cache_restore'] = restore(args.baseline, args.root / 'cached.xlsm', args.output)
    args.metadata.write_text(json.dumps(meta, ensure_ascii=False, default=str, indent=2) + '\n')
    print('Saved:', args.output, flush=True)
