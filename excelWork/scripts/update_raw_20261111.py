"""Import the November export without treating its balancing entry as proven interest.

Source fields and historical accounting choices are preserved. The 12,802 entry
is carried as submitted, explicitly flagged as an unresolved balance adjustment;
this import does not decide its economic classification or rewrite September.
"""
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json

import openpyxl

from redesign_reserve import ExtendedPackage
from raw_sort_id_20261101 import END, reorder
from update_postpaid_20261102 import asset, key, compact
from update_raw_20261104 import classify
from update_raw_20261101 import WRAP, SHARED
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore


def read_source(path):
    w = openpyxl.load_workbook(path, data_only=True)
    rows = iter(w.active.values)
    headers = next(rows)
    expected = {'날짜', '자산', '분류', '소분류', '내용', '금액(원)', '수입/지출', '메모', '금액', '화폐'}
    assert set(headers) in [expected, expected | {'태그'}]
    assert len(headers) == len(set(headers))
    result = []
    for row, values in enumerate(rows, 2):
        v = dict(zip(headers, values))
        if not v['날짜']:
            continue
        assert isinstance(v['날짜'], datetime) and v['날짜'].year == 2026 and v['날짜'].month == 11
        assert v['화폐'] == 'KRW' and v['금액(원)'] == v['금액']
        # Do not silently lose newly populated export columns.
        assert not v.get('태그'), ('Nonempty tag requires a storage decision', row)
        result.append(dict(source_row=row, date=v['날짜'], asset=asset(v['자산']),
                           category=v['분류'], sub=v['소분류'] or None, title=v['내용'],
                           amount=v['금액(원)'], kind=v['수입/지출'], memo=v['메모'] or None))
    w.close()
    return result


def identity(item):
    # A shared account may record several purchases at the exact same second.
    return (*key(item), compact(item['title']))


def plan(baseline, source):
    w = openpyxl.load_workbook(baseline, data_only=True)
    raw = w['RAW']; existing = defaultdict(list); old = {}; ids = {}
    for row in range(2, END + 1):
        ident = int(raw.cell(row, 42).value); ids[ident] = row
        date = raw.cell(row, 1).value
        if not isinstance(date, datetime):
            continue
        item = dict(date=date, asset=asset(raw.cell(row, 3).value), title=raw.cell(row, 6).value,
                    amount=raw.cell(row, 7).value, kind=raw.cell(row, 8).value)
        existing[identity(item)].append(ident); old[ident] = item
    assert set(ids) == set(range(2, END + 1))
    items = read_source(source); mappings = []; new = []; used = set()
    assert len({identity(item) for item in items}) == len(items)
    for item in items:
        hits = existing.get(identity(item), [])
        assert len(hits) <= 1, ('Ambiguous identity', item, hits)
        if hits:
            ident = hits[0]; assert ident not in used; used.add(ident)
            mappings.append(dict(source_row=item['source_row'], transaction_id=ident, status='preserved'))
        else:
            new.append(dict(item))
    missing = [i for i, item in old.items() if item['date'].month == 11 and i not in used]
    assert not missing, ('Changed/missing November records need reconciliation', missing)
    empty = sorted(i for i, row in ids.items() if not raw.cell(row, 1).value)
    assert len(empty) >= len(new)
    new.sort(key=lambda item: (item['date'], item['source_row']))
    for ident, item in zip(empty, new):
        item.update(transaction_id=ident, input_row=ids[ident])
        mappings.append(dict(source_row=item['source_row'], transaction_id=ident, status='added'))
    w.close()
    return dict(new=new, source_mapping=sorted(mappings, key=lambda item:item['source_row']),
                source_records=len(items), old_raw_count=len(old),
                old_november_count=sum(v['date'].month == 11 for v in old.values()),
                source_income_12802_is_unresolved_adjustment=True,
                original_income_fields_preserved_not_verified=True,
                september_historical_inputs_not_changed=True)


def prepare(baseline, source, output):
    meta = plan(baseline, source)
    p = ExtendedPackage(baseline)
    w = openpyxl.load_workbook(baseline, data_only=True)
    ledger_row = max(r for r in range(64, 501) if w['잔액'][f'A{r}'].value) + 1
    links = []
    for item in meta['new']:
        r = item['input_row']
        for c, field in [('A','date'),('C','asset'),('D','category'),('E','sub'),('F','title'),('G','amount'),('H','kind'),('I','memo')]:
            p.write('RAW', f'{c}{r}', item[field])
        classification = classify(item)
        if item['title'] in ['G마켓(1/6)', '코나아이(MONA)']:
            # Installment notation / telecom vendor, never its free-text memo.
            # No unsupported assumption that MONA replaces the U-mobile plan.
            classification = '고정비'
            p.write('RAW', f'AQ{r}', 26)
        if classification:
            p.write('RAW', f'J{r}', classification)
        if item['asset'] in SHARED:
            p.write('RAW', f'K{r}', 0); p.write('RAW', f'L{r}', '공동')
        incoming_transfer = item['category'] == WRAP and item['asset'] != WRAP
        if item['asset'] != WRAP and not incoming_transfer:
            continue
        assert item['amount'] > 0 and ledger_row <= 500
        incoming = incoming_transfer or item['kind'] == '수입'
        transfer = item['kind'].startswith('이체')
        if incoming_transfer:
            p.write('RAW', f'S{r}', '입금')
        sender, recipient = (item['asset'], WRAP) if incoming_transfer else (('외부', WRAP) if incoming else (WRAP, asset(item['category']) if transfer else '외부'))
        title = item['title']
        if item['amount'] == 12802 and incoming:
            title = '잔액 맞춤 입력 12,802원 · 원인 미확인(이자 확정 아님)'
        inputs = {'A':item['date'], 'B':'계좌이체' if transfer else ('입금' if incoming else '출금'),
                  'C':'외부' if incoming else '생활비', 'D':'생활비' if incoming else ('외부' if transfer else '지출'),
                  'E':item['amount'], 'F':title, 'J':sender, 'K':recipient,
                  'L':'WRAP계좌이체' if transfer else 'WRAP입출금', 'W':item['transaction_id']}
        for c, value in inputs.items():
            p.write('잔액', f'{c}{ledger_row}', value)
        links.append(dict(transaction_id=item['transaction_id'], ledger_row=ledger_row, amount=item['amount'], direction='입금' if incoming else '출금', provisional=item['amount']==12802 and incoming))
        ledger_row += 1
    p.write('11월', 'A70', '11/11 RAW 82건 · 12,802원은 잔액 맞춤 입력(원인 미확인). 원본 수입 입력을 보존했으나 확정 이자로 보지 않습니다.')
    p.write('잔액', 'D3', '11/11 RAW 반영 · 11월 입금 12,802원은 원인 미확인 잔액 조정. 원본 수입 입력 보존이며 이자 확정 아님.')
    p.write('사용안내', 'A26', '11/11 갱신 · 11월 RAW 82건. WRAP 12,802원은 원인 미확인 차액이며 원본 수입 입력을 잠정 보존합니다.')
    p.save(output); w.close()
    w = openpyxl.load_workbook(output, data_only=True)
    meta['sort'] = reorder(p, w, 'descending'); w.close(); p.save(output)
    w = openpyxl.load_workbook(output, data_only=True)
    for row, rr in enumerate(range(2, 12), 8):
        p.write('입력·연결', f'H{row}', w['RAW'].cell(rr, 42).value)
    w.close(); p.save(output); meta['wrap_links'] = links
    gc.collect()
    return meta


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    for name in ['baseline','source','output','root','metadata']:
        parser.add_argument(name, type=Path)
    a = parser.parse_args()
    candidate = a.root / 'candidate.xlsm'
    meta = prepare(a.baseline, a.source, candidate)
    a.metadata.write_text(json.dumps(meta, ensure_ascii=False, default=str, indent=2)+'\n')
    print('Prepared', len(meta['new']), 'new records;', len(meta['wrap_links']), 'WRAP links', flush=True)
    calculated = recalculate(candidate, a.root / 'calculation')
    import_changed_caches(candidate, calculated, a.root / 'cached.xlsm')
    result = restore(a.baseline, a.root / 'cached.xlsm', a.output)
    print('Memo cache restoration:', result, flush=True)
