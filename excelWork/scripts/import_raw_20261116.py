"""Nov 16 source reconciliation and persisted, automatically refreshed blocks.

Use stable transaction IDs; bank-confirmed historical corrections take priority
over stale export values. No financial classification depends on a memo.
Patch the XLSM package, never save it through a spreadsheet application.
"""
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json

import openpyxl
from lxml import etree as ET
from openpyxl.utils import get_column_letter as col

from import_raw_20261115 import plan as previous_plan, source_rows, CASH, NEW_CLASS_CODES
from update_raw_20261104 import classify
from update_raw_20261101 import SHARED
from redesign_reserve import ExtendedPackage
from raw_sort_id_20261101 import reorder
from raw_auto_link_20261026 import conditional
from update_20261015 import Q
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
    """Store visible fills, and keep short conditional rules for future edits.

    Existing rules rendered in LibreOffice, but no November base fills were
    saved. Storing both representations also supplies the grid to previews
    that do not evaluate conditional formatting. Only 11/12 calculators change.
    """
    w = openpyxl.load_workbook(candidate, data_only=True)
    p = ExtendedPackage(candidate)
    styles = p.tree('xl/styles.xml')
    xfs = styles.find(Q('cellXfs'))
    dxfs = styles.find(Q('dxfs'))
    def style_key(node):
        # Ignore serialization-only namespace declarations and attribute order.
        return (node.tag, tuple(sorted(node.attrib.items())), node.text,
                tuple(style_key(child) for child in node))

    keys = {style_key(x): i for i, x in enumerate(xfs)}
    fill_styles = {}

    def style_with_fill(style, fill, border):
        key = (style, fill, border)
        if key not in fill_styles:
            xf = deepcopy(xfs[style])
            xf.set('fillId', str(fill))
            xf.set('applyFill', '1')
            xf.set('borderId', str(border))
            xf.set('applyBorder', '1')
            serial = style_key(xf)
            if serial not in keys:
                keys[serial] = len(xfs)
                xfs.append(xf)
            fill_styles[key] = keys[serial]
        return fill_styles[key]

    result = {}
    for month in [11, 12]:
        name = f'{month}월계산기'
        sheet = p.sheet(name)
        old_rules = sheet.findall(Q('conditionalFormatting'))
        assert len(old_rules) == 9
        existing_dxfs = [int(x.find(Q('cfRule')).get('dxfId')) for x in old_rules]
        by_rgb = {dxfs[i].find(Q('fill')).find(Q('patternFill')).find(Q('fgColor')).get('rgb'): i
                  for i in existing_dxfs}
        palette_rgb = {r: w[name][f'B{r}'].fill.fgColor.rgb for r in range(3, 11)}
        color_dxfs = {r: by_rgb[rgb] for r, rgb in palette_rgb.items()}
        white_dxf = by_rgb['FFFFFFFF']
        white = dxfs[white_dxf].find(Q('fill'))
        assert white.find(Q('patternFill')).find(Q('fgColor')).get('rgb') == 'FFFFFFFF'
        fills = styles.find(Q('fills'))
        white_ids = [i for i, fill in enumerate(fills) if
                     ET.tostring(fill) == ET.tostring(white)]
        if not white_ids:
            white_ids = [len(fills)]
            fills.append(deepcopy(white))
            fills.set('count', str(len(fills)))
        palette_fills = {r: int(xfs[int(p.cell(name, f'B{r}').get('s'))].get('fillId'))
                         for r in range(3, 11)}
        counts = {r: w[name][f'C{r}'].value for r in range(3, 11)}
        assert all(isinstance(n, (int, float)) and n >= 0 and int(n) == n for n in counts.values())
        total = sum(counts.values())
        assert total <= 8000
        cumulative = {}
        running = 0
        for r, n in counts.items():
            running += n
            cumulative[r] = running
        # Helpers are positions only. They do not contain transaction/memo data.
        for index, (r, c) in enumerate(slots(), 1):
            addr = f'{col(c)}{r}'
            cell = p.cell(name, addr)
            category = next((k for k, end in cumulative.items() if index <= end), None)
            fill = palette_fills[category] if category is not None else white_ids[0]
            start = 15 if r < 145 else 165
            template = p.cell(name, f'{col(c)}{20 + (r-start)%5}')
            border = int(xfs[int(template.get('s'))].get('borderId'))
            cell.set('s', str(style_with_fill(int(cell.get('s', '0')), fill, border)))
            p.write(name, f'{col(c + 31)}{r}', index)
        p.write(name, 'AE2', 0)
        for r in range(3, 11):
            p.write(name, f'AE{r}', f'=SUM($C$3:C{r})', True)
        for node in old_rules:
            sheet.remove(node)
        for r in range(3, 11):
            conditional(p, name, 'B15:U144 B165:U434',
                        f'AND(AG15>$AE${r-1},AG15<=$AE${r})', color_dxfs[r], r-2)
        conditional(p, name, 'B15:U144 B165:U434',
                    'AG15>SUM($C$3:$C$10)', white_dxf, 9)
        cols = sheet.find(Q('cols'))
        for node in list(cols):
            if int(node.get('min')) >= 33:
                cols.remove(node)
        ET.SubElement(cols, Q('col'), min='33', max='52', hidden='1', width='12', customWidth='1')
        sheet.find(Q('dimension')).set('ref', 'A1:AZ434')
        p.write(name, 'A2', 'RAW 입력·수정·삭제 시 자동 갱신됩니다. 1칸=1,000원 · 별도 버튼 없이 색상과 합계가 바뀝니다.')
        p.write(name, 'A154', 'RAW 자동 연동 · 색상 칸과 합계를 함께 갱신합니다. 금액은 RAW에서 수정하세요.')
        formulas = [n.find(Q('cfRule')).find(Q('formula')).text
                    for n in sheet.findall(Q('conditionalFormatting'))]
        result[str(month)] = dict(counts=counts, physical_filled_cells=total,
                                  conditional_rules=9, max_formula_length=max(map(len, formulas)), capacity=8000)
    xfs.set('count', str(len(xfs)))
    p.save(output)
    w.close()
    gc.collect()
    return result


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
