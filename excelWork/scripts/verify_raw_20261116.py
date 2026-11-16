"""Check the import, actual stored block fills, and RAW edit/delete behavior."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json
import os
import subprocess

import fitz
import openpyxl
from lxml import etree as ET
from openpyxl.styles import PatternFill
from openpyxl.worksheet.properties import PageSetupProperties

from import_raw_20261116 import plan, source_rows
from redesign_reserve import ExtendedPackage
from verify_reserve_redesign import assert_no_errors, recalculate
from verify_raw_sort_20261101 import equal, memo_checks
from verify_reserve_cash_alignment import verify_month
from balance_selected_ledger import check_view
from validate_excel_structure import validate_package
from update_20261015 import Q


def index_of(r, c):
    start, offset = (15, 0) if r < 145 else (165, 2600)
    return offset + ((r-start)//5*10 + (c-2)//2)*10 + (c-2)%2*5 + (r-start)%5 + 1


def grid_coordinates():
    return [(r, c) for start, stop in [(15, 145), (165, 435)]
            for r in range(start, stop) for c in range(2, 22)]


def render_colors(path, folder, month=11):
    """Render real conditional formatting with base fills removed in a fixture.

    Cached category counts come from the full RAW recalculation. The delivery
    package is never saved with openpyxl or LibreOffice. PDF fill area provides
    an independent count of visible cells, including refunds/empty categories.
    """
    folder.mkdir()
    source = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    name = f'{month}월계산기'
    for s in list(source):
        if s.title != name:
            source.remove(s)
    s = source[name]
    source.active = 0
    source.defined_names.clear()
    for row in s:
        for cell in row:
            if cell.data_type == 'f' and not (cell.column == 31 and 3 <= cell.row <= 10):
                cell.value = cached[name][cell.coordinate].value
    # No physical colors: the PDF must be painted by the conditional rules.
    for r, c in grid_coordinates():
        s.cell(r, c).fill = PatternFill(patternType='solid', fgColor='FFFFFFFF')
    s.print_area = 'B15:U144'
    s.print_title_rows = None
    s.print_title_cols = None
    s.page_setup.paperSize = s.PAPERSIZE_A4
    s.page_setup.orientation = 'portrait'
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 1
    s.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    s.row_breaks.brk = []
    s.col_breaks.brk = []
    for hf in [s.oddHeader, s.oddFooter, s.evenHeader, s.evenFooter, s.firstHeader, s.firstFooter]:
        for part in [hf.left, hf.center, hf.right]:
            part.text = None
    file = folder/'grid.xlsx'
    source.save(file)
    source.close()
    palette = {r: cached[name][f'B{r}'].fill.fgColor.rgb for r in range(3, 11)}
    expected = {r: cached[name][f'C{r}'].value for r in range(3, 11)}
    cached.close()
    run = subprocess.run(['soffice', '-env:UserInstallation='+(folder/'profile').as_uri(),
                          '--headless', '--convert-to', 'pdf', '--outdir', str(folder), str(file)],
                         capture_output=True, text=True, timeout=90,
                         env={**os.environ, 'XDG_CACHE_HOME': str(folder/'cache')})
    assert run.returncode == 0 and (folder/'grid.pdf').is_file(), run.stderr
    doc = fitz.open(folder/'grid.pdf')
    assert len(doc) == 1, len(doc)
    page = doc[0]
    areas = Counter()
    rectangles = []
    for drawing in page.get_drawings():
        fill = drawing['fill']
        if fill is None:
            continue
        rgb = ''.join(f'{round(v*255):02X}' for v in fill)
        if rgb not in {v[-6:] for v in palette.values()}:
            continue
        rect = drawing['rect']
        areas[rgb] += rect.width*rect.height
        rectangles.append((rect.width, rect.height))
    # Calc merges adjacent cells with the same color into row rectangles.
    # Derive one cell's size from the complete printed 20-column grid bounds.
    drawings = [d for d in page.get_drawings() if d['fill'] is not None]
    xs = [d['rect'].x0 for d in drawings] + [d['rect'].x1 for d in drawings]
    cell_width = (max(xs)-min(xs))/20
    cell_height = min(h for width, h in rectangles)
    # PDF fills overlap at their edges by about .02 pt. Count whole columns
    # and rows in each merged rectangle, avoiding accumulated edge-area error.
    cells = Counter()
    for drawing in page.get_drawings():
        fill = drawing['fill']
        if fill is None:
            continue
        rgb = ''.join(f'{round(v*255):02X}' for v in fill)
        if rgb in {v[-6:] for v in palette.values()}:
            rect = drawing['rect']
            cells[rgb] += round(rect.width/cell_width)*round(rect.height/cell_height)
    actual = {r: cells[code[-6:]] for r, code in palette.items()}
    for r, n in expected.items():
        assert actual[r] == n, (month, r, actual[r], n)
    page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(folder/'grid.png')
    doc.close()
    return dict(month=month, conditional_colors_rendered=True, base_fills_removed_in_fixture=True,
                expected_cells=expected, rendered_cells={r: round(n) for r, n in actual.items()},
                pdf_pages=1, renderer='LibreOffice 26.8 / PDF drawing areas')


def verify(baseline, final, source, metadata, report_path):
    meta = json.loads(metadata.read_text())
    a = openpyxl.load_workbook(baseline, data_only=True)
    w = openpyxl.load_workbook(final, data_only=True)
    assert_no_errors(w)
    raw, old = w['RAW'], a['RAW']
    ids = {int(raw.cell(r,42).value): r for r in range(2,4002)}
    prev = {int(old.cell(r,42).value): r for r in range(2,4002)}
    assert set(ids) == set(prev) == set(range(2,4002))
    rows = [r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    assert len(rows) == 2741 and len(meta['new']) == 8 and len(meta['patches']) == 3
    assert Counter(raw.cell(r,1).value.month for r in rows)[11] == 140
    assert raw.tables['RawData2026'].ref == 'A1:AQ4001'
    assert [(s.title, s.sheet_state) for s in w] == [(s.title, s.sheet_state) for s in a]
    assert '검증' not in w.sheetnames
    allowed = {p['transaction_id']: set(p['changes']) for p in meta['patches']}
    checked = 0
    for ident, pr in prev.items():
        if not isinstance(old.cell(pr,1).value,datetime):
            continue
        for c, key in [(1,'A'),(3,'C'),(4,'D'),(5,'E'),(6,'F'),(7,'G'),(8,'H'),
                       (9,'I'),(10,'J'),(11,'K'),(12,'L'),(13,'M'),(43,'AQ')]:
            if key not in allowed.get(ident,set()):
                equal(old.cell(pr,c).value, raw.cell(ids[ident],c).value, ('preserved',ident,key))
                checked += 1
    src = {v['source_row']: v for v in source_rows(source)}
    assert len(src) == len(meta['source_mapping']) == 414
    assert {v['source_row'] for v in meta['source_mapping']} == set(src)
    for link in meta['source_mapping']:
        if link['status'] in ['previous_cumulative_interest_preserved', 'statement_amount_and_date_preserved',
                              'statement_amount_preserved', 'cumulative_interest_split_635']:
            continue
        r = ids[link['transaction_id']]
        v = src[link['source_row']]
        equal(raw.cell(r,7).value, v['amount'], ('source amount',v['source_row']))
        equal(raw.cell(r,9).value, v['memo'], ('source memo',v['source_row']))
    refund = [r for r in rows if raw.cell(r,6).value == 'NH농협카드 연회비' and raw.cell(r,7).value == -6000]
    assert refund == [ids[2708]] and raw.cell(ids[2708],3).value == '토스뱅크'
    assert raw.cell(ids[2708],1).value.day == 16
    assert raw.cell(ids[2734],7).value == 158
    assert raw.cell(ids[2706],7).value - raw.cell(ids[2705],7).value == 158
    for month in [*range(1,11),12]:
        for addr in ['I33','I38','C40','E74','E75','E79','H43']:
            equal(a[f'{month}월'][addr].value, w[f'{month}월'][addr].value, (month,addr))
    nov = w['11월']
    assert sum(v['amount'] for v in meta['new'] if v['kind'] == '지출') == 32187
    assert nov['I33'].value == 1262901 == a['11월']['I33'].value + 32187 + 1
    assert nov['I38'].value == w['11월계산기']['W12'].value == 1284000
    assert nov['E74'].value == 626515 and nov['E75'].value == 1400978
    assert nov['E79'].value == nov['E78'].value == 1889416
    # The refund moves dates: no second refund and no second card payment cost.
    for header in [3,8,13,18,23,28]:
        for c in range(2,9):
            dt = nov.cell(header,c).value
            if not isinstance(dt,datetime) or dt.month != 11:
                continue
            delta = {12:1,13:6000,16:26187}.get(dt.day,0)
            rounded = {13:6000,16:27000}.get(dt.day,0)
            equal(nov.cell(header+2,c).value or 0, (a['11월'].cell(header+2,c).value or 0)+delta, ('daily actual',dt.day))
            equal(nov.cell(header+1,c).value or 0, (a['11월'].cell(header+1,c).value or 0)+rounded, ('daily rounded',dt.day))
    ledger = w['잔액']
    assert sum(v['amount'] for v in meta['bank_links']) == 2214709
    assert ledger['G16'].value == 5048217 == a['잔액']['G16'].value-2214709
    assert ledger['G7'].value == 2898217
    assert ledger['G17'].value == w['예비비']['I13'].value == 2150000
    assert '5,047,690' in ledger['D3'].value and '527' in ledger['D3'].value
    reserves = verify_month(w,11)
    view = check_view(w,11)
    memos = memo_checks(w)
    for row in range(64,501):
        if ledger[f'A{row}'].value:
            assert ledger[f'I{row}'].value == 'OK', row
    physical = {}
    for month in [11,12]:
        s = w[f'{month}월계산기']
        palette = {s[f'B{r}'].fill.fgColor.rgb:r for r in range(3,11)}
        counts = Counter()
        for r,c in grid_coordinates():
            assert s.cell(r,c+31).value == index_of(r,c)
            color = s.cell(r,c).fill.fgColor.rgb
            if color in palette:
                counts[palette[color]] += 1
        expected = {r:s[f'C{r}'].value for r in range(3,11)}
        assert all(counts[r] == n for r,n in expected.items()), (month,counts,expected)
        assert sum(counts.values())*1000 == s['W12'].value
        physical[str(month)] = dict(counts=expected, stored_colored_cells=sum(counts.values()))
    p, q = ExtendedPackage(baseline), ExtendedPackage(final)
    assert p.parts.keys() == q.parts.keys()
    protected = [n for n in p.parts if n.endswith(('.bin','.rels')) or n.startswith(
        ('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n] == q.parts[n] for n in protected)
    before, after = p.tree('xl/styles.xml'), q.tree('xl/styles.xml')
    for component in before:
        new = after.find(component.tag)
        if component.tag in [Q('cellXfs'),Q('fills')]:
            assert len(new) >= len(component)
            assert all(ET.tostring(x) == ET.tostring(new[i]) for i,x in enumerate(component))
        else:
            assert ET.tostring(component) == ET.tostring(new), component.tag
    for month in range(1,13):
        for suffix in ['월','월계산기']:
            name = f'{month}{suffix}'
            old_cf = p.sheet(name).findall(Q('conditionalFormatting'))
            new_cf = q.sheet(name).findall(Q('conditionalFormatting'))
            if suffix == '월계산기' and month in [11,12]:
                assert len(new_cf) == 9
                assert all(len(x.find(Q('cfRule')).find(Q('formula')).text) < 50 for x in new_cf)
            else:
                assert list(map(ET.tostring,old_cf)) == list(map(ET.tostring,new_cf)), name
    structure = validate_package(final)
    assert not structure['errors'], structure
    result = dict(as_of='2026-11-16', workbook_sha256=sha256(final.read_bytes()).hexdigest(),
                  source_sha256=sha256(source.read_bytes()).hexdigest(), source_records=414,
                  total_raw_records=len(rows), new_records=8, updated_records=3,
                  raw_by_month=dict(Counter(raw.cell(r,1).value.strftime('%Y-%m') for r in rows)),
                  historical_input_cells_preserved=checked, formula_errors=0,
                  november={addr:nov[addr].value for addr in ['I33','I38','E74','E75','E78','E79']},
                  refund_count=1, revised_late_fee=158, managed_cash_total=5048217,
                  naver_cash=5047690, wrap_residual=527, purpose_available=2150000,
                  monthly_memos=memos, reserves=reserves, selected_ledger=view,
                  calculator_physical_fills=physical, vba_controls_relationships_preserved=True,
                  original_style_ids_preserved=True, structure=structure, metadata=meta,
                  source_original_local_only=True, native_excel_ui_executed=False)
    a.close();w.close();gc.collect()
    again,_ = plan(final,source)
    assert not again['new'] and not again['patches'] and not again['derived'], again
    result.update(repeat_new_records=0, repeat_updates=0)
    report_path.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Financial, stored colors, memos, structure and repeat-import checks passed.',flush=True)
    return result


def scenario(final,root,report):
    w = openpyxl.load_workbook(final,data_only=True)
    ids = {int(w['RAW'].cell(r,42).value):r for r in range(2,4002)}
    empty = next(r for r in range(2,4002) if not w['RAW'].cell(r,1).value)
    def edit(p):
        rr = ids[2738]
        p.write('RAW',f'G{rr}',11888)
        if w['RAW'].cell(rr,11).value not in [None,'']:
            p.write('RAW',f'K{rr}',11888)
        for c in ['A',*list('CDEFGHIJKLMN'),'AQ']:
            p.write('RAW',f'{c}{ids[2740]}',None)
        p.write('RAW',f'I{ids[2708]}',None)
        for c,value in dict(A=datetime(2026,12,1),C='현대네이버카드',D='식비',E='식대',
                            F='검증용 신규 결제',G=1000,H='지출',J='변동비').items():
            p.write('RAW',f'{c}{empty}',value)
    calculated = recalculate(final,root/'raw_edit_scenario',edit)
    w.close()
    s = openpyxl.load_workbook(calculated,data_only=True)
    assert_no_errors(s)
    assert s['11월']['I33'].value == 1258402
    assert s['11월']['I38'].value == s['11월계산기']['W12'].value == 1279000
    assert s['12월']['I33'].value == 1000
    assert s['12월계산기']['W12'].value == s['12월']['I38'].value == 1000
    assert s['RAW'].cell(ids[2708],40).value in [None,'']
    assert s['잔액']['G16'].value == 5048217
    memo_checks(s)
    s.close()
    report['current_color_render'] = render_colors(final,root/'render_current')
    report['edited_color_render'] = render_colors(calculated,root/'render_edited')
    report['december_color_render'] = render_colors(calculated,root/'render_december',12)
    report['raw_edit_delete_new_month_and_memo_clear_scenario_passed'] = True
    print('RAW edit/delete, December addition, memo clear and PDF colors passed.',flush=True)


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    for name in ['baseline','final','source','metadata','root','report']:
        p.add_argument(name,type=Path)
    a = p.parse_args()
    report = verify(a.baseline,a.final,a.source,a.metadata,a.report)
    scenario(a.final,a.root,report)
    a.report.write_text(json.dumps(report,ensure_ascii=False,default=str,indent=2)+'\n')
