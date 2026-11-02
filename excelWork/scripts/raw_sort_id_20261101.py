"""Enable native RAW table sorting without moving transaction relationships.

Patch the existing XLSM ZIP; never save it with openpyxl. Immutable IDs replace
physical RAW row pointers. Fully qualified #This Row references are understood
by both Excel and the disposable recalculation engine (local @ refs are not).
"""
from collections import Counter
from copy import deepcopy
from pathlib import Path
import argparse
import json
import re

from lxml import etree as ET
import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries

from redesign_reserve import ExtendedPackage
from update_20261015 import Q

TABLE = 'RawData2026'
END = 4001
ID = '거래ID'
IDS = f'{TABLE}[{ID}]'


def outside_strings(text, fn):
    """Do not rewrite addresses embedded in memos or HYPERLINK strings."""
    return ''.join(part if i % 2 else fn(part) for i, part in
                   enumerate(re.split(r'("(?:[^"]|"")*")', text)))


def field(header):
    return f'{TABLE}[{header}]'


def current(header):
    return f'{TABLE}[[#This Row],[{header}]]'


def lookup(header, ident):
    return f'INDEX({field(header)},MATCH({ident},{IDS},0))'


def rewrite_raw(formula, row, headers):
    tokens = Tokenizer('=' + formula.lstrip('=')).items
    for token in tokens:
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            match = re.fullmatch(r'\$?([A-Z]+)\$?(\d+)', token.value)
            if match:
                assert int(match[2]) == row, (row, formula, token.value)
                token.value = current(headers[column_index_from_string(match[1])])
    text = ''.join(t.value for t in tokens)
    return outside_strings(text, lambda x: re.sub(r'\bROW\(\)', current(ID), x, flags=re.I))


def duplicate_formula():
    key = current('거래키')
    escaped = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({key},"~","~~"),"*","~*"),"?","~?")'
    equal = '*'.join(f'({field(h)}={current(h)})' for h in
                     ['유효일시','자산','분류','소분류','내용','실금액','수입/지출','동일거래 구분'])
    return (f'IF({key}="","",IF(LEN({escaped})<=250,'
            f'COUNTIFS({field("거래키")},{escaped},{IDS},"<="&{current(ID)}),'
            f'SUMPRODUCT({equal}*({IDS}<={current(ID)}))))')


def memo_sequence():
    month, memo = current('월'), current('연동메모')
    return (f'IF(OR({month}="",{memo}=""),"",IF(YEAR({month})<>2026,"",'
            f'MONTH({month})*10000+COUNTIFS({field("월")},{month},'
            f'{field("연동메모")},"?*",{IDS},"<="&{current(ID)})))')


def index_calls(text, callback):
    """Rewrite nested INDEX calls; respect quoted text and table brackets."""
    out, pos = [], 0
    while pos < len(text):
        if text[pos] == '"':
            end = pos+1
            while end < len(text):
                if text[end] == '"':
                    if end+1 < len(text) and text[end+1] == '"':
                        end += 2; continue
                    end += 1; break
                end += 1
            out.append(text[pos:end]); pos = end; continue
        match = re.match(r'INDEX\(', text[pos:], re.I)
        if not match:
            out.append(text[pos]); pos += 1; continue
        start, depth, bracket, quoted, args, arg_start = pos+6, 1, 0, False, [], pos+6
        i = start
        while i < len(text):
            char = text[i]
            if char == '"':
                if quoted and i+1 < len(text) and text[i+1] == '"':
                    i += 2; continue
                quoted = not quoted
            elif not quoted:
                if char == '[': bracket += 1
                elif char == ']': bracket -= 1
                elif not bracket:
                    if char == '(': depth += 1
                    elif char == ')':
                        depth -= 1
                        if depth == 0:
                            args.append(text[arg_start:i]); break
                    elif char == ',' and depth == 1:
                        args.append(text[arg_start:i]); arg_start = i+1
            i += 1
        assert depth == 0, text[pos:]
        args = [index_calls(a, callback) for a in args]
        replacement = callback(args)
        out.append(replacement if replacement is not None else 'INDEX('+','.join(args)+')')
        pos = i+1
    return ''.join(out)


def rewrite_external(formula, headers):
    def indexed(args):
        if len(args) != 2: return None
        source, row = args
        raw_column = re.fullmatch(r"(?:RAW|'RAW')!\$?([A-Z]+):\$?\1", source)
        if raw_column:
            return lookup(headers[column_index_from_string(raw_column[1])], row)
        column = re.fullmatch(TABLE+r'\[([^\[\]]+)\]', source)
        if column:
            if row.isdigit():
                return lookup(column[1], str(int(row)+1))
            shifted = re.fullmatch(r'(\$?[A-Z]+\$?\d+)-1', row)
            if shifted:
                return lookup(column[1], shifted[1])
        return None
    result = index_calls(formula.lstrip('='), indexed)
    def direct(text):
        text = re.sub(r"(?:RAW|'RAW')!\$?([A-Z]+)\$?2:\$?\1\$?10018",
                      lambda m: field(headers[column_index_from_string(m[1])]), text)
        return re.sub(r"(?:RAW|'RAW')!\$?([A-Z]+)\$?(\d+)",
                      lambda m: lookup(headers[column_index_from_string(m[1])], m[2]), text)
    return outside_strings(result, direct)


def prepare(source, output):
    p = ExtendedPackage(source)
    w = openpyxl.load_workbook(source)
    raw = w['RAW']
    assert raw.tables[TABLE].ref == 'A1:AO4001', 'Run once on the pre-sort baseline only'
    headers = {c: raw.cell(1,c).value for c in range(1,42)}
    headers[42] = ID
    changes = Counter()
    for sheet in w:
        for row in sheet:
            for c in row:
                value = getattr(c.value, 'text', c.value)
                if not isinstance(value,str) or not value.startswith('='): continue
                if sheet.title == 'RAW':
                    if c.column > 41: continue  # physical blank-row navigation is intentional
                    if c.column == 17: new = duplicate_formula()
                    elif c.column == 41: new = memo_sequence()
                    else: new = rewrite_raw(value, c.row, headers)
                else:
                    if 'RAW!' not in value and "'RAW'!" not in value and TABLE not in value: continue
                    new = rewrite_external(value, headers)
                if new != value.lstrip('='):
                    old_formula = p.cell(sheet.title,c.coordinate).find(Q('f'))
                    attributes = dict(old_formula.attrib) if old_formula is not None else {}
                    p.write(sheet.title,c.coordinate,new,True)
                    # Preserve array metadata. Expanded shared formulas become ordinary formulas.
                    if attributes.get('t') == 'array':
                        p.cell(sheet.title,c.coordinate).find(Q('f')).attrib.update(attributes)
                    changes[sheet.title] += 1
    for r in range(2,END+1):
        p.write('RAW',f'AP{r}',r)
        p.cell('RAW',f'AP{r}').set('s',p.cell('RAW',f'N{r}').get('s','0'))
    p.write('RAW','AP1',ID)
    p.cell('RAW','AP1').set('s',p.cell('RAW','N1').get('s','0'))
    tree = p.sheet('RAW')
    protection = tree.find(Q('sheetProtection'))
    assert protection is not None
    tree.remove(protection)  # Excel refuses native sort if ANY body cells remain locked.
    cols = tree.find(Q('cols'))
    hidden = next(c for c in cols if c.get('min') == '15')
    hidden.set('max','41')
    visible = ET.Element(Q('col'),min='42',max='42',width='12',customWidth='1')
    cols.insert(list(cols).index(hidden)+1,visible)
    if tree.find(Q('autoFilter')) is not None:
        tree.find(Q('autoFilter')).set('ref',f'A1:AP{END}')
    validation = tree.find(Q('dataValidations'))
    dv = ET.SubElement(validation,Q('dataValidation'),type='custom',allowBlank='0',
                       showErrorMessage='1',errorStyle='stop',errorTitle='거래ID는 변경하지 마세요',
                       error='거래ID는 연결용 고정값입니다. 정렬할 때 거래와 함께 이동합니다.',sqref=f'AP2:AP{END}')
    ET.SubElement(dv,Q('formula1')).text = (f'AND(ISNUMBER(AP2),AP2=INT(AP2),AP2>=2,AP2<={END},'
                                            f'COUNTIF($AP$2:$AP${END},AP2)=1)')
    validation.set('count',str(len(validation)))
    table_part = next(n for n in p.parts if n.startswith('xl/tables/') and n.endswith('.xml')
                      and ET.fromstring(p.parts[n]).get('displayName') == TABLE)
    table = p.tree(table_part)
    table.set('ref',f'A1:AP{END}')
    table.find(Q('autoFilter')).set('ref',f'A1:AP{END}')
    for old in list(table.find(Q('autoFilter'))):
        if old.tag == Q('sortState'): table.find(Q('autoFilter')).remove(old)
    columns = table.find(Q('tableColumns'))
    for column in columns:
        f = column.find(Q('calculatedColumnFormula'))
        if f is None: continue
        idx = int(column.get('id'))
        if idx == 17: f.text = duplicate_formula()
        elif idx == 41: f.text = memo_sequence()
        else: f.text = rewrite_raw(f.text,2,headers)
    ET.SubElement(columns,Q('tableColumn'),id='42',name=ID)
    columns.set('count','42')
    for addr,value in {
        'AQ1':'RAW 입력 · 월별 필터 / 날짜 정렬',
        'AR12':'필터를 전체로 해제하세요. 정렬 후에도 빈 행에 새 거래를 추가합니다.',
        'AR14':'4,000건 준비 · 거래ID/수식 유지, 행삭제·다른 거래 덮어쓰기 금지',
        'AQ16':'날짜 정렬',
        'AR16':'A열 ▼ → 오래된 항목순 / 최신 항목순. 거래ID가 연결을 유지합니다.',
        'AQ18':'거래ID (AP열)',
        'AR18':'수정 금지. 예비비 F·잔액 W/X 연결은 행번호 대신 이 ID를 사용합니다.',
    }.items():
        p.write('RAW',addr,value)
        if addr.endswith('18'):
            p.cell('RAW',addr).set('s',p.cell('RAW',addr[:2]+'17').get('s','0'))
    merge = tree.find(Q('mergeCells'))
    ET.SubElement(merge,Q('mergeCell'),ref='AR18:AV18');merge.set('count',str(len(merge)))
    p.write('예비비','F99','RAW 거래ID')
    p.write('예비비','A75','새 사용은 아래 사용등록에 RAW 거래ID(AP열)·예비비 항목을 지정합니다. 정산은 사용ID·원장행·금액을 연결합니다.')
    p.write('잔액','W63','RAW 기본 거래ID')
    p.write('잔액','X63','RAW 연결 거래ID')
    book = p.tree('xl/workbook.xml')
    calc = book.find(Q('calcPr'))
    if calc is not None:
        calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1');calc.set('calcMode','auto')
    p.save(output);w.close()
    return {'formula_cells_rewritten':dict(changes),'table_ref':f'A1:AP{END}',
            'stable_ids':4000,'raw_sheet_protection_removed_for_native_sort':True,
            'id_validation_is_guardrail_not_paste_protection':True}


def reorder(p, baseline_values, mode):
    """Disposable full-table permutation; never move columns beyond the table."""
    from datetime import datetime
    from openpyxl.utils.datetime import to_excel
    data_end = range_boundaries(baseline_values['RAW'].tables[TABLE].ref)[2]
    ids = {r: int(baseline_values['RAW'].cell(r,42).value) for r in range(2,END+1)}
    def key(r):
        value = baseline_values['RAW'].cell(r,1).value
        if mode == 'id': return ids[r],
        if mode == 'reverse_ids': return -ids[r],
        if value in (None,''): return 1,0,ids[r]
        value = to_excel(value) if isinstance(value,datetime) else float(value)
        return 0,value if mode == 'ascending' else -value,ids[r]
    order = sorted(range(2,END+1),key=key)
    data = p.sheet('RAW').find(Q('sheetData'))
    rows = {int(r.get('r')):r for r in data}
    stored = {}
    for r in range(2,END+1):
        stored[r] = [deepcopy(c) for c in rows[r] if column_index_from_string(re.match('[A-Z]+',c.get('r'))[0]) <= data_end]
        for c in list(rows[r]):
            if column_index_from_string(re.match('[A-Z]+',c.get('r'))[0]) <= data_end: rows[r].remove(c)
    for target, source in enumerate(order,2):
        for i,c in enumerate(stored[source]):
            c.set('r',re.sub(r'\d+$',str(target),c.get('r')))
            rows[target].insert(i,c)
    p.cell_maps.pop('RAW',None);p.row_maps.pop('RAW',None)
    return {'mode':mode,'moved_rows':sum(r!=source for r,source in enumerate(order,2)),
            'first_id':ids[order[0]],'last_nonblank_id':next(ids[r] for r in reversed(order) if baseline_values['RAW'].cell(r,1).value)}


if __name__ == '__main__':
    parser = argparse.ArgumentParser();parser.add_argument('source',type=Path);parser.add_argument('output',type=Path)
    args = parser.parse_args();print(json.dumps(prepare(args.source,args.output),ensure_ascii=False,indent=2))
