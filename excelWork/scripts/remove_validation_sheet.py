"""Remove the user-facing validation tab without round-tripping an XLSM.

API: ``remove_validation_sheet(package) -> report`` mutates an ExtendedPackage
in memory; the caller saves it. The CLI accepts source/output/report paths.
Remaining sheet IDs, worksheet part names, controls and VBA bytes are retained.
Functional formula references to the removed sheet cause a fail-closed error;
navigation links are redirected to the reserve sheet instead.
"""
from collections import defaultdict
import argparse
import json
from pathlib import Path
import posixpath
import re

from lxml import etree as ET

from redesign_reserve import ExtendedPackage
from update_20261015 import NS, PNS, RNS, Q

APP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
FORMULA_TAGS = {'f', 'formula', 'formula1', 'formula2', 'definedName'}
SHEET_TOKEN = re.compile(r"(?:'((?:[^']|'')+)'|([^\s\"',+*/=()<>!]+))!")


def _read_tree(p, name):
    return p.trees[name] if name in p.trees else ET.fromstring(
        p.parts[name], ET.XMLParser(resolve_entities=False))


def _source_of_rels(name):
    if name == '_rels/.rels':
        return ''
    parent, filename = posixpath.split(name)
    return posixpath.join(posixpath.dirname(parent), filename[:-5])


def _rels_for(source):
    return posixpath.join(posixpath.dirname(source), '_rels',
                          posixpath.basename(source) + '.rels')


def _target(source, relationship):
    target = relationship.get('Target', '').split('#', 1)[0]
    if relationship.get('TargetMode') == 'External' or not target:
        return None
    return (posixpath.normpath(target.lstrip('/')) if target.startswith('/')
            else posixpath.normpath(posixpath.join(posixpath.dirname(source), target)))


def _references_sheet(text, sheet_name, names):
    """Recognize normal refs and 3-D spans, including quoted Korean names."""
    for match in SHEET_TOKEN.finditer(text or ''):
        token = (match.group(1) or match.group(2)).replace("''", "'").lstrip('#')
        if '[' in token:  # Do not mistake an external workbook for this workbook.
            continue
        if token == sheet_name:
            return True
        if ':' in token:
            start, end = token.split(':', 1)
            start, end = start.strip("'"), end.strip("'")
            if start in names and end in names:
                lo, hi = sorted((names.index(start), names.index(end)))
                if lo <= names.index(sheet_name) <= hi:
                    return True
    return False


def inspect_dependencies(p, sheet_name='검증'):
    """Return formula dependencies and navigation without changing package bytes."""
    book = _read_tree(p, 'xl/workbook.xml')
    sheets = list(book.find(Q('sheets')))
    names = [s.get('name') for s in sheets]
    if sheet_name not in names:
        raise ValueError(f'Worksheet is absent: {sheet_name}')
    removed_index = names.index(sheet_name)
    removed_part = p.sheets[sheet_name]
    dependencies, navigation = [], []
    for part in p.parts:
        if part == removed_part or not part.endswith(('.xml', '.rels')):
            continue
        tree = _read_tree(p, part)
        for node in tree.iter():
            local = ET.QName(node).localname
            if local in FORMULA_TAGS and _references_sheet(node.text, sheet_name, names):
                if local == 'definedName' and node.get('localSheetId') == str(removed_index):
                    continue
                if local != 'definedName' and (node.text or '').lstrip('= ').upper().startswith('HYPERLINK('):
                    navigation.append({'part': part, 'kind': 'formula', 'text': node.text})
                else:
                    dependencies.append({'part': part, 'tag': local, 'text': node.text})
            if local == 'hyperlink' and _references_sheet(node.get('location'), sheet_name, names):
                navigation.append({'part': part, 'kind': 'location', 'text': node.get('location')})
            if local == 'Relationship' and node.get('Type', '').endswith('/hyperlink'):
                if _references_sheet(node.get('Target'), sheet_name, names):
                    navigation.append({'part': part, 'kind': 'target', 'text': node.get('Target')})
    return {'sheet': sheet_name, 'sheet_index': removed_index,
            'worksheet_part': removed_part, 'functional_dependencies': dependencies,
            'navigation': navigation}


def _owned_parts(p, root):
    """Only delete reachable ancillary parts with no other retained owner."""
    children, parents = defaultdict(set), defaultdict(set)
    for part in p.parts:
        if not part.endswith('.rels'):
            continue
        source = _source_of_rels(part)
        for rel in _read_tree(p, part):
            target = _target(source, rel)
            if target is not None:
                children[source].add(target)
                parents[target].add(source)
    candidates, pending = set(), [root]
    while pending:
        part = pending.pop()
        if part in candidates or part not in p.parts:
            continue
        candidates.add(part)
        pending.extend(children[part])
    # The workbook itself is the intentional owner being detached from root.
    retained = {part for part in candidates - {root}
                if parents[part] - candidates}
    pending = list(retained)
    while pending:
        part = pending.pop()
        for child in children[part] & candidates:
            if child not in retained and child != root:
                retained.add(child)
                pending.append(child)
    removed = candidates - retained
    removed.update(_rels_for(part) for part in list(removed)
                   if _rels_for(part) in p.parts)
    return removed, retained


def _rewrite_ref(text, old, new):
    # Navigation to a removed sheet goes to a stable replacement start cell.
    old_quoted = "'" + old.replace("'", "''") + "'"
    new_quoted = "'" + new.replace("'", "''") + "'"
    return re.sub(r"(?:" + re.escape(old_quoted) + '|' + re.escape(old) + r')!',
                  new_quoted + '!', text)


def _update_app_properties(p, sheet_name, removed_index, old_count):
    if 'docProps/app.xml' not in p.parts:
        return
    app = p.tree('docProps/app.xml')
    ns = {'a': APP_NS, 'v': VT_NS}
    heading = app.find('a:HeadingPairs/v:vector', ns)
    titles = app.find('a:TitlesOfParts/v:vector', ns)
    if heading is None or titles is None:
        return
    # Heading labels are localized, so identify the first worksheet segment by
    # its actual count/titles instead of assuming an English "Worksheets" label.
    if len(heading) < 2 or len(titles) < old_count:
        raise ValueError('Unexpected app.xml HeadingPairs/TitlesOfParts layout')
    sheet_count = heading[1].find('{'+VT_NS+'}i4')
    if sheet_count is None or int(sheet_count.text) != old_count:
        raise ValueError('app.xml worksheet count differs from workbook.xml')
    if titles[removed_index].text != sheet_name:
        raise ValueError('app.xml worksheet title order differs from workbook.xml')
    titles.remove(titles[removed_index])
    sheet_count.text = str(old_count - 1)
    removed_named_titles = 0
    for node in list(titles)[old_count-1:]:
        if (node.text or '').startswith((sheet_name+'!', "'"+sheet_name.replace("'", "''")+"'!")):
            titles.remove(node)
            removed_named_titles += 1
    if removed_named_titles and len(heading) >= 4:
        named_count = heading[3].find('{'+VT_NS+'}i4')
        if named_count is not None:
            named_count.text = str(int(named_count.text) - removed_named_titles)
    titles.set('size', str(len(titles)))


def remove_validation_sheet(p, sheet_name='검증', replacement='예비비'):
    report = inspect_dependencies(p, sheet_name)
    if report['functional_dependencies']:
        raise ValueError('Functional references must be resolved before removing '
                         f'{sheet_name}: {report["functional_dependencies"]}')
    if replacement not in p.sheets or replacement == sheet_name:
        raise ValueError('Replacement navigation worksheet must remain present')
    removed_part, index = report['worksheet_part'], report['sheet_index']
    book = p.tree('xl/workbook.xml')
    sheets = book.find(Q('sheets'))
    old_count = len(sheets)
    if old_count <= 1:
        raise ValueError('Cannot remove the last worksheet')
    before_ids = {s.get('name'): (s.get('sheetId'), s.get('{'+RNS+'}id')) for s in sheets}
    rid = sheets[index].get('{'+RNS+'}id')
    removed_parts, retained_shared = _owned_parts(p, removed_part)
    rels = p.tree('xl/_rels/workbook.xml.rels')
    for rel in list(rels):
        if rel.get('Id') == rid:
            rels.remove(rel)
        # A stale calc chain is metadata, not a calculation source; Excel will
        # rebuild it. Removing it is safer than keeping entries for a dead sheet.
        elif rel.get('Type', '').endswith('/calcChain'):
            target = _target('xl/workbook.xml', rel)
            if target in p.parts:
                removed_parts.add(target)
            rels.remove(rel)
    sheets.remove(sheets[index])
    names = book.find(Q('definedNames'))
    removed_names = []
    if names is not None:
        for node in list(names):
            local_id = node.get('localSheetId')
            if local_id is not None:
                local_id = int(local_id)
                if local_id == index:
                    removed_names.append(dict(node.attrib))
                    names.remove(node)
                elif local_id > index:
                    node.set('localSheetId', str(local_id - 1))
    for view in book.iter(Q('workbookView')):
        for attr in ['activeTab', 'firstSheet']:
            value = view.get(attr)
            if value is not None:
                current = int(value)
                view.set(attr, str(min(current - (current > index), old_count-2)))
    _update_app_properties(p, sheet_name, index, old_count)
    ctypes = p.tree('[Content_Types].xml')
    for node in list(ctypes):
        if node.get('PartName', '').lstrip('/') in removed_parts:
            ctypes.remove(node)
    updates = []
    remaining_names = [s.get('name') for s in sheets] + [sheet_name]
    strings = None
    if 'xl/sharedStrings.xml' in p.parts:
        strings = [''.join(s.itertext()) for s in _read_tree(p, 'xl/sharedStrings.xml')]
    for part in list(p.parts):
        if part in removed_parts or not part.endswith(('.xml', '.rels')):
            continue
        tree = _read_tree(p, part)
        changed = False
        for node in tree.iter():
            local = ET.QName(node).localname
            if local in FORMULA_TAGS and _references_sheet(node.text, sheet_name, remaining_names):
                node.text = _rewrite_ref(node.text, sheet_name, replacement)
                changed = True
                updates.append({'part': part, 'kind': 'formula'})
            attr = 'location' if local == 'hyperlink' else (
                'Target' if local == 'Relationship' and node.get('Type', '').endswith('/hyperlink') else None)
            if attr and _references_sheet(node.get(attr), sheet_name, remaining_names):
                node.set(attr, _rewrite_ref(node.get(attr), sheet_name, replacement))
                changed = True
                updates.append({'part': part, 'kind': attr})
            # Update only live cells with instructions to visit the removed tab.
            # Do not rewrite the shared-string table or historical raw material.
            if local == 'c' and part.startswith('xl/worksheets/'):
                value = node.find(Q('v'))
                text = (strings[int(value.text)] if strings is not None and node.get('t') == 's' and value is not None
                        else ''.join(node.find(Q('is')).itertext()) if node.find(Q('is')) is not None else '')
                if sheet_name + ' 탭' in text:
                    for child in list(node):
                        if child.tag in [Q('f'), Q('v'), Q('is')]:
                            node.remove(child)
                    node.set('t', 'inlineStr')
                    new_text = ET.SubElement(ET.SubElement(node, Q('is')), Q('t'))
                    new_text.text = text.replace(sheet_name+' 탭과 월간 TRUE/FALSE를 확인하세요.',
                                                 '검증은 별도 기록으로 관리하며, 예비비·잔액 탭에서 현황을 확인하세요.')
                    if new_text.text == text:
                        new_text.text = text.replace(sheet_name+' 탭', replacement+' 탭')
                    changed = True
                    updates.append({'part': part, 'kind': 'instruction', 'cell': node.get('r')})
        if changed:
            p.trees[part] = tree
    for part in removed_parts:
        p.parts.pop(part, None)
        p.trees.pop(part, None)
    p.sheets.pop(sheet_name)
    p.cell_maps.pop(sheet_name, None)
    p.row_maps.pop(sheet_name, None)
    after_ids = {s.get('name'): (s.get('sheetId'), s.get('{'+RNS+'}id')) for s in sheets}
    assert all(before_ids[name] == ids for name, ids in after_ids.items())
    report.update({'remaining_worksheet_count': len(sheets),
                   'removed_parts': sorted(removed_parts),
                   'retained_shared_ancillary_parts': sorted(retained_shared),
                   'removed_local_defined_names': removed_names,
                   'navigation_and_instruction_updates': updates,
                   'remaining_sheet_ids_and_relationship_ids_preserved': True,
                   'vba_parts_modified': False})
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    if args.source.resolve() == args.output.resolve():
        raise ValueError('Use a separate output path to preserve the input')
    p = ExtendedPackage(args.source)
    report = remove_validation_sheet(p)
    p.save(args.output)
    if args.report:
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
