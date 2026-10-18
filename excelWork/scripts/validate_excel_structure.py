"""Focused sheetPr XSD and structural checks for an Excel OOXML package.

This deliberately is NOT a complete ECMA-376 / ISO 29500 validator and does
not claim that Excel or VBA was executed. It checks XML well-formedness, the
worksheet structures relevant to this workbook, and the strict sheetPr child
order that permissive spreadsheet readers can silently accept or normalize.

Usage: python validate_excel_structure.py workbook.xlsm
The JSON result has worksheet_count, xml_part_count, and errors. Exit 0 means
these focused checks passed; it does not establish complete OOXML conformance.
"""

from collections import Counter
from functools import lru_cache
import argparse
import json
from pathlib import Path
import re
from zipfile import BadZipFile, ZipFile

from lxml import etree as ET


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = lambda name: "{" + NS + "}" + name
MAX_ROW, MAX_COL = 1048576, 16384
CELL_REF = re.compile(r"\$?([A-Z]{1,3})\$?([1-9][0-9]*)\Z")
WORKSHEET_SEQUENCE = (
    "sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr "
    "sheetProtection protectedRanges scenarios autoFilter sortState "
    "dataConsolidate customSheetViews mergeCells phoneticPr "
    "conditionalFormatting dataValidations hyperlinks printOptions pageMargins "
    "pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches "
    "ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture "
    "oleObjects controls webPublishItems tableParts extLst"
).split()
WORKSHEET_RANK = {name: n for n, name in enumerate(WORKSHEET_SEQUENCE)}


@lru_cache(maxsize=1)
def _properties_schema():
    path = Path(__file__).parent / "schemas" / "worksheet-properties.xsd"
    parser = ET.XMLParser(resolve_entities=False, no_network=True)
    return ET.XMLSchema(ET.parse(str(path), parser))


def _coordinate(ref):
    """Return a bounded (column, row); absolute markers are accepted in ranges."""
    match = CELL_REF.fullmatch(ref or "")
    if not match:
        raise ValueError(f"invalid cell reference {ref!r}")
    letters, row = match.groups()
    column = 0
    for char in letters:
        column = column * 26 + ord(char) - ord("A") + 1
    row = int(row)
    if not 1 <= column <= MAX_COL or not 1 <= row <= MAX_ROW:
        raise ValueError(f"out-of-range cell reference {ref!r}")
    return column, row


def _range(ref):
    bounds = (ref or "").split(":")
    if len(bounds) not in (1, 2):
        raise ValueError(f"invalid range {ref!r}")
    first = _coordinate(bounds[0])
    last = _coordinate(bounds[-1])
    if first[0] > last[0] or first[1] > last[1]:
        raise ValueError(f"reversed range {ref!r}")
    return first[0], first[1], last[0], last[1]


def _check_sequence(element, names, error, label, repeatable=()):
    rank = {name: n for n, name in enumerate(names)}
    seen = Counter()
    previous = -1
    for child in element:
        if not isinstance(child.tag, str):
            continue
        qualified = ET.QName(child)
        if qualified.namespace != NS:
            # Extension namespaces are outside this checker's coverage.
            continue
        name = qualified.localname
        if name not in rank:
            error(f"{label}: unsupported SpreadsheetML child {name!r}")
            continue
        seen[name] += 1
        if seen[name] > 1 and name not in repeatable:
            error(f"{label}: duplicate {name!r}")
        if rank[name] < previous:
            error(f"{label}: {name!r} is out of schema sequence")
        previous = rank[name]


def _check_count(element, error):
    count = element.get("count")
    if count is None:
        return
    label = ET.QName(element).localname
    try:
        number = int(count)
        if number < 0 or number != len(element):
            error(f"{label}: count={count!r} does not match {len(element)} children")
    except ValueError:
        error(f"{label}: invalid count {count!r}")


def _check_rows(root, error):
    data = root.find(Q("sheetData"))
    if data is None:
        error("worksheet: missing required sheetData")
        return
    previous_row = 0
    seen_rows = set()
    for row in data:
        if row.tag != Q("row"):
            error("sheetData: child is not a SpreadsheetML row")
            continue
        try:
            number = int(row.get("r", previous_row + 1))
        except ValueError:
            error(f"row: invalid index {row.get('r')!r}")
            continue
        if not 1 <= number <= MAX_ROW:
            error(f"row: index {number} is outside 1..{MAX_ROW}")
        if number in seen_rows:
            error(f"sheetData: duplicate row {number}")
        elif number <= previous_row:
            error(f"sheetData: row {number} is not in ascending order")
        seen_rows.add(number)
        previous_row = number
        previous_col = 0
        seen_cols = set()
        for cell in row:
            if cell.tag == Q("extLst"):
                continue
            if cell.tag != Q("c"):
                error(f"row {number}: child is not a SpreadsheetML cell")
                continue
            address = cell.get("r")
            try:
                column, cell_row = _coordinate(address) if address else (previous_col + 1, number)
            except ValueError as exc:
                error(f"row {number}: {exc}")
                continue
            if column > MAX_COL:
                error(f"row {number}: inferred cell exceeds column {MAX_COL}")
            if cell_row != number:
                error(f"cell {address}: belongs to row {cell_row}, not parent row {number}")
            if column in seen_cols:
                error(f"row {number}: duplicate cell column {column}")
            elif column <= previous_col:
                error(f"row {number}: cell {address!r} is not in ascending column order")
            seen_cols.add(column)
            previous_col = column
            _check_sequence(cell, ("f", "v", "is", "extLst"), error, f"cell {address or number}")
            formula = cell.find(Q("f"))
            if formula is not None and (formula.text or "").startswith("="):
                error(f"cell {address}: OOXML formula must not start with '='")


def _check_columns(root, error):
    previous_max = 0
    for columns in root.findall(Q("cols")):
        for column in columns:
            if column.tag != Q("col"):
                error("cols: child is not a SpreadsheetML col")
                continue
            try:
                first, last = int(column.get("min")), int(column.get("max"))
            except (TypeError, ValueError):
                error(f"col: invalid min/max {column.get('min')!r}/{column.get('max')!r}")
                continue
            if not 1 <= first <= last <= MAX_COL:
                error(f"col: invalid interval {first}..{last}")
            if first <= previous_max:
                error(f"col: interval {first}..{last} overlaps or precedes an earlier interval")
            previous_max = last


def _check_ranges_and_counts(root, error):
    merged = []
    for container in root.findall(Q("mergeCells")):
        _check_count(container, error)
        for child in container:
            if child.tag != Q("mergeCell"):
                error("mergeCells: child is not mergeCell")
                continue
            ref = child.get("ref")
            try:
                left, top, right, bottom = _range(ref)
            except ValueError as exc:
                error(f"mergeCell: {exc}")
                continue
            for old_ref, old_left, old_top, old_right, old_bottom in merged:
                if left <= old_right and old_left <= right and top <= old_bottom and old_top <= bottom:
                    error(f"mergeCells: {ref!r} overlaps {old_ref!r}")
                    break
            merged.append((ref, left, top, right, bottom))
    for container in root.findall(Q("dataValidations")):
        _check_count(container, error)
        for validation in container:
            if validation.tag != Q("dataValidation"):
                error("dataValidations: child is not dataValidation")
                continue
            references = (validation.get("sqref") or "").split()
            if not references:
                error("dataValidation: missing or empty sqref")
            for ref in references:
                try:
                    _range(ref)
                except ValueError as exc:
                    error(f"dataValidation: {exc}")
            _check_sequence(validation, ("formula1", "formula2"), error, "dataValidation")
    for container in root.findall(Q("tableParts")):
        _check_count(container, error)
        for child in container:
            if child.tag != Q("tablePart"):
                error("tableParts: child is not tablePart")


def _check_worksheet(root, error):
    _check_sequence(root, WORKSHEET_SEQUENCE, error, "worksheet", ("cols", "conditionalFormatting"))
    properties = root.find(Q("sheetPr"))
    if properties is not None:
        schema = _properties_schema()
        if not schema.validate(properties):
            messages = "; ".join(item.message for item in schema.error_log)
            error(f"sheetPr focused XSD: {messages}")
    dimension = root.find(Q("dimension"))
    if dimension is not None:
        try:
            _range(dimension.get("ref"))
        except ValueError as exc:
            error(f"dimension: {exc}")
    _check_rows(root, error)
    _check_columns(root, error)
    _check_ranges_and_counts(root, error)


def validate_package(path):
    """Return focused validation evidence without modifying the input package."""
    report = {"worksheet_count": 0, "xml_part_count": 0, "errors": []}

    def add(part, message):
        report["errors"].append({"part": part, "message": message})

    try:
        with ZipFile(path) as archive:
            duplicates = [name for name, count in Counter(archive.namelist()).items() if count > 1]
            for name in duplicates:
                add(name, "duplicate ZIP member name")
            damaged = archive.testzip()
            if damaged is not None:
                add(damaged, "ZIP CRC check failed")
            for name in archive.namelist():
                if not name.endswith((".xml", ".rels", ".vml")):
                    continue
                report["xml_part_count"] += 1
                try:
                    parser = ET.XMLParser(resolve_entities=False, no_network=True)
                    root = ET.fromstring(archive.read(name), parser)
                except ET.XMLSyntaxError as exc:
                    add(name, f"XML is not well-formed: {exc}")
                    continue
                if root.tag == Q("worksheet"):
                    report["worksheet_count"] += 1
                    _check_worksheet(root, lambda message, part=name: add(part, message))
    except (OSError, BadZipFile, RuntimeError) as exc:
        add("[package]", str(exc))
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    result = validate_package(args.workbook)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(bool(result["errors"]))
