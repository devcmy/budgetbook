"""Independent regression tests for the focused OOXML structure checker.

Every test creates a minimal worksheet ZIP fixture in its own temporary
directory; no real budget workbook or hard-coded snapshot path is needed.
Run with: python -m unittest discover -s excelWork/scripts \
    -p 'test_validate_excel_structure.py' -v
"""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

if __package__:
    from .validate_excel_structure import NS, validate_package
else:
    from validate_excel_structure import NS, validate_package


class ValidateExcelStructureTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory(prefix="budgetbook-ooxml-test-")
        self.addCleanup(self.temporary.cleanup)
        self.workbook = Path(self.temporary.name) / "fixture.xlsx"

    def check_fixture(self, body, expected_message=None):
        # This deliberately is a minimal ZIP, not a complete Office package:
        # the checker documents that full OPC/ECMA conformance is out of scope.
        xml = f'<worksheet xmlns="{NS}">{body}</worksheet>'
        with ZipFile(self.workbook, "w") as archive:
            archive.writestr("xl/worksheets/sheet1.xml", xml)
        report = validate_package(self.workbook)
        self.assertEqual(set(report), {"worksheet_count", "xml_part_count", "errors"})
        self.assertEqual(report["worksheet_count"], 1)
        self.assertEqual(report["xml_part_count"], 1)
        if expected_message is None:
            self.assertEqual(report["errors"], [])
        else:
            self.assertEqual(len(report["errors"]), 1, report)
            error = report["errors"][0]
            self.assertEqual(error["part"], "xl/worksheets/sheet1.xml")
            self.assertIn(expected_message, error["message"])
        return report

    def test_valid_sheetpr_order_and_attributes(self):
        self.check_fixture(
            '<sheetPr><tabColor rgb="FF276C72"/>'
            '<outlinePr summaryBelow="1"/>'
            '<pageSetUpPr fitToPage="1"/></sheetPr><sheetData/>'
        )

    def test_sheetpr_wrong_child_order_is_rejected(self):
        self.check_fixture(
            '<sheetPr><pageSetUpPr/><outlinePr/></sheetPr><sheetData/>',
            "sheetPr focused XSD",
        )

    def test_sheetpr_invalid_boolean_is_rejected(self):
        self.check_fixture(
            '<sheetPr><outlinePr summaryBelow="yes"/></sheetPr><sheetData/>',
            "sheetPr focused XSD",
        )

    def test_sheetpr_unknown_unqualified_attribute_is_rejected(self):
        self.check_fixture(
            '<sheetPr unexpected="1"/><sheetData/>',
            "sheetPr focused XSD",
        )

    def test_sheetpr_future_namespace_attribute_is_allowed(self):
        self.check_fixture(
            '<sheetPr xmlns:f="urn:future" f:extra="accepted"/><sheetData/>'
        )

    def test_sheetpr_invalid_argb_length_is_rejected(self):
        self.check_fixture(
            '<sheetPr><tabColor rgb="112233"/></sheetPr><sheetData/>',
            "sheetPr focused XSD",
        )

    def test_duplicate_row_is_rejected(self):
        self.check_fixture(
            '<sheetData><row r="1"/><row r="1"/></sheetData>',
            "duplicate row 1",
        )

    def test_duplicate_cell_is_rejected(self):
        self.check_fixture(
            '<sheetData><row r="1"><c r="A1"/><c r="A1"/></row></sheetData>',
            "duplicate cell column 1",
        )

    def test_cell_in_wrong_parent_row_is_rejected(self):
        self.check_fixture(
            '<sheetData><row r="1"><c r="A2"/></row></sheetData>',
            "belongs to row 2, not parent row 1",
        )

    def test_overlapping_column_intervals_are_rejected(self):
        self.check_fixture(
            '<cols><col min="1" max="2"/><col min="2" max="3"/></cols>'
            '<sheetData/>',
            "overlaps or precedes an earlier interval",
        )

    def test_overlapping_merged_ranges_are_rejected(self):
        self.check_fixture(
            '<sheetData/><mergeCells count="2">'
            '<mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>',
            "overlaps",
        )

    def test_wrong_validation_count_is_rejected(self):
        self.check_fixture(
            '<sheetData/><dataValidations count="2">'
            '<dataValidation sqref="A1"><formula1>1</formula1>'
            '</dataValidation></dataValidations>',
            "dataValidations: count='2' does not match 1 children",
        )

    def test_wrong_tableparts_count_is_rejected(self):
        self.check_fixture(
            '<sheetData/><tableParts count="1"/>',
            "tableParts: count='1' does not match 0 children",
        )

    def test_worksheet_root_child_order_is_enforced(self):
        self.check_fixture(
            '<sheetData/><dimension ref="A1"/>',
            "worksheet: 'dimension' is out of schema sequence",
        )


if __name__ == "__main__":
    unittest.main()
