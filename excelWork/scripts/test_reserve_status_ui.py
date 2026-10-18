"""Small reference-migration tests, independent of the user's XLSM."""
import unittest

from reserve_status_ui import _move_reference, _move_formula
from reserve_cash_alignment import allocation, settlement_permission


class ReserveReferenceTests(unittest.TestCase):
    def test_only_display_rows_move(self):
        self.assertEqual(_move_reference('$AB36'), '$AB46')
        self.assertEqual(_move_reference('$B$33'), '$B$43')
        self.assertEqual(_move_reference('A32:O69'), 'A42:O79')
        self.assertEqual(_move_reference('$A$100:$Z$399'), '$A$100:$Z$399')

    def test_local_helpers_and_filters_move(self):
        self.assertEqual(_move_formula('IF($AB36=0,"",$B$33)', local=True), 'IF($AB46=0,"",$B$43)')

    def test_other_sheet_ranges_do_not_move(self):
        f = 'INDEX(잔액!$F$64:$F$500,$AB36)'
        self.assertEqual(_move_formula(f, local=True), 'INDEX(잔액!$F$64:$F$500,$AB46)')

    def test_external_explicit_reserve_reference_moves(self):
        self.assertEqual(_move_formula("'예비비'!$B$33+예비비!AC50"), "'예비비'!$B$43+예비비!AC60")

    def test_quoted_display_text_is_not_a_reference(self):
        self.assertEqual(_move_formula('IF(B33="B33","A32",B34)', local=True), 'IF(B43="B33","A32",B44)')

    def test_register_and_raw_references_stay_fixed(self):
        f = 'SUMIFS($O$430:$O$729,$B$430:$B$729,A107)+RAW!F36'
        self.assertEqual(_move_formula(f, local=True), f)

    def test_allocation_has_previous_and_current_month_indices(self):
        self.assertIn('MONTH($B$3)+1', allocation('$A10', '$B$3'))
        self.assertNotIn('MONTH($B$3)+1', allocation('$A10', '$B$3', previous=True))
        self.assertIn('$BI$81:$BU$86', allocation('$A10', '$B$3'))

    def test_bill_source_keeps_narrow_historical_exceptions_and_caps(self):
        f = settlement_permission(440)
        self.assertIn('F440="결제대금비"', f)
        self.assertIn('B440="U008",C440=208', f)
        self.assertIn('B440="U008",C440=223', f)
        self.assertIn('<=39000', f)
        self.assertIn('<=INDEX(잔액!$Y:$Y,MAX(64,C440))', f)


if __name__ == '__main__':
    unittest.main()
