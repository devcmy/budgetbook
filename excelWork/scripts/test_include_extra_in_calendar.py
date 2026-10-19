"""Once-only calendar inclusion and per-transaction rounding regressions."""
from datetime import datetime
import unittest

import openpyxl

from include_extra_in_calendar_20261019 import daily_formula, extra_delta, grid, rounded, usage_totals, variance_formula


class IncludeExtraInCalendarTests(unittest.TestCase):
    def test_positive_amount_rounds_up(self):
        self.assertEqual(rounded(110075), 111000)

    def test_negative_adjustment_rounds_toward_positive_infinity(self):
        self.assertEqual(rounded(-1300), -1000)

    def test_exact_thousand_does_not_grow(self):
        self.assertEqual(rounded(150000), 150000)

    def test_blank_amount_is_zero(self):
        self.assertEqual(extra_delta(None, None), (0, 0))

    def test_uncovered_premium_is_added(self):
        self.assertEqual(extra_delta(110075, 0), (110075, 111000))

    def test_already_covered_dutch_purchase_is_not_added_twice(self):
        self.assertEqual(extra_delta(17000, 17000), (0, 0))

    def test_refund_of_covered_purchase_adjusts_once(self):
        self.assertEqual(extra_delta(16000, 17000), (-1000, -1000))

    def test_rounded_difference_is_per_transaction_not_per_day(self):
        self.assertEqual(extra_delta(1100, 1000), (100, 1000))
        self.assertEqual(sum(extra_delta(n, 0)[1] for n in [4200, 1700]), 7000)

    def test_usage_totals_never_add_extra_again(self):
        self.assertEqual(usage_totals(470022, 1692425, 103600), (2266047, 2266047))

    def test_september_extra_moves_into_base_not_larger_final(self):
        self.assertEqual(usage_totals(1244246, 2073851, 0), (3318097, 3318097))

    def test_blank_day_stays_blank_until_an_expense_is_entered(self):
        self.assertTrue(daily_formula(None, 'B18', 'H').startswith('=IF(COUNTIFS('))
        self.assertIn('=0,"",', daily_formula(None, 'B18', 'H'))

    def test_existing_formula_is_preserved_as_base(self):
        self.assertTrue(daily_formula('=SUM(A1:A3)', 'B18', 'H').startswith('=(SUM(A1:A3))+SUMIFS('))

    def test_dates_use_whole_day_bounds_for_timed_transactions(self):
        formula = daily_formula(1000, 'B18', 'AB')
        self.assertIn('">="&B18', formula)
        self.assertIn('"<"&B18+1', formula)
        self.assertIn('$AB$52:$AB$61', formula)

    def test_grid_excludes_adjacent_month_dates(self):
        sheet = openpyxl.Workbook().active
        sheet['B3'] = datetime(2026, 9, 30)
        sheet['C3'] = datetime(2026, 10, 1)
        self.assertEqual([day for day, *_ in grid(sheet, 10)], [datetime(2026, 10, 1)])

    def test_variance_ignores_formula_empty_strings(self):
        formula = variance_formula('H4', 'H5')
        self.assertIn('COUNT(H4,H5)', formula)
        self.assertNotIn('COUNTA', formula)
        self.assertIn('SUM(H4)-SUM(H5)', formula)


if __name__ == '__main__':
    unittest.main()
