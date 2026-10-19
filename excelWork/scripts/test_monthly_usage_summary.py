"""Regression tests for actual/planned expenses and gross advance payments."""
from datetime import datetime
import unittest

from monthly_usage_summary import (
    ACTUAL_STATUSES, ADVANCE_CATEGORIES, LABELS, advance_formula, advance_total,
    fixed_formula, fixed_total, summary_formulas, totals,
)


def record(month=10, category='대납', kind='지출', amount=51800, status='정산대기'):
    return (datetime(2026, month, 12), datetime(2026, month, 1), '카드',
            '여행', None, '철도', amount, kind, None, category, 0, status)


class MonthlyUsageSummaryTests(unittest.TestCase):
    def test_actual_accepts_each_recognized_status(self):
        self.assertEqual(fixed_total([(s, 100) for s in ACTUAL_STATUSES]), 400)

    def test_actual_excludes_pending_and_future_placeholders(self):
        self.assertEqual(fixed_total([('예상', 100), ('R', 200), (None, 300)]), 0)

    def test_plan_basis_explicitly_includes_expected(self):
        self.assertEqual(fixed_total([('확정', 100), ('예상', 200)], 'actual_and_planned'), 300)

    def test_fixed_blank_and_negative_adjustments(self):
        self.assertEqual(fixed_total([('확정', None), ('결제', -100)]), -100)

    def test_fixed_invalid_basis_is_rejected(self):
        with self.assertRaises(ValueError):
            fixed_total([], 'silent_guess')

    def test_advance_uses_original_amount_not_zero_personal_burden(self):
        self.assertEqual(advance_total([record()], datetime(2026, 10, 1)), 51800)

    def test_advance_supports_explicit_classifications(self):
        rows = [record(category=c) for c in ADVANCE_CATEGORIES]
        self.assertEqual(advance_total(rows, datetime(2026, 10, 1)), 207200)

    def test_advance_excludes_other_month(self):
        self.assertEqual(advance_total([record(month=9)], datetime(2026, 10, 1)), 0)

    def test_transfer_topup_is_not_a_new_advance(self):
        self.assertEqual(advance_total([record(category='내부이체', kind='이체지출')], datetime(2026, 10, 1)), 0)

    def test_calendar_consumption_is_not_an_advance(self):
        self.assertEqual(advance_total([record(category='변동비')], datetime(2026, 10, 1)), 0)

    def test_recovery_is_not_an_original_payment(self):
        self.assertEqual(advance_total([record(kind='수입')], datetime(2026, 10, 1)), 0)

    def test_gross_payments_are_not_outstanding_receivables(self):
        self.assertEqual(advance_total([record(status='정산완료')], datetime(2026, 10, 1)), 51800)

    def test_base_and_final_are_exact_requested_components(self):
        self.assertEqual(totals(470022, 1582350, 103600, 110075), (2155972, 2266047))

    def test_additional_already_in_calendar_is_not_added_twice(self):
        self.assertEqual(totals(0, 17000, 0, 17000 - 17000), (17000, 17000))

    def test_formula_uses_gross_column_and_dynamic_month(self):
        expression = advance_formula()
        self.assertEqual(expression.count('SUMIFS('), len(ADVANCE_CATEGORIES))
        self.assertIn('RAW!$G$2:$G$10018', expression)
        self.assertIn('$T$1', expression)
        self.assertNotIn('정산대기', expression)
        self.assertNotIn('RAW!$K$', expression)

    def test_actual_formula_does_not_add_the_forecast_total(self):
        self.assertNotIn('Q25', fixed_formula('actual'))
        self.assertIn('확정', fixed_formula('actual'))
        self.assertEqual(fixed_formula('actual_and_planned'), '=SUM(Q4:Q24)')

    def test_seven_rows_show_both_fixed_cost_bases(self):
        self.assertEqual(len(LABELS), 7)
        self.assertEqual(LABELS[1:3], ('고정비 · 실제 확정', '고정비 · 실적+미결제 예정'))

    def test_two_fixed_rows_have_separate_formulas(self):
        formulas = summary_formulas()
        self.assertEqual(formulas['E74'], fixed_formula('actual'))
        self.assertEqual(formulas['E75'], fixed_formula('actual_and_planned'))

    def test_actual_totals_never_add_inclusive_fixed_row_or_tetris(self):
        formulas = summary_formulas()
        self.assertTrue(formulas['E78'].startswith('=E74+E76+'))
        self.assertNotIn('E75', formulas['E78'])
        self.assertNotIn('E77', formulas['E78'])
        self.assertEqual(formulas['E79'], '=E78+H39')


if __name__ == '__main__':
    unittest.main()
