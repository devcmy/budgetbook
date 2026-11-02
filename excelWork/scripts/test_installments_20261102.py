import unittest
from installments_20261102 import plan_rows,alias_key,future_slot_formula


class InstallmentPlanTests(unittest.TestCase):
    def test_buds_rolls_forward(self):
        plans=[x for x in plan_rows() if x['name']=='갤럭시버즈4프로']
        self.assertEqual([(p['month'],p['installment'],p['total'],p['amount']) for p in plans],[(11,2,36,11004),(12,3,36,11004)])

    def test_finished_applecare_is_not_a_plan(self):
        self.assertFalse(any('애플케어' in p['name'] for p in plan_rows()))

    def test_active_phone_airpods_kept(self):
        self.assertEqual([(p['installment'],p['total']) for p in plan_rows() if p['name']=='휴대폰'],[(15,22),(16,22)])
        self.assertEqual([(p['installment'],p['total']) for p in plan_rows() if p['name']=='에어팟프로3'],[(7,36),(8,36)])

    def test_same_vendor_products_do_not_collide(self):
        keys=[alias_key(p['month'],p['asset'],p['title']) for p in plan_rows()]
        self.assertEqual(len(keys),len(set(keys)))
        self.assertEqual(len(plan_rows()),6)

    def test_spaces_in_titles_do_not_matter(self):
        self.assertEqual(alias_key(11,'롯데 365카드','롯데카드 띵샵 (2/36)'),alias_key(11,'롯데365카드','롯데카드띵샵(2/36)'))

    def test_formula_does_not_classify_by_memo_or_physical_row(self):
        f=future_slot_formula('0')
        self.assertNotIn('[메모]',f);self.assertNotIn('[연동메모]',f);self.assertNotIn('ROW()',f)

    def test_unknown_vendor_series_does_not_overwrite_airpods(self):
        self.assertIn('"띵샵"',future_slot_formula('0'))
        self.assertIn('),26,',future_slot_formula('0'))

    def test_historical_override_is_retained_outside_future_guard(self):
        self.assertTrue(future_slot_formula('LEGACY_VALUE').endswith(',LEGACY_VALUE)'))


if __name__=='__main__':unittest.main()
