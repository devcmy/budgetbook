import unittest

from memo_only_20261026 import (classification,duplicate_formula,fixed_note,
                               joined_rows,key_formula,memo_formula,memo_index_formula)


class MemoOnlyTests(unittest.TestCase):
    def test_classification_does_not_read_memo(self):
        f=classification(2494,28)
        self.assertNotIn('I2494',f);self.assertNotIn('SEARCH("대납"',f)
    def test_business_fixed_rules_stay(self):self.assertIn('"고정비"',classification(2494,28))
    def test_business_transfer_rules_stay(self):self.assertIn('"자산형성-ISA"',classification(2494,28))
    def test_key_not_memo(self):self.assertNotIn('I2494',key_formula(2494))
    def test_key_has_explicit_discriminator(self):self.assertIn('M2494',key_formula(2494))
    def test_long_duplicate_fallback_not_memo(self):self.assertNotIn('I2494',duplicate_formula(2494))
    def test_long_duplicate_fallback_has_discriminator(self):self.assertIn('$M$2:M2494',duplicate_formula(2494))
    def test_memo_source_is_single_column(self):
        f=memo_formula(2494);self.assertIn('I2494',f);self.assertNotIn('M2494',f)
    def test_empty_record_hides_memo(self):self.assertIn('A2494=""',memo_formula(2494))
    def test_index_for_display_only(self):self.assertIn('$AN$2:AN2494',memo_index_formula(2494))
    def test_fixed_notes_are_raw_display(self):self.assertIn('[연동메모]',fixed_note(13))
    def test_monthly_joins_do_not_duplicate_sources(self):self.assertEqual(joined_rows([4,4]).count('RAW!AN4'),1)
    def test_empty_join_is_empty(self):self.assertEqual(joined_rows([]),'=""')


if __name__=='__main__':unittest.main()
