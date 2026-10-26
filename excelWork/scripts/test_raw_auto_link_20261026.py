"""Unit coverage for the all-month RAW engine (integration tests are separate)."""
from datetime import datetime
import unittest

import openpyxl

from raw_auto_link_20261026 import (BACK,CAPACITY,HEADERS,PREPARED_END,RR,ceil_units,
    default_class,fixed_slot,palette_for,row_formulas,slot_formula)


class RawAutoLinkTests(unittest.TestCase):
    def setUp(self):
        self.sheet=openpyxl.Workbook().active
        for r,label in {13:'신한은행',14:'Uber One',15:'ChatGPT',16:'로켓와우',18:'T우주+YouTube',23:'에어팟프로3',24:'애플케어플러스'}.items():self.sheet[f'L{r}']=label

    def test_round_positive(self):self.assertEqual(ceil_units(3101),4)
    def test_round_negative(self):self.assertEqual(ceil_units(-1300),-1)
    def test_zero(self):self.assertEqual(ceil_units(0),0)
    def test_whole_thousand(self):self.assertEqual(ceil_units(11000),11)
    def test_fixed_insurance(self):self.assertEqual(fixed_slot('롯데손해보험',self.sheet),7)
    def test_fixed_apple(self):self.assertEqual(fixed_slot('Apple Appstore',self.sheet),19)
    def test_fixed_chatgpt(self):self.assertEqual(fixed_slot('ChatGPT 구독',self.sheet),15)
    def test_fixed_uber(self):self.assertEqual(fixed_slot('우버택시 멤버쉽',self.sheet),14)
    def test_installments_with_new_number(self):self.assertEqual(fixed_slot('11번가 (15/22)',self.sheet),22)
    def test_new_fixed_is_not_lost(self):self.assertEqual(fixed_slot('새 정기비',self.sheet),26)
    def test_food_palette(self):self.assertEqual(palette_for('변동비','식비',None),3)
    def test_housing_not_medical(self):self.assertEqual(palette_for('이사부대','주거',None),4)
    def test_meeting_palette(self):self.assertEqual(palette_for('약속·만남','여행/약속/공사','만남'),8)
    def test_smoking_and_travel_palette(self):self.assertEqual(palette_for('담배','기타',None),5)
    def test_health_palette(self):self.assertEqual(palette_for('변동비','건강/보험/의료',None),7)
    def test_fashion_palette(self):self.assertEqual(palette_for('변동비','패션/미용',None),6)
    def test_gifts_palette(self):self.assertEqual(palette_for('경조·선물','기타',None),10)
    def test_table_headers_unique(self):self.assertEqual(len(HEADERS),len(set(HEADERS)))
    def test_filter_month_second_column(self):self.assertEqual(HEADERS[1],'월')
    def test_prepared_input_has_headroom(self):self.assertGreater(PREPARED_END,2493+1000)
    def test_visual_capacity_above_original(self):self.assertGreater(CAPACITY,2600)
    def test_structured_ranges_not_old_fixed_end(self):self.assertEqual(RR('Y'),'RawData2026[달력금액]')
    def test_base_index_first_cell(self):
        self.assertIn('ROW()-15',slot_formula());self.assertIn('COLUMN()-2',slot_formula())
    def test_overflow_index_offset(self):self.assertTrue(slot_formula(165,2600).startswith('(2600+'))
    def test_default_class_handles_transfer(self):self.assertIn('LEFT(H2494,2)="이체"',default_class(2494,30))
    def test_default_class_keeps_investment_out_of_living(self):self.assertIn('D2494="투자","투자"',default_class(2494,30))
    def test_default_class_handles_advance(self):self.assertIn('대리결제',default_class(2494,30))
    def test_month_comes_from_date(self):self.assertIn('DATE(YEAR(O2494),MONTH(O2494),1)',row_formulas(2494,30,300,set(),2493)['B'])
    def test_missing_date_does_not_break_blank_rows(self):self.assertIn('IF(A2494="",0,',row_formulas(2494,30,300,set(),2493)['Y'])
    def test_explicit_classification_overrides_default(self):self.assertIn('IF(J2494="",AJ2494,J2494)',row_formulas(2494,30,300,set(),2493)['T'])
    def test_explicit_zero_personal_burden_is_respected(self):self.assertIn('IF(K2494="",AL2494,K2494)',row_formulas(2494,30,300,set(),2493)['W'])
    def test_dedupe_key_includes_memo(self):self.assertIn('I2494',row_formulas(2494,30,300,set(),2493)['P'])
    def test_long_duplicate_criteria_has_exact_fallback(self):self.assertIn('SUMPRODUCT(',row_formulas(2494,30,300,set(),2493)['Q'])
    def test_invalid_input_has_visible_state(self):self.assertIn('날짜 확인',row_formulas(2494,30,300,set(),2493)['N'])
    def test_new_bank_rows_do_not_pretend_ledger_is_updated(self):self.assertIn('잔액원장 별도',row_formulas(2494,30,300,set(),2493)['N'])
    def test_advance_uses_original_not_zero_burden(self):self.assertIn('V2494,0',row_formulas(2494,30,300,set(),2493)['AG'])
    def test_literal_wildcards_are_escaped_for_duplicate_checks(self):
        f=row_formulas(2494,30,300,set(),2493)['Q'];self.assertIn('"~*"',f);self.assertIn('"~?"',f)
    def test_duplicate_fields_have_unambiguous_lengths(self):self.assertIn('LEN(F2494)&":"&F2494',row_formulas(2494,30,300,set(),2493)['P'])
    def test_investment_account_transfer_is_separate_from_living(self):self.assertIn('"증권",D2494',default_class(2494,30))
    def test_contributions_do_not_count_incoming_transfers(self):self.assertNotIn('H2494<>"수입"',row_formulas(2494,30,300,set(),2493)['AI'])


if __name__=='__main__':unittest.main()
