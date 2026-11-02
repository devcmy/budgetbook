"""Business predicates and formula boundary tests for charge-month transport."""
import unittest
from update_postpaid_20261102 import program_card,usage,settlement,classification_formula,slot_formula,fixed_amount_formula,DEFERRED


def record(**kw):
    return {'asset':'토스뱅크','category':'토스뱅크K패스체크카드','title':'카드정산 결제','kind':'이체지출',**kw}


class PostpaidTests(unittest.TestCase):
    def test_program_aliases(self):
        for value in ['토스뱅크K패스체크카드','신한K-패스체크카드','토스뱅크모두의카드','모두의 카드','k패스']:
            self.assertTrue(program_card(value),value)

    def test_other_cards_do_not_become_fixed_settlements(self):
        self.assertFalse(settlement(record(category='현대카드')))
        self.assertFalse(settlement(record(category='롯데365카드')))

    def test_outbound_charge_recognized(self):
        for kind in ['이체지출','이체출금','지출']:self.assertTrue(settlement(record(kind=kind)))

    def test_incoming_counterpart_not_counted(self):
        for kind in ['이체입금','이체수입','수입']:self.assertFalse(settlement(record(kind=kind)))

    def test_usage_is_evidence_not_second_charge(self):
        item=record(asset='토스뱅크K패스체크카드',category='교통/차량',title='9월 후불교통',kind='지출')
        self.assertTrue(usage(item));self.assertFalse(settlement(item))

    def test_ordinary_card_purchase_not_transport_evidence(self):
        self.assertFalse(usage(record(asset='토스뱅크K패스체크카드',title='식당',kind='지출')))

    def test_refunds_remain_income(self):
        self.assertFalse(usage(record(asset='토스뱅크K패스체크카드',title='후불교통',kind='수입')))
        self.assertFalse(settlement(record(title='K패스 환급금',kind='수입')))

    def test_memo_words_never_classify(self):
        for memo in ['대납','후불교통','모두의카드','날짜 미기록','']:
            item=record(category='현대카드',memo=memo)
            self.assertFalse(settlement(item))

    def test_formulas_have_no_memo_dependency(self):
        for formula in [classification_formula('IF(A2="","","변동비")'),slot_formula('0'),fixed_amount_formula()]:
            self.assertNotIn('[메모]',formula)
            self.assertNotIn('[연동메모]',formula)
            self.assertNotIn('ROW()',formula)

    def test_duplicate_gate_is_required_for_fixed_amount(self):
        self.assertIn('[동일거래순번]]=1',fixed_amount_formula())

    def test_transit_slot_is_explicit(self):
        self.assertIn(',11,',slot_formula('0'))
        self.assertIn(DEFERRED,classification_formula('0'))


if __name__=='__main__':unittest.main()
