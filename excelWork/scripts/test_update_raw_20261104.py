from datetime import datetime
import unittest

from update_raw_20261104 import classify,correction_key,corrected_burden


class NovemberDeltaTests(unittest.TestCase):
    def item(self,**updates):
        base=dict(date=datetime(2026,11,2,11,3),asset='롯데365카드',category='기타',sub=None,
                  title='기타',amount=5000,kind='지출',memo=None)
        base.update(updates);return base

    def test_amount_correction_identity_excludes_amount_and_memo(self):
        self.assertEqual(correction_key(self.item()),correction_key(self.item(amount=4500,memo='changed')))

    def test_different_time_is_separate_transaction(self):
        self.assertNotEqual(correction_key(self.item()),correction_key(self.item(date=datetime(2026,11,2,11,4))))

    def test_full_personal_burden_tracks_corrected_amount(self):
        self.assertEqual(corrected_burden(5000,5000,4500),4500)
        self.assertIsNone(corrected_burden(5000,None,4500))

    def test_partial_burden_cannot_be_guessed(self):
        with self.assertRaises(AssertionError):corrected_burden(5000,2500,4500)

    def test_phone_installment_not_one_off_electronics(self):
        self.assertIsNone(classify(self.item(category='가전/전자/잡화',title='11번가(15/22)')))

    def test_buds_installment_not_gift_expense(self):
        self.assertIsNone(classify(self.item(category='경조사/회비/선물/대인',title='롯데카드 띵샵(2/36)')))

    def test_real_one_off_electronics_and_memo_independence(self):
        a=self.item(category='가전/전자/잡화',title='네이버플러스 스토어',memo='가습기')
        b=dict(a,memo='할부 혹은 대납이라는 메모')
        self.assertEqual(classify(a),'일회성전자');self.assertEqual(classify(a),classify(b))


if __name__=='__main__':unittest.main()
