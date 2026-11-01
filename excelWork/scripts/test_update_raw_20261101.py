from datetime import datetime
import unittest

from update_raw_20261101 import asset,explicit_class,key


class OctoberAppendTests(unittest.TestCase):
    def item(self,**changes):
        values={'asset':'우리카드의정석SUPER','category':'식비','sub':'식대','title':'점심','kind':'지출','memo':None}
        values.update(changes);return values
    def test_joint_account_alias(self):self.assertEqual(asset('20260606'),'모임통장')
    def test_travel_account_is_shared_income(self):self.assertEqual(explicit_class(self.item(asset='여행통장',kind='수입')),'공금')
    def test_meeting_income_not_salary(self):self.assertEqual(explicit_class(self.item(asset='모임통장',kind='수입')),'공금')
    def test_meeting_category(self):self.assertEqual(explicit_class(self.item(category='여행/약속/공사',sub='약속')),'약속·만남')
    def test_gift_category(self):self.assertEqual(explicit_class(self.item(category='경조사/회비/선물/대인')),'경조·선물')
    def test_electronic_category(self):self.assertEqual(explicit_class(self.item(category='가전/전자/잡화')),'일회성전자')
    def test_installment_not_gift(self):self.assertIsNone(explicit_class(self.item(category='경조사/회비/선물/대인',title='롯데카드 띵샵(1/36)')))
    def test_memo_does_not_classify_advance(self):self.assertEqual(explicit_class(self.item(memo='대납·대리결제')),explicit_class(self.item()))
    def test_memo_not_identity(self):
        a={'date':datetime(2026,10,31),'asset':'국민은행','amount':1000,'kind':'수입','memo':'처음'}
        b={**a,'memo':'변경'};self.assertEqual(key(a),key(b))
    def test_salary_keeps_default_income_rules(self):self.assertIsNone(explicit_class(self.item(asset='국민은행',category='월급',kind='수입',title='월급')))


if __name__=='__main__':unittest.main()
