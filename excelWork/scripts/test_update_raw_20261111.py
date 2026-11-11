from datetime import datetime
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import patch

from update_raw_20261111 import read_source, identity


class NovemberExportTests(TestCase):
    def read(self, tag=False, tag_value='', date=None):
        headers=['날짜','자산','분류','소분류','내용','금액(원)','수입/지출','메모','금액','화폐']
        values=[date or datetime(2026,11,1),'20260606','식비','','식사',1000,'지출','원본\n메모',1000,'KRW']
        if tag:
            headers.insert(7,'태그');values.insert(7,tag_value)
        workbook=SimpleNamespace(active=SimpleNamespace(values=[headers,values]),close=lambda:None)
        with patch('update_raw_20261111.openpyxl.load_workbook',return_value=workbook):
            return read_source('fixture.xlsx')[0]

    def test_original_ten_columns(self):
        self.assertEqual(self.read()['memo'],'원본\n메모')

    def test_new_tag_column_does_not_shift_memo(self):
        self.assertEqual(self.read(tag=True)['memo'],'원본\n메모')

    def test_populated_tag_is_not_silently_lost(self):
        with self.assertRaises(AssertionError):self.read(tag=True,tag_value='보존 필요')

    def test_non_november_change_stops(self):
        with self.assertRaises(AssertionError):self.read(date=datetime(2026,10,1))

    def test_shared_asset_normalization(self):
        self.assertEqual(self.read()['asset'],'모임통장')

    def test_identity_ignores_memo_but_distinguishes_same_time_merchants(self):
        item=self.read()
        self.assertEqual(identity(item),identity(dict(item,memo='수정')))
        self.assertNotEqual(identity(item),identity(dict(item,title='다른 상점')))
