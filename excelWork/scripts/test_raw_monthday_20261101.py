import unittest
from lxml import etree as ET

from raw_monthday_memo_20261101 import DATE_FORMAT,date_style
from update_20261015 import Q


class MonthDayDisplayTests(unittest.TestCase):
    def source(self):
        xf=ET.Element(Q('xf'),fontId='3',fillId='4',borderId='5',numFmtId='168',xfId='0')
        ET.SubElement(xf,Q('protection'),locked='0')
        return xf
    def test_month_day_format(self):self.assertEqual(DATE_FORMAT,'m"월" d"일"')
    def test_no_visible_year_or_time(self):
        self.assertNotIn('y',DATE_FORMAT);self.assertNotIn('h',DATE_FORMAT)
    def test_original_style_not_mutated(self):
        xf=self.source();date_style(xf,200);self.assertEqual(xf.get('numFmtId'),'168')
    def test_number_format_applied(self):
        xf=date_style(self.source(),200);self.assertEqual(xf.get('numFmtId'),'200');self.assertEqual(xf.get('applyNumberFormat'),'1')
    def test_input_protection_and_colors_preserved(self):
        xf=date_style(self.source(),200)
        self.assertEqual(xf.find(Q('protection')).get('locked'),'0')
        self.assertEqual([xf.get(k) for k in ['fontId','fillId','borderId']],['3','4','5'])


if __name__=='__main__':unittest.main()
