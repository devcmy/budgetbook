"""Pure formula migration regressions (delivery ZIP tests run separately)."""
import unittest
from raw_sort_id_20261101 import rewrite_external,rewrite_raw,duplicate_formula,memo_sequence,lookup,current

HEADERS={1:'일시',2:'월',6:'내용',7:'금액(KRW)',8:'수입/지출',9:'메모',10:'재분류',11:'본인부담',40:'연동메모',42:'거래ID'}


class SortFormulaTests(unittest.TestCase):
    def test_row_becomes_id_not_display_position(self):
        self.assertEqual(rewrite_raw('=IF(A7="",0,ROW())',7,HEADERS),f'IF({current("일시")}="",0,{current("거래ID")})')

    def test_quoted_addresses_are_not_executed_or_rewritten(self):
        self.assertEqual(rewrite_external('="RAW!A356"&RAW!K356',HEADERS),'"RAW!A356"&'+lookup('본인부담','356'))

    def test_register_index_is_id_lookup(self):
        self.assertEqual(rewrite_external('=IFERROR(INDEX(RAW!$K:$K,MAX(2,F100)),"")',HEADERS),'IFERROR('+lookup('본인부담','MAX(2,F100)')+',"")')

    def test_static_table_index_offset_is_removed(self):
        self.assertEqual(rewrite_external('=INDEX(RawData2026[연동메모],355)',HEADERS),lookup('연동메모','356'))

    def test_register_table_index_offset_is_removed(self):
        self.assertEqual(rewrite_external('=INDEX(RawData2026[연동메모],W248-1)',HEADERS),lookup('연동메모','W248'))

    def test_existing_key_match_stays_intact(self):
        text='INDEX(RawData2026[연동메모],MATCH(10001,RawData2026[메모순번],0))'
        self.assertEqual(rewrite_external('='+text,HEADERS),text)

    def test_full_raw_ranges_are_bounded_table_columns(self):
        self.assertEqual(rewrite_external('=SUM(RAW!$G$2:$G$10018)',HEADERS),'SUM(RawData2026[금액(KRW)])')

    def test_local_external_mapping_keeps_fixed_other_sheet(self):
        text='=A7+\'_RAW자동연동\'!$A$2'
        self.assertEqual(rewrite_raw(text,7,HEADERS),current('일시')+'+\'_RAW자동연동\'!$A$2')

    def test_duplicate_and_memo_ranks_use_id(self):
        for text in [duplicate_formula(),memo_sequence()]:
            self.assertIn('RawData2026[거래ID]',text)
            self.assertIn(current('거래ID'),text)
            self.assertNotIn('ROW()',text)

    def test_memo_not_duplicate_or_classification_criterion(self):
        self.assertNotIn('[메모]',duplicate_formula())
        self.assertNotIn('[연동메모]',duplicate_formula())


if __name__=='__main__': unittest.main()
