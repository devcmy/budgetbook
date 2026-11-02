"""Formula and full-table sort invariants for native guided entry."""
from datetime import datetime
from types import SimpleNamespace
import unittest

from lxml import etree as ET
import openpyxl
from openpyxl.worksheet.table import Table

from entry_flow_20261102 import manual_slot_formula,manual_status_formula,plan_formula,plan_status,actual_count,candidate_condition,candidate_index_formula,source_position,PLAN,CODES,MODES
from raw_sort_id_20261101 import reorder,END
from update_20261015 import Q


class EntryFlowTests(unittest.TestCase):
    def test_explicit_slot_code_not_memo(self):
        formula=manual_slot_formula('BASE_SLOT')
        self.assertIn('[고정비 연결코드]',formula)
        self.assertIn('_ManualFixedCodes',formula)
        self.assertNotIn('[메모]',formula)
        self.assertNotIn('[연동메모]',formula)

    def test_large_original_formula_is_not_duplicated(self):
        self.assertEqual(manual_slot_formula('UNIQUE_BASE').count('UNIQUE_BASE'),1)

    def test_slot_needs_fixed_classification(self):
        self.assertIn('[적용분류]]="고정비"',manual_slot_formula('0'))
        self.assertIn('J 재분류를 고정비로',manual_status_formula('"기존"'))

    def test_ordinary_transfers_not_eligible_for_manual_expense(self):
        formula=manual_slot_formula('0')
        self.assertIn('카드정산결제',formula)
        self.assertIn('K패스',formula)
        self.assertIn('[수입/지출]]="지출"',formula)

    def test_plans_outside_memo_merge(self):
        self.assertEqual(PLAN,'AB')
        self.assertIn('AB24',plan_formula('OLD',24))
        self.assertNotIn('W24',plan_formula('OLD',24))

    def test_plan_only_applies_without_actuals(self):
        formula=plan_formula('OLD_ACTUAL',20)
        self.assertIn('=0)',formula)
        self.assertTrue(formula.endswith(',OLD_ACTUAL)'))
        self.assertIn('"예상"',plan_status('OLD_STATUS',20))

    def test_actual_count_checks_duplicates(self):
        self.assertIn('[동일거래순번],1',actual_count(20))
        self.assertIn('[수입/지출],"지출"',actual_count(20))

    def test_memo_search_is_only_navigation_not_classification(self):
        self.assertIn('[내용]',candidate_condition())
        self.assertIn('[메모]',candidate_condition())
        self.assertNotIn('[메모]',manual_slot_formula('ORIGINAL'))

    def test_candidate_function_has_office_compatible_prefix(self):
        self.assertIn('_xlfn.AGGREGATE(',candidate_index_formula(1))
        self.assertIn('RawData2026[거래ID]',source_position())

    def test_record_lookup_uses_stable_id(self):
        self.assertIn('MATCH($B$7,RawData2026[거래ID],0)',source_position())

    def test_all_entry_modes_present(self):
        self.assertIn('신규 RAW 등록',MODES)
        self.assertIn('고정비 예정 변경',MODES)
        self.assertIn('메모 수정',MODES)
        self.assertIn('WRAP 원장 연결',MODES)
        self.assertIn('예비비 사용 연결',MODES)
        self.assertIn('공동·대납 정산 확인',MODES)
        self.assertIn('WRAP 내부 계정이동',MODES)

    def test_fixed_codes_exclude_total_row(self):
        self.assertNotIn(25,CODES)
        self.assertEqual(len(CODES),22)

    def test_new_manual_column_moves_with_sorted_transaction(self):
        wb=openpyxl.Workbook();s=wb.active;s.title='RAW'
        s.add_table(Table(displayName='RawData2026',ref='A1:AQ4001'))
        root=ET.Element(Q('worksheet'));data=ET.SubElement(root,Q('sheetData'))
        for r in range(2,END+1):
            s.cell(r,42,r)
            row=ET.SubElement(data,Q('row'),r=str(r))
            if r in [2,3]:
                s.cell(r,1,datetime(2026,11,r-1));s.cell(r,43,20 if r==2 else 24)
                c=ET.SubElement(row,Q('c'),r=f'A{r}');ET.SubElement(c,Q('v')).text=str(r)
            c=ET.SubElement(row,Q('c'),r=f'AP{r}');ET.SubElement(c,Q('v')).text=str(r)
            if r in [2,3]:
                c=ET.SubElement(row,Q('c'),r=f'AQ{r}');ET.SubElement(c,Q('v')).text=str(s.cell(r,43).value)
            if r==2:
                c=ET.SubElement(row,Q('c'),r='AR2');ET.SubElement(c,Q('v')).text='GUIDE'
        package=SimpleNamespace(sheet=lambda name:root,cell_maps={},row_maps={})
        result=reorder(package,wb,'descending')
        row=data.find(Q('row'))
        cells={c.get('r'):c.find(Q('v')).text for c in row}
        self.assertEqual(cells['AP2'],'3');self.assertEqual(cells['AQ2'],'24')
        self.assertEqual(cells['AR2'],'GUIDE');self.assertGreater(result['moved_rows'],0)
        wb.close()


if __name__=='__main__':unittest.main()
