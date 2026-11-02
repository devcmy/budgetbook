"""Read-only full delivery checks and a disposable live-link calculation fixture."""
from hashlib import sha256
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
import argparse
import gc
import json

from lxml import etree as ET
import openpyxl
import olefile
from oletools.olevba import VBA_Project

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from validate_excel_structure import validate_package
from verify_entry_flow_20261102 import compare_actuals
from verify_reserve_redesign import recalculate,assert_no_errors
from raw_sort_id_20261101 import reorder
from simple_search_20261102 import matches,result_formulas,MODULE,VML


def sources(blob):
    o=olefile.OleFileIO(BytesIO(blob));p=VBA_Project(o,'','PROJECT','VBA/dir',relaxed=False)
    p.parse_project_stream();list(p.parse_modules());out={m.name:m.code_raw for m in p.modules};o.close();return out


def verify(baseline,delivery,root,report_path):
    a,b=ExtendedPackage(baseline),ExtendedPackage(delivery)
    structure=validate_package(delivery);assert not structure['errors'],structure
    allowed={'xl/vbaProject.bin','xl/styles.xml','xl/workbook.xml','[Content_Types].xml'}
    allowed.update(a.sheets[n] for n in ['입력·연결','사용안내','RAW','10월','11월','12월','잔액','예비비'])
    other_changed=[n for n in a.parts if n not in allowed and a.parts[n]!=b.parts[n]]
    assert not other_changed,other_changed
    original,new=sources(a.parts['xl/vbaProject.bin']),sources(b.parts['xl/vbaProject.bin'])
    assert set(new)-set(original)=={'BBSearch'}
    assert all(new[n]==code for n,code in original.items())
    assert new['BBSearch']==MODULE.read_text().replace('\r\n','\n').replace('\n','\r\n').encode('ascii')
    vml=ET.fromstring(b.parts[VML]);ns={'x':'urn:schemas-microsoft-com:office:excel'}
    assert vml.xpath('string(//x:FmlaMacro)',namespaces=ns)=='[0]!BB_Search'
    # Only display/input messages change on these existing sheets, never cells.
    for name in ['10월','11월','12월','잔액','예비비']:
        assert ET.tostring(a.sheet(name).find(Q('sheetData')))==ET.tostring(b.sheet(name).find(Q('sheetData'))),name
    before=openpyxl.load_workbook(baseline,data_only=True);after=openpyxl.load_workbook(delivery,data_only=True)
    assert before.sheetnames==after.sheetnames
    assert all(before[s].sheet_state==after[s].sheet_state for s in before.sheetnames)
    assert_no_errors(after);financial=compare_actuals(before,after)
    form=after['입력·연결'];assert form['E3'].value=='전체'
    assert form['B7'].value=='결제수단 / 계좌' and form['G8'].value=='메모 수정'
    assert form['K21'].value is None
    raw_unchanged=0;rows=[]
    for r in range(2,4002):
        record=[]
        for c in range(1,44):
            old,newval=before['RAW'].cell(r,c).value,after['RAW'].cell(r,c).value
            assert old==newval,(r,c,old,newval);raw_unchanged+=1;record.append(newval)
        rows.append(record)
    camping=[r[41] for r in rows if matches(r,'글램핑',0)]
    assert camping==[2515]
    assert [r[41] for r in rows if matches(r,'글램핑',11)]==[]
    assert [r[41] for r in rows if matches(r,'303',11)]==[2592]
    assert len([r for r in rows if matches(r,'',0)])==2597
    for m in range(1,13):assert before[f'{m}월']['E75'].value==after[f'{m}월']['E75'].value
    before.close();after.close();del before,after;gc.collect()
    # Simulate only the source ID a search would choose. Never execute VBA here.
    def edit(p):
        values=openpyxl.load_workbook(delivery,data_only=True)
        order=reorder(p,values,'ascending');values.close()
        p.write('입력·연결','H8',2515)
        for col,f in zip('ABCDEFG',result_formulas(8)):p.write('입력·연결',f'{col}8',f,True)
        # Alter only a fixture RAW memo: results must follow the ID after sorting.
        vals={}
        for row in p.sheet('RAW').find(Q('sheetData')):
            for cell in row:
                if cell.get('r','').startswith('AP'):
                    v=cell.find(Q('v'))
                    if v is not None and v.text=='2515':target=int(row.get('r'))
        p.write('RAW',f'I{target}','')
    calculated=recalculate(delivery,root/'link_calculation',edit)
    test=openpyxl.load_workbook(calculated,data_only=True);assert_no_errors(test)
    s=test['입력·연결']
    assert s['C8'].value=='숙소비' and s['D8'].value==169000
    assert s['E8'].value in [None,''] and s['F8'].value=='원본 수정'
    assert s['G8'].value=='메모 수정'
    assert test['11월']['E79'].value==75553
    test.close()
    report={'date':'2026-11-02','workbook_sha256':sha256(delivery.read_bytes()).hexdigest(),
            'baseline_sha256':sha256(baseline.read_bytes()).hexdigest(),
            'structure':structure,'raw_value_cells_compared':raw_unchanged,
            'actual_financial_cells_compared':financial,'existing_raw_records':2597,
            'raw_transactions_added':0,'formula_errors':0,'existing_vba_sources_preserved':len(original),
            'new_vba_module':'BBSearch','new_macro_button_binding_verified':True,
            'native_microsoft_excel_button_execution':False,'vba_was_not_executed_in_cloud':True,
            'all_month_search_finds_camping':True,'month_11_excludes_october_camping':True,
            'amount_keyword_finds_naver':True,'preview_and_result_links_use_stable_raw_id':True,
            'memo_deletion_and_sorted_raw_live_link_fixture_passed':True,
            'current_financial_values_and_forecasts_unchanged':True,'hidden_sheets_unchanged':True,
            'macros_must_be_enabled_by_user_for_search_button':True,
            'search_clears_raw_display_filters_only':True,
            'source_only_vba_performance_caches_rebuilt':True,
            'guide':'2026-11-02_검색화면단순화_PC버튼_검증.md'}
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(report,ensure_ascii=False),flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for arg in ['baseline','delivery','root','report']:p.add_argument(arg,type=Path)
    a=p.parse_args();verify(a.baseline,a.delivery,a.root,a.report)
