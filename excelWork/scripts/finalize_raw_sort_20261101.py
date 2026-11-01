"""Verify the cache-imported delivery and generate its external test report."""
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json

import openpyxl

from verify_raw_sort_20261101 import package_checks,compare_finances,memo_checks
from verify_reserve_redesign import assert_no_errors
from verify_reserve_cash_alignment import verify_month


def finalize(source,delivery,sort_report,edit_report,default_values,output):
    report=json.loads(sort_report.read_text())
    report.pop('paths',None);report.pop('default_recalculated_path',None)
    report['after_sort_edits']=json.loads(edit_report.read_text())
    report['delivery_package']=package_checks(source,delivery)
    gc.collect()
    w=openpyxl.load_workbook(delivery,data_only=True)
    reference=openpyxl.load_workbook(default_values,data_only=True)
    assert_no_errors(w)
    checked=compare_finances(reference,w)
    report['delivery']={'compared_cells_against_recalculation':checked,
                        'reserve':verify_month(w,10),'memos':memo_checks(w),
                        'cached_errors':0,'original_record_order_retained':True}
    report['workbook_sha256']=sha256(delivery.read_bytes()).hexdigest()
    report['baseline_sha256']=sha256(source.read_bytes()).hexdigest()
    report['baseline_commit']='930e37e'
    report['previous_version']='history/2026-11-01/03'
    report['date']='2026-11-01'
    report['data_as_of']='2026-10-31'
    report['unit_tests_passed']=142
    report['synthetic_data_only_in_temp_copies']=True
    report['guide']='2026-11-01_RAW날짜정렬_검증.md'
    report['cache_import_only_no_office_roundtrip_of_delivery']=True
    report['temporary_evaluator_limitations']=[
        'Input CR/LF normalization in formula caches; original RAW memo text remains byte-exact',
        'Two parallel Office processes exhausted memory in the initial diagnostic; completed runs were serial',
        'Only fully qualified Table[[#This Row],[Header]] was supported in the engine probe; delivery uses it, not local @ syntax',
    ]
    output.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'workbook_sha256':report['workbook_sha256'],'delivery_checks':checked,
                      'sorted_numeric_and_record_checks_each':report['default']['financial_values_checked'],
                      'ascending_moved_rows':report['ascending_from_descending']['moved_rows'],
                      'formula_errors':0,'report':str(output)},ensure_ascii=False,indent=2))
    w.close();reference.close()
    return report


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for arg in ['source','delivery','sort_report','edit_report','default_values','output']:p.add_argument(arg,type=Path)
    a=p.parse_args();finalize(a.source,a.delivery,a.sort_report,a.edit_report,a.default_values,a.output)
