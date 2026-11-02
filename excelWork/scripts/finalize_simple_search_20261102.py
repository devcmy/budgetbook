"""Update current audit pointers after successful read-only search verification."""
from hashlib import sha256
from pathlib import Path
import argparse
import json


def finalize(delivery,report_path,audit_path):
    report=json.loads(report_path.read_text());digest=sha256(delivery.read_bytes()).hexdigest()
    assert report['workbook_sha256']==digest
    assert not report['structure']['errors']
    assert report['memo_deletion_and_sorted_raw_live_link_fixture_passed']
    assert report['existing_vba_sources_preserved']==13
    report.update(unit_test_count=181,baseline_commit='96d097e97d2867ffb71d2c2cfdec9413331911ac',
                  snapshot='history/2026-11-02/04')
    audit=json.loads(audit_path.read_text())
    audit.update(current_verification=report_path.name,workbook_sha256=digest,ui_updated_at='2026-11-02')
    audit['entry_flow_20261102']['scope']='Historical guided form; replaced by the simple desktop button search. Actual data, manual RAW AQ mapping and monthly AB plans are preserved.'
    audit['simple_search_20261102']=report
    for path,value in [(report_path,report),(audit_path,audit)]:
        path.write_text(json.dumps(value,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({'workbook_sha256':digest,'current_verification':report_path.name},ensure_ascii=False))


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for arg in ['delivery','report','audit']:p.add_argument(arg,type=Path)
    a=p.parse_args();finalize(a.delivery,a.report,a.audit)
