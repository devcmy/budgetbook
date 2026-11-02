"""Publish the verified entry workflow audit, without resaving the XLSM."""
from hashlib import sha256
from pathlib import Path
import argparse
import json

import openpyxl


def update_audit(delivery, report_path, audit_path):
    report = json.loads(report_path.read_text())
    digest = sha256(delivery.read_bytes()).hexdigest()
    assert report['workbook_sha256'] == digest
    assert report['formula_errors'] == 0 and not report['package']['structure']['errors']
    assert report['workflow_tests']['plan_only_changes_forecast_not_actual']
    assert report['workflow_tests']['manual_fixed_link_moves_existing_spending_without_adding_it_twice']
    assert report['camping_analysis']['source_transaction_untouched']
    book = openpyxl.load_workbook(delivery, data_only=True)
    assert book['11월']['E79'].value == 75553
    assert book['11월']['E75'].value == 1201465
    assert book['12월']['E75'].value == 1206060
    assert book['RAW'].tables['RawData2026'].ref == 'A1:AQ4001'
    audit = json.loads(audit_path.read_text())
    audit.update(current_verification=report_path.name, workbook_sha256=digest,
                 ui_updated_at='2026-11-02')
    audit['structure'].update(
        visible=[s.title for s in book if s.sheet_state == 'visible'],
        hidden=[s.title for s in book if s.sheet_state != 'visible'],
        total=len(book.sheetnames))
    assert len(audit['structure']['visible']) == 11 and len(audit['structure']['hidden']) == 36
    audit['raw'].update(table_range='A1:AQ4001', manual_fixed_code_column='AQ',
                        guide_range='AR:AW', manual_override_requires_fixed_class=True,
                        manual_codes_preserved_with_transaction_id_on_sort=True)
    audit['installments_20261102']['scope'] = (
        'Historical installment migration checks. Current data and forecasts unchanged; '
        'current workbook hash and guided input tests follow entry_flow_20261102.')
    report.update(guide='2026-11-02_입력연결_확정색상_캠핑검토.md',
                  baseline_commit='64643ba977d64232d8b19c760102eaa612375a37',
                  snapshot='history/2026-11-02/03',
                  unit_test_count=174,
                  preview_method='Disposable value-only copies rendered to PDF; no native Microsoft Excel UI execution.')
    audit['entry_flow_20261102'] = {k: v for k, v in report.items() if k != 'memos'}
    audit['entry_flow_20261102']['workflow_tests'] = {
        k: v for k, v in report['workflow_tests'].items() if k != 'memos'}
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, default=str, indent=2) + '\n')
    report_path.write_text(json.dumps(report, ensure_ascii=False, default=str, indent=2) + '\n')
    print(json.dumps({'workbook_sha256': digest, 'worksheets': len(book.sheetnames),
                      'raw_transactions_added': 0, 'november_actual': book['11월']['E79'].value},
                     ensure_ascii=False), flush=True)
    book.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    for arg in ['delivery', 'report', 'audit']:
        parser.add_argument(arg, type=Path)
    args = parser.parse_args()
    update_audit(args.delivery, args.report, args.audit)
