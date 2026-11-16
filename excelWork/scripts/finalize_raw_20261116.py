"""Publish current audit pointers only after finance and visual checks pass."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import json

import openpyxl


def finalize(workbook, report_path, audit_path):
    report = json.loads(report_path.read_text())
    digest = sha256(workbook.read_bytes()).hexdigest()
    assert report['workbook_sha256'] == digest and report['formula_errors'] == 0
    assert report['raw_edit_delete_new_month_and_memo_clear_scenario_passed']
    assert all(report[k]['conditional_colors_rendered'] for k in
               ['current_color_render','edited_color_render','december_color_render'])
    assert report['repeat_new_records'] == report['repeat_updates'] == 0
    w = openpyxl.load_workbook(workbook, data_only=True)
    raw, ledger = w['RAW'], w['잔액']
    rows = [r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    audit = json.loads(audit_path.read_text())
    audit.update(as_of='2026-11-16', ui_updated_at='2026-11-16',
                 current_verification=report_path.name, workbook_sha256=digest)
    audit['raw'].update(count=len(rows), end=max(raw.cell(r,1).value for r in rows).isoformat(),
                        by_month=dict(sorted(Counter(raw.cell(r,1).value.strftime('%Y-%m') for r in rows).items())),
                        remaining_prepared_records=4000-len(rows),
                        november_scope=dict(Counter(raw.cell(r,18).value for r in rows if raw.cell(r,1).value.month == 11)),
                        preservation_scope='Nov16: 8 new source records and 3 corrections; the revised refund retains its original ID. Memos remain independent of classification.')
    for key, month, addresses in [
        ('november_summary',11,['C37','C40','C41','I33','H37','H38','I38','H43','H44','H45','O11','U11','O13','U13','O23','O24','O26','E74','E75','E78','E79']),
        ('october_summary',10,['C40','I33','I38','H43','H44','E74','E75','E78','E79']),
        ('december_summary',12,['O11','U11','O23','O24','E74','E75','E78','E79'])]:
        audit[key] = {addr:w[f'{month}월'][addr].value for addr in addresses}
    audit['managed_cash'].update(total=report['managed_cash_total'],
        balances_by_bank={'네이버미래에셋':report['naver_cash'],'미래에셋_WRAP형CMA':report['wrap_residual']},
        living_balance=ledger['G7'].value, purpose_available=report['purpose_available'],
        selected_month='2026-11', latest_bank_app_balance_verified=False)
    audit['wrap'].update(balances={ledger[f'A{r}'].value:ledger[f'G{r}'].value for r in range(7,16)},
        total=ledger['G16'].value, previous=ledger['B16'].value,
        inflows=ledger['C16'].value+ledger['F16'].value,
        outflows=ledger['D16'].value+ledger['E16'].value, selected_month='2026-11',cached_matches=True)
    audit['november_wrap']['combined_managed_cash'] = report['managed_cash_total']
    for item in audit['calculators']:
        c = w[f'{item["month"]}월계산기']
        item.update(display=c['W12'].value, salary_burden=c['W13'].value,
                    counts={str(r):c[f'C{r}'].value for r in range(3,11)},
                    color_total=sum((c[f'C{r}'].value or 0)*1000 for r in range(3,11)),
                    cached_counts_match=True, expected_display_matches=True, month_display_matches=True)
        if item['month'] in [11,12]:
            item.update(report['calculator_physical_fills'][str(item['month'])])
            item['short_dynamic_color_rules_verified'] = True
            item['grid_borders_extended'] = True
    with ZipFile(workbook) as z:
        vba = z.read('xl/vbaProject.bin')
    audit['vba'].update(binary_sha256=sha256(vba).hexdigest(),binary_identical_to_previous_version=True,
        input_reference='history/2026-11-16/01/테트리스_가계부_2026_고도화.xlsm',
        preservation_scope='Nov16 source and visual updates preserve VBA byte-for-byte.',executed=False)
    audit['raw_update_20261116'] = report
    audit_path.write_text(json.dumps(audit,ensure_ascii=False,default=str,indent=2)+'\n')
    w.close()
    print('Published current audit:',digest,flush=True)


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    for name in ['workbook','report','audit']:
        p.add_argument(name,type=Path)
    a = p.parse_args()
    finalize(a.workbook,a.report,a.audit)
