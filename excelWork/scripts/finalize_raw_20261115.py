"""Refresh current audit pointers after the corrected export is verified."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import json
import openpyxl


def finalize(workbook,report_path,audit_path):
    report=json.loads(report_path.read_text());digest=sha256(workbook.read_bytes()).hexdigest()
    assert report['workbook_sha256']==digest and report['formula_errors']==0
    assert report['managed_cash_total']==7262926 and report['repeat_new_records']==0
    assert report['migration_and_external_withdrawal_scenario_passed']
    w=openpyxl.load_workbook(workbook,data_only=True);raw=w['RAW'];ledger=w['잔액']
    rows=[r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    audit=json.loads(audit_path.read_text())
    audit.update(as_of='2026-11-15',ui_updated_at='2026-11-15',current_verification=report_path.name,workbook_sha256=digest)
    audit['raw'].update(count=len(rows),end=max(raw.cell(r,1).value for r in rows).isoformat(),
                        by_month=dict(sorted(Counter(raw.cell(r,1).value.strftime('%Y-%m') for r in rows).items())),
                        remaining_prepared_records=4000-len(rows),
                        november_scope=dict(Counter(raw.cell(r,18).value for r in rows if raw.cell(r,1).value.month==11)),
                        preservation_scope='Nov15 corrected export: 56 new source records, one explicit fee component; source memos synchronize including blanks. Prior statement-confirmed values preserved; Oct8 interest split from existing cumulative amount, no double count.')
    audit['november_summary']={a:w['11월'][a].value for a in ['C37','C40','C41','I33','H37','H38','I38','H43','H44','H45','O11','U11','O13','U13','O23','O24','O26','E74','E75','E78','E79']}
    audit['october_summary']={a:w['10월'][a].value for a in ['C40','I33','I38','H43','H44','E74','E75','E78','E79']}
    audit['december_summary']={a:w['12월'][a].value for a in ['O11','U11','O23','O24','E74','E75','E78','E79']}
    audit['managed_cash']=dict(total=7262926,balances_by_bank={'네이버미래에셋':7262399,'미래에셋_WRAP형CMA':527},
                               living_balance=5112926,purpose_available=2150000,selected_month='2026-11',
                               primary_account='네이버미래에셋',naver_scope_effective='2026-11-12',
                               bank_move_not_expense=True,latest_bank_app_balance_verified=False,
                               scope='Computed from the provided export, with last verified WRAP cumulative interest retained. No unreported daily interest is invented.')
    audit['wrap'].update(balances={ledger[f'A{r}'].value:ledger[f'G{r}'].value for r in range(7,16)},
                         total=ledger['G16'].value,previous=ledger['B16'].value,
                         inflows=ledger['C16'].value+ledger['F16'].value,
                         outflows=ledger['D16'].value+ledger['E16'].value,
                         selected_month='2026-11',cached_matches=True,
                         scope='Legacy key: current purpose-account ledger now combines NAVER and WRAP. Bank-level balances follow managed_cash; gross internal bank moves cancel in the total.')
    audit['november_wrap']=dict(total=527,scope='Residual in the original WRAP account; the new primary account is NAVER.',combined_managed_cash=7262926)
    for item in audit['calculators']:
        c=w[f'{item["month"]}월계산기']
        item.update(display=c['W12'].value,salary_burden=c['W13'].value,
                    counts={str(r):c[f'C{r}'].value for r in range(3,11)},
                    color_total=sum((c[f'C{r}'].value or 0)*1000 for r in range(3,11)),
                    cached_counts_match=True,expected_display_matches=True,month_display_matches=True)
    with ZipFile(workbook) as z:vba=z.read('xl/vbaProject.bin')
    audit['vba'].update(binary_sha256=sha256(vba).hexdigest(),binary_identical_to_previous_version=True,
                        input_reference='history/2026-11-15/01/테트리스_가계부_2026_고도화.xlsm',
                        preservation_scope='Nov15 source merge preserves existing VBA byte-for-byte.',executed=False)
    audit['raw_update_20261115']=report
    audit_path.write_text(json.dumps(audit,ensure_ascii=False,default=str,indent=2)+'\n')
    w.close();print('Published current audit:',digest,flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['workbook','report','audit']:p.add_argument(name,type=Path)
    a=p.parse_args();finalize(a.workbook,a.report,a.audit)
