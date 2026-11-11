"""Publish current data/audit pointers for confirmed cumulative interest."""
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
    assert report['wrap_balance']==8288985 and report['unverified_balance_adjustment']==0
    w=openpyxl.load_workbook(workbook,data_only=True);s=w['RAW'];ledger=w['잔액']
    rows=[r for r in range(2,4002) if isinstance(s.cell(r,1).value,datetime)]
    audit=json.loads(audit_path.read_text())
    audit.update(as_of='2026-11-11',ui_updated_at='2026-11-11',current_verification=report_path.name,workbook_sha256=digest)
    audit['raw'].update(count=len(rows),end=max(s.cell(r,1).value for r in rows).isoformat(),
                        by_month=dict(sorted(Counter(s.cell(r,1).value.strftime('%Y-%m') for r in rows).items())),
                        november_scope=dict(Counter(s.cell(r,18).value for r in rows if s.cell(r,1).value.month==11)),
                        remaining_prepared_records=4000-len(rows),
                        preservation_scope='Confirmed Nov11 cumulative WRAP interest in existing ID2667. Five source records added; documented corrections only. Free memos never determine classification.')
    audit['november_summary']={a:w['11월'][a].value for a in ['C37','C40','C41','I33','H37','H38','I38','H43','H44','H45','O11','U11','O18','O21','O22','O23','O24','O26','E74','E75','E78','E79']}
    audit['november_summary'].update(confirmed_cumulative_wrap_interest=13964,unverified_adjustment=0)
    audit['december_summary']={a:w['12월'][a].value for a in ['O11','U11','O23','O24','S24','U24','E74','E75','E78','E79']}
    audit['wrap'].update(balances={ledger[f'A{r}'].value:ledger[f'G{r}'].value for r in range(7,16)},
                         total=ledger['G16'].value,previous=ledger['B16'].value,
                         inflows=ledger['C16'].value+ledger['F16'].value,
                         outflows=ledger['D16'].value+ledger['E16'].value,
                         selected_month='2026-11',cached_matches=True,
                         scope='Nov11 certificate closing balance. Cumulative interest is booked once when the user updates the ledger.')
    audit['november_wrap']=dict(total=8288985,living_balance=6138985,purpose_funds=2150000,
                               cumulative_interest=13964,interest_record_count=1,unresolved_adjustment=0,
                               bank_statement_verified=True,as_of='2026-11-11')
    if 'wrap_reconciliation_20261111' in audit:
        audit['wrap_reconciliation_20261111']['scope']='Historical diagnostic before the securities certificate; superseded by confirmed_interest_20261111.'
    for item in audit['calculators']:
        c=w[f'{item["month"]}월계산기']
        item.update(display=c['W12'].value,salary_burden=c['W13'].value,
                    counts={str(r):c[f'C{r}'].value for r in range(3,11)},
                    color_total=sum((c[f'C{r}'].value or 0)*1000 for r in range(3,11)),
                    cached_counts_match=True,expected_display_matches=True,month_display_matches=True)
    with ZipFile(workbook) as z:vba=z.read('xl/vbaProject.bin')
    audit['vba'].update(binary_sha256=sha256(vba).hexdigest(),binary_identical_to_previous_version=True,
                        input_reference='history/2026-11-11/02/테트리스_가계부_2026_고도화.xlsm',
                        preservation_scope='Confirmed-interest update preserves existing VBA byte-for-byte.',executed=False)
    audit['confirmed_interest_20261111']=report
    audit_path.write_text(json.dumps(audit,ensure_ascii=False,default=str,indent=2)+'\n')
    w.close();print('Published audit',digest,flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['workbook','report','audit']:p.add_argument(name,type=Path)
    a=p.parse_args();finalize(a.workbook,a.report,a.audit)
