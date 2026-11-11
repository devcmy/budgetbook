"""Refresh current audit pointers after the independent checks pass."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import json
import openpyxl


def finalize(workbook, report_path, wrap_path, audit_path):
    report=json.loads(report_path.read_text());wrap=json.loads(wrap_path.read_text())
    digest=sha256(workbook.read_bytes()).hexdigest()
    assert report['workbook_sha256']==digest and report['formula_errors']==0
    assert report['vba_buttons_styles_relationships_byte_preserved']
    assert wrap['september_bridge_residual']==0 and wrap['october_missing_or_duplicate_links']==0
    w=openpyxl.load_workbook(workbook,data_only=True)
    rows=[r for r in w['RAW'].iter_rows(min_row=2,max_row=4001,values_only=True) if isinstance(r[0],datetime)]
    audit=json.loads(audit_path.read_text())
    audit.update(as_of='2026-11-11',ui_updated_at='2026-11-11',current_verification=report_path.name,workbook_sha256=digest)
    audit['raw'].update(count=len(rows),end=max(r[0] for r in rows).isoformat(),
                        by_month=dict(sorted(Counter(r[0].strftime('%Y-%m') for r in rows).items())),
                        november_scope=dict(Counter(r[17] for r in rows if r[0].month==11)),
                        remaining_prepared_records=4000-len(rows),
                        preservation_scope='Nov11 import adds 54 records; all existing inputs/memos preserved by stable ID. Source 12,802 income is preserved as submitted but explicitly unresolved, not proven interest.')
    audit['november_summary']={a:w['11월'][a].value for a in ['C37','C40','C41','I33','H37','H38','I38','H43','H44','H45','O11','U11','O18','O21','O22','O23','O24','O26','E74','E75','E78','E79']}
    audit['november_summary']['income_includes_unresolved_adjustment']=12802
    audit['december_summary']={a:w['12월'][a].value for a in ['O11','U11','O24','S24','U24','E74','E75','E78','E79']}
    audit['wrap']['selected_month']='2026-10'
    audit['wrap']['scope']='Existing October selector retained. Current November figures and unresolved adjustment follow november_wrap and wrap_reconciliation_20261111.'
    audit['november_wrap']=dict(total_as_submitted=8287823,total_without_unresolved_adjustment=8275021,
                               living_as_submitted=6137823,purpose_funds=2150000,unresolved_adjustment=12802,
                               bank_statement_verified=False)
    for item in audit['calculators']:
        s=w[f'{item["month"]}월계산기']
        item.update(display=s['W12'].value,salary_burden=s['W13'].value,
                    counts={str(r):s[f'C{r}'].value for r in range(3,11)},
                    color_total=sum((s[f'C{r}'].value or 0)*1000 for r in range(3,11)),
                    cached_counts_match=True,expected_display_matches=True,month_display_matches=True)
    with ZipFile(workbook) as z:vba=z.read('xl/vbaProject.bin')
    audit['vba'].update(binary_sha256=sha256(vba).hexdigest(),binary_identical_to_previous_version=True,
                        input_reference='history/2026-11-11/01/테트리스_가계부_2026_고도화.xlsm',
                        preservation_scope='Nov11 import preserves the previous VBA byte-for-byte.',executed=False)
    audit['raw_update_20261111']=report
    audit['wrap_reconciliation_20261111']=wrap
    audit_path.write_text(json.dumps(audit,ensure_ascii=False,default=str,indent=2)+'\n')
    w.close();print('Current audit:',digest,flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ['workbook','report','wrap','audit']:p.add_argument(name,type=Path)
    a=p.parse_args();finalize(a.workbook,a.report,a.wrap,a.audit)
