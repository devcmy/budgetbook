"""Publish current data/audit pointers without reserializing the workbook."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import json

import openpyxl

from update_raw_20261104 import plan


def finalize(delivery,source,report_path,audit_path):
    report=json.loads(report_path.read_text());digest=sha256(delivery.read_bytes()).hexdigest()
    assert digest==report['workbook_sha256'] and report['formula_errors']==0
    assert report['vba_buttons_styles_relationships_byte_preserved']
    repeat=plan(delivery,source)
    assert not repeat['new'] and not repeat['corrections']
    report.update(guide='2026-11-04_11월RAW추가_검증.md',unit_test_count=188,
                  previous_commit='27ab9f5e66fc4e2754a8d86e92b76184e9542a1d',
                  previous_version='history/2026-11-04/01',raw_attachment='raw/2026-11-04/'+source.name,
                  repeated_import_new_records=0,repeated_import_corrections=0)
    w=openpyxl.load_workbook(delivery,data_only=True);raw=w['RAW']
    rows=[r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    audit=json.loads(audit_path.read_text())
    audit.update(as_of='2026-11-04',ui_updated_at='2026-11-04',current_verification=report_path.name,workbook_sha256=digest)
    audit['raw'].update(count=len(rows),end=max(raw.cell(r,1).value for r in rows).isoformat(),
                        by_month=dict(sorted(Counter(raw.cell(r,1).value.strftime('%Y-%m') for r in rows).items())),
                        november_scope=dict(Counter(raw.cell(r,18).value for r in rows if raw.cell(r,1).value.month==11)),
                        remaining_prepared_records=4000-len(rows),
                        future_dated_invoices_confirmed_actual_on='2026-11-04',
                        preservation_scope='Historical inputs/memos preserved by stable ID except reconciled 11/2 transaction ID2597: amount and full personal burden 5000 -> 4500. No original booking dates changed.')
    audit['november_summary']={a:w['11월'][a].value for a in ['C37','C40','C41','I33','H37','H38','I38','H43','H44','H45','O11','U11','O18','O21','O22','O23','O24','S24','U24','E74','E75','E78','E79']}
    audit['december_summary']={a:w['12월'][a].value for a in ['O24','S24','U24','E74','E75','E78','E79']}
    for item in audit['calculators']:
        s=w[f'{item["month"]}월계산기']
        item.update(display=s['W12'].value,salary_burden=s['W13'].value,
                    counts={str(r):s[f'C{r}'].value for r in range(3,11)},
                    color_total=sum((s[f'C{r}'].value or 0)*1000 for r in range(3,11)),
                    cached_counts_match=True,expected_display_matches=True,month_display_matches=True)
    with ZipFile(delivery) as z:vba=z.read('xl/vbaProject.bin')
    audit['vba'].update(binary_sha256=sha256(vba).hexdigest(),binary_identical_to_previous_version=True,
                        binary_identical_to_input=True,binary_identical_to_user_attachment=False,
                        input_reference='history/2026-11-04/01/테트리스_가계부_2026_고도화.xlsm',
                        preservation_scope='Current Nov4 update preserves the Nov2 search-button VBA binary exactly; historical source audits remain historical.',executed=False)
    audit['simple_search_20261102']['scope']='Historical search-screen implementation tests. Search VBA, controls and styles unchanged; current data and workbook hash follow raw_update_20261104.'
    audit['raw_update_20261104']={k:v for k,v in report.items() if k!='memos'}
    for path,value in [(report_path,report),(audit_path,audit)]:
        path.write_text(json.dumps(value,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({'workbook_sha256':digest,'raw_records':len(rows),'november_usage':w['11월']['E79'].value,
                      'repeat_import_new_records':0},ensure_ascii=False),flush=True)
    w.close()


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['delivery','source','report','audit']:p.add_argument(a,type=Path)
    a=p.parse_args();finalize(a.delivery,a.source,a.report,a.audit)
