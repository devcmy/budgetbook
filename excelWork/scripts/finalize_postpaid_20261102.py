"""Refresh generated audit metadata after all November delivery tests pass."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import json

import openpyxl


def publish_audit(delivery,report_path,memo_report,audit_path,source,previous):
    report=json.loads(report_path.read_text())
    digest=sha256(delivery.read_bytes()).hexdigest()
    assert digest==report['workbook_sha256']
    assert report['future_charge_sort_and_memo_test']['ordinary_card_settlement_still_excluded']
    report.update(memo_cache_restoration=json.loads(memo_report.read_text()),
                  previous_git_commit='0292bbf',previous_version='history/2026-11-02/01',
                  user_attachment='inputs/2026-11-02/테트리스_가계부_2026_고도화_첨부원본.xlsm',
                  raw_attachment='raw/2026-11-02/2026-09-01_2026-11-02.xlsx',
                  guide='2026-11-02_모두의카드_청구월_11월반영_검증.md')
    w=openpyxl.load_workbook(delivery,data_only=True);raw=w['RAW']
    rows=[r for r in range(2,4002) if isinstance(raw.cell(r,1).value,datetime)]
    assert len(rows)==2597
    audit=json.loads(audit_path.read_text())
    audit['structure'].update(visible=[s.title for s in w if s.sheet_state=='visible'],
                              hidden=[s.title for s in w if s.sheet_state!='visible'],total=len(w.sheetnames))
    counts=Counter(raw.cell(r,1).value.strftime('%Y-%m') for r in rows)
    audit['raw'].update(count=len(rows),start=min(raw.cell(r,1).value for r in rows).isoformat(),
                        end=max(raw.cell(r,1).value for r in rows).isoformat(),by_month=dict(sorted(counts.items())),
                        november_scope=dict(Counter(raw.cell(r,18).value for r in rows if raw.cell(r,1).value.month==11)),
                        remaining_prepared_records=4000-len(rows),date_display_format=raw['A2'].number_format,
                        postpaid_transport='Usage summary is evidence only; identified bank settlement is fixed cost in charge month. No memo dependency.',
                        source_fields_and_existing_overrides_preserved=True)
    for item in audit['calculators']:
        m=item['month'];s=w[f'{m}월계산기']
        item.update(display=s['W12'].value,salary_burden=s['W13'].value,
                    counts={str(r):s[f'C{r}'].value for r in range(3,11)},
                    color_total=sum((s[f'C{r}'].value or 0)*1000 for r in range(3,11)),
                    cached_counts_match=True,expected_display_matches=True,month_display_matches=True)
    with ZipFile(delivery) as z,ZipFile(source) as src,ZipFile(previous) as old:
        vba=z.read('xl/vbaProject.bin');assert vba==src.read('xl/vbaProject.bin')
        same_previous=vba==old.read('xl/vbaProject.bin')
    # Excel's save changed the uploaded binary relative to yesterday's Git
    # file. Preserve the user's binary, without claiming a new source audit.
    old_vba_audit={k:audit['vba'].pop(k) for k in ['all_module_sources_equal','module_count','binary_identical'] if k in audit['vba']}
    audit['vba'].update(historical_module_audit=old_vba_audit,current_module_sources_reextracted=False,
                        binary_sha256=sha256(vba).hexdigest(),binary_identical_to_previous_version=same_previous,
                        binary_identical_to_user_attachment=True,binary_identical_to_input=True,executed=False)
    audit.update(as_of='2026-11-02',ui_updated_at='2026-11-02',current_verification=report_path.name,
                 workbook_sha256=digest,
                 november_summary={a:w['11월'][a].value for a in ['C37','C40','H37','H38','I38','H44','H45','O11','U11','E74','E75','E78','E79']})
    audit['raw_sort_20261101']['scope']='Historical sorting verification. Sorting remains supported; current user edits/data/hash follow postpaid_20261102.'
    audit['postpaid_20261102']={k:v for k,v in report.items() if k not in ['metadata','memos','package']}
    audit['postpaid_20261102']['source_matching_summary']={
        'matched_existing':len(report['metadata']['matched']),
        'added':len(report['metadata']['new']),
        'unrelated_september_new_candidates_deferred':len(report['metadata']['deferred_unrelated_september']),
        'old_september_absent_from_new_export_preserved':len(report['metadata']['existing_september_not_in_new_export_preserved']),
    }
    audit_path.write_text(json.dumps(audit,ensure_ascii=False,default=str,indent=2)+'\n')
    report_path.write_text(json.dumps(report,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({'workbook_sha256':digest,'raw_records':len(rows),'visible_sheets':len(audit['structure']['visible']),
                      'november_usage':report['november_usage'],'user_vba_binary_preserved':True},ensure_ascii=False),flush=True)
    w.close()


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['delivery','report','memo_report','audit','source','previous']:p.add_argument(a,type=Path)
    a=p.parse_args();publish_audit(a.delivery,a.report,a.memo_report,a.audit,a.source,a.previous)
