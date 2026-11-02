"""Publish the already-verified installment forecast report and audit pointer."""
from hashlib import sha256
from pathlib import Path
import argparse
import json

import openpyxl


def update_audit(delivery,report_path,audit_path):
    report=json.loads(report_path.read_text());digest=sha256(delivery.read_bytes()).hexdigest()
    assert report['workbook_sha256']==digest
    assert report['future_invoice_tests']['buds_airpods_phone_mapped_separately']
    w=openpyxl.load_workbook(delivery,data_only=True)
    assert w['11월']['E75'].value==1201465 and w['11월']['E79'].value==75553
    audit=json.loads(audit_path.read_text())
    audit.update(current_verification=report_path.name,workbook_sha256=digest,ui_updated_at='2026-11-02')
    audit['raw']['installment_product_matching']='Explicit month+asset+normalized title/installment aliases for Nov/Dec 2026; no free memo dependency; existing RAW values unchanged.'
    audit['november_summary'].update({a:w['11월'][a].value for a in ['O24','S24','U24','E74','E75','E78','E79']})
    audit['december_summary']={a:w['12월'][a].value for a in ['O24','S24','U24','E74','E75','E78','E79']}
    audit['postpaid_20261102']['scope']='Historical transport/November input checks; actual data unchanged, current forecasts and file hash follow installments_20261102.'
    report['guide']='2026-11-02_할부이월_종료항목정리_검증.md'
    audit['installments_20261102']={k:v for k,v in report.items() if k not in ['memos']}
    audit_path.write_text(json.dumps(audit,ensure_ascii=False,default=str,indent=2)+'\n')
    report_path.write_text(json.dumps(report,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({'workbook_sha256':digest,'november_forecast':w['11월']['E75'].value,
                      'november_actual':w['11월']['E79'].value,'december_forecast':w['12월']['E75'].value,
                      'raw_transactions_added':0},ensure_ascii=False),flush=True)
    w.close()


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['delivery','report','audit']:p.add_argument(a,type=Path)
    a=p.parse_args();update_audit(a.delivery,a.report,a.audit)
