"""Merge only the new November export, preserving stable IDs and source memos.

The user explicitly confirmed the three future-dated installment invoices as
actual confirmed charges on Nov 4. Original booking dates are not rewritten.
No historical migrations or VBA edits are rerun.
"""
from collections import defaultdict,Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json

import openpyxl

from redesign_reserve import ExtendedPackage
from update_postpaid_20261102 import read_source,key,compact
from update_raw_20261101 import explicit_class,WRAP,SHARED
from raw_sort_id_20261101 import END,reorder
from update_20261015 import Q
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore

TODAY=datetime(2026,11,4)


def correction_key(item):
    return (item['date'],item['asset'],item['kind'],item['category'],item['sub'] or None,compact(item['title']))


def classify(item):
    # Regular/installment evidence has precedence over an export's electronics
    # or gift category. Memo product names never drive financial classification.
    title=compact(item['title'])
    if item['category']=='정기구독' or title.startswith('11번가') or '띵샵' in title or '애플케어' in title:
        return None
    return explicit_class(item)


def corrected_burden(original_amount,original_burden,new_amount):
    if original_burden in [None,'']:return original_burden
    assert original_burden==original_amount,'Partial burden requires user confirmation'
    return new_amount


def plan(baseline,source):
    w=openpyxl.load_workbook(baseline,data_only=True);raw=w['RAW']
    exact=defaultdict(list);loose=defaultdict(list);ids={};old={}
    for r in range(2,END+1):
        ident=int(raw.cell(r,42).value);ids[ident]=r;date=raw.cell(r,1).value
        if not isinstance(date,datetime):continue
        item=dict(date=date,asset=raw.cell(r,3).value,category=raw.cell(r,4).value,
                  sub=raw.cell(r,5).value,title=raw.cell(r,6).value,amount=raw.cell(r,7).value,
                  kind=raw.cell(r,8).value,memo=raw.cell(r,9).value)
        exact[key(item)].append(ident);loose[correction_key(item)].append(ident);old[ident]=item
    assert set(ids)==set(range(2,END+1))
    items=read_source(source);mappings=[];new=[];corrections=[];labels=[];used=set()
    source_exact=Counter(key(item) for item in items)
    source_loose=Counter(correction_key(item) for item in items)
    for item in items:
        hits=exact.get(key(item),[])
        assert len(hits)<=1,('Ambiguous exact identity',item['source_row'],hits)
        if hits:
            ident=hits[0];assert ident not in used;used.add(ident)
            mappings.append(dict(source_row=item['source_row'],transaction_id=ident,status='preserved'))
            if any((old[ident][k] or None)!=(item[k] or None) for k in ['category','sub','title','memo']):labels.append(ident)
            continue
        assert item['date'].month==11,('Unrequested October addition/change',item)
        possible=loose.get(correction_key(item),[])
        assert len(possible)<=1,('Ambiguous correction',item,possible)
        if possible:
            ident=possible[0];assert ident not in used;used.add(ident)
            assert source_loose[correction_key(item)]==1 and source_exact[key(old[ident])]==0
            r=ids[ident];assert raw.cell(r,12).value=='본인'
            corrections.append(dict(item,transaction_id=ident,input_row=r,old_amount=old[ident]['amount'],
                                    old_burden=raw.cell(r,11).value,
                                    new_burden=corrected_burden(old[ident]['amount'],raw.cell(r,11).value,item['amount'])))
            mappings.append(dict(source_row=item['source_row'],transaction_id=ident,status='amount_corrected'))
        else:new.append(dict(item))
    new.sort(key=lambda item:(item['date'],item['source_row']))
    empty=sorted(i for i,r in ids.items() if raw.cell(r,1).value in [None,''])
    assert len(empty)>=len(new)
    for ident,item in zip(empty,new):
        item.update(transaction_id=ident,input_row=ids[ident])
        mappings.append(dict(source_row=item['source_row'],transaction_id=ident,status='added'))
    missing=[ident for ident,item in old.items() if item['date'].month in [10,11] and ident not in used]
    assert not missing,('Old Oct/Nov transactions absent from export',missing)
    futures=[dict(transaction_id=item['transaction_id'],source_row=item['source_row'],date=item['date'],amount=item['amount'],title=item['title'])
             for item in new if item['date'].date()>TODAY.date()]
    assert all('띵샵' in compact(item['title']) or compact(item['title']).startswith('11번가') for item in futures)
    assert not any(item['asset']==WRAP or item['category']==WRAP or item['asset'] in SHARED for item in new)
    out=dict(new=new,corrections=corrections,source_mapping=sorted(mappings,key=lambda item:item['source_row']),
             source_records=len(items),source_counts=dict(Counter(item['date'].strftime('%Y-%m') for item in items)),
             old_raw_count=len(old),old_november_count=sum(item['date'].month==11 for item in old.values()),
             old_labels_and_memos_preserved_ids=labels,future_invoices=futures,
             future_invoice_actual_status_user_confirmed_on='2026-11-04',new_wrap_or_reserve_links=0)
    w.close();return out


def prepare(baseline,source,output,confirmed_future=False):
    meta=plan(baseline,source)
    assert not meta['future_invoices'] or confirmed_future,'Future invoice actual status must be confirmed'
    p=ExtendedPackage(baseline)
    for item in meta['new']:
        row=item['input_row']
        for c,k in [('A','date'),('C','asset'),('D','category'),('E','sub'),('F','title'),('G','amount'),('H','kind'),('I','memo')]:
            p.write('RAW',f'{c}{row}',item[k])
        classification=classify(item)
        if classification:p.write('RAW',f'J{row}',classification)
    for item in meta['corrections']:
        row=item['input_row'];p.write('RAW',f'G{row}',item['amount'])
        p.write('RAW',f'K{row}',item['new_burden'])
    p.write('11월','A70','11/4 첨부 RAW 28건 · 신규 20건 + 기존 1건 금액 정정. 11/5·12·24 할부는 사용자 확인된 청구 확정액입니다.')
    p.write('사용안내','A26','11/4 갱신 · 11월 RAW 28건. 11/5·12·24 할부는 청구 확정으로 반영하되 원본 기재일을 유지합니다.')
    p.save(output)
    values=openpyxl.load_workbook(output,data_only=True)
    meta['sort']=reorder(p,values,'descending');values.close()
    p.save(output)
    values=openpyxl.load_workbook(output,data_only=True)
    recent=[values['RAW'].cell(row,42).value for row in range(2,12)]
    for row,ident in enumerate(recent,8):p.write('입력·연결',f'H{row}',ident)
    values.close();p.save(output);gc.collect();return meta


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for name in ['baseline','source','output','root','metadata']:parser.add_argument(name,type=Path)
    parser.add_argument('--confirmed-future-invoices',action='store_true')
    a=parser.parse_args();candidate=a.root/'candidate.xlsm'
    metadata=prepare(a.baseline,a.source,candidate,a.confirmed_future_invoices)
    a.metadata.write_text(json.dumps(metadata,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Prepared November delta:',len(metadata['new']),'new,',len(metadata['corrections']),'amount correction.',flush=True)
    calculated=recalculate(candidate,a.root/'calculation')
    changes=import_changed_caches(candidate,calculated,a.root/'cached.xlsm')
    memos=restore(a.baseline,a.root/'cached.xlsm',a.output)
    print(json.dumps({'cache_changes':{k:len(v) for k,v in changes.items()},'memo_restoration':memos,
                      'workbook_sha256':sha256(a.output.read_bytes()).hexdigest()},ensure_ascii=False),flush=True)
