"""Confirm cumulative interest and reconcile November statement without duplication.

Keep the user's grouped purchases, classifications, and memos. Explicit source
groups also distinguish transport billing, bank-side funding, and point tender.
"""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json

import openpyxl

from redesign_reserve import ExtendedPackage
from raw_sort_id_20261101 import reorder
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore

# Actual bank-export worksheet rows -> existing immutable transaction IDs.
# A group can span several source transactions or several user-split records.
GROUPS=[([2],[2670]),([3],[2671]),([4],[2672]),([5],[2668]),([6],[2669]),
        ([7,12],[2662]),([8],[2662]),([9],[2665,2666]),([10],[2664]),([11],[2663]),
        ([15],[2660]),([16],[2659]),([17],[2657]),([18],[2658]),([19],[2655,2656]),
        ([20],[2654]),([21],[2653]),([22],[2650]),([23],[2651]),([24],[2652]),
        ([25],[2649]),([26],[2648]),([27],[2647]),([28],[2646]),([29],[2645]),
        ([30],[2644]),([31],[2643]),([32],[2640,2641]),([33],[2635]),([34],[2639]),
        ([35],[2638]),([36],[2637]),([37],[2631]),([38],[2630]),([39],[2629]),
        ([40,41],[2628]),([42],[2627]),([43],[2626]),([44],[2622]),([45],[2625]),
        ([46],[2623,2624]),([47],[2616]),([48],[2611]),([49],[2615]),([50],[2613]),
        ([51],[2614]),([52],[2612]),([53],[2610]),([54],[2609]),([55],[2608]),
        ([56],[2606]),([57],[2607]),([58],[2604]),([59],[2603,2605]),([60],[2602]),
        ([61],[2601]),([62],[2599]),([63],[2598]),([64],[2600]),([65],[2597]),
        ([66],[2596]),([67],[2595]),([68],[2594]),([69],[2593]),([71],[2592]),([75],[2591])]
AMOUNT_CORRECTIONS={2652:(5000,5200),2658:(54985,54983),2616:(10200,10250)}
INTEREST_ID=2667
INTEREST_DATE=datetime(2026,11,11,17,5,22)
INTEREST_AMOUNT=13964


def read_actual(path):
    w=openpyxl.load_workbook(path,data_only=True)
    s=w['가계부 내역']
    assert tuple(next(s.values))==('날짜','시간','타입','대분류','소분류','내용','금액','화폐','결제수단','메모')
    rows={i:r for i,r in enumerate(s.values,1) if i>1 and isinstance(r[0],datetime) and r[0].month==11}
    assert set(rows)==set(range(2,76)) and all(r[7]=='KRW' for r in rows.values())
    assert rows[13][5]=='정진규' and rows[14][5]=='마이신한포인트'
    assert rows[70][6]==-4597 and rows[70][8]=='네이버페이 간편결제(포인트)'
    w.close();return rows


def plan(baseline,source):
    actual=read_actual(source)
    w=openpyxl.load_workbook(baseline,data_only=True);s=w['RAW']
    ids={int(s.cell(r,42).value):r for r in range(2,4002)}
    assert len(ids)==4000 and set(ids)==set(range(2,4002))
    records={i:tuple(s.cell(r,c).value for c in range(1,44)) for i,r in ids.items() if isinstance(s.cell(r,1).value,datetime)}
    corrections=[]
    for ident,(old,new) in AMOUNT_CORRECTIONS.items():
        r=records[ident];assert r[6] in [old,new] and r[10]==r[6]
        if r[6]!=new:corrections.append(dict(transaction_id=ident,old_amount=r[6],new_amount=new))
    coverage=[];covered=set()
    for source_rows,ref_ids in GROUPS:
        assert not covered.intersection(source_rows);covered.update(source_rows)
        amount=sum(abs(actual[r][6]) for r in source_rows)
        expected=sum(AMOUNT_CORRECTIONS[i][1] if i in AMOUNT_CORRECTIONS else records[i][6] for i in ref_ids)
        assert amount==expected,(source_rows,ref_ids,amount,expected)
        if source_rows!=[18]:
            assert {actual[r][0].date() for r in source_rows}=={records[i][0].date() for i in ref_ids}
        status='linked_existing'
        if source_rows in [[8],[22]]:status='funding_counterpart_already_linked'
        if source_rows==[7,12]:status='transport_bill_components_already_linked'
        coverage.append(dict(source_rows=source_rows,transaction_ids=ref_ids,status=status,source_amount=amount))
    new=[]
    for sr in [13,14,72,73,74]:
        v=actual[sr];date=datetime.combine(v[0].date(),v[1])
        if sr==13:
            # Same recurring sender and account family as Sep/Oct travel dues.
            item=dict(date=date,asset='여행통장',category='공금',sub=None,title=v[5],amount=v[6],kind='수입',memo=v[9],classification='공금',burden=0,payment='공동')
        elif sr==14:
            item=dict(date=date,asset='신한은행',category='금융',sub=None,title=v[5],amount=v[6],kind='수입',memo=v[9],classification='이자·환급',burden=v[6],payment='본인')
        else:
            item=dict(date=date,asset='토스뱅크K패스체크카드',category='교통/차량',sub=None,title=v[5],amount=-v[6],kind='지출',memo=v[9],classification='후불교통 사용기록',burden=-v[6],payment='본인')
        hits=[i for i,r in records.items() if (r[0],r[2],r[5],r[6],r[7])==(item['date'],item['asset'],item['title'],item['amount'],item['kind'])]
        assert len(hits)<=1
        if hits:
            coverage.append(dict(source_rows=[sr],transaction_ids=hits,status='already_imported',source_amount=abs(v[6])))
        else:new.append(dict(item,source_row=sr))
        covered.add(sr)
    empty=sorted(i for i,r in ids.items() if not s.cell(r,1).value)
    new.sort(key=lambda v:(v['date'],v['source_row']))
    for item,ident in zip(new,empty):
        item.update(transaction_id=ident,input_row=ids[ident])
        coverage.append(dict(source_rows=[item['source_row']],transaction_ids=[ident],status='added',source_amount=item['amount']))
    assert len(empty)>=len(new)
    covered.add(70);coverage.append(dict(source_rows=[70],transaction_ids=[2592],status='point_tender_archived_no_additional_cash_expense',source_amount=4597))
    assert covered==set(actual)
    current=records[INTEREST_ID]
    assert current[6] in [12802,INTEREST_AMOUNT] and current[10]==current[6]
    meta=dict(source_rows=74,new=new,corrections=corrections,coverage=coverage,
              old_raw_count=len(records),interest_transaction_id=INTEREST_ID,
              interest_old_amount=current[6],interest_new_amount=INTEREST_AMOUNT,
              interest_changed=current[6]!=INTEREST_AMOUNT or current[0]!=INTEREST_DATE,
              interest_recorded_once_on_entry_date=True,
              date_correction=dict(transaction_id=2658,source_date='2026-11-09',old_date=records[2658][0]),
              shared_travel_income_assumption='Nov10 same recurring sender/account pattern as Sep10/Oct7 travel dues; excluded from personal income.',
              october_other_bank_income_2120_not_imported_scope='Current request covers November.')
    w.close();return meta,ids


def prepare(baseline,source,output):
    meta,ids=plan(baseline,source);actual=read_actual(source)
    p=ExtendedPackage(baseline)
    for v in meta['corrections']:
        r=ids[v['transaction_id']]
        for c in ['G','K']:p.write('RAW',f'{c}{r}',v['new_amount'])
    v=actual[18];p.write('RAW',f'A{ids[2658]}',datetime.combine(v[0].date(),v[1]))
    r=ids[INTEREST_ID]
    for c,v in dict(A=INTEREST_DATE,F='미래에셋 WRAP 이자수익',G=INTEREST_AMOUNT,K=INTEREST_AMOUNT,
                    I='11/11 가계부 작성 시 누계 입력. 10/13 정산 이후 11/11까지 세금·수수료 차감 후 CMA MMW 운용수익. 일별로 중복 입력하지 않음.').items():
        p.write('RAW',f'{c}{r}',v)
    for item in meta['new']:
        r=item['input_row']
        for c,k in [('A','date'),('C','asset'),('D','category'),('E','sub'),('F','title'),('G','amount'),('H','kind'),('I','memo'),('J','classification'),('K','burden'),('L','payment')]:
            p.write('RAW',f'{c}{r}',item[k])
    for c,v in dict(A=INTEREST_DATE,E=INTEREST_AMOUNT,F='미래에셋 WRAP 이자수익 · 11/11까지 누계').items():p.write('잔액',f'{c}261',v)
    for name in ['잔액','예비비']:p.write(name,'B3',datetime(2026,11,1))
    p.write('잔액','D3','11/11 거래내역증명서 대조 · WRAP 8,288,985원. 이자는 가계부 작성일에 세후 누계로 1회 입력합니다.')
    p.write('11월','A70','11/11까지 갱신 · WRAP 이자수익은 작성일 누계 입력. 실제 원장과 기존 거래를 연결하고 누락 수입·금액 차이를 반영했습니다.')
    p.write('사용안내','A26','11/11 거래내역증명서 확인 · WRAP 8,288,985원. 이자수익은 매일 입력하지 않고 가계부 작성일에 직전 정산 이후 누계로 반영합니다.')
    p.save(output)
    w=openpyxl.load_workbook(output,data_only=True);meta['sort']=reorder(p,w,'descending');w.close();p.save(output)
    w=openpyxl.load_workbook(output,data_only=True)
    for r,source_row in enumerate(range(2,12),8):p.write('입력·연결',f'H{r}',w['RAW'].cell(source_row,42).value)
    w.close();p.save(output);gc.collect();return meta


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for name in ['baseline','source','output','root','metadata']:parser.add_argument(name,type=Path)
    a=parser.parse_args();candidate=a.root/'candidate.xlsm'
    meta=prepare(a.baseline,a.source,candidate)
    a.metadata.write_text(json.dumps(meta,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Prepared',len(meta['new']),'new,',len(meta['corrections']),'amount corrections; interest',INTEREST_AMOUNT,flush=True)
    calculated=recalculate(candidate,a.root/'calculation')
    import_changed_caches(candidate,calculated,a.root/'cached.xlsm')
    print(restore(a.baseline,a.root/'cached.xlsm',a.output),flush=True)
