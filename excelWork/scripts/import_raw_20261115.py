"""Merge the corrected export and carry purpose balances to the new cash account.

Previous bank-confirmed corrections win over unchanged old export values.
All lookup/link fields use transaction IDs. Memos are synchronized as text only.
"""
from collections import defaultdict,Counter
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json
import re

import openpyxl

from redesign_reserve import ExtendedPackage
from raw_sort_id_20261101 import reorder,current
from update_raw_20261104 import classify
from update_raw_20261101 import SHARED,WRAP
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore
from update_20261015 import Q

NAVER='네이버미래에셋'
CASH={WRAP,NAVER}
NEW_CLASS_CODES={'What Brandon Thinks':('고정비',26),
                 '롯데프리미엄아울렛 이천점(1/3)':('고정비',26)}


def source_rows(path):
    w=openpyxl.load_workbook(path,data_only=True);s=w.active
    assert tuple(next(s.values))==('날짜','자산','분류','소분류','내용','금액(원)','수입/지출','태그','메모','금액','화폐')
    out=[]
    for sr,v in enumerate(s.iter_rows(min_row=2,values_only=True),2):
        if not v[0]:continue
        assert isinstance(v[0],datetime) and v[0].year==2026 and v[0].month in [10,11]
        assert v[10]=='KRW' and v[5]==v[9]
        out.append(dict(source_row=sr,date=v[0],asset='모임통장' if v[1]=='20260606' else v[1],
                        category=v[2],sub=v[3] or None,title=v[4],amount=v[5],kind=v[6],tag=v[7] or None,memo=v[8] or None))
    w.close();return out


def identity(v):return v['date'],v['asset'],v['amount'],v['kind']


def plan(baseline,source):
    w=openpyxl.load_workbook(baseline,data_only=True);s=w['RAW']
    physical={int(s.cell(r,42).value):r for r in range(2,4002)}
    old={};exact=defaultdict(list)
    for ident,r in physical.items():
        when=s.cell(r,1).value
        if not isinstance(when,datetime):continue
        v=dict(date=when,asset=s.cell(r,3).value,category=s.cell(r,4).value,sub=s.cell(r,5).value,
               title=s.cell(r,6).value,amount=s.cell(r,7).value,kind=s.cell(r,8).value,memo=s.cell(r,9).value)
        old[ident]=v;exact[identity(v)].append(ident)
    assert len(physical)==4000 and set(physical)==set(range(2,4002))
    rows=source_rows(source);new=[];links=[];patches=[];used=set();memo_edits=[]
    for v in rows:
        # These are unchanged source values superseded by statement evidence.
        special=None
        if v['asset']==WRAP and v['title']=='미래에셋 WRAP 이자' and v['amount']==12802:special=(2667,'previous_cumulative_interest_preserved')
        if v['title']=='G마켓(1/6)':special=(2658,'statement_amount_and_date_preserved')
        if v['title']=='롯데카드 띵샵(7/36)':special=(2616,'statement_amount_preserved')
        if v['date'].date()==datetime(2026,11,9).date() and v['date'].time().isoformat()=='21:50:00.890000' and v['title']=='기타':special=(2652,'statement_amount_preserved')
        if v['asset']==WRAP and v['date'].day==12 and v['date'].month==10 and v['title']=='미래에셋 WRAP 이자':special=(2442,'cumulative_interest_split_635')
        if special:
            ident,status=special;assert ident in old and ident not in used;used.add(ident)
            links.append(dict(source_row=v['source_row'],transaction_id=ident,status=status));continue
        candidates=[i for i in exact.get(identity(v),[]) if i not in used]
        status='exact'
        if not candidates:
            same=[i for i,o in old.items() if i not in used and o['date'].date()==v['date'].date() and o['asset']==v['asset'] and o['title']==v['title'] and o['kind']==v['kind']]
            near=[i for i in same if abs((old[i]['date']-v['date']).total_seconds())<2]
            if near:candidates=sorted(near,key=lambda i:abs((old[i]['date']-v['date']).total_seconds()))[:1];status='timestamp_or_amount_update'
            elif len(same)==1:candidates=same;status='same_day_update'
        if not candidates:
            # Two source card assignments were explicitly corrected by the user.
            cross=[i for i,o in old.items() if i not in used and i in [2398,2366] and o['title']==v['title'] and o['amount']==v['amount'] and abs((o['date']-v['date']).total_seconds())<2]
            if cross:candidates=cross;status='source_card_assignment_update'
        assert len(candidates)<=1,('Ambiguous identity',v,candidates)
        if candidates:
            ident=candidates[0];used.add(ident);o=old[ident]
            changes={}
            if o['amount']!=v['amount']:
                rr=physical[ident];burden=s.cell(rr,11).value
                assert burden==o['amount'],('Partial burden needs a decision',ident,burden)
                changes.update(G=v['amount'],K=v['amount'])
            if status=='source_card_assignment_update':changes['C']=v['asset']
            if ident in [2502,2615] and o['date']!=v['date']:changes['A']=v['date']
            # Preserve confirmed classification overrides; copy the actual memo,
            # including an intentional blank, without using it as a criterion.
            if (o['memo'] or None)!=(v['memo'] or None):changes['I']=v['memo'];memo_edits.append(ident)
            if changes:patches.append(dict(transaction_id=ident,changes=changes,old_amount=o['amount']))
            links.append(dict(source_row=v['source_row'],transaction_id=ident,status=status))
        else:new.append(dict(v))
    assert len({identity(v) for v in new})==len(new)
    empty=sorted(i for i,r in physical.items() if not s.cell(r,1).value)
    new.sort(key=lambda v:(v['date'],v['source_row']))
    for ident,v in zip(empty,new):
        v.update(transaction_id=ident,input_row=physical[ident]);links.append(dict(source_row=v['source_row'],transaction_id=ident,status='added'))
    assert len(empty)>len(new)
    # The numeric fee component has its own explicit expense record and no
    # dependency on the source memo after this reconciliation.
    fee_date=datetime(2026,11,12,20,54,0,146000)
    fee_hits=[i for i,o in old.items() if o['title']=='현대네이버카드 지연금' and o['date']==fee_date and o['amount']==157]
    derived=[]
    if not fee_hits:
        ident=empty[len(new)]
        derived=[dict(date=fee_date,asset='카카오뱅크',category='금융',sub=None,title='현대네이버카드 지연금',amount=157,kind='지출',memo=None,tag=None,
                      transaction_id=ident,input_row=physical[ident],source_row=None,classification='변동비',
                      evidence='Source settlement 312654 minus funded principal 312497; source separately identifies the late charge.')]
    # 635 was already included in the earlier 8,557 cumulative entry. Split
    # it into its own source record without changing total interest or cash.
    split_new=[v for v in new if v['date'].month==10 and v['date'].day==8 and v['asset']==WRAP and v['amount']==635 and v['kind']=='수입']
    if split_new:
        assert len(split_new)==1 and old[2442]['amount']==8557
        patches.append(dict(transaction_id=2442,changes=dict(G=7922,K=7922,I='10/12 누계 8,557원 중 10/8 예탁금이용료 635원을 별도 원장으로 분리한 잔여 금액.'),old_amount=8557))
    assert not any(v['title']=='정산용' and v['amount']==-1960 for v in rows)
    w.close()
    return dict(source_count=len(rows),source_by_month=dict(Counter(v['date'].month for v in rows)),
                old_count=len(old),new=new,derived=derived,patches=patches,source_mapping=links,
                memo_edits=memo_edits,source_tags=[dict(source_row=v['source_row'],tag=v['tag']) for v in rows if v['tag']],
                retained_old_ids=[i for i,o in old.items() if o['date'].month in [10,11] and i not in used],
                managed_accounts=sorted(CASH),migration_user_confirmed=True),physical


def cash(ref,when=None):
    naver=f'{ref}="{NAVER}"'
    if when:naver=f'AND({naver},{when}>=DATE(2026,11,12))'
    return f'OR({ref}="{WRAP}",{naver})'


def display_labels(p):
    for addr,value in {'H6':'적립계획','I6':'사용계획','A32':'관리계좌 잔액 합계',
                       'A34':'선택월 관리계좌 내역 · 입금·사용 및 계좌이체'}.items():p.write('잔액',addr,value)
    for addr in ['A36','G36']:
        cell=p.cell('잔액',addr);f=cell.find(Q('f'));v=cell.find(Q('v'))
        assert f is not None
        f.text=f.text.replace('WRAP','관리계좌')
        if v is not None and v.text:v.text=v.text.replace('WRAP','관리계좌')


def extend_accounts(p):
    a,c,k,h=[current(name) for name in ['일시','자산','분류','수입/지출']]
    for r in range(2,4002):
        p.write('RAW',f'R{r}',f'=IF({a}="","",IF(OR({cash(c,a)},{cash(k,a)}),"WRAP",IF(OR(ISNUMBER(SEARCH("카드",{c})),{c}="KB쿠팡와우"),"카드사용","타계좌")))',True)
        p.write('RAW',f'S{r}',f'=IF({current("잔액 관리범위")}<>"WRAP","",IF(AND({cash(c,a)},{cash(k,a)}),"내부이체",IF(OR({h}="수입",{h}="이체입금",{h}="이체수입",{cash(k,a)}),"입금","출금")))',True)
    for r in range(64,501):
        internal=f'AND({cash(f"J{r}",f"A{r}")},{cash(f"K{r}",f"A{r}")},J{r}<>K{r})'
        scope=f'OR(AND(L{r}="원본WRAP",A{r}<DATE(2026,10,1)),{cash(f"J{r}",f"A{r}")},{cash(f"K{r}",f"A{r}")},AND(COUNTIF($A$7:$A$15,C{r})>0,COUNTIF($A$7:$A$15,D{r})>0,C{r}<>D{r}))'
        p.write('잔액',f'I{r}',f'=IF(COUNTA(A{r}:E{r})=0,"",IF(NOT({scope}),"관리계좌 범위 확인",IF(OR(NOT(ISNUMBER(A{r})),B{r}="",C{r}="",D{r}="",NOT(ISNUMBER(E{r})),E{r}<=0,AND(C{r}=D{r},NOT({internal}))),"입력 확인","OK")))',True)
    # Family principal return may arrive at either of the managed accounts.
    for cell in p.sheet('_예비비계산').iter(Q('c')):
        f=cell.find(Q('f'))
        if f is not None and f.text:
            f.text=re.sub(r'INDEX\(잔액!\$K:\$K,BA\d+\)="미래에셋_WRAP형CMA"',lambda m:cash(m.group(0).split('=')[0]),f.text)
    replacements={'A1':'잔액 | 관리계좌 안의 목적별 자금','A4':'관리계좌: 네이버미래에셋 + 기존 WRAP. 두 계좌 간 이동은 총잔액을 바꾸지 않으며 타은행으로 보낸 돈은 송금 때 차감합니다.',
                  'G6':'조회월 관리계좌 잔액','A16':'관리계좌 잔액 합계','A17':'관리계좌 목적자금','A20':'관리계좌의 계정별 월말 잔액',
                  'A31':'네이버미래에셋·기존 WRAP만 합산합니다. 다른 통장의 잔액과 공동계좌 보관금은 포함하지 않습니다.',
                  'A33':'목적계정 사용가능액 + 결제대금비 등 모든 계정 = 관리계좌 총액입니다. 관리계좌 사이 송금은 소비가 아닙니다.',
                  'A62':'선택월 관리계좌 원장 | 조회월 자동 연동 · 전체 입력자료는 왼쪽 +로 펼침'}
    for addr,text in replacements.items():p.write('잔액',addr,text)
    display_labels(p)
    nav=f'SUMIFS($Y$64:$Y$500,$K$64:$K$500,"{NAVER}",$A$64:$A$500,"<"&EDATE($B$3,1))-SUMIFS($Y$64:$Y$500,$J$64:$J$500,"{NAVER}",$A$64:$A$500,"<"&EDATE($B$3,1))'
    p.write('잔액','D3',f'="조회월 기준 · 네이버미래에셋 "&TEXT({nav},"#,##0")&"원 / 기존 WRAP "&TEXT(G16-({nav}),"#,##0")&"원"',True)


def prepare(baseline,source,output):
    meta,ids=plan(baseline,source);p=ExtendedPackage(baseline)
    w=openpyxl.load_workbook(baseline,data_only=True)
    existing_links={w['잔액'][f'W{r}'].value:r for r in range(64,501) if w['잔액'][f'W{r}'].value}
    ledger_row=max(r for r in range(64,501) if w['잔액'][f'A{r}'].value)+1
    for patch in meta['patches']:
        ident=patch['transaction_id']
        for c,value in patch['changes'].items():p.write('RAW',f'{c}{ids[ident]}',value)
        if ident in existing_links and 'G' in patch['changes']:p.write('잔액',f'E{existing_links[ident]}',patch['changes']['G'])
    bank_links=[]
    for v in [*meta['new'],*meta['derived']]:
        r=v['input_row']
        for c,k in [('A','date'),('C','asset'),('D','category'),('E','sub'),('F','title'),('G','amount'),('H','kind'),('I','memo')]:p.write('RAW',f'{c}{r}',v[k])
        cl=v.get('classification') or classify(v)
        if v['title'] in NEW_CLASS_CODES:cl,code=NEW_CLASS_CODES[v['title']];p.write('RAW',f'AQ{r}',code)
        if cl:p.write('RAW',f'J{r}',cl)
        if v['asset'] in SHARED:p.write('RAW',f'K{r}',0);p.write('RAW',f'L{r}','공동')
        src,dst=v['asset'],v['category'];inside=src in CASH and dst in CASH
        if src not in CASH and dst not in CASH:continue
        incoming=src not in CASH or v['kind']=='수입'
        transfer=v['kind'].startswith('이체')
        sender,recipient=(src,dst) if transfer else (('외부',src) if incoming else (src,'외부'))
        inputs=dict(A=v['date'],B='계좌이체' if transfer else ('입금' if incoming else '출금'),
                    C='외부' if incoming else '생활비',D='생활비' if incoming or inside else ('외부' if transfer else '지출'),
                    E=v['amount'],F=f'{src} → {dst}' if transfer else v['title'],J=sender,K=recipient,
                    L='WRAP계좌이체' if transfer else 'WRAP입출금',W=v['transaction_id'])
        assert inputs['E']>0 and ledger_row<=500
        for c,value in inputs.items():p.write('잔액',f'{c}{ledger_row}',value)
        bank_links.append(dict(ledger_row=ledger_row,transaction_id=v['transaction_id'],internal=inside,
                               direction=0 if inside else (1 if incoming else -1),amount=v['amount']))
        ledger_row+=1
    extend_accounts(p)
    p.write('10월','A70','11/15 수정 원장 반영 · 우리카드 할인 16,974원. 정산용 보정 없이 카드 청구합계와 대조했습니다.')
    p.write('11월','A70','11/15까지 반영 · 네이버미래에셋으로 관리계좌 이전. KTX 취소는 0원, 이자는 직전 확정 누계 보존, 이체·공금·중복 기록은 소비에 재가산하지 않습니다.')
    p.write('사용안내','A14','잔액·예비비는 네이버미래에셋과 기존 WRAP을 함께 관리합니다. 계좌 간 이동은 지출이 아니며 기존 목적별 배분을 유지합니다.')
    p.write('사용안내','A26','11/15 갱신 · 수정 원장과 취소·환급 반영. 관리계좌 합계/개별 계좌 잔액은 잔액 탭 조회월로 확인합니다. 원본 메모는 RAW I열에서 수정합니다.')
    for sheet in ['잔액','예비비']:p.write(sheet,'B3',datetime(2026,11,1))
    p.save(output);w.close()
    values=openpyxl.load_workbook(output,data_only=True);meta['sort']=reorder(p,values,'descending');values.close();p.save(output)
    values=openpyxl.load_workbook(output,data_only=True)
    for r,rr in enumerate(range(2,12),8):p.write('입력·연결',f'H{r}',values['RAW'].cell(rr,42).value)
    values.close();p.save(output);meta['bank_links']=bank_links;gc.collect();return meta


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for name in ['baseline','source','output','root','metadata']:a.add_argument(name,type=Path)
    a.add_argument('--plan-only',action='store_true');args=a.parse_args()
    if args.plan_only:
        meta,_=plan(args.baseline,args.source)
    else:
        meta=prepare(args.baseline,args.source,args.root/'candidate.xlsm')
    args.metadata.write_text(json.dumps(meta,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Plan:',len(meta['new']),'new,',len(meta['derived']),'derived,',len(meta['patches']),'updated',flush=True)
    if not args.plan_only:
        calc=recalculate(args.root/'candidate.xlsm',args.root/'calculation')
        import_changed_caches(args.root/'candidate.xlsm',calc,args.root/'cached.xlsm')
        print(restore(args.baseline,args.root/'cached.xlsm',args.output),flush=True)
