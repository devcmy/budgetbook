"""After-sort input/memo/duplicate tests, all on disposable copies only."""
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json

import openpyxl
from raw_sort_id_20261101 import END,reorder
from redesign_reserve import ExtendedPackage
from verify_reserve_redesign import recalculate,assert_no_errors
from verify_reserve_cash_alignment import verify_month
from verify_raw_sort_20261101 import equal,memo_checks

AMOUNT = 34003
TETRIS = 36000
NOTES = {2474:'정렬 후 더치 메모 수정 시험',2475:'정렬 후 WRAP 메모 수정 시험',356:'정렬 후 예비비 메모 수정 시험'}


def core_snapshot(w):
    result={}
    for m in range(1,13):
        for sn,addrs in [(f'{m}월',['I33','H33','H37','I37','H38','I38','H44','I44']),
                         (f'{m}월계산기',['W12'])]:
            for a in addrs:result[sn,a]=w[sn][a].value
    for sn,addrs in [('잔액',['G7','G8','G9','G10','G11','G12','G16']),('예비비',['F13','I13'])]:
        for a in addrs: result[sn,a]=w[sn][a].value
    return result


def main(candidate,work):
    values=openpyxl.load_workbook(candidate,data_only=True)
    before=core_snapshot(values)
    p=ExtendedPackage(candidate)
    byid={values['RAW'].cell(r,42).value:r for r in range(2,END+1)}
    for ident,memo in NOTES.items():p.write('RAW',f'I{byid[ident]}',memo)
    # Exact transaction duplicate, with a different memo. Identity never reads I.
    old=byid[2474];duplicate=byid[2590]
    for col in [1,*range(3,14)]:
        original=values['RAW'].cell(old,col).value
        from openpyxl.utils import get_column_letter
        p.write('RAW',f'{get_column_letter(col)}{duplicate}',original)
    p.write('RAW',f'I{duplicate}','중복본 메모는 별개지만 거래는 중복')
    for ident,day,amount,title in [(2591,1,12001,'정렬 후 새 거래 시험'),
                                   (2592,2,22002,'긴 거래키 중복 시험 '+'가'*320),
                                   (2593,2,22002,'긴 거래키 중복 시험 '+'가'*320)]:
        r=byid[ident]
        for col,value in {'A':datetime(2026,11,day,12),'C':'현대카드','D':'식비','E':None,
                          'F':title,'G':amount,'H':'지출','I':f'새 입력 메모 {ident}','M':None}.items():
            p.write('RAW',f'{col}{r}',value)
    # Put higher IDs before lower ones: duplicated transactions are encountered
    # before their original, proving that Q does not depend on display order.
    reorder(p,values,'reverse_ids');p.save(work/'edited.xlsm')
    values.close();del values,p;gc.collect()
    calculated=recalculate(work/'edited.xlsm',work/'edited_calc')
    w=openpyxl.load_workbook(calculated,data_only=True);assert_no_errors(w)
    raw=w['RAW'];rows={raw.cell(r,42).value:r for r in range(2,END+1)}
    for ident,rank in {2474:1,2590:2,2591:1,2592:1,2593:2}.items():
        equal(raw.cell(rows[ident],17).value,rank,('duplicate rank',ident))
    assert len(raw.cell(rows[2592],16).value)>250
    for ident in [2590,2593]:equal(raw.cell(rows[ident],25).value,0,('duplicate excluded',ident))
    equal(w['11월']['I33'].value,AMOUNT,'new month actual')
    equal(w['11월계산기']['W12'].value,TETRIS,'new month calculator')
    after=core_snapshot(w)
    for key,value in before.items():
        if key[0].startswith('11월'):continue
        equal(after[key],value,('unchanged core',*key))
    equal(w['10월']['I53'].value,NOTES[2474],'extra memo follows ID')
    equal(w['잔액']['G248'].value,NOTES[2475],'ledger memo follows ID')
    equal(w['예비비']['H102'].value,NOTES[356],'reserve memo follows ID')
    report={'new_month_actual':AMOUNT,'new_month_tetris':TETRIS,
            'short_and_long_key_duplicates_excluded':True,'duplicate_keeper_independent_of_display_order':True,
            'linked_memo_edits_after_sort':True,'edited_reserve':verify_month(w,10),
            'edited_memos':memo_checks(w)}
    p=ExtendedPackage(work/'edited.xlsm')
    for ident in NOTES:p.write('RAW',f'I{rows[ident]}',None)
    p.save(work/'cleared.xlsm');w.close();del w,p;gc.collect()
    calculated2=recalculate(work/'cleared.xlsm',work/'cleared_calc')
    w=openpyxl.load_workbook(calculated2,data_only=True);assert_no_errors(w)
    for sn,addr in [('10월','I53'),('잔액','G248'),('예비비','H102')]:
        equal(w[sn][addr].value,None,('cleared source memo disappears',sn,addr))
    for key,value in after.items():equal(core_snapshot(w)[key],value,('memo clear financial invariant',*key))
    report['cleared_memos']=memo_checks(w)
    report['memo_delete_after_sort_updates_month_ledger_reserve_without_financial_change']=True
    w.close();return report


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('candidate',type=Path);parser.add_argument('work',type=Path);parser.add_argument('report',type=Path)
    args=parser.parse_args();args.work.mkdir(exist_ok=False)
    report=main(args.candidate,args.work)
    args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print('After-sort edits, new inputs, duplicates and memo deletion passed',flush=True)
