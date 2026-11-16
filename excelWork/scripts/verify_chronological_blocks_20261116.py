"""Verify chronological placement, RAW sort/date edits and internal funding."""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import gc
import json

import fitz
import openpyxl
from lxml import etree as ET
from openpyxl.utils.datetime import from_excel

from chronological_blocks_20261116 import expected_blocks,slots,allocate_october,BACK,END
from redesign_reserve import ExtendedPackage
from raw_sort_id_20261101 import reorder
from verify_reserve_redesign import recalculate,assert_no_errors
from verify_reserve_cash_alignment import verify_month
from verify_raw_sort_20261101 import equal,memo_checks
from verify_raw_20261116 import render_colors
from balance_selected_ledger import check_view
from validate_excel_structure import validate_package
from update_20261015 import Q


def chronology(w,month,stored=False):
    records,sequence=expected_blocks(w,month)
    s=w[f'{month}월계산기'];back=w[BACK];base=56+(month-11)*14
    colors={s[f'B{r}'].fill.fgColor.rgb:r for r in range(3,11)}
    for i,record in enumerate(records,2):
        dt=back.cell(i,base+6).value
        if isinstance(dt,(int,float)):dt=from_excel(dt)
        equal(dt,record['date'],('chronological date',month,i))
        assert back.cell(i,base+7).value==record['id'],('chronological ID',month,i)
        assert back.cell(i,base+8).value==record['palette']
        assert back.cell(i,base+9).value==record['signed']
        assert back.cell(i,base+11).value==record['keep'],('refund units',month,record)
    assert back.cell(1,base+5).value==len(records)
    assert back.cell(1,base+13).value==len(sequence)
    for i,(r,c) in enumerate(slots()):
        want=sequence[i]['palette'] if i<len(sequence) else 0
        assert s.cell(r,c+31).value==want,('automatic chronological color',month,i+1,want,s.cell(r,c+31).value)
        if stored:
            assert colors.get(s.cell(r,c).fill.fgColor.rgb,0)==want,('stored chronological color',month,i+1)
    counts=Counter(x['palette'] for x in sequence)
    assert all(s[f'C{r}'].value==counts[r] for r in range(3,11))
    assert s['W12'].value==len(sequence)*1000==w[f'{month}월']['I38'].value
    assert [(x['date'],x['id']) for x in sequence]==sorted((x['date'],x['id']) for x in sequence)
    return dict(transaction_count=len(records),colored_cells=len(sequence),category_counts=dict(counts),
                first_transactions=records[:8],all_block_dates_ascending=True,
                raw_sort_independent=True,stored_color_sequence_verified=stored)


def rendered_order(path,folder,month):
    # The common renderer clears all base fills; rules must supply the colors.
    result=render_colors(path,folder,month)
    w=openpyxl.load_workbook(path,data_only=True);s=w[f'{month}월계산기']
    palette={s[f'B{r}'].fill.fgColor.rgb[-6:]:r for r in range(3,11)}
    doc=fitz.open(folder/'grid.pdf');d=doc[0].get_drawings()
    filled=[x for x in d if x['fill'] is not None]
    bounds=(min(x['rect'].x0 for x in filled),min(x['rect'].y0 for x in filled),
            max(x['rect'].x1 for x in filled),max(x['rect'].y1 for x in filled))
    x0,y0,x1,y1=bounds;cw=(x1-x0)/20;rh=(y1-y0)/130
    color_rectangles=[]
    for drawing in filled:
        rgb=''.join(f'{round(v*255):02X}' for v in drawing['fill'])
        if rgb in palette:color_rectangles.append((drawing['rect'],palette[rgb]))
    checked=0
    for r in range(15,145):
        for c in range(2,22):
            pt=fitz.Point(x0+(c-2+.5)*cw,y0+(r-15+.5)*rh)
            actual=next((pal for rect,pal in reversed(color_rectangles) if pt in rect),0)
            assert actual==s.cell(r,c+31).value,('rendered placement',month,r,c,actual,s.cell(r,c+31).value)
            checked+=1
    doc.close();w.close()
    result['individual_positions_verified']=checked
    return result


def verify(baseline,final,metadata):
    a=openpyxl.load_workbook(baseline,data_only=True);w=openpyxl.load_workbook(final,data_only=True)
    assert_no_errors(w)
    assert [(s.title,s.sheet_state) for s in a]==[(s.title,s.sheet_state) for s in w]
    assert w['잔액']['B3'].value==w['예비비']['B3'].value==datetime(2026,10,1)
    for r in range(2,4002):
        for c in range(1,44):equal(a['RAW'].cell(r,c).value,w['RAW'].cell(r,c).value,('RAW preserved',r,c))
    for month in range(1,13):
        for addr in ['C37','C38','C39','C40','C41','I33','I38','H37','E74','E75','E78','E79','H43']:
            equal(a[f'{month}월'][addr].value,w[f'{month}월'][addr].value,(month,addr))
        for addr in [*[f'C{r}' for r in range(3,11)],'W12','W13']:
            equal(a[f'{month}월계산기'][addr].value,w[f'{month}월계산기'][addr].value,('calculator totals',month,addr))
    l=w['잔액'];assert l['G16'].value==8434741 and l['G7'].value==6084741
    assert l['G17'].value==w['예비비']['I13'].value==2350000
    for r,amount,balance in [(9,100000,1500000),(10,50000,500000),(11,50000,350000)]:
        assert l[f'C{r}'].value==amount and l[f'G{r}'].value==balance
        assert l[f'J{r}'].value==0,('October plan fully funded',r)
    meta=json.loads(metadata.read_text());assert sum(x['amount'] for x in meta['allocations'])==200000
    assert len(meta['allocations'])==3
    for item in meta['allocations']:
        r=item['row']
        assert l[f'B{r}'].value=='적립' and l[f'C{r}'].value=='생활비'
        assert l[f'I{r}'].value=='OK' and l[f'Y{r}'].value==item['amount']
        assert l[f'J{r}'].value==l[f'K{r}'].value=='미래에셋_WRAP형CMA'
        assert l[f'W{r}'].value is None and l[f'X{r}'].value is None
    reserves=verify_month(w,10);view=check_view(w,10);memos=memo_checks(w)
    calculators={str(m):chronology(w,m,stored=True) for m in [11,12]}
    assert calculators['11']['first_transactions'][0]['id']==2593
    assert [x['palette'] for x in expected_blocks(w,11)[1][:13]]==[9]*3+[7]*5+[4]*5
    p,q=ExtendedPackage(baseline),ExtendedPackage(final)
    assert p.parts.keys()==q.parts.keys()
    protected=[n for n in p.parts if n.endswith(('.bin','.rels')) or n.startswith(('xl/drawings/','xl/media/','xl/ctrlProps/','xl/activeX/'))]
    assert all(p.parts[n]==q.parts[n] for n in protected)
    old_styles=p.tree('xl/styles.xml');new_styles=q.tree('xl/styles.xml')
    for node in old_styles:
        other=new_styles.find(node.tag)
        if node.tag==Q('cellXfs'):
            assert len(other)>=len(node)
            assert all(ET.tostring(x)==ET.tostring(other[i]) for i,x in enumerate(node))
        else:assert ET.tostring(node)==ET.tostring(other),node.tag
    for month in range(1,11):
        name=f'{month}월계산기'
        assert p.parts[p.sheets[name]]==q.parts[q.sheets[name]],('older calculator preserved',month)
    structure=validate_package(final);assert not structure['errors'],structure
    assert not allocate_october(ExtendedPackage(final),final),'Repeat allocation must add nothing'
    result=dict(as_of='2026-11-16',workbook_sha256=sha256(final.read_bytes()).hexdigest(),
                current_verification='chronological_blocks_20261116_validation.json',
                calculators=calculators,october=dict(cma_total=8434741,living=6084741,purpose_available=2350000,
                                                   reserve=reserves,ledger=view,allocations=meta['allocations']),
                all_raw_values_preserved=True,all_monthly_financial_totals_preserved=True,
                older_calculators_byte_identical=True,vba_controls_relationships_preserved=True,
                original_styles_preserved=True,structure=structure,formula_errors=0,monthly_memos=memos,
                repeat_allocation_new_records=0,native_excel_ui_executed=False,
                previous_commit='4996ccc',previous_version='history/2026-11-16/02')
    w.close();a.close();gc.collect()
    print('Financial preservation, chronological cells and October allocation passed.',flush=True)
    return result


def scenario(final,root,report):
    p=ExtendedPackage(final);w=openpyxl.load_workbook(final,data_only=True)
    reorder(p,w,'ascending');w.close();p.save(root/'sorted.xlsm')
    w=openpyxl.load_workbook(root/'sorted.xlsm',data_only=True)
    ids={int(w['RAW'].cell(r,42).value):r for r in range(2,4002)}
    empty=[r for r in range(2,4002) if not w['RAW'].cell(r,1).value]
    def edit(p):
        p.write('RAW',f'A{ids[2738]}',datetime(2026,11,2,10))
        for c in ['A',*list('CDEFGHIJKLMN'),'AQ']:p.write('RAW',f'{c}{ids[2740]}',None)
        for r,day,amount,category in [(empty[0],2,2000,'기타'),(empty[1],1,1000,'식비')]:
            for c,value in dict(A=datetime(2026,12,day,12),C='현대네이버카드',D=category,F='날짜순 검증',G=amount,H='지출',J='변동비').items():p.write('RAW',f'{c}{r}',value)
        for sheet in ['잔액','예비비']:p.write(sheet,'B3',datetime(2026,11,1))
    calc=recalculate(root/'sorted.xlsm',root/'scenario',edit);w.close()
    w=openpyxl.load_workbook(calc,data_only=True);assert_no_errors(w)
    nov=chronology(w,11);dec=chronology(w,12)
    assert w['11월']['I33'].value==1257401 and w['11월계산기']['W12'].value==1278000
    assert w['12월계산기']['W12'].value==3000
    assert dec['first_transactions'][0]['date'].day==1
    assert w['잔액']['G16'].value==5048217 and w['잔액']['G7'].value==2698217
    assert w['잔액']['G17'].value==2350000
    verify_month(w,11);memo_checks(w);w.close()
    report['november_after_sort_date_edit_and_delete']=nov
    report['december_new_entries']=dec
    report['november_balances']=dict(cma_total=5048217,living=2698217,purpose_available=2350000,naver=5047690,wrap=527)
    report['current_render']=rendered_order(final,root/'render_current',11)
    report['edited_render']=rendered_order(calc,root/'render_edited',11)
    report['december_render']=rendered_order(calc,root/'render_december',12)
    report['raw_sort_date_edit_delete_and_next_month_scenario_passed']=True
    print('Sort, date edit, deletion, December, and individual PDF color positions passed.',flush=True)


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for name in ['baseline','final','metadata','root','report']:a.add_argument(name,type=Path)
    args=a.parse_args();args.root.mkdir(exist_ok=True)
    result=verify(args.baseline,args.final,args.metadata)
    scenario(args.final,args.root,result)
    args.report.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
