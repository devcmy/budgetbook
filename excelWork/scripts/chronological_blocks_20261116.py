"""Date/time-ordered calculator blocks and the authorized October allocation.

Color denotes the category; placement follows date, time, then stable ID.
Refunds retract the preceding blocks of the same category. All input scans use
table columns, so native RAW sorting does not change transaction identities.
"""
from collections import Counter
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import argparse
import gc
import json

import openpyxl
from lxml import etree as ET
from openpyxl.utils import get_column_letter as col

from redesign_reserve import ExtendedPackage, FUNDS
from raw_auto_link_20261026 import conditional
from update_20261015 import Q
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore

BACK = '_RAW자동연동'
END = 4011  # 4,000 RAW rows and ten additional-expense differences.


def slots():
    for start, stop in [(15,145),(165,435)]:
        for r0 in range(start,stop,5):
            for c0 in range(2,22,2):
                for c in range(c0,c0+2):
                    for r in range(r0,r0+5):
                        yield r,c


def expected_blocks(w, month):
    records = []
    for values in w['RAW'].iter_rows(min_row=2,max_row=4001,values_only=True):
        dt, units = values[14], values[25]
        if not isinstance(dt,datetime) or dt.year != 2026 or dt.month != month or not units:
            continue
        records.append(dict(date=dt,id=int(values[41]),palette=int(values[23]),signed=int(units)))
    s = w[f'{month}월']
    for r in range(52,62):
        dt, units = s[f'A{r}'].value, s[f'AB{r}'].value
        if isinstance(dt,datetime) and dt.month == month and units:
            records.append(dict(date=dt,id=4002+r-52,palette=int(s[f'AC{r}'].value),signed=int(units)))
    records.sort(key=lambda v:(v['date'],v['id']))
    debt = Counter()
    for i,record in enumerate(records):
        p,n = record['palette'],record['signed']
        record['keep'] = 0
        if n > 0:
            taken = min(n,debt[p]);debt[p] -= taken
            record['keep'] = n-taken
        else:
            remaining = -n
            for previous in reversed(records[:i]):
                if previous['palette'] != p:
                    continue
                taken = min(previous['keep'],remaining)
                previous['keep'] -= taken;remaining -= taken
                if not remaining:
                    break
            debt[p] += remaining
    assert not any(debt.values()), ('Unmatched negative category total',month,debt)
    sequence = [record for record in records for _ in range(record['keep'])]
    assert len(sequence) <= 8000
    return records, sequence


def install_helpers(p,month):
    base = 56+(month-11)*14
    names = [col(base+i) for i in range(14)]
    date,ident,palette,units,rank,offset,dt,uid,pal,signed,running,keep,start,end = names
    rg = lambda c: f'${c}$2:${c}${END}'
    p.write(BACK,f'{offset}1',f'=COUNTIF({rg(date)},">0")',True)
    p.write(BACK,f'{end}1',f'=SUM({rg(keep)})',True)
    for r in range(2,END+1):
        index = r-1
        if index <= 4000:
            eligible = f'AND(INDEX(RawData2026[월],{index})=DATE(2026,{month},1),INDEX(RawData2026[달력칸수],{index})<>0)'
            formula = f'=IF({eligible},INDEX(RawData2026[유효일시],{index}),0)'
            sources = {ident:f'INDEX(RawData2026[거래ID],{index})',
                       palette:f'INDEX(RawData2026[계산기분류],{index})',
                       units:f'INDEX(RawData2026[달력칸수],{index})'}
        else:
            mr = 52+index-4001
            formula = f'=IF(AND(ISNUMBER(\'{month}월\'!A{mr}),\'{month}월\'!A{mr}>=DATE(2026,{month},1),\'{month}월\'!A{mr}<EDATE(DATE(2026,{month},1),1),\'{month}월\'!AB{mr}<>0),\'{month}월\'!A{mr},0)'
            sources = {ident:str(index+1),palette:f"'{month}월'!AC{mr}",units:f"'{month}월'!AB{mr}"}
        p.write(BACK,f'{date}{r}',formula,True)
        for c,source in sources.items():
            p.write(BACK,f'{c}{r}',f'=IF({date}{r}=0,0,{source})',True)
        p.write(BACK,f'{rank}{r}',f'=IF({date}{r}=0,0,COUNTIFS({rg(date)},">0",{rg(date)},"<"&{date}{r})+COUNTIFS({rg(date)},{date}{r},{rg(ident)},"<="&{ident}{r}))',True)
        p.write(BACK,f'{offset}{r}',f'=IF({index}>${offset}$1,0,MATCH({index},{rg(rank)},0))',True)
        for c,source in [(dt,date),(uid,ident),(pal,palette),(signed,units)]:
            p.write(BACK,f'{c}{r}',f'=IF({offset}{r}=0,0,INDEX({rg(source)},{offset}{r}))',True)
        p.write(BACK,f'{running}{r}',f'=IF({pal}{r}=0,0,SUMIF(${pal}$2:{pal}{r},{pal}{r},${signed}$2:{signed}{r}))',True)
        p.write(BACK,f'{keep}{r}',f'=IF({signed}{r}<=0,0,MAX(0,_xlfn.MINIFS(${running}{r}:${running}${END},${pal}{r}:${pal}${END},{pal}{r})-MAX(0,{running}{r}-{signed}{r})))',True)
        p.write(BACK,f'{start}{r}',f'=SUM(${keep}$2:{keep}{r})-{keep}{r}+1',True)
        p.write(BACK,f'{end}{r}',f'=SUM(${keep}$2:{keep}{r})',True)
    return dict(base=base,columns=names,total=f"'{BACK}'!${end}$1",
                starts=f"'{BACK}'!{rg(start)}",palettes=f"'{BACK}'!{rg(pal)}")


def install(candidate,output):
    w = openpyxl.load_workbook(candidate,data_only=True)
    p = ExtendedPackage(candidate)
    styles = p.tree('xl/styles.xml');xfs=styles.find(Q('cellXfs'));fills=styles.find(Q('fills'));dxfs=styles.find(Q('dxfs'))
    def key(node):
        return node.tag,tuple(sorted(node.attrib.items())),node.text,tuple(key(x) for x in node)
    cache = {key(x):i for i,x in enumerate(xfs)}
    def filled_style(original,fill,border):
        xf = deepcopy(xfs[original]);xf.set('fillId',str(fill));xf.set('applyFill','1')
        xf.set('borderId',str(border));xf.set('applyBorder','1')
        serial = key(xf)
        if serial not in cache:
            cache[serial]=len(xfs);xfs.append(xf)
        return cache[serial]
    result={}
    for month in [11,12]:
        name=f'{month}월계산기';sheet=p.sheet(name)
        records,sequence=expected_blocks(w,month)
        actual=Counter(x['palette'] for x in sequence)
        assert all(actual[r]==w[name][f'C{r}'].value for r in range(3,11))
        helpers=install_helpers(p,month)
        old=sheet.findall(Q('conditionalFormatting'))
        dxf_by_rgb={dxfs[int(n.find(Q('cfRule')).get('dxfId'))].find(Q('fill')).find(Q('patternFill')).find(Q('fgColor')).get('rgb'):int(n.find(Q('cfRule')).get('dxfId')) for n in old}
        color_fill={r:int(xfs[int(p.cell(name,f'B{r}').get('s'))].get('fillId')) for r in range(3,11)}
        palette_rgb={r:w[name][f'B{r}'].fill.fgColor.rgb for r in range(3,11)}
        white_fill=next(i for i,f in enumerate(fills) if f.find(Q('patternFill')) is not None and f.find(Q('patternFill')).get('patternType')=='solid' and f.find(Q('patternFill')).find(Q('fgColor')) is not None and f.find(Q('patternFill')).find(Q('fgColor')).get('rgb')=='FFFFFFFF')
        for index,(r,c) in enumerate(slots(),1):
            cell=p.cell(name,f'{col(c)}{r}')
            palette_value=sequence[index-1]['palette'] if index<=len(sequence) else 0
            row_start=15 if r<145 else 165
            template=p.cell(name,f'{col(c)}{20+(r-row_start)%5}')
            border=int(xfs[int(template.get('s'))].get('borderId'))
            cell.set('s',str(filled_style(int(cell.get('s','0')),color_fill[palette_value] if palette_value else white_fill,border)))
            p.write(name,f'{col(c+31)}{r}',f'=IF({index}>{helpers["total"]},0,LOOKUP({index},{helpers["starts"]},{helpers["palettes"]}))',True)
        for node in old:sheet.remove(node)
        for r in range(3,11):
            conditional(p,name,'B15:U144 B165:U434',f'AG15={r}',dxf_by_rgb[palette_rgb[r]],r-2)
        conditional(p,name,'B15:U144 B165:U434','AG15=0',dxf_by_rgb['FFFFFFFF'],9)
        p.write(name,'A2','1일 → 말일, 같은 날은 결제시간순으로 이어집니다. 색상은 거래 분류 · 1칸=1,000원입니다.')
        p.write(name,'A154','RAW 날짜·금액·분류 수정 시 날짜순 배치와 합계가 자동 갱신됩니다. 환급은 같은 분류의 앞선 칸을 상계합니다.')
        p.write(name,'AF3',f'={helpers["total"]}',True)
        columns=sheet.find(Q('cols'))
        for node in list(columns):
            if int(node.get('min'))>=33:columns.remove(node)
        ET.SubElement(columns,Q('col'),min='33',max='52',hidden='1',width='12',customWidth='1')
        sheet.find(Q('dimension')).set('ref','A1:AZ434')
        result[str(month)]=dict(counts=dict(actual),total_blocks=len(sequence),order='date_time_stable_id',
                               helpers=helpers,records=records,refund_method='same_category_preceding_blocks',
                               first_blocks=[dict(x) for x in sequence[:15]])
    xfs.set('count',str(len(xfs)))
    dim=p.sheet(BACK).find(Q('dimension'))
    _,_,last_col,last_row=openpyxl.utils.cell.range_boundaries(dim.get('ref'))
    dim.set('ref',f'A1:{col(max(83,last_col))}{max(31699,last_row)}')
    p.save(output);w.close();gc.collect()
    return result


def allocate_october(p,baseline):
    w=openpyxl.load_workbook(baseline,data_only=True);s=w['잔액'];reserve=w['예비비']
    planned={fund:sum(reserve.cell(r,13).value or 0 for r in range(20,29) if reserve.cell(r,1).value==fund) for fund in FUNDS}
    assert planned==dict(zip(FUNDS,[0,100000,50000,50000,0])),planned
    existing=Counter()
    for r in range(64,501):
        date=s[f'A{r}'].value
        if isinstance(date,datetime) and date.year==2026 and date.month==10 and s[f'B{r}'].value=='적립' and s[f'C{r}'].value=='생활비':
            existing[s[f'D{r}'].value]+=s[f'E{r}'].value
    row=max(r for r in range(64,501) if s[f'A{r}'].value)+1;added=[]
    for fund,amount in planned.items():
        missing=amount-existing[fund]
        assert missing>=0
        if not missing:continue
        assert row<=500
        inputs=dict(A=datetime(2026,10,31,23,59,59),B='적립',C='생활비',D=fund,E=missing,
                    F=f'10월 {fund} 적립 · CMA 내부 배분(월말 기록)',
                    J='미래에셋_WRAP형CMA',K='미래에셋_WRAP형CMA',L='WRAP내계정이동')
        for c,value in inputs.items():p.write('잔액',f'{c}{row}',value)
        added.append(dict(row=row,fund=fund,amount=missing,date='2026-10-31',raw_transaction_created=False));row+=1
    for sheet in ['잔액','예비비']:p.write(sheet,'B3',datetime(2026,10,1))
    w.close();return added


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for name in ['baseline','root','output','metadata']:a.add_argument(name,type=Path)
    args=a.parse_args();args.root.mkdir(exist_ok=True)
    p=ExtendedPackage(args.baseline);allocations=allocate_october(p,args.baseline)
    p.write('사용안내','A26','11/16 보완 · 계산기는 날짜·시간순 배치입니다. 10월 목적자금 200,000원은 생활비에서 내부 적립하며 CMA 총액은 유지합니다.')
    p.save(args.root/'allocated.xlsm')
    meta=dict(calculators=install(args.root/'allocated.xlsm',args.root/'candidate.xlsm'),allocations=allocations)
    print('Prepared date order and allocations.',flush=True)
    calc=recalculate(args.root/'candidate.xlsm',args.root/'calculation')
    import_changed_caches(args.root/'candidate.xlsm',calc,args.root/'cached.xlsm')
    meta['memo_cache_restore']=restore(args.baseline,args.root/'cached.xlsm',args.output)
    args.metadata.write_text(json.dumps(meta,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Saved',args.output,flush=True)
