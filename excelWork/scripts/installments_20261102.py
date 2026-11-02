"""Roll active installments into forecasts; remove a completed installment.

Known monthly installment aliases are explicit metadata, never inferred from
RAW I. Historical amounts/overrides and all actual ledger transactions survive.
Only November/December forecasts and an empty October finished row change.
"""
from datetime import datetime
from pathlib import Path
import argparse
import json

from lxml import etree as ET
import openpyxl

from redesign_reserve import ExtendedPackage
from update_20261015 import Q
from raw_sort_id_20261101 import current,field,TABLE,END

BACK='_RAW자동연동'
NORMAL=lambda value:str(value or '').replace(' ','')
PLAN_END=7


def plan_rows():
    result=[]
    for month in [11,12]:
        products=[('휴대폰',22,month+4,22,74700,f'11번가({month+4}/22)'),
                  ('에어팟프로3',23,month-4,36,10200,f'롯데카드 띵샵({month-4}/36)'),
                  ('갤럭시버즈4프로',24,month-9,36,11004,f'롯데카드 띵샵({month-9}/36)')]
        for name,slot,installment,total,amount,title in products:
            result.append({'month':month,'name':name,'slot':slot,'installment':installment,'total':total,
                           'amount':amount,'asset':'롯데365카드','title':title,'row':len(result)+2})
    return result


def alias_key(month,asset,title):return f'{month}|{NORMAL(asset)}|{NORMAL(title)}'


def future_slot_formula(original):
    title=f'SUBSTITUTE({current("내용")}," ","")'
    account=f'SUBSTITUTE({current("자산")}," ","")'
    key=f'MONTH({current("유효일시")})&"|"&{account}&"|"&{title}'
    guard=(f'AND({current("적용분류")}="고정비",IFERROR(YEAR({current("유효일시")}),0)=2026,'
           f'IFERROR(MONTH({current("유효일시")}),0)>=11)')
    fallback=f'IF(ISNUMBER(SEARCH("띵샵",{title})),26,{original.lstrip("=")})'
    lookup=f'IFERROR(VLOOKUP({key},\'{BACK}\'!$AT$2:$AU${PLAN_END},2,FALSE),{fallback})'
    future=f'IF(ISNUMBER(SEARCH("애플케어플러스",{title})),26,{lookup})'
    return f'IF({guard},{future},{original.lstrip("=")})'


def raw_total(slot):return f'SUMIFS({field("고정금액")},{field("월")},$T$1,{field("고정항목")},{slot})'
def raw_count(slot):return f'COUNTIFS({field("월")},$T$1,{field("고정항목")},{slot},{field("적용분류")},"고정비",{field("수입/지출")},"지출",{field("동일거래순번")},1)'


def prepare(baseline,output):
    w=openpyxl.load_workbook(baseline,data_only=True);formulas=openpyxl.load_workbook(baseline)
    raw=w['RAW'];byid={raw.cell(r,42).value:r for r in range(2,END+1)}
    assert raw.cell(byid[2526],6).value=='롯데카드 띵샵(1/36)' and raw.cell(byid[2526],7).value==11004
    assert raw.cell(byid[2091],6).value=='애플케어플러스(3/3)' and raw.cell(byid[2091],1).value.month==9
    assert w['10월']['O24'].value==0 and w['11월']['O24'].value==w['12월']['O24'].value==26400
    p=ExtendedPackage(baseline);plans=plan_rows()
    for c,title in [('AT','할부 청구 식별키 · 메모 아님'),('AU','고정항목'),('AV','상품'),('AW','회차'),('AX','예상금액'),('AY','청구월')]:
        p.write(BACK,c+'1',title)
    for item in plans:
        r=item['row']
        for c,value in {'AT':alias_key(item['month'],item['asset'],item['title']),'AU':item['slot'],
                        'AV':item['name'],'AW':f'{item["installment"]}/{item["total"]}회차',
                        'AX':item['amount'],'AY':datetime(2026,item['month'],1)}.items():p.write(BACK,f'{c}{r}',value)
    p.sheet(BACK).find(Q('dimension')).set('ref','A1:AY31699')
    updated=0
    for r in range(2,END+1):
        old=formulas['RAW'][f'AA{r}'].value
        if not isinstance(old,str) or not old.startswith('='):continue
        p.write('RAW',f'AA{r}',future_slot_formula(old),True);updated+=1
    table_part=next(n for n in p.parts if n.startswith('xl/tables/') and n.endswith('.xml')
                    and ET.fromstring(p.parts[n]).get('displayName')==TABLE)
    table=p.tree(table_part)
    item=next(c for c in table.find(Q('tableColumns')) if c.get('name')=='고정항목')
    f=item.find(Q('calculatedColumnFormula'));f.text=future_slot_formula(f.text)
    # Keep October's historical Buds charge in its original Other slot. Only
    # remove the finished, zero-cost AppleCare display; no row insertion/deletion.
    for c in ['N','P','S']:p.write('10월',c+'24',None)
    for c,label in [('L','종료 후 추가청구'),('M','LOCA 365'),('T','롯데카드')]:
        p.write('10월',c+'24',f'IF({raw_count(24)}>0,"{label}","")',True)
    p.write('10월','O24',f'IF({raw_count(24)}>0,{raw_total(24)},"")',True)
    p.write('10월','U24',f'IF({raw_count(24)}>0,"확정","")',True)
    # Q/R formulas already return blank when O/P have no numbers.
    for m in [11,12]:
        sn=f'{m}월';buds=next(x for x in plans if x['month']==m and x['slot']==24)
        p.write(sn,'L24',buds['name']);p.write(sn,'M24','LOCA 365');p.write(sn,'N24',24)
        p.write(sn,'P24',None);p.write(sn,'T24','롯데카드')
        p.write(sn,'O24',f'IF({raw_count(24)}>0,{raw_total(24)},\'{BACK}\'!AX{buds["row"]})',True)
        p.write(sn,'U24',f'IF({raw_count(24)}>0,"확정","예상")',True)
        for x in [v for v in plans if v['month']==m]:
            p.write(sn,f'S{x["slot"]}',f"'{BACK}'!AW{x['row']}",True)
    p.tree('xl/workbook.xml').find(Q('calcPr')).set('fullCalcOnLoad','1')
    p.save(output);w.close();formulas.close()
    return {'date':'2026-11-02','plans':plans,'raw_slot_formulas_updated':updated,
            'completed_installment':{'name':'애플케어플러스','final_raw_id':2091,'final_month':'2026-09','final_installment':'3/3'},
            'buds_first_raw_id':2526,'buds_october_historical_other_slot_retained':True,
            'raw_transactions_added':0,'free_memo_used_as_product_identifier':False}


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['baseline','output','metadata']:p.add_argument(a,type=Path)
    a=p.parse_args();result=prepare(a.baseline,a.output)
    a.metadata.write_text(json.dumps(result,ensure_ascii=False,default=str,indent=2)+'\n')
    print(json.dumps({k:v for k,v in result.items() if k!='plans'},ensure_ascii=False),flush=True)
