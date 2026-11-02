"""Preserve exact note line breaks when importing disposable Office caches.

The evaluator normalizes literal CR/LF. Restore unchanged display strings from
the uploaded source, and rebuild RAW-driven memo caches from actual RAW I.
This edits cached values only, never user inputs, classification, or formulas.
"""
from collections import defaultdict,Counter
from io import BytesIO
from pathlib import Path
import math
import re

from lxml import etree as ET
from openpyxl.reader.strings import read_string_table
from openpyxl.utils.datetime import from_excel

from redesign_reserve import ExtendedPackage
from update_20261015 import Q


def restore(baseline,delivery,output):
    before,p=ExtendedPackage(baseline),ExtendedPackage(delivery)
    strings={}
    def value(package,cell):
        if cell is None:return None
        t=cell.get('t');v=cell.find(Q('v'))
        if t=='inlineStr':return ''.join(x.text or '' for x in cell.iter(Q('t')))
        if v is None or v.text is None:return None
        if t=='s':
            if id(package) not in strings:
                strings[id(package)]=read_string_table(BytesIO(package.parts['xl/sharedStrings.xml']))
            return strings[id(package)][int(v.text)]
        if t=='str':return v.text
        if t in [None,'n']:return float(v.text)
        return v.text
    changes=Counter()
    def cache(sheet,addr,text):
        cell=p.cell(sheet,addr)
        assert cell.find(Q('f')) is not None,(sheet,addr)
        text=None if text in [None,''] else str(text)
        if value(p,cell)==text:return
        for child in list(cell):
            if child.tag in [Q('v'),Q('is')]:cell.remove(child)
        cell.attrib.pop('t',None)
        if text is not None:
            cell.set('t','str');ET.SubElement(cell,Q('v')).text=text
        changes[sheet]+=1
    normalized=lambda x:x.replace('\r','').replace('\n','')
    for sheet in p.sheets:
        if sheet=='RAW':continue
        if sheet not in before.sheets:continue
        source={c.get('r'):c for c in before.sheet(sheet).iter(Q('c'))}
        for cell in p.sheet(sheet).iter(Q('c')):
            if cell.find(Q('f')) is None:continue
            addr=cell.get('r');original=source.get(addr)
            if original is None or original.find(Q('f')) is None:continue
            old,new=value(before,original),value(p,cell)
            if isinstance(old,str) and isinstance(new,str) and ('\n' in old or '\r' in old) and normalized(old)==normalized(new):
                cache(sheet,addr,old)
        before.trees.pop(before.sheets[sheet],None)
    notes={};fixed=defaultdict(list);monthly=defaultdict(list)
    for r in range(2,4002):
        get=lambda c:value(p,p.cell('RAW',f'{c}{r}'))
        memo=get('I') if get('A') is not None else None
        cache('RAW',f'AN{r}',memo)
        ident=int(get('AP'));notes[ident]=memo
        month=get('B')
        if month is None or memo in [None,'']:continue
        m=from_excel(month).month
        monthly[m].append((ident,memo))
        if get('T')=='고정비' and get('Q')==1:fixed[m,int(get('AA'))].append(memo)
    for m in range(1,13):
        sheet=f'{m}월';calc=f'{m}월계산기'
        ordered=sorted(monthly[m]);count=len(ordered)
        try:page=max(1,min(int(value(p,p.cell(sheet,'W40')) or 1),max(1,math.ceil(count/5))))
        except (TypeError,ValueError):page=1
        for i in range(5):
            index=(page-1)*5+i;text=ordered[index][1] if index<count else None
            cache(sheet,f'W{35+i}',text);cache(calc,f'B{158+i}',text)
        for slot in [*range(4,25),26]:cache(sheet,f'V{slot}','\n'.join(fixed[m,slot]))
        for r in range(52,62):
            f=p.cell(sheet,f'I{r}').find(Q('f'))
            if f is None:continue
            ids=list(dict.fromkeys(int(v) for v in re.findall(r'INDEX\(RawData2026\[연동메모\],MATCH\((\d+),RawData2026\[거래ID\],0\)\)',f.text or '')))
            if ids:cache(sheet,f'I{r}','\n'.join(notes[i] for i in ids if notes.get(i)))
    for r in range(64,501):
        if p.cell('잔액',f'G{r}').find(Q('f')) is None:continue
        a,b=value(p,p.cell('잔액',f'W{r}')),value(p,p.cell('잔액',f'X{r}'))
        ids=list(dict.fromkeys(int(v) for v in [a,b] if isinstance(v,(int,float)) and v>=2))
        if ids:cache('잔액',f'G{r}','\n'.join(notes.get(i) or '' for i in ids if notes.get(i)))
    for r in range(100,400):
        if p.cell('예비비',f'H{r}').find(Q('f')) is None:continue
        ident=value(p,p.cell('예비비',f'F{r}'))
        if isinstance(ident,(int,float)) and ident>=2:cache('예비비',f'H{r}',notes.get(int(ident)))
    if '입력·연결' in p.sheets:
        active=value(p,p.cell('입력·연결','K21'))
        selected=value(p,p.cell('입력·연결','B7'))
        cache('입력·연결','B16',notes.get(int(selected)) if active and isinstance(selected,(int,float)) else None)
    p.trees.pop('xl/workbook.xml',None);p.save(output)
    return {'cache_texts_restored_by_sheet':dict(changes),'raw_memos_and_linebreaks_preserved':True,
            'only_formula_text_caches_modified':True,'input_values_and_formulas_unchanged':True}
