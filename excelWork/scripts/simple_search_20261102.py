"""Replace the guided form with a small desktop-Excel VBA search screen.

Patch ZIP parts only. Search output links follow immutable AP transaction IDs.
The VBA module reads RAW but never changes financial source cells.
"""
from copy import deepcopy
from datetime import datetime
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import argparse
import json

from lxml import etree as ET
import openpyxl
from openpyxl.reader.strings import read_string_table
import xlsxwriter

from redesign_reserve import ExtendedPackage, append_styles
from update_20261015 import Q, PNS
from vba_search_project_20261102 import add_search

SCREEN='입력·연결'
VML='xl/drawings/vmlBBSearch20261102.vml'
REL='xl/worksheets/_rels/sheet47.xml.rels'
MODULE=Path(__file__).with_name('BBSearch.bas')


def result_formulas(row):
    position=f'IFERROR(MATCH($H{row},RAW!$AP$2:$AP$4001,0),0)'
    formulas=[]
    for col in ['A','C','F','G']:
        formulas.append(f'=IFERROR(INDEX(RAW!${col}$2:${col}$4001,{position}),"")')
    memo=f'INDEX(RAW!$I$2:$I$4001,{position})'
    formulas.append(f'=IFERROR(IF({memo}="","",{memo}),"")')
    for col,label in [('A','원본 수정'),('I','메모 수정')]:
        formulas.append(f'=IF({position}=0,"",HYPERLINK("#\'RAW\'!{col}"&({position}+1),"{label}"))')
    return formulas


def matches(row,query,month=0):
    """Independent Python counterpart of the small VBA predicate, for fixtures."""
    date=row[0]
    if not isinstance(date,datetime):return False
    if month and date.month!=month:return False
    text=' '.join(str(row[c] or '') for c in [2,5,8,6])
    return query.strip().casefold() in text.casefold()


def author(path,baseline):
    values=openpyxl.load_workbook(baseline,data_only=True)
    raw=values['RAW'];preview=[]
    for r in range(2,4002):
        if isinstance(raw.cell(r,1).value,datetime):
            preview.append((r,[raw.cell(r,c).value for c in [1,3,6,7,9]],raw.cell(r,42).value))
        if len(preview)==10:break
    w=xlsxwriter.Workbook(path);s=w.add_worksheet(SCREEN)
    common={'font_name':'맑은 고딕','font_size':11,'valign':'vcenter'}
    fmt=lambda **kw:w.add_format({**common,**kw})
    title=fmt(bold=True,font_color='FFFFFF',bg_color='17324D',font_size=15)
    hint=fmt(font_color='52677D',text_wrap=True)
    label=fmt(bold=True,font_color='17324D')
    edit=fmt(bg_color='FFF4D9',border=1,border_color='D4DFE8')
    stat=fmt(bg_color='EDF3F8',font_color='17324D')
    header=fmt(bold=True,bg_color='DCE9EF',border=1,border_color='C7D6E0')
    body=fmt(text_wrap=True,bottom=1,bottom_color='E0E7ED')
    date=fmt(num_format='m"월" d"일" hh:mm',bottom=1,bottom_color='E0E7ED')
    money=fmt(num_format='#,##0;[Red]-#,##0',bottom=1,bottom_color='E0E7ED')
    link=fmt(font_color='176D91',underline=True,bottom=1,bottom_color='E0E7ED')
    nav=fmt(font_color='176D91',underline=True)
    s.hide_gridlines(2);s.set_zoom(90);s.freeze_panes(7,0)
    for col,width in [('A',20),('B',26),('C',34),('D',15),('E',48),('F',15),('G',15)]:s.set_column(f'{col}:{col}',width)
    s.set_column('H:J',None,None,{'hidden':True})
    s.merge_range('A1:G1','거래 검색',title);s.set_row(0,37)
    s.merge_range('A2:G2','검색어를 입력하고 검색을 누르세요. 월은 기본 전체입니다. 결과에서 원본 수정 또는 메모 수정을 누릅니다.',hint);s.set_row(1,29)
    s.write('A3','검색어',label);s.merge_range('B3:C3','',edit)
    s.write('D3','조회월',label);s.write('E3','전체',edit);s.set_row(2,36)
    s.data_validation('E3',{'validate':'list','source':['전체']+[f'{m}월' for m in range(1,13)]})
    s.insert_button('F3',{'macro':'BB_Search','caption':'검색','width':172,'height':37,'x_offset':0,'y_offset':3})
    s.merge_range('A4:G4','최근 10건 미리보기 · 검색을 누르면 선택한 조건으로 전체 RAW를 조회합니다.',stat);s.set_row(3,28)
    s.merge_range('A5:G5','버튼에는 매크로 사용이 필요합니다. 보안상 실행할 수 없으면 RAW 탭의 필터를 사용하세요. 검색 결과의 금액·메모는 원본을 따라갑니다.',hint);s.set_row(4,31)
    for addr,target,text in [('A6',"'RAW'!AS5",'새 거래 입력'),('B6',"'11월'!AB4",'11월 예정액'),('C6',"'12월'!AB4",'12월 예정액'),('D6',"'잔액'!A64",'WRAP 원장'),('E6',"'예비비'!A100",'예비비 사용등록')]:
        if addr=='A6':s.write_formula(addr,'=HYPERLINK("#\'RAW\'!A"&(MATCH(TRUE,INDEX(RawData2026[일시]="",0),0)+1),"새 거래 입력")',nav,'새 거래 입력')
        else:s.write_url(addr,'internal:'+target,nav,text)
    s.write_row('A7',['날짜','결제수단 / 계좌','거래내용','금액','메모','원본 수정','메모 수정'],header);s.set_row(6,29)
    # Preformat output only. Values/formulas are written by a search, not by RAW posting.
    for r in range(7,4007):
        s.set_row(r,31)
        for c,f in enumerate([date,body,body,money,body,link,link]):s.write_blank(r,c,None,f)
    for row,(source,record,ident) in enumerate(preview,8):
        s.write(row-1,7,ident)
        cache=[*record,'원본 수정','메모 수정']
        cache[0]=xlsxwriter.utility._datetime_to_excel_datetime(cache[0],False,False)
        for c,formula in enumerate(result_formulas(row)):
            s.write_formula(row-1,c,formula,[date,body,body,money,body,link,link][c],'' if cache[c] is None else cache[c])
    for addr,value in [('J1',len(preview)),('J2','검색 결과 {n}건 · 원본 수정 / 메모 수정으로 이동'),('J3','원본 수정'),('J4','메모 수정'),('J5','검색 결과가 없습니다. 월을 전체로 선택하거나 검색어를 바꿔보세요.')]:s.write(addr,value)
    w.define_name('BB_Query',f"='{SCREEN}'!$B$3");w.define_name('BB_Month',f"='{SCREEN}'!$E$3")
    s.set_selection('B3');s.activate();s.set_landscape();s.fit_to_pages(1,1);s.print_area('A1:G17')
    w.close();values.close()


def simplify(baseline,output,authored):
    author(authored,baseline)
    p=ExtendedPackage(baseline);part=p.sheets[SCREEN]
    with ZipFile(authored) as z:
        stylemap=append_styles(p,z);tree=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        strings=read_string_table(BytesIO(z.read('xl/sharedStrings.xml')))
        for c in tree.iter(Q('c')):
            if c.get('s') is not None:c.set('s',str(stylemap[int(c.get('s'))]))
            if c.get('t')=='s':
                v=c.find(Q('v'));text=strings[int(v.text)];c.remove(v);c.set('t','inlineStr')
                ET.SubElement(ET.SubElement(c,Q('is')),Q('t')).text=text
        p.trees[part]=tree
        rels=ET.fromstring(z.read('xl/worksheets/_rels/sheet1.xml.rels'))
        for rel in rels:
            assert rel.get('Type').endswith('/vmlDrawing')
            rel.set('Target','../drawings/'+Path(VML).name)
        relpart=str(Path(part).parent/'_rels'/(Path(part).name+'.rels'))
        assert relpart not in p.parts and VML not in p.parts
        p.parts[relpart]=ET.tostring(rels,encoding='utf-8',xml_declaration=True)
        p.parts[VML]=z.read('xl/drawings/vmlDrawing1.vml')
    ct=p.tree('[Content_Types].xml');ns=ET.QName(ct).namespace
    if not any(x.get('Extension')=='vml' for x in ct):
        ET.SubElement(ct,'{'+ns+'}Default',Extension='vml',ContentType='application/vnd.openxmlformats-officedocument.vmlDrawing')
    book=p.tree('xl/workbook.xml');names=book.find(Q('definedNames'))
    for node in list(names):
        if node.get('name') in ['_FormFixedChoices','BB_Query','BB_Month']:names.remove(node)
    for name,ref in [('BB_Query',f"'{SCREEN}'!$B$3"),('BB_Month',f"'{SCREEN}'!$E$3")]:
        ET.SubElement(names,Q('definedName'),name=name).text=ref
    for name in p.sheets:
        if name==SCREEN:continue
        dvs=p.sheet(name).find(Q('dataValidations'))
        if dvs is None:
            p.trees.pop(p.sheets[name],None);continue
        touched=False
        for dv in dvs:
            if dv.get('promptTitle')=='입력·연결 먼저 확인':
                touched=True
                dv.set('promptTitle','원본 거래에서 수정')
                prompt=dv.get('prompt','')
                if '입력·연결' in prompt or '화면' in prompt:
                    dv.set('prompt','거래 검색 화면에서 원본 거래를 찾은 뒤 RAW를 수정하세요. 예정액은 11·12월 AB, 메모는 RAW I입니다.')
                error=dv.get('error','')
                if '입력·연결' in error:dv.set('error','자동 계산식을 덮어쓰지 마세요. 거래 검색에서 RAW 원본을 찾거나 11·12월 AB 예정입력 칸을 사용하세요.')
        if not touched:p.trees.pop(p.sheets[name],None)
    # These existing source positions remain valid; only their navigation labels change.
    p.write('사용안내','A27',f'=HYPERLINK("#\'{SCREEN}\'!B3","거래 검색 → 원본 RAW 수정 · 신규 거래는 RAW 빈 행에 입력")',True)
    p.cell('사용안내','A27').set('t','str');ET.SubElement(p.cell('사용안내','A27'),Q('v')).text='거래 검색 → 원본 RAW 수정 · 신규 거래는 RAW 빈 행에 입력'
    p.write('RAW','AR19','거래 검색');p.write('RAW','AS19',f'=HYPERLINK("#\'{SCREEN}\'!B3","검색어·월 선택 → 검색 → 원본 수정")',True)
    p.cell('RAW','AS19').set('t','str');ET.SubElement(p.cell('RAW','AS19'),Q('v')).text='검색어·월 선택 → 검색 → 원본 수정'
    source=MODULE.read_text().replace('\r\n','\n').replace('\n','\r\n').encode('ascii')
    p.parts['xl/vbaProject.bin'],vba=add_search(p.parts['xl/vbaProject.bin'],source)
    p.save(output)
    return {'workbook_sha256':sha256(Path(output).read_bytes()).hexdigest(),
            'baseline_sha256':sha256(Path(baseline).read_bytes()).hexdigest(),
            'screen':SCREEN,'search_query':'B3','month':'E3','default_month':'전체',
            'button_macro':'BB_Search','button_vml':VML,'manual_id_confirmation_removed':True,
            'search_filters_title_memo_asset_amount':True,'search_does_not_write_financial_source_cells':True,
            'native_excel_execution':False,'vba':vba}


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for arg in ['baseline','output','authored','report']:a.add_argument(arg,type=Path)
    args=a.parse_args();result=simplify(args.baseline,args.output,args.authored)
    args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(result,ensure_ascii=False),flush=True)
