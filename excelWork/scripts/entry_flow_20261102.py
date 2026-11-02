"""Native Excel guided entry, explicit RAW fixed-slot links, and plan inputs.

The new screen is a navigator/search form, NOT another financial register.
Actuals remain in RAW, forecasts in monthly AB, and memo text in RAW I.
No VBA event handler, automatic account inference, or parallel expense posting.
"""
from collections import Counter
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile
import argparse
import json
import re

from lxml import etree as ET
import openpyxl
from openpyxl.styles import Font,PatternFill,Border,Side,Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as col,column_index_from_string as colindex,range_boundaries

from redesign_reserve import ExtendedPackage,append_styles
from update_20261015 import Q,RNS,PNS
from raw_sort_id_20261101 import current,field,TABLE,END
from update_postpaid_20261102 import settlement_condition
from raw_auto_link_20261026 import insert_before_conditional

SCREEN='입력·연결';BACK='_RAW자동연동';LINK='고정비 연결코드';PLAN='AB'
CODES=[*range(4,25),26]
ACTUAL=['확정','결제','RAW 확인','기존 입력']
MODES=['신규 RAW 등록','기존 RAW 수정·연결','고정비 예정 변경','메모 수정','WRAP 원장 연결','예비비 사용 연결','공동·대납 정산 확인','WRAP 내부 계정이동']
MONEY='#,##0;[Red]-#,##0;–'


def manual_slot_formula(original):
    code=f'IFERROR(VALUE({current(LINK)}),0)'
    valid=f'ISNUMBER(MATCH({code},_ManualFixedCodes,0))'
    eligible=f'OR({current("수입/지출")}="지출",{settlement_condition()})'
    return f'IF(AND({current("적용분류")}="고정비",{eligible},{valid}),{code},{original.lstrip("=")})'


def manual_status_formula(original):
    code=f'IFERROR(VALUE({current(LINK)}),0)'
    valid=f'ISNUMBER(MATCH({code},_ManualFixedCodes,0))'
    eligible=f'OR({current("수입/지출")}="지출",{settlement_condition()})'
    return (f'IF({current("일시")}="","",IF({current(LINK)}="",{original.lstrip("=")},'
            f'IF(NOT({valid}),"확인: 고정비 연결코드",IF({current("적용분류")}<>"고정비",'
            f'"연결 대기: J 재분류를 고정비로",IF(NOT({eligible}),"확인: 이체/수입 연결 유형",'
            f'({original.lstrip("=")})&" · 수기 항목 연결")))))')


def actual_count(slot):
    base=f'{field("월")},$T$1,{field("고정항목")},{slot},{field("적용분류")},"고정비",{field("동일거래순번")},1'
    return f'(COUNTIFS({base},{field("수입/지출")},"지출")+COUNTIFS({base},{field("고정금액")},"<>0"))'


def plan_formula(original,slot):
    return f'IF(AND(ISNUMBER({PLAN}{slot}),{actual_count(slot)}=0),{PLAN}{slot},{original.lstrip("=")})'


def plan_status(original,slot):
    return f'IF(AND(ISNUMBER({PLAN}{slot}),{actual_count(slot)}=0),"예상",{original.lstrip("=")})'


def source_position():return 'IFERROR(MATCH($B$7,RawData2026[거래ID],0),0)'


def candidate_condition():
    return (f'({field("월")}=$G$3)*ISNUMBER({field("유효일시")})'
            f'*IF($B$4="",1,{field("실금액")}=$B$4)'
            f'*IF($G$4="",1,IFERROR(INT({field("유효일시")}),0)=IFERROR(INT($G$4),0))'
            f'*IF($B$5="",1,{field("자산")}=$B$5)'
            f'*IF($G$5="",1,((ISNUMBER(SEARCH($G$5,{field("내용")})))+(ISNUMBER(SEARCH($G$5,{field("메모")}))))>0)')


def candidate_index_formula(rank):
    # AGGREGATE is stored with the OOXML future-function prefix. Without it
    # LibreOffice reports #NAME? (the IFERROR would hide all candidate rows).
    return f'IFERROR(_xlfn.AGGREGATE(15,6,(ROW(RAW!$A$2:$A$4001)-1)/({candidate_condition()}),{rank}),"")'


def build_screen(path):
    w=openpyxl.Workbook();s=w.active;s.title=SCREEN
    navy='17324D';ink='243B53';yellow='FFF4D9';blue='EDF3F8';green='E8F3ED'
    thin=Side(style='thin',color='D4DFE8')
    for row in s.iter_rows(min_row=1,max_row=45,min_col=1,max_col=13):
        for c in row:
            c.font=Font(name='맑은 고딕',size=10,color=ink)
            c.alignment=Alignment(vertical='center',wrap_text=True)
    def band(row,text):
        s.merge_cells(start_row=row,start_column=1,end_row=row,end_column=10)
        s.cell(row,1,text);s.cell(row,1).font=Font(name='맑은 고딕',size=12,bold=True,color='FFFFFF')
        s.cell(row,1).fill=PatternFill('solid',fgColor=navy);s.row_dimensions[row].height=29
    def region(addr,end,value,editable=False,fmt=None):
        s.merge_cells(f'{addr}:{end}');c=s[addr];c.value=value
        c.fill=PatternFill('solid',fgColor=yellow if editable else blue)
        c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        if fmt:c.number_format=fmt
    def dv(kind,formula,addr,prompt,error=None):
        d=DataValidation(type=kind,formula1=formula,allow_blank=True,showInputMessage=True,
                         showErrorMessage=bool(error),errorStyle='stop',promptTitle='입력·연결',prompt=prompt,
                         errorTitle='입력 확인',error=error)
        d.add(addr);s.add_data_validation(d)
    band(1,'입력·연결 · 금액은 원본에서 한 번만 등록합니다')
    region('A2','J2','① 처리유형 → ② RAW 후보 조회 → ③ 거래ID 확인 → ④ 원본 수정 / 예정 입력')
    for addr,label in {'A3':'처리유형','F3':'조회월','A4':'거래금액(검색)','F4':'거래일(검색)',
                       'A5':'결제수단(검색)','F5':'내용/메모 검색','A7':'선택 RAW ID','F7':'고정비 항목',
                       'A8':'입력 위치 확인','F8':'확인용 ID¹','A11':'선택 거래일','F11':'원금',
                       'A12':'결제수단','F12':'본인부담','A13':'원본 내용','A14':'적용분류',
                       'F14':'고정비 코드','A15':'수입/지출','F15':'현재 월 집계',
                       'A16':'RAW 메모(조회)','A17':'잔액 관리범위','F17':'기존 연결'}.items():s[addr]=label
    region('B3','D3',MODES[1],True);region('G3','J3',datetime(2026,11,1),True,'yyyy-mm')
    region('B4','D4',None,True,MONEY);region('G4','J4',None,True,'yyyy-mm-dd')
    region('B5','D5',None,True);region('G5','J5',None,True)
    region('B7','D7',2592,True,'0');region('G7','J7',None,True)
    region('B8','D8','원본 내용을 아래에서 먼저 확인하세요');region('G8','J8',None,True,'0')
    dv('list','"'+','.join(MODES)+'"','B3','새 거래/기존 거래/예정/메모/원장/예비비를 먼저 고르세요.')
    dv('list','_EntryMonths','G3','조회할 월을 선택하세요.','목록의 2026년 조회월을 선택하세요.')
    dv('list','_EntryAssets','B5','비워두면 모든 자산을 조회합니다. 새 자산명 직접 입력도 가능합니다.')
    dv('list','_FormFixedChoices','G7','선택월의 고정비 항목을 선택하세요.')
    for addr in ['B7','G8']:
        d=DataValidation(type='whole',operator='between',formula1=0,formula2=END,allow_blank=True,
                         showInputMessage=True,showErrorMessage=True,errorStyle='stop',
                         promptTitle='RAW 거래ID',prompt='기존 거래는 확인한 ID를 다시 입력. 신규는 0. 행번호가 아닙니다.',
                         errorTitle='거래ID 확인',error='0~4001의 정수를 입력하세요.')
        d.add(addr);s.add_data_validation(d)
    for i,code in enumerate(CODES,2):
        k=f'INDIRECT("\'"&MONTH($G$3)&"월\'!K{code}")'
        l=f'INDIRECT("\'"&MONTH($G$3)&"월\'!L{code}")'
        label='"기타 고정비"' if code==26 else f'IF({l}="",IF({k}="","미사용 항목",{k}),IF({k}="",{l},{k}&" / "&{l}))'
        s[f'M{i}']=f'=IFERROR(TEXT({code},"00")&" · "&{label},"조회월 선택 필요")'
    idx=source_position()
    s['K21']=f'=IF(OR($B$3="고정비 예정 변경",$B$3="신규 RAW 등록",$B$3="WRAP 내부 계정이동"),"",IFERROR(IF(INDEX({field("일시")},{idx})="","",{idx}),""))'
    s['K22']='=IFERROR(IF(COUNTIF(_FormFixedChoices,$G$7)=1,VALUE(LEFT($G$7,2)),0),0)'
    s['K23']='=IF($K$21="",0,COUNTIF(잔액!$W$64:$W$500,$B$7)+COUNTIF(잔액!$X$64:$X$500,$B$7))'
    s['K24']='=IFERROR(IF(AND($K$22>0,MONTH($G$3)>=11),INDIRECT("\'"&MONTH($G$3)&"월\'!U"&$K$22),""),"")'
    s['K26']=(f'=IFERROR(AND($K$21<>"",$G$8=$B$7,INDEX({field("월")},$K$21)=$G$3,'
               f'OR($B$4="",INDEX({field("실금액")},$K$21)=$B$4),'
               f'OR($G$4="",INT(INDEX({field("유효일시")},$K$21))=INT($G$4)),OR($B$5="",INDEX({field("자산")},$K$21)=$B$5)),FALSE)')
    s['K29']=f'=SUMPRODUCT({candidate_condition()})'
    status=('=IF($B$3="신규 RAW 등록",IF($K$29=0,"같은 조건의 후보가 없습니다. 새 거래임을 확인하고 확인용 ID에 0을 입력하세요.",'
            '"유사 RAW 후보 "&$K$29&"건: 기존 거래인지 먼저 확인하세요. 실제 별개 거래일 때만 신규 등록합니다."),'
            'IF($B$3="고정비 예정 변경",IF(MONTH($G$3)<11,"마감월은 실제 RAW 정정만 지원합니다. 예정은 11~12월에서 입력하세요.",'
            'IF($K$22=0,"고정비 항목을 선택하세요.",IF(OR($K$24="확정",$K$24="결제",$K$24="기존 입력"),'
            '"이미 실제 확정된 항목입니다. 예정으로 실제 금액을 바꾸지 않습니다.","예정은 실제 거래가 아닙니다. AB 예정입력 칸에서 작성하세요."))),'
            'IF($K$21="","후보에서 실제 RAW 거래ID를 선택하세요.",IF(NOT($K$26),"조회 조건과 원본을 확인하고 확인용 ID에 같은 거래ID를 다시 입력하세요.",'
            '"RAW "&$B$7&" 확인됨 · 아래 원본 입력 위치로 이동해 수정하세요. 이 화면 자체는 금액을 추가하지 않습니다."))))')
    status='=IF($B$3="WRAP 내부 계정이동","은행 거래 없는 내부 분류이동만 선택합니다. RAW는 만들지 않고 확인용 ID에 0으로 확인합니다.",'+status[1:]+')'
    region('A9','J9',status);s.row_dimensions[9].height=43
    for addr,end,header,fmt in [('B11','D11','일시','yyyy-mm-dd'),('G11','J11','실금액',MONEY),
                               ('B12','D12','자산',None),('G12','J12','적용부담',MONEY),('B13','J13','내용',None),
                               ('B14','D14','적용분류',None),('G14','J14','고정항목','0'),('B15','D15','수입/지출',None)]:
        region(addr,end,f'=IF($K$21="","",INDEX({field(header)},$K$21))',fmt=fmt)
    region('G15','J15',f'=IF($K$21="","","달력 "&TEXT(INDEX({field("달력금액")},$K$21),"#,##0")&" / 고정 "&TEXT(INDEX({field("고정금액")},$K$21),"#,##0"))')
    region('B16','J16',f'=IF($K$21="","",INDEX({field("연동메모")},$K$21))')
    region('B17','D17',f'=IF($K$21="","",INDEX({field("잔액 관리범위")},$K$21))')
    region('G17','J17','=IF($K$21="","","원장 "&$K$23&"건 / 예비비 "&COUNTIF(예비비!$F$100:$F$399,$B$7)&"건")')
    band(18,'등록 절차 · 원본을 수정해야 반영됩니다')
    default='"거래ID 확인 → RAW J에서 분류 선택 → 고정비라면 AQ 연결코드 선택. 고정비 항목 코드: "&IF($K$22>0,$K$22,"먼저 항목 선택")&". 달력/월별 수식에 금액을 직접 덮어쓰지 않습니다."'
    steps={
        '신규 RAW 등록':'후보 확인 후 새 거래가 맞으면 확인용 ID에 0 → RAW 빈 행 → A 일시/C~I 거래 입력 → 분류 확인. 실제 별개의 동일 거래는 M 구분을 다르게 쓰며 메모로 구별하지 않습니다.',
        '고정비 예정 변경':'조회월·고정비 항목 선택 → AB 예정입력으로 이동. RAW 청구가 생기면 실제 금액이 우선하며 예정은 다시 더하지 않습니다.',
        '메모 수정':'거래ID/원본 확인 → 확인용 ID 입력 → RAW I 메모 수정. 분류나 금액은 메모로 결정하지 않습니다.',
        'WRAP 원장 연결':'RAW가 WRAP 입출금인지 확인 → 기존 연결 건 확인 → 잔액 입력 원장을 +로 펼쳐 W/X에 거래ID 및 실제 계정을 지정. 자동 추정하지 않습니다.',
        '예비비 사용 연결':'실제 예비비 부담인지 확인 → 예비비 사용등록에서 항목·본인부담·F 거래ID 지정. 카드대금/정산은 별도 원장과 연결합니다.',
        '공동·대납 정산 확인':'본인 최종부담·받을 금액·실제 회수입금을 원거래와 대조합니다. 분담 전에는 추정 금액으로 원장을 바꾸지 않습니다. 본인부담 K 수정/회수 분류는 확인 후 진행합니다.',
        'WRAP 내부 계정이동':'은행 거래가 없는 WRAP 계정 간 분류이동입니다. 확인용 ID에 0 → 잔액 입력 원장에 보낸/받은 계정과 금액 지정. RAW W/X 연결은 넣지 않고 총현금이 보존되는지 확인합니다.',
    }
    for mode,text in reversed(list(steps.items())):default=f'IF($B$3="{mode}","{text}",{default})'
    region('A19','J20','='+default)
    s.row_dimensions[19].height=45;s.row_dimensions[20].height=22
    def link(addr,end,column,label,approved=True):
        target=f'"#\'RAW\'!{column}"&($K$21+1)'
        test='$K$26' if approved else '$K$21<>""'
        region(addr,end,f'=IF({test},HYPERLINK({target},"{label}"),"원본과 거래ID를 먼저 확인하세요")')
    link('A22','E22','A','① 선택 RAW 원본 확인',False)
    target=f'"#\'RAW\'!A"&(MATCH(TRUE,INDEX({field("일시")}="",0),0)+1)'
    region('F22','J22',f'=IF(AND($B$3="신규 RAW 등록",$G$8=0,ISNUMBER($G$8)),IFERROR(HYPERLINK({target},"② 새 RAW 거래 등록"),"RAW 준비영역이 가득찼습니다. 먼저 확장하세요."),"신규 등록: 확인용 ID에 0을 입력하세요")')
    link('A23','E23','J','③ RAW 재분류(J) 입력')
    link('F23','J23','AQ','④ RAW 고정비 연결코드(AQ) 입력')
    link('A24','E24','I','RAW 메모(I) 수정')
    region('F24','J24','=IF(AND($B$3="고정비 예정 변경",$K$22>0,MONTH($G$3)>=11,NOT(OR($K$24="확정",$K$24="결제",$K$24="기존 입력"))),HYPERLINK("#\'"&MONTH($G$3)&"월\'!AB"&$K$22,"예정입력(AB열)로 이동"),"예정 항목/월과 확정 여부를 확인하세요")')
    region('A25','E25','=HYPERLINK("#\'"&MONTH($G$3)&"월\'!A1","월별 결과 확인")')
    region('F25','J25','=IF($K$26,"기존 WRAP 원장 연결 "&$K$23&"건 · 같은 거래ID 중복등록 금지","")')
    region('A26','E26','=IF(OR(AND($B$3="WRAP 내부 계정이동",ISNUMBER($G$8),$G$8=0),AND($K$26,$B$17="WRAP")),HYPERLINK("#\'잔액\'!A64","잔액 입력 원장 · +로 펼쳐 등록"),"WRAP 입출금인지 확인하세요. 카드소비를 재차 차감하지 않습니다.")')
    region('F26','J26','=IF($K$26,HYPERLINK("#\'예비비\'!A100","예비비 사용등록 · +로 펼쳐 등록"),"RAW ID 확인 후 예비비 연결")')
    link('A27','E27','K','본인부담(K) 확인 · 정산 확정 후 수정')
    region('F27','J27','공동 정산의 본인 부담과 회수금은 확인 후 구분합니다.')
    region('A28','J29','¹ 기존 거래는 선택 ID를 확인용 ID에 다시 입력, 신규/내부이동은 0입니다. 검색창은 금융 원장이 아닙니다. 실제 새 거래는 RAW에 등록합니다. 예정은 11~12월 AB열, 메모는 RAW I입니다. 원장/예비비 계정은 확인 없이 추정하지 않습니다.')
    band(31,'RAW 후보 · 조회 조건을 좁힌 뒤 ID를 선택하세요')
    s['A31']='="RAW 후보 · "&TEXT($K$29,"0")&"건 (앞 10건 표시)"'
    for addr,text in {'A32':'거래ID','B32':'일시','C32':'자산','D32':'거래내용','G32':'원금','H32':'분류','I32':'고정코드','J32':'수입/지출'}.items():
        s[addr]=text;s[addr].fill=PatternFill('solid',fgColor=green);s[addr].font=Font(name='맑은 고딕',size=10,bold=True,color=ink)
    s.merge_cells('D32:F32')
    for r in range(33,43):
        s[f'K{r}']='='+candidate_index_formula(r-32)
        for c,header in [('A','거래ID'),('B','일시'),('C','자산'),('D','내용'),('G','실금액'),('H','적용분류'),('I','고정항목'),('J','수입/지출')]:
            s[f'{c}{r}']=f'=IF($K{r}="","",INDEX({field(header)},$K{r}))'
        s.merge_cells(f'D{r}:F{r}');s[f'B{r}'].number_format='m/d hh:mm';s[f'G{r}'].number_format=MONEY;s.row_dimensions[r].height=30
    region('A44','J45','월별 자동계산 칸은 수기 금액으로 덮어쓰지 않습니다. Excel 입력 안내가 표시되면 이 화면에서 유형과 RAW 거래ID부터 확인하세요. 붙여넣기는 유효성 검사를 우회할 수 있으므로 원본/예정 칸만 사용합니다.')
    for c,width in {'A':17,'B':18,'C':19,'D':15,'E':10,'F':17,'G':16,'H':13,'I':13,'J':16}.items():s.column_dimensions[c].width=width
    s['N1']=0;s['N1'].number_format=';;;'
    s.column_dimensions.group('K','N',hidden=True)
    for r in [2,28,29,44,45]:s.row_dimensions[r].height=28
    for r in [3,4,5,7,8,11,12,13,14,15,16,17,22,23,24,25,26,27]:s.row_dimensions[r].height=28
    s.freeze_panes='D10';s.sheet_view.zoomScale=85;s.sheet_view.tabSelected=True
    s.sheet_view.selection[0].activeCell='B3';s.sheet_view.selection[0].sqref='B3'
    s.print_area='A1:J45';s.sheet_properties.pageSetUpPr.fitToPage=True;s.page_setup.orientation='landscape';s.page_setup.paperSize=s.PAPERSIZE_A3;s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=1
    w.save(path)
    styles={name:s[addr].style_id for name,addr in [('input','B3'),('money','B4'),('month','G3'),('label','A3'),('hidden','N1')]}
    w.close();return styles


def relocate_raw_guide(p):
    tree=p.sheet('RAW')
    for row in tree.find(Q('sheetData')):
        for c in row:
            address=c.get('r');index=colindex(re.match('[A-Z]+',address)[0])
            if 43<=index<=48:
                c.set('r',col(index+1)+re.search(r'\d+',address)[0])
                f=c.find(Q('f'))
                if f is not None:f.text=re.sub(r'\bAR5\b','AS5',f.text or '')
    merges=tree.find(Q('mergeCells'))
    for m in merges:
        x1,y1,x2,y2=range_boundaries(m.get('ref'))
        if x1>=43:m.set('ref',f'{col(x1+1)}{y1}:{col(x2+1)}{y2}')
    for c in tree.find(Q('cols')):
        if int(c.get('min'))>=43:c.set('min',str(int(c.get('min'))+1));c.set('max',str(int(c.get('max'))+1))
    columns=tree.find(Q('cols'))
    node=ET.Element(Q('col'),min='43',max='43',width='19',customWidth='1')
    pos=next(i for i,c in enumerate(columns) if int(c.get('min'))>43);columns.insert(pos,node)
    for dv in tree.find(Q('dataValidations')):
        f=dv.find(Q('formula1'))
        if f is not None and f.text and '$AQ$50' in f.text:f.text=f.text.replace('$AQ$','$AR$')
    tree.find(Q('dimension')).set('ref','A1:AW4001')
    if tree.find(Q('autoFilter')) is not None:tree.find(Q('autoFilter')).set('ref','A1:AQ4001')
    p.cell_maps.pop('RAW',None);p.row_maps.pop('RAW',None)


def named(p,name,ref):
    book=p.tree('xl/workbook.xml');nodes=book.find(Q('definedNames'))
    if nodes is None:nodes=ET.SubElement(book,Q('definedNames'))
    assert not any(n.get('name')==name for n in nodes),name
    ET.SubElement(nodes,Q('definedName'),name=name).text=ref


def add_dv(p,sheet,areas,formula,*,kind='custom',prompt='',error='',formula2=None,operator=None):
    tree=p.sheet(sheet);dvs=tree.find(Q('dataValidations'))
    if dvs is None:dvs=ET.Element(Q('dataValidations'),count='0');insert_before_conditional(tree,dvs)
    attrs=dict(type=kind,sqref=areas,allowBlank='1',showInputMessage='1',showErrorMessage='1',errorStyle='stop',
               promptTitle='입력·연결 먼저 확인',prompt=prompt,errorTitle='자동 연동을 유지하세요',error=error)
    if operator:attrs['operator']=operator
    formula=re.sub(r'(?<![A-Za-z0-9_.])ISFORMULA\(','_xlfn.ISFORMULA(',formula)
    dv=ET.SubElement(dvs,Q('dataValidation'),**attrs);ET.SubElement(dv,Q('formula1')).text=formula
    if formula2 is not None:ET.SubElement(dv,Q('formula2')).text=str(formula2)
    dvs.set('count',str(len(dvs)))


def new_dxf(p,color,font):
    styles=p.tree('xl/styles.xml');dxfs=styles.find(Q('dxfs'));idx=len(dxfs)
    d=ET.SubElement(dxfs,Q('dxf'));f=ET.SubElement(d,Q('font'));ET.SubElement(f,Q('color'),rgb=font)
    fill=ET.SubElement(d,Q('fill'));pat=ET.SubElement(fill,Q('patternFill'),patternType='solid')
    ET.SubElement(pat,Q('fgColor'),rgb=color);ET.SubElement(pat,Q('bgColor'),rgb=color)
    dxfs.set('count',str(len(dxfs)));return idx


def conditional(p,sheet,areas,formula,dxf,priority):
    node=ET.Element(Q('conditionalFormatting'),sqref=areas)
    rule=ET.SubElement(node,Q('cfRule'),type='expression',dxfId=str(dxf),priority=str(priority),stopIfTrue='1')
    ET.SubElement(rule,Q('formula')).text=formula;insert_before_conditional(p.sheet(sheet),node)


def prepare(baseline,output,author):
    authored_styles=build_screen(author)
    p=ExtendedPackage(baseline);w=openpyxl.load_workbook(baseline)
    with ZipFile(author) as z:
        stylemap=append_styles(p,z)
        tree=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        for c in tree.iter(Q('c')):
            if c.get('s') is not None:c.set('s',str(stylemap[int(c.get('s'))]))
    book=p.tree('xl/workbook.xml');sheets=book.find(Q('sheets'));sid=max(int(s.get('sheetId')) for s in sheets)+1
    part=f'xl/worksheets/sheet{sid}.xml';rid='rIdEntryFlow20261102';rels=p.tree('xl/_rels/workbook.xml.rels')
    assert SCREEN not in p.sheets and part not in p.parts
    ET.SubElement(rels,'{'+PNS+'}Relationship',Id=rid,Type=RNS+'/worksheet',Target=f'worksheets/sheet{sid}.xml')
    ET.SubElement(sheets,Q('sheet'),name=SCREEN,sheetId=str(sid),attrib={'{'+RNS+'}id':rid})
    p.parts[part]=b'';p.trees[part]=tree;p.sheets[SCREEN]=part
    ct=p.tree('[Content_Types].xml');ET.SubElement(ct,'{'+ET.QName(ct).namespace+'}Override',PartName='/'+part,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    app=p.tree('docProps/app.xml');ans=ET.QName(app).namespace;vns='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
    count=app.find(f'{{{ans}}}HeadingPairs/{{{vns}}}vector/{{{vns}}}variant[2]/{{{vns}}}i4')
    if count is not None:count.text=str(int(count.text)+1)
    titles=app.find(f'{{{ans}}}TitlesOfParts/{{{vns}}}vector')
    ET.SubElement(titles,'{'+vns+'}lpstr').text=SCREEN;titles.set('size',str(len(titles)))
    for sheet in list(w.sheetnames):
        view=p.sheet(sheet).find(Q('sheetViews'))
        if view is not None:
            for v in view:
                if v.get('tabSelected')=='1':v.set('tabSelected','0')
    book.find(Q('bookViews'))[0].set('activeTab',str(len(sheets)-1))
    for name,ref in [('_ManualFixedCodes',f"'{BACK}'!$BA$2:$BA$23"),('_EntryMonths',f"'{BACK}'!$BB$2:$BB$13"),
                     ('_FormFixedChoices',f"'{SCREEN}'!$M$2:$M$23")]:named(p,name,ref)
    for name,header in [('_ERids','거래ID'),('_ERdate','유효일시'),('_ERscope','잔액 관리범위'),('_ERkind','수입/지출')]:
        named(p,name,field(header))
    for i,code in enumerate(CODES,2):
        p.write(BACK,f'AZ{i}',f'{code:02} · 고정비 연결코드');p.write(BACK,f'BA{i}',code)
    p.write(BACK,'AZ1','수기 고정비 연결 · 메모와 분리');p.write(BACK,'BA1','허용 고정코드');p.write(BACK,'BB1','조회월');p.write(BACK,'BC1','자산 목록')
    for m in range(1,13):
        p.write(BACK,f'BB{m+1}',datetime(2026,m,1))
        p.cell(BACK,f'BB{m+1}').set('s',str(stylemap[authored_styles['month']]))
    assets=sorted({str(w['RAW'].cell(r,3).value) for r in range(2,END+1) if w['RAW'].cell(r,3).value is not None})
    for r,asset in enumerate(assets,2):p.write(BACK,f'BC{r}',asset)
    named(p,'_EntryAssets',f"'{BACK}'!$BC$2:$BC${len(assets)+1}")
    p.sheet(BACK).find(Q('dimension')).set('ref','A1:BC31699')
    relocate_raw_guide(p)
    p.write('RAW','AQ1',LINK);p.cell('RAW','AQ1').set('s',p.cell('RAW','J1').get('s'))
    for r in range(2,END+1):p.cell('RAW',f'AQ{r}').set('s',str(stylemap[authored_styles['input']]))
    p.write('RAW','AR19','수기 입력·연결');p.write('RAW','AS19',f'=HYPERLINK("#\'{SCREEN}\'!B3","입력 유형과 RAW ID부터 확인 →")',True)
    merges=p.sheet('RAW').find(Q('mergeCells'));ET.SubElement(merges,Q('mergeCell'),ref='AS19:AW19');merges.set('count',str(len(merges)))
    add_dv(p,'RAW','AQ2:AQ4001','_ManualFixedCodes',kind='list',prompt='고정비로 재분류(J)한 뒤 연결코드를 고르세요. 입력·연결 화면에서 선택월 항목을 먼저 확인합니다.',error='허용된 고정비 연결코드를 선택하세요.')
    add_dv(p,'RAW','A2:A4001','TRUE()',prompt='신규인지 기존 거래인지 입력·연결 화면에서 먼저 확인하세요. RAW 직접 입력은 즉시 자동 연동되며 중복 후보를 확인해야 합니다.',error='')
    changed=Counter()
    for r in range(2,END+1):
        old=w['RAW'][f'AA{r}'].value
        if isinstance(old,str) and old.startswith('='):p.write('RAW',f'AA{r}',manual_slot_formula(old),True);changed['AA']+=1
        p.write('RAW',f'N{r}',manual_status_formula(w['RAW'][f'N{r}'].value),True);changed['N']+=1
    tablepart=next(n for n in p.parts if n.startswith('xl/tables/') and n.endswith('.xml') and ET.fromstring(p.parts[n]).get('displayName')==TABLE)
    table=p.tree(tablepart);table.set('ref','A1:AQ4001');table.find(Q('autoFilter')).set('ref','A1:AQ4001')
    for node in table.iter(Q('sortState')):
        node.set('ref',node.get('ref','A1:AP4001').replace(':AP',':AQ'))
    columns=table.find(Q('tableColumns'))
    for c in columns:
        f=c.find(Q('calculatedColumnFormula'))
        if f is not None and c.get('name')=='고정항목':f.text=manual_slot_formula(f.text)
        if f is not None and c.get('name')=='입력확인':f.text=manual_status_formula(f.text)
    ET.SubElement(columns,Q('tableColumn'),id='43',name=LINK);columns.set('count','43')
    green=new_dxf(p,'FFE2F0D9','FF167A45');yellow=new_dxf(p,'FFFFF2CC','FF9C6500');red=new_dxf(p,'FFFCE4D6','FF9C0006')
    cf_report={}
    for month in [10,11,12]:
        sn=f'{month}월';sheet=p.sheet(sn)
        for cf in list(sheet.findall(Q('conditionalFormatting'))):
            if cf.get('sqref')=='O18:O38' and all('$T18' in (r.find(Q('formula')).text or '') for r in cf):sheet.remove(cf)
        existing=[rule for cf in sheet.findall(Q('conditionalFormatting')) for rule in cf]
        for i,rule in enumerate(existing,7):rule.set('priority',str(i))
        conditional(p,sn,'O4:R24 U4:U24','OR($U4="확정",$U4="결제",$U4="RAW 확인")',green,1)
        conditional(p,sn,'O4:R24 U4:U24','$U4="예상"',yellow,2)
        conditional(p,sn,'O26:R26 U26','OR($U26="확정",$U26="결제",$U26="RAW 확인")',green,3)
        conditional(p,sn,'O26:R26 U26','$U26="예상"',yellow,4)
        dvs=sheet.find(Q('dataValidations'))
        if dvs is not None:
            for d in list(dvs):
                if d.get('sqref')=='T4:T24' and d.find(Q('formula1')) is not None and d.find(Q('formula1')).text=='"R,확정"':dvs.remove(d)
            dvs.set('count',str(len(dvs)))
        for area,first in [('O4:O24','O4'),('O26','O26'),('U4:U24','U4'),('U26','U26')]:
            add_dv(p,sn,area,f'ISFORMULA({first})',prompt='실제 금액/상태는 RAW에서 연동됩니다. 수기는 입력·연결 화면에서 RAW ID를 먼저 확인하세요.',error='자동 수식을 덮어쓰지 마세요. 입력·연결에서 RAW를 수정하거나 AB 예정입력 칸을 사용하세요.')
        for column in ['D','E']:
            add_dv(p,sn,f'{column}52:{column}61',f'OR({column}52="",ISFORMULA({column}52))',prompt='추가 실지출도 RAW 원본에 등록해야 합니다. 같은 돈을 이 표에 다시 입력하지 않습니다.',error='새 거래는 RAW에 등록하세요. 기존 거래라면 RAW ID를 찾아 원본을 수정하세요.')
        for start in [4,9,14,19,24,29]:
            add_dv(p,sn,f'B{start}:H{start+1}',f'ISFORMULA(B{start})',
                   prompt='달력 금액은 RAW에서 자동 계산합니다. 신규/기존 여부와 거래ID를 입력·연결에서 확인하세요.',
                   error='달력 계산값을 직접 덮어쓰지 마세요. RAW 원본을 수정하면 달력과 계산기가 함께 갱신됩니다.')
        addresses=[f'C{r}' for r in range(37,42)]+[f'{c}{r}' for c in ['H','I'] for r in range(37,46)]
        addresses += [f'N{r}' for r in [*range(37,44),57,58,59,*range(64,69)]]
        addresses += [f'V{r}' for r in [*range(4,25),26]]
        for addr in addresses:
            if w[sn][addr].data_type=='f':
                add_dv(p,sn,addr,f'ISFORMULA({addr})',prompt='자동 연동 칸입니다. 실제 거래는 RAW, 연결 메모는 RAW I, 예비비는 사용등록에서 수정하세요.',
                       error='연동 수식은 직접 덮어쓰지 마세요. 입력·연결에서 원본 위치를 먼저 확인하세요.')
        cf_report[sn]={'confirmed_status_column':'U','green_dxf':green,'old_payment_method_validation_removed':True}
    for month in [11,12]:
        sn=f'{month}월'
        p.write(sn,f'{PLAN}3','예정 입력');p.cell(sn,f'{PLAN}3').set('s',p.cell(sn,'U3').get('s'))
        for r in CODES:
            old_o=w[sn][f'O{r}'].value;old_u=w[sn][f'U{r}'].value
            p.write(sn,f'O{r}',plan_formula(old_o,r),True);p.write(sn,f'U{r}',plan_status(old_u,r),True)
            p.cell(sn,f'{PLAN}{r}').set('s',str(stylemap[authored_styles['money']]))
            blocked='OR('+','.join(f'U{r}="{x}"' for x in ACTUAL)+')'
            add_dv(p,sn,f'{PLAN}{r}',f'AND(ISNUMBER({PLAN}{r}),{PLAN}{r}>=0,{PLAN}{r}<=1000000000000,NOT({blocked}))',
                   prompt='RAW 청구가 없는 예정액만 입력합니다. 실제 청구가 오면 실제 금액이 우선하며 예정은 재가산하지 않습니다.',
                   error='0 이상 예정액을 입력하세요. 이미 RAW 확정이 있으면 RAW 원본을 수정하세요.')
        for addr,column in [('H37','Q'),('E74','Q'),('I37','R')]:
            old=w[sn][addr].value
            assert old.endswith('+'+column+'26'),(sn,addr,old)
            condition='OR('+','.join(f'U26="{x}"' for x in ACTUAL)+')'
            p.write(sn,addr,old[:-(len(column)+3)]+f'+IF({condition},{column}26,0)',True)
        columns=p.sheet(sn).find(Q('cols'))
        # AB was hidden calculation space. Keep the memo merges V:AA intact,
        # and hide only AB52:AB61 auxiliary units by number format.
        for c in list(columns):
            if int(c.get('min'))<=28<=int(c.get('max')):
                lo,hi=int(c.get('min')),int(c.get('max'));pos=list(columns).index(c);columns.remove(c)
                for a,b in [(lo,27),(29,hi)]:
                    if a<=b:n=deepcopy(c);n.set('min',str(a));n.set('max',str(b));columns.insert(pos,n);pos+=1
        node=ET.Element(Q('col'),min='28',max='28',width='16',customWidth='1')
        pos=next((i for i,c in enumerate(columns) if int(c.get('min'))>28),len(columns));columns.insert(pos,node)
        for r in range(52,62):p.cell(sn,f'AB{r}').set('s',str(stylemap[authored_styles['hidden']]))
    for c in ['W','X']:
        ref=f'{c}64'
        formula=(f'OR({ref}="",IFERROR(AND(ISNUMBER({ref}),COUNTIF($W$64:$X$500,{ref})=1,'
                 f'INDEX(_ERscope,MATCH({ref},_ERids,0))="WRAP",'
                 f'ISNUMBER(INDEX(_ERdate,MATCH({ref},_ERids,0)))),FALSE))')
        add_dv(p,'잔액',f'{c}64:{c}500',formula,prompt='실제 WRAP 입출금의 RAW 거래ID만 한 번 연결합니다. 내부 분류이동은 비워두세요.',
               error='이미 연결된 ID, 미등록 RAW 또는 WRAP이 아닌 거래입니다. 카드소비를 WRAP에서 재차 차감하지 마세요.')
    ref='F100'
    formula=(f'OR({ref}="",IFERROR(AND(ISNUMBER({ref}),COUNTIF($F$100:$F$399,{ref})=1,'
             f'INDEX(_ERkind,MATCH({ref},_ERids,0))="지출",'
             f'ISNUMBER(INDEX(_ERdate,MATCH({ref},_ERids,0)))),FALSE))')
    add_dv(p,'예비비','F100:F399',formula,prompt='RAW 지출 거래ID를 한 번 연결합니다. 본인부담/예비비 항목은 별도 확인하세요.',
           error='지출 RAW ID인지, 사용등록에 중복 연결되지 않았는지 확인하세요.')
    p.write('사용안내','A27',f'=HYPERLINK("#\'{SCREEN}\'!B3","수기 입력·RAW 연결·예정·메모·원장 등록은 입력·연결 화면부터 시작하세요 →")',True)
    calc=book.find(Q('calcPr'));calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    p.save(output);w.close()
    return {'sheet':SCREEN,'raw_table_ref':'A1:AQ4001','raw_manual_fixed_code_column':'AQ','guide_moved_to':'AR:AW',
            'helper_formula_updates':dict(changed),'color_rules':cf_report,'green_fill':'FFE2F0D9',
            'monthly_forecast_input':'11월/12월 AB4:AB24,AB26','transactions_added':0,
            'no_vba_popups_or_macro_execution':True,'input_form_is_navigation_not_parallel_expense_register':True,
            'confirmation_bound_to_selected_transaction_id':True,'excel_paste_can_bypass_data_validation':True}


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for a in ['baseline','output','author','metadata']:p.add_argument(a,type=Path)
    a=p.parse_args();result=prepare(a.baseline,a.output,a.author)
    a.metadata.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(result,ensure_ascii=False),flush=True)
