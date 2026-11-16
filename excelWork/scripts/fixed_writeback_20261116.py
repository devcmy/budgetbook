"""Install PC Excel fixed-cost writeback without changing recorded finances."""
from collections import defaultdict
from copy import deepcopy
from io import BytesIO
from pathlib import Path
import argparse
import gc
import json
import struct

import olefile
import openpyxl
from lxml import etree as ET
from oletools.olevba import VBA_Project,VBA_Module

from redesign_reserve import ExtendedPackage
from raw_auto_link_20261026 import insert_before_conditional
from update_20261015 import Q
from vba_search_project_20261102 import compress,compound,rec
from verify_reserve_redesign import recalculate
from update_cash_details_20261019 import import_changed_caches
from preserve_memo_caches_20261102 import restore

META='_RAW자동연동'
FIRST=32002
SLOTS=[*range(4,25),26]
MODULE=Path(__file__).with_name('BBFixed.bas')
EVENTS='''
Private Sub Workbook_Open()
    Dim wasSaved As Boolean
    wasSaved = Me.Saved
    On Error GoTo Failed
    BBFixed.BB_SnapshotRaw
    BBFixed.BB_FixedEditable True
    If wasSaved Then Me.Saved = True
    Exit Sub
Failed:
    MsgBox "고정비 편집을 시작하지 못했습니다: " & Err.Description, vbExclamation
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo Failed
    BBFixed.BB_FixedEditable False
    Exit Sub
Failed:
    Cancel = True
    MsgBox "고정비 연결을 보호하지 못해 저장을 중단했습니다: " & Err.Description, vbExclamation
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim wasSaved As Boolean
    wasSaved = Me.Saved
    On Error GoTo Failed
    BBFixed.BB_FixedEditable True
    If Success And wasSaved Then Me.Saved = True
    Exit Sub
Failed:
    MsgBox "고정비 편집을 다시 열지 못했습니다. 파일을 다시 열어 주세요.", vbExclamation
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    BBFixed.BB_FixedChange Sh, Target
End Sub
'''


def source_modules(binary):
    ole=olefile.OleFileIO(BytesIO(binary));project=VBA_Project(ole,'','PROJECT','VBA/dir',relaxed=False)
    project.parse_project_stream();list(project.parse_modules())
    result={m.name:m.code_raw for m in project.modules};ole.close();return result


def add_module(binary,workbook_name):
    ole=olefile.OleFileIO(BytesIO(binary));streams={tuple(p):ole.openstream(p).read() for p in ole.listdir()}
    project=VBA_Project(ole,'','PROJECT','VBA/dir',relaxed=False);project.parse_project_stream()
    raw=project.dir_stream.getvalue();stream=project.dir_stream;sizepos=stream.tell()
    assert struct.unpack('<I',stream.read(4))[0]==2
    count=struct.unpack('<H',stream.read(2))[0];assert stream.read(6)==rec(0x13,b'\x00\x00')[:6];stream.read(2)
    prefix=bytearray(raw[:stream.tell()]);struct.pack_into('<H',prefix,sizepos+4,count+1)
    records=[];before={};changed_workbook=False
    for i in range(count):
        start=stream.tell();m=VBA_Module(project,stream,i);end=stream.tell()
        assert m.name!='BBFixed','Already installed'
        before[m.name]=m.code_raw
        record=bytearray(raw[start:end]);needle=rec(0x31,b'\x00'*4)[:6]
        assert record.count(needle)==1
        at=record.index(needle)+6;record[at:at+4]=b'\x00'*4;records.append(bytes(record))
        source=m.code_raw
        if m.name==workbook_name:
            text=source.decode(project.codec)
            assert 'Workbook_SheetChange' not in text and 'Workbook_Open' not in text
            if 'Option Explicit' not in text:text+='\r\nOption Explicit\r\n'
            text+=EVENTS.replace('\r\n','\n').replace('\n','\r\n')
            source=text.encode(project.codec);changed_workbook=True
        streams['VBA',m.streamname]=compress(source)
    assert changed_workbook
    name=b'BBFixed';uni='BBFixed'.encode('utf-16le')
    extra=(rec(0x19,name)+rec(0x47,uni)+rec(0x1A,name)+rec(0x32,uni)+rec(0x1C)+rec(0x48)
           +rec(0x31,struct.pack('<I',0))+rec(0x1E,struct.pack('<I',0))+rec(0x2C,b'\xff\xff')+rec(0x21)+rec(0x2B))
    streams['VBA','dir']=compress(bytes(prefix)+b''.join(records)+extra+raw[stream.tell():])
    streams['VBA','BBFixed']=compress(MODULE.read_text().replace('\r\n','\n').replace('\n','\r\n').encode(project.codec))
    streams['VBA','_VBA_PROJECT']=b'\xcc\x61\xff\xff\x00\x00\x00'
    for path in list(streams):
        if path[-1].startswith('__SRP_'):del streams[path]
    text=streams['PROJECT',].decode(project.codec)
    text=text.replace('Name="VBAProject"','Module=BBFixed\r\nName="VBAProject"')
    text+='BBFixed=0, 0, 0, 0, C\r\n';streams['PROJECT',]=text.encode(project.codec)
    wm=streams['PROJECTwm',];assert wm.endswith(b'\x00\x00')
    streams['PROJECTwm',]=wm[:-2]+name+b'\x00'+uni+b'\x00\x00\x00\x00'
    metadata={}
    for entry in ole.direntries:
        if entry is not None and entry.entry_type in [1,5]:
            path=() if entry.entry_type==5 else (entry.name,)
            metadata[path]={'clsid':entry.clsid,'state':entry.dwUserFlags,'create':entry.createTime,'modify':entry.modifyTime}
    result=compound(streams,metadata);ole.close()
    after=source_modules(result)
    assert set(after)==set(before)|{'BBFixed'}
    assert all(after[k]==v for k,v in before.items() if k!=workbook_name)
    assert after[workbook_name].startswith(before[workbook_name])
    return result,dict(module_count=len(after),preserved_modules=len(before)-1,
                       workbook_event_module=workbook_name,source_codec=project.codec,
                       source_only_rebuild=True,native_excel_macro_executed=False)


def candidate_ids(w,month,slot):
    return [int(v[41]) for v in w['RAW'].iter_rows(min_row=2,max_row=4001,values_only=True)
            if v[0] and v[0].year==2026 and v[0].month==month and v[19]=='고정비'
            and v[20]=='본인' and v[16]==1 and v[26]==slot]


def install(baseline,output):
    p=ExtendedPackage(baseline);values=openpyxl.load_workbook(baseline,data_only=True)
    formulas=openpyxl.load_workbook(baseline,data_only=False)
    groups=defaultdict(list)
    for v in values['RAW'].iter_rows(min_row=2,max_row=4001,values_only=True):
        if v[0] and v[19]=='고정비' and v[20]=='본인' and v[16]==1 and v[26]:groups[v[0].month,v[26]].append(v)
    styles=p.tree('xl/styles.xml');xfs=styles.find(Q('cellXfs'))
    def key(n):return n.tag,tuple(sorted(n.attrib.items())),n.text,tuple(key(x) for x in n)
    cache={key(x):i for i,x in enumerate(xfs)}
    def protection(style,locked):
        xf=deepcopy(xfs[style]);node=xf.find(Q('protection'))
        if node is None:node=ET.SubElement(xf,Q('protection'))
        node.set('locked',str(int(locked)));xf.set('applyProtection','1');k=key(xf)
        if k not in cache:cache[k]=len(xfs);xfs.append(xf)
        return cache[k]
    offsets=[]
    for m in range(1,13):
        name=f'{m}월';sheet=p.sheet(name)
        for slot in SLOTS:
            r=FIRST+(m-1)*23+slot-4
            rows=groups.get((m,slot),[]);actual=sum(v[27] or 0 for v in rows)
            old=values[name][f'O{slot}'].value or 0;delta=old-actual if rows else 0
            assert isinstance(old,(int,float)),(m,slot,old)
            original=formulas[name][f'O{slot}'].value
            fallback=original[1:] if isinstance(original,str) and original.startswith('=') else str(original or 0)
            count=f'COUNTIFS(RawData2026[월],$T$1,RawData2026[고정항목],{slot},RawData2026[적용분류],"고정비",RawData2026[동일거래순번],1,RawData2026[적용결제],"본인")'
            total=f'SUMIFS(RawData2026[고정금액],RawData2026[월],$T$1,RawData2026[고정항목],{slot})'
            formula=f'=IF({count}>0,{total}+({delta}),{fallback})'
            p.write(name,f'O{slot}',formula,True)
            condition=f'(RawData2026[월]=$T$1)*(RawData2026[고정항목]={slot})*(RawData2026[적용분류]="고정비")*(RawData2026[동일거래순번]=1)'
            memo=f'=IFERROR(_xlfn.TEXTJOIN(CHAR(10),TRUE,_xlfn._xlws.FILTER(RawData2026[연동메모],{condition},"")),"메모 길이 초과 · RAW에서 확인")'
            # Scalar FILTER/TEXTJOIN is editable without a legacy CSE array.
            memo_node=p.cell(name,f'V{slot}').find(Q('f'))
            memo_node.attrib.pop('t',None);memo_node.attrib.pop('ref',None)
            memo_node.text=memo[1:]
            for c,value in [(85,m),(86,slot),(87,delta),(88,formula),(89,memo)]:
                p.write(META,f'{openpyxl.utils.get_column_letter(c)}{r}',value)
            if delta:offsets.append(dict(month=m,slot=slot,historical_offset=delta))
        # Formulas stay locked; literal/blank input cells retain editability.
        for cell in sheet.iter(Q('c')):
            addr=cell.get('r');col,row=openpyxl.utils.cell.coordinate_from_string(addr)
            row=int(row);locked=cell.find(Q('f')) is not None or (row in SLOTS and col in ['O','V','W','X','Y','Z','AA'])
            cell.set('s',str(protection(int(cell.get('s','0')),locked)))
        if m>=11:
            for slot in SLOTS:
                cell=p.cell(name,f'AB{slot}');cell.set('s',str(protection(int(cell.get('s','0')),False)))
        node=sheet.find(Q('sheetProtection'))
        if node is None:node=ET.Element(Q('sheetProtection'));insert_before_conditional(sheet,node)
        node.attrib.clear();node.attrib.update(dict(sheet='1',objects='0',scenarios='0',formatCells='0',formatColumns='0',formatRows='0',selectLockedCells='0',selectUnlockedCells='0',sort='0',autoFilter='0'))
        dvs=sheet.find(Q('dataValidations'))
        if dvs is None:dvs=ET.Element(Q('dataValidations'));insert_before_conditional(sheet,dvs)
        for dv in list(dvs):
            ranges=list(openpyxl.worksheet.cell_range.MultiCellRange(dv.get('sqref')).ranges)
            if all(x.min_col==x.max_col and x.min_col in [15,22] and x.min_row>=4 and x.max_row<=26 for x in ranges):dvs.remove(dv)
        for sqref,kind,f1,f2,prompt in [
            ('O4:O24 O26','whole','0','1000000000000','청구액 수정 → 연결된 RAW에 반영합니다. 여러 건이면 수정할 거래를 선택합니다.'),
            ('V4:V24 V26','textLength','0','32000','메모 수정·삭제 → 선택한 RAW 메모에 반영합니다. 메모는 분류에 쓰지 않습니다.')]:
            dv=ET.SubElement(dvs,Q('dataValidation'),type=kind,operator='between',sqref=sqref,allowBlank='1',showInputMessage='1',showErrorMessage='1',errorStyle='stop',promptTitle='고정비 양방향 수정',prompt=prompt,errorTitle='입력 확인',error='청구액은 0원 이상 정수, 메모는 32,000자 이내로 입력하세요.')
            ET.SubElement(dv,Q('formula1')).text=f1;ET.SubElement(dv,Q('formula2')).text=f2
        dvs.set('count',str(len(dvs)))
        p.write(name,'K32','PC Excel에서 매크로 사용: 청구액(O)·메모(V)를 직접 수정하면 RAW에 반영됩니다. 합산 항목은 거래를 선택하며 예정액은 따로 관리합니다.')
    xfs.set('count',str(len(xfs)))
    for row in values['RAW'].iter_rows(min_row=2,max_row=4001,values_only=True):
        ident=int(row[41])
        p.write(META,f'CG{33000+ident}',row[6] or 0)
        p.write(META,f'CH{33000+ident}',row[10] or 0)
    p.sheet(META).find(Q('dimension')).set('ref','A1:CK37001')
    p.write('사용안내','A26','고정비 양방향 수정: PC Excel에서 매크로를 켜면 월별 청구액(O)·메모(V)를 직접 수정할 수 있습니다. RAW 수정도 월별에 반영됩니다. 날짜순 계산기와 10월 내부 적립은 유지합니다.')
    book=p.tree('xl/workbook.xml');codename=book.find(Q('workbookPr')).get('codeName')
    p.parts['xl/vbaProject.bin'],vba=add_module(p.parts['xl/vbaProject.bin'],codename)
    p.save(output);values.close();formulas.close();gc.collect()
    return dict(vba=vba,monthly_sheets=12,fields=['O 청구액','V 메모'],historical_offsets=offsets,
                saved_formula_cells_locked=True,forecast_separate=True,metadata_start=FIRST)


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for name in ['baseline','root','output','metadata']:a.add_argument(name,type=Path)
    args=a.parse_args();args.root.mkdir(exist_ok=True)
    meta=install(args.baseline,args.root/'candidate.xlsm');print('Installed fixed-cost writeback.',flush=True)
    calc=recalculate(args.root/'candidate.xlsm',args.root/'calculation')
    import_changed_caches(args.root/'candidate.xlsm',calc,args.root/'cached.xlsm')
    meta['memo_cache_restore']=restore(args.baseline,args.root/'cached.xlsm',args.output)
    args.metadata.write_text(json.dumps(meta,ensure_ascii=False,default=str,indent=2)+'\n')
    print('Saved',args.output,flush=True)
