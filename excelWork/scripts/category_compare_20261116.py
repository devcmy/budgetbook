"""Inventory existing source/app, Excel and calculator categories, read-only.

Do not propose or apply a migration. Account names in transfer-category fields
are explicitly separated from income/expense categories. Memo is never read.
"""
from collections import Counter,defaultdict
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import argparse
import json

import openpyxl
from openpyxl.styles import Alignment,Font,PatternFill


def domain(kind):
    return '계좌이체 대상' if kind.startswith('이체') else kind


def inventory(source):
    w=openpyxl.load_workbook(source,data_only=True);raw=w['RAW'];rows=[]
    labels={r:w['11월계산기'][f'A{r}'].value for r in range(3,11)}
    colors={r:w['11월계산기'][f'B{r}'].fill.fgColor.rgb[-6:] for r in range(3,11)}
    for v in raw.iter_rows(min_row=2,max_row=4001,values_only=True):
        if not isinstance(v[0],datetime):continue
        assert v[9]==v[19],('Visible and effective classification differ',v[41])
        route=labels[int(v[23])] if v[24] else '미포함'
        rows.append(dict(id=int(v[41]),month=v[0].month,source=v[3],sub=v[4] or '(없음)',
                         kind=v[7],domain=domain(v[7]),excel=v[9],calculator=route,
                         zero_units=bool(v[24] and not v[25]),title=v[5],fixed_slot=v[26]))
    w.close();assert len(rows)==2741
    return rows,labels,colors


def build(source,output,report_path,markdown):
    digest=sha256(source.read_bytes()).hexdigest();rows,labels,colors=inventory(source)
    wb=openpyxl.Workbook();wb.remove(wb.active)
    def table(name,headers,records,widths):
        s=wb.create_sheet(name);s.append(headers)
        for record in records:s.append(record)
        for row in s:
            for cell in row:
                if isinstance(cell.value,str):cell.data_type='s'
                cell.font=Font(name='맑은 고딕',size=10)
                cell.alignment=Alignment(vertical='top',wrap_text=True)
        for cell in s[1]:
            cell.font=Font(name='맑은 고딕',size=10,bold=True,color='FFFFFF')
            cell.fill=PatternFill('solid',fgColor='234B6A')
        s.row_dimensions[1].height=28;s.freeze_panes='A2';s.auto_filter.ref=s.dimensions
        s.sheet_view.showGridLines=False
        for i,width in enumerate(widths,1):s.column_dimensions[openpyxl.utils.get_column_letter(i)].width=width
        for r in range(2,s.max_row+1):
            lines=max(len(str(c.value or '').splitlines()) for c in s[r])
            s.row_dimensions[r].height=min(300,max(30,lines*15))
            if r%2==0:
                for cell in s[r]:cell.fill=PatternFill('solid',fgColor='EFF4F7')
        return s
    join=lambda values:'\n'.join(sorted(set(values)))
    categories=sorted({r['source'] for r in rows if r['domain']!='계좌이체 대상'})
    overview=[]
    for category in categories:
        vs=[r for r in rows if r['source']==category and r['domain']!='계좌이체 대상']
        excluded=[r['excel'] for r in vs if r['calculator']=='미포함']
        overview.append([category,join(r['sub'] for r in vs),join(r['excel'] for r in vs),
                         join(r['calculator'] for r in vs if r['calculator']!='미포함') or '미포함',
                         join(excluded) or '(없음)',len(vs)])
    table('3가지 비교',['편한가계부 분류','편한가계부 소분류','현재 엑셀 재분류','현재 계산기 색상','계산기 미포함 재분류','RAW 건수'],overview,[30,30,28,24,28,12])
    app=Counter((r['domain'],r['source'],r['sub']) for r in rows)
    table('편한가계부 목록',['유형','분류 또는 이체 대상','소분류','사용 건수'],
          [list(k)+[n] for k,n in sorted(app.items())],[21,34,30,14])
    excel=[]
    for classification in sorted({r['excel'] for r in rows}):
        vs=[r for r in rows if r['excel']==classification]
        excel.append([classification,len(vs),join(r['source'] for r in vs),join(r['calculator'] for r in vs)])
    table('엑셀 목록',['현재 재분류(J열)','RAW 건수','연결된 원본 분류·이체 대상','연결된 계산기 색상 또는 제외'],excel,[26,14,48,33])
    calc=[]
    for key,label in labels.items():
        vs=[r for r in rows if r['calculator']==label]
        calc.append([label,join(r['source'] for r in vs),join(r['excel'] for r in vs),len(vs),
                     'RAW 달력금액이 0이 아닌 거래의 분류 경로. 0칸 환급도 경로에 포함.'])
    s=table('계산기 목록',['실제 범례 이름','연결된 원본 분류','연결된 엑셀 재분류','RAW 경로 건수','집계 설명'],calc,[24,40,30,17,62])
    for r,key in enumerate(labels,2):
        s.cell(r,1).fill=PatternFill('solid',fgColor=colors[key])
        s.cell(r,1).font=Font(name='맑은 고딕',size=10,bold=True,color='FFFFFF' if key in [6,10] else '000000')
    mappings=defaultdict(list)
    for r in rows:mappings[(r['domain'],r['source'],r['sub'],r['excel'],r['calculator'])].append(r)
    details=[]
    for key,vs in sorted(mappings.items()):
        examples=list(dict.fromkeys(r['title'] for r in sorted(vs,key=lambda r:(r['month'],r['id']),reverse=True)))[:4] if key[0]=='지출' else []
        details.append([*key,len(vs),'\n'.join(examples),' · '.join(str(r['id']) for r in vs[:10]),
                        sum(r['zero_units'] for r in vs)])
    table('상세 매칭',['유형','원본 분류·이체 대상','원본 소분류','현재 엑셀 재분류','계산기 분류 또는 제외','RAW 건수','거래내용 예시','거래ID 예시','0칸 경로 건수'],details,[20,32,28,24,22,12,68,45,16])
    transfers=[]
    for target in sorted({r['source'] for r in rows if r['domain']=='계좌이체 대상'}):
        vs=[r for r in rows if r['source']==target and r['domain']=='계좌이체 대상']
        transfers.append([target,len(vs),join(r['excel'] for r in vs),join(r['calculator'] for r in vs)])
    table('이체 대상 목록',['원본 분류란에 적힌 계좌·카드명','RAW 건수','현재 엑셀 재분류','계산기 분류 또는 제외'],transfers,[40,14,35,25])
    fixed=defaultdict(list)
    for r in rows:
        if r['excel']=='고정비':fixed[r['source'],r['title']].append(r)
    table('고정비 원본내용',['원본 분류·이체 대상','실제 RAW 거래내용','사용 건수','등장 월','고정비 연결코드','현재 재분류','계산기'],
          [[a,b,len(vs),', '.join(str(m) for m in sorted({v['month'] for v in vs})),
            ', '.join(str(n) for n in sorted({v['fixed_slot'] for v in vs})),'고정비',join(v['calculator'] for v in vs)]
           for (a,b),vs in sorted(fixed.items())],[32,48,12,32,22,15,20])
    note=[
        ['기준','운영 기준본 RAW 2,741건, 2026년 1~11월. 확인된 미래 할부 날짜 포함.'],
        ['원본 범위','편한가계부 전체 설정 목록이 아니라 전달된 거래에 등장한 분류·소분류입니다.'],
        ['표기','현재 RAW의 원본 분류(D)·소분류(E), 재분류(J), 자동 계산기분류(X)와 달력 집계 여부(Y)를 대조했습니다.'],
        ['이체','이체 행의 분류란은 계좌·카드명일 수 있어 일반 카테고리 22개와 이체 대상 47개를 분리했습니다.'],
        ['계산기','미포함은 해당 RAW가 현재 달력금액(Y)에 들어가지 않는 경우입니다. 고정비·이체·공동자금·수입·중복·과거 별도 처리·0원 등이 포함됩니다.'],
        ['0칸','환급 등 천원 올림 후 0칸인 거래도 금액이 달력에 들어가면 분류 경로에 표시합니다. 색상 경로 건수는 칸수·금액 합계가 아닙니다.'],
        ['과거 항목','목록에는 종료된 구독·할부·과거 이사 항목도 포함됩니다. 목록에 있다는 이유로 현재 고정비에 다시 등록하지 않습니다.'],
        ['비교만 수행','새 분류안이나 자동 매칭 규칙을 적용하지 않았으며 운영 가계부는 수정하지 않았습니다.'],
        ['확인할 연결','현재 자동 분류에서 담배가 여행/이동으로, 일부 가전·전자 거래가 기타로 연결됩니다. 현재 상태를 그대로 표시합니다.'],
        ['원본 해시',digest],
    ]
    table('읽는 방법',['항목','설명'],note,[24,110])
    output.parent.mkdir(parents=True,exist_ok=True);wb.save(output);wb.close()
    check=openpyxl.load_workbook(output,data_only=True)
    assert sum(check['상세 매칭'].cell(r,6).value for r in range(2,check['상세 매칭'].max_row+1))==len(rows)
    assert check['엑셀 목록'].max_row-1==36 and check['계산기 목록'].max_row-1==8
    assert check['이체 대상 목록'].max_row-1==47 and check['3가지 비교'].max_row-1==22
    assert sha256(source.read_bytes()).hexdigest()==digest
    check.close()
    result=dict(as_of='2026-11-16',workbook_sha256=digest,operating_workbook_unchanged=True,
                raw_records=len(rows),by_month=dict(sorted(Counter(r['month'] for r in rows).items())),
                source_income_expense_categories=22,source_transfer_targets=47,excel_categories=36,
                calculator_categories=list(labels.values()),source_category_subcategory_pairs=len({(r['source'],r['sub']) for r in rows}),
                detailed_mapping_rows=len(details),all_raw_rows_covered=True,
                comparison_workbook_sha256=sha256(output.read_bytes()).hexdigest(),
                source_categories=categories,excel_category_counts=dict(sorted(Counter(r['excel'] for r in rows).items())),
                new_taxonomy_applied=False,memos_used_for_classification=False)
    report_path.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    lines=['# 편한가계부 · 엑셀 · 테트리스 카테고리 비교','',
           '운영 RAW 2,741건에 실제 등장한 카테고리의 비교입니다. 현재 명칭과 연결을 그대로 추출했으며 새 분류안을 적용하지 않았습니다.',
           '',f'전체 비교 파일: [{output.name}]({output.name})','',
           '일반 수입·지출 원본 카테고리 22개, 별도 이체 대상 47개, 엑셀 재분류 36개, 계산기 범례 8개입니다. 앱에 설정돼 있지만 전달 거래에 쓰이지 않은 카테고리는 알 수 없습니다.','',
           '| 편한가계부 원본 분류 | 현재 엑셀 재분류 | 현재 계산기 분류 |','| --- | --- | --- |']
    for r in overview:lines.append('| '+' | '.join([r[0],r[2].replace('\n',' · '),r[3].replace('\n',' · ')])+ ' |')
    lines += ['', '계산기 열은 해당 분류에서 실제 RAW 달력 집계로 들어가는 색상입니다. 같은 원본 분류라도 고정비·대납·공동자금·수입·과거 별도 처리 등은 미포함일 수 있습니다. 상세 파일에서 각 연결을 분리해 볼 수 있습니다.',
              '', '현재 담배→여행/이동, 일부 가전/전자/잡화→기타처럼 의미가 어긋나는 연결도 수정하지 않고 그대로 표시했습니다.',
              '', '완료된 애플케어·과거 할부·이사 항목도 전체 기간 목록에 포함되므로 현재 운영 항목과 구분해야 합니다.',
              '', '검증: `category_compare_20261116_validation.json`. 운영 엑셀의 SHA-256은 비교 전후 동일합니다.']
    markdown.write_text('\n'.join(lines)+'\n')
    print(json.dumps({k:result[k] for k in ['raw_records','source_income_expense_categories','excel_categories','detailed_mapping_rows','operating_workbook_unchanged']},ensure_ascii=False))


if __name__=='__main__':
    a=argparse.ArgumentParser()
    for name in ['source','output','report','markdown']:a.add_argument(name,type=Path)
    args=a.parse_args();build(args.source,args.output,args.report,args.markdown)
