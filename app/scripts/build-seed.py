"""Read-only extraction. Preserve RAW; seed only reconciled October operational records."""
import json, pathlib, hashlib, shutil
import openpyxl
ROOT=pathlib.Path(__file__).resolve().parents[1]
SOURCE=ROOT.parent.parent/'household-budget-history/source/ledger-handoff'
raw_path=SOURCE/'evidence/latest_raw_records.json'
raw=json.loads(raw_path.read_text())
snapshot=json.loads((ROOT/'ios/TetrisBudget/snapshot.json').read_text())
batch=hashlib.sha256(raw_path.read_bytes()).hexdigest()
rows=[r for r in raw if str(r['일시']).startswith('2026-10')]
links={2327:16,2328:17,2350:10,2352:8,2354:14,2361:19}
entries=[]
for r in rows:
    id=r['source_row']; account=r['자산']; fixed=r['재분류']=='고정비'
    op='jointActivity' if account=='20260606' else 'jointAllocate' if id==2346 else 'refund' if id in [2329,2345,2358] else 'income' if r['수입/지출']=='수입' else 'expense'
    funding='living' if account=='미래에셋_WRAP형CMA' else 'joint' if account=='20260606' else 'card' if ('카드' in account or account=='KB쿠팡와우') else 'bank'
    entries.append(dict(id=id,date=r['일시'],month='2026-10',title=r['내용'],account=account,category=r['재분류'],amount=r['금액(KRW)'],operation=op,status='unconfirmed' if fixed else 'confirmed',funding=funding,destination='bank',expenseClass='fixed' if fixed else 'variable',reserveID=None,linkedID=None,attributedMonth=None,alreadyNet=False,displayAmount=None,planID=f'oct-fixed-{links[id]}' if id in links else None,note=' · '.join(str(r[k]) for k in ['추가입력','비고'] if r[k]),sourceKind=r['수입/지출'],sourceKey=f'{batch}:{id}'))
bucket={'가족행사비':'family','여행비':'travel','상환비':'repay','경조사비':'condolence','특별비':'other'}
reserves=[dict(r,bucket=bucket[r['group']]) for r in snapshot['reserves']]
# K is the item, L is the detail. Blank K cells continue the preceding group.
# Reading only L loses e.g. 관리비 / 도시가스 and makes both titles identical.
workbook=openpyxl.load_workbook(SOURCE/'sources/가계부/테트리스_가계부_2026_복구.xlsx',data_only=True,read_only=True)
sheet=workbook['10월']; item=''; plans=[]; fixed_labels=[]
for i,r in enumerate(snapshot['fixed']):
    row=i+4
    item=str(sheet.cell(row,11).value or item)
    detail=str(sheet.cell(row,12).value or '')
    title=' · '.join([x for x in [item,detail] if x])
    assert title and r['account']==sheet.cell(row,13).value
    plans.append(dict(id=f'oct-fixed-{i}',month='2026-10',type='fixed',**dict(r,title=title)))
    fixed_labels.append(dict(id=f'oct-fixed-{i}',sheet='10월',itemCell=f'K{row}',detailCell=f'L{row}',item=item,detail=detail,legacyTitle=detail or item,title=title))
workbook.close()
(ROOT/'verification/fixed-item-labels.json').write_text(json.dumps(fixed_labels,ensure_ascii=False,indent=2)+'\n')
plans.append(dict(id='oct-salary',month='2026-10',type='income',title='급여',amount=2776050,status='예상',account='급여 계좌'))
for m in ['2026-10','2026-11','2026-12']:
    plans.append(dict(id=f'reserve-{m}',month=m,type='reserve',title='예비비 적립계획',amount=200000,status='예상',account='생활비'))
for m in ['2026-11','2026-12']:
    plans.append(dict(id=f'family-{m}',month=m,type='familyContribution',title='가족계좌 납입 (10월부터 중단)',amount=200000,status='예상',account='가족계좌'))
state=dict(schemaVersion=1,asOf=snapshot['asOf'],baseMonth='2026-10',nextID=100000,bankBase=dict(living=6330881,family=350000,travel=1400000,repay=300000,condolence=100000,other=0),reserves=reserves,plans=plans,transactions=entries,originals=entries,originalTetrisOutflow=snapshot['tetrisOutflow'],originalTetrisRemaining=snapshot['tetrisRemaining'],originalDaily=snapshot['dailyAmounts'],sourceArchive=raw,originalPlans=plans,audit=[])
path=ROOT/'ios/TetrisBudget/seed-ledger.json'
path.write_text(json.dumps(state,ensure_ascii=False,indent=2)+'\n')
shutil.copyfile(ROOT/'core/ledger-engine.js', ROOT/'ios/TetrisBudget/ledger-engine.js')
print(json.dumps(dict(operational=len(entries),archive=len(raw),sha256=batch,source_unchanged=True),ensure_ascii=False))
