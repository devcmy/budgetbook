"""Exercise the actual preview UI via set_content without external navigation.
This managed browser disallows navigation: storage is an explicit test double.
Storage-write/read and app reload semantics are tested, not physical browser persistence.
"""
from pathlib import Path
import json,tempfile,shutil
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parent
HTML=(ROOT/'테트리스_가계부_기능테스트.html').read_text()
ORIGIN=(ROOT/'테트리스_가계부_기능테스트.html').as_uri()
checks=[];errors=[];requests=[]
profile=tempfile.mkdtemp(prefix='tetris-budget-browser-')
context=None
legacy_seed=json.loads((ROOT.parent/'ios/TetrisBudget/seed-ledger.json').read_text())
fixed_labels=json.loads((ROOT.parent/'verification/fixed-item-labels.json').read_text())
for plans in [legacy_seed['plans'],legacy_seed['originalPlans']]:
 for label in fixed_labels:
  next(p for p in plans if p['id']==label['id'])['title']=label['legacyTitle']
legacy_seed['personalization']={'names':{},'compact':True,'reduceTransparency':False}
storage_snapshot={'tetris-budget-functional-v1':json.dumps(legacy_seed,ensure_ascii=False)}
with sync_playwright() as p:
 def start():
  c=p.chromium.launch_persistent_context(profile,executable_path='/usr/bin/chromium',headless=True,viewport={'width':440,'height':956},locale='ko-KR',timezone_id='Asia/Seoul',args=['--no-sandbox','--disable-crashpad-for-testing','--disable-background-networking'])
  c.route('**/*',lambda r:(requests.append(r.request.url),r.abort()))
  page=c.pages[0] if c.pages else c.new_page()
  page.set_default_timeout(5000);page.on('pageerror',lambda error:errors.append(str(error)))
  page.evaluate("""initial=>{window.__testStorage={...initial};Object.defineProperty(window,'localStorage',{value:{getItem:key=>window.__testStorage[key]??null,setItem:(key,value)=>{window.__testStorage[key]=String(value)},removeItem:key=>delete window.__testStorage[key]}});} """,storage_snapshot)
  page.set_content(HTML,wait_until='domcontentloaded')
  page.wait_for_function('Boolean(window.BudgetPlayground?.summary)')
  return c,page
 try:
  context,page=start()
  def summary():return page.evaluate('window.BudgetPlayground.summary')
  def state():return page.evaluate('window.BudgetPlayground.state')
  def tab(t):page.locator('#tabbar [data-tab="'+t+'"]').click()
  def close():page.locator('[data-action="close"]').first.click()
  def record(title):page.locator('#screen [data-record]').filter(has_text=title).first.click()
  assert page.evaluate('window.BudgetPlayground.storageMode')=='persistent'
  assert page.locator('#tabbar button').count()==5
  assert summary()['remaining']==1195895 and summary()['wrap']==8480881
  before_month_switch=state()
  page.get_by_role('button',name='이전 달',exact=True).click()
  assert page.evaluate('window.BudgetPlayground.month')=='2026-09'
  assert '2026년 9월' in page.locator('.month-title').inner_text()
  assert page.get_by_role('dialog').count()==0
  assert summary()['historicalPending'] and summary()['archivedCount']>0
  assert state()==before_month_switch and summary()['wrap']==8480881
  page.get_by_role('button',name='다음 달',exact=True).click()
  assert page.evaluate('window.BudgetPlayground.month')=='2026-10'
  assert summary()['remaining']==1195895
  def period(value):
   page.locator('[data-modal="month"]').click();page.locator('#period').fill(value)
   page.get_by_role('button',name='이 달 보기',exact=True).click()
  period('2026-01');page.get_by_role('button',name='이전 달',exact=True).click()
  assert page.evaluate('window.BudgetPlayground.month')=='2025-12'
  page.get_by_role('button',name='다음 달',exact=True).click()
  assert page.evaluate('window.BudgetPlayground.month')=='2026-01'
  period('2024-03');page.get_by_role('button',name='이전 달',exact=True).click()
  assert page.evaluate('window.BudgetPlayground.month')=='2024-02'
  period('0001-01');assert page.get_by_role('button',name='이전 달',exact=True).is_disabled()
  period('9999-12');assert page.get_by_role('button',name='다음 달',exact=True).is_disabled()
  period('2026-10');assert state()==before_month_switch
  checks.append('이전·다음 달 한 번 탭, 월 선택 병행, 연도 전환·윤년·연도 경계, 원장·잔액 유지')
  tab('records')
  groups=page.locator('#record-groups .record-day')
  assert groups.count()==5
  for group in groups.all():
   date=group.get_attribute('data-date')
   rows=[t for t in summary()['transactions'] if t['date'].startswith(date)]
   income=sum(t['recordIncome'] for t in rows);expense=sum(t['recordExpense'] for t in rows)
   assert f'{income:,}원' in group.locator('.record-day-totals .income-text').inner_text()
   assert f'{expense:,}원' in group.locator('.record-day-totals .expense-text').inner_text()
  page.locator('#search').fill('세이돈부리')
  assert page.locator('#record-groups .record-day').count()==1
  assert '11,000원' in page.locator('.record-day-totals .expense-text').inner_text()
  page.locator('#search').fill('');checks.append('기록 날짜별 확정 수입·지출 합계와 검색 결과 합계, 이체·납부 중복 제외')
  tab('settings')
  plans=page.locator('.plan-group')
  assert plans.count()==3 and plans.evaluate_all('els=>els.every(e=>!e.open)')
  assert '2,776,050' in page.locator('[data-plan-type="income"] summary').inner_text()
  assert '1,248,055' in page.locator('[data-plan-type="fixed"] summary').inner_text()
  assert '200,000' in page.locator('[data-plan-type="reserve"] summary').inner_text()
  page.locator('[data-plan-type="fixed"] summary').click()
  assert page.locator('[data-plan-type="fixed"] [data-plan]').count()==21
  page.locator('[data-plan="oct-fixed-1"]').click()
  assert page.locator('#title').input_value()=='관리비 · 9월사용분(전달)'
  close();checks.append('설정 예상·계획 3종류 합계, 접기·펼치기, 21개 고정비 항목 수정 연결')
  tab('balance');assert 'WRAP 잔액 구성' in page.locator('#screen').inner_text()
  assert '같은 WRAP 계좌' in page.locator('#screen').inner_text()
  checks.append('잔액 구성 명칭과 생활비·목적자금 설명 표시')
  tab('home')
  page.locator('[data-focus="cash"]').click()
  assert '11,900원' in page.get_by_role('dialog').inner_text()
  assert '결산에 추가할 금액' in page.get_by_role('dialog').inner_text()
  close();checks.append('초기 5탭, 엑셀 결산/WRAP/비예비비 중복 방지')
  page.locator('[data-focus="variable"]').click()
  dialog=page.get_by_role('dialog')
  groups=dialog.locator('.variable-day')
  assert groups.count()==5
  assert groups.evaluate_all('els=>els.map(e=>e.dataset.date)')==['2026-10-05','2026-10-04','2026-10-03','2026-10-02','2026-10-01']
  assert dialog.locator('.variable-transaction').count()==19
  assert '286,580' in dialog.locator('.variable-overview').inner_text()
  for day,amount,weekday in [(5,120890,'월요일'),(4,112890,'일요일'),(3,5000,'토요일'),(2,30500,'금요일'),(1,17300,'목요일')]:
   group=dialog.locator(f'.variable-day[data-date="2026-10-0{day}"]')
   assert f'{amount:,}원' in group.locator('.variable-day-heading').inner_text()
   assert weekday in group.locator('.variable-day-heading').inner_text()
   assert sum(int(e.inner_text().replace(',','').replace('원','')) for e in group.locator('.variable-amount').all())==amount
  cu=dialog.locator('.variable-transaction').filter(has_text='종량제봉투, 과자')
  assert cu.count()==1 and '잡화/소품/가구/생필품' in cu.locator('.variable-category').inner_text()
  assert '롯데365카드' in cu.locator('.variable-account').inner_text()
  assert 'RAW 내역' not in cu.inner_text()
  (ROOT/'screenshots').mkdir(exist_ok=True)
  page.screenshot(path=str(ROOT/'screenshots/variable.png'))
  cu.click()
  assert '종량제봉투, 과자' in page.get_by_role('dialog').inner_text()
  assert 'RAW 내역' in page.get_by_role('dialog').inner_text()
  page.get_by_role('button',name='기록 수정',exact=True).click()
  assert page.locator('#title').input_value()=='CU'
  close();checks.append('비고정비 날짜별 내림차순 묶음, 요일·일 합계·19건 합산, 원본 분류·메모·카드 표시 및 상세·수정 연결')
  # A month-only expense belongs to a separate final group, including negative corrections.
  page.locator('[data-action="add"]').first.click()
  page.locator('#title').fill('날짜 미확인 분류 테스트')
  page.locator('#amount').fill('3000');page.locator('#unknownDate').check()
  page.get_by_role('button',name='거래 저장',exact=True).click()
  tab('home');page.locator('[data-focus="variable"]').click()
  unknown=page.get_by_role('dialog').locator('.variable-day').last
  assert unknown.get_attribute('data-date')==''
  assert '날짜 미확인' in unknown.inner_text() and '3,000원' in unknown.inner_text()
  assert summary()['monthlyFocus']['variable']==289580
  unknown.locator('[data-record]').click();page.get_by_role('button',name='기록 수정',exact=True).click()
  page.locator('#amount').fill('-3000');page.get_by_role('button',name='거래 저장',exact=True).click()
  tab('home');page.locator('[data-focus="variable"]').click()
  unknown=page.get_by_role('dialog').locator('.variable-day').last
  assert '-3,000원' in unknown.inner_text() and unknown.locator('.variable-amount.credit').count()==1
  assert summary()['monthlyFocus']['variable']==283580
  unknown.locator('[data-record]').click();page.get_by_role('button',name='거래 삭제',exact=True).click()
  page.get_by_role('button',name='거래 삭제',exact=True).click()
  assert summary()['monthlyFocus']['variable']==286580
  tab('home');checks.append('월만 입력한 거래는 날짜 미확인 마지막 묶음, 음수 정정 금액 부호 및 전체 합계 반영')
  page.locator('[data-focus="fixed"]').click()
  dialog=page.get_by_role('dialog')
  assert '관리비 · 9월사용분(전달)' in dialog.inner_text()
  assert '도시가스 · 9월사용분(전달)' in dialog.inner_text()
  assert '보험비 · 이번달' in dialog.inner_text()
  assert summary()['tetrisPending']==False
  for label in fixed_labels:
   assert next(p for p in summary()['fixed'] if p['planID']==label['id'])['title']==label['title']
  (ROOT/'screenshots').mkdir(exist_ok=True)
  page.screenshot(path=str(ROOT/'screenshots/fixed.png'))
  dialog.locator('[data-fixed="1"]').click()
  assert page.locator('#title').input_value()=='관리비 · 9월사용분(전달)'
  assert page.locator('#amount').input_value()=='170350'
  close();checks.append('이전 저장본 고정비 21건 항목·세부 복원, 관리비/도시가스 구분, 실제 입력 제목 연결')
  # Actual form validation and live multi-screen recalculation.
  page.locator('[data-action="add"]').first.click()
  page.locator('#amount').fill('1.5');page.locator('#title').fill('커피 기능 테스트')
  page.get_by_role('button',name='거래 저장',exact=True).click()
  assert '정수' in page.locator('#form-error').inner_text()
  page.locator('#amount').fill('10000');page.locator('#funding').select_option('living')
  page.get_by_role('button',name='거래 저장',exact=True).click()
  assert summary()['wrap']==8470881 and summary()['remaining']==1185895
  assert summary()['monthlyFocus']['unplannedCash']==21900
  record('커피 기능 테스트');page.get_by_role('button',name='기록 수정',exact=True).click()
  page.locator('#amount').fill('12000');page.get_by_role('button',name='거래 저장',exact=True).click()
  assert summary()['wrap']==8468881 and summary()['remaining']==1183895
  page.locator('#search').fill('커피 기능 테스트');assert page.locator('#screen [data-record]').count()==1
  record('커피 기능 테스트');page.get_by_role('button',name='거래 삭제',exact=True).click()
  page.get_by_role('button',name='거래 삭제',exact=True).click()
  assert summary()['wrap']==8480881 and summary()['remaining']==1195895
  checks.append('실제 입력·금액 검증·수정·검색·확인 후 삭제가 계산에 반영')
  # Reserve funding vs spending.
  tab('reserve');page.get_by_role('button',name='실제적립 입력',exact=True).click()
  page.get_by_role('button',name='거래 저장',exact=True).click()
  assert summary()['wrap']==8480881 and summary()['living']==6230881 and summary()['reserve']==2250000
  assert summary()['reserveMonthActual']==100000
  tab('reserve');page.get_by_role('button',name='사용 입력',exact=True).click()
  page.get_by_role('button',name='거래 저장',exact=True).click()
  assert summary()['wrap']==8440881 and summary()['reserve']==2210000
  assert summary()['monthlyFocus']['reserveUse']==40000
  checks.append('예비비 실제적립 총잔액 유지 / 직접 사용은 결산과 WRAP에 한 번 반영')
  # Fixed actual replaces estimate.
  tab('home');page.locator('[data-focus="fixed"]').click()
  page.get_by_role('dialog').get_by_role('button').filter(has_text='월세').click()
  page.locator('#amount').fill('400000');page.locator('#title').fill('실제 월세')
  page.get_by_role('button',name='거래 저장',exact=True).click()
  assert summary()['monthlyFocus']['fixed']==1257055
  checks.append('고정비 예상 391,000원을 실제 400,000원으로 대체, 증가는 9,000원')
  # Settings aliases and preferences, backed by real browser disk storage.
  tab('settings');page.locator('[data-rename="tab.home"]').click()
  page.locator('#alias').fill('   ');page.get_by_role('button',name='이름 저장',exact=True).click()
  assert page.locator('#form-error').inner_text()
  page.locator('#alias').fill('내 월간장부');page.get_by_role('button',name='이름 저장',exact=True).click()
  page.get_by_text('이번달 메뉴',exact=True).click();page.locator('[data-rename="monthly.cash"]').click()
  page.locator('#alias').fill('계좌 현금지출');page.get_by_role('button',name='이름 저장',exact=True).click()
  page.locator('[data-pref="reduceTransparency"]').check()
  assert page.locator('body').evaluate("e=>e.classList.contains('reduce-transparency')")
  saved=state();saved_summary=summary();storage_snapshot=page.evaluate("window.__testStorage");context.close();context,page=start()
  assert state()==saved and summary()==saved_summary
  assert '내 월간장부' in page.locator('#tabbar').inner_text()
  assert '계좌 현금지출' in page.locator('[data-focus="cash"]').inner_text()
  checks.append('탭·메뉴 이름 검증, 투명도 설정, 저장소 테스트 대역 쓰기·새 화면 읽기 복원')
  # Export actual downloaded JSON, change, preview restore, cancel, restore.
  tab('settings')
  with page.expect_download() as download:
   page.locator('[data-action="export"]').first.click()
  backup=ROOT/'verification-backup.json';download.value.save_as(backup)
  exported=json.loads(backup.read_text());assert exported==state();assert len(exported['sourceArchive'])==2366
  page.locator('[data-action="add"]').first.click();page.locator('#amount').fill('3333');page.locator('#title').fill('복원 전 임시 거래')
  page.get_by_role('button',name='거래 저장',exact=True).click();changed=state();assert changed!=exported
  page.locator('#backup-file').set_input_files(backup)
  assert '이 백업으로 복원할까요' in page.get_by_role('dialog').inner_text()
  close();assert state()==changed
  page.locator('#backup-file').set_input_files(backup);page.get_by_role('button',name='이 백업으로 복원',exact=True).click()
  assert state()==exported
  bad=ROOT/'verification-bad.json';bad.write_text('{broken')
  page.locator('#backup-file').set_input_files(bad);assert page.get_by_role('dialog',name='처리하지 못했어요').count()==1
  assert state()==exported;close();checks.append('실제 JSON 다운로드·취소·복원, 손상 파일 거절, 전체 RAW 보존')
  # Card use and payment live flow; cross-month refund attribution required.
  tab('records');page.locator('[data-action="add"]').first.click()
  page.locator('#title').fill('카드 소비 테스트');page.locator('#amount').fill('50000');page.locator('#funding').select_option('card')
  page.get_by_role('button',name='거래 저장',exact=True).click();purchase=state()['transactions'][-1];before=summary()['outflow']
  page.locator('[data-action="add"]').first.click();page.locator('#operation').select_option('cardPayment')
  page.locator('#date').fill('2026-11-06T12:00');page.locator('#title').fill('다음달 카드대금');page.locator('#amount').fill('50000');page.locator('#funding').select_option('living');page.locator('#linkedID').select_option(str(purchase['id']))
  page.get_by_role('button',name='거래 저장',exact=True).click();assert summary()['actualOutflow']==0 and summary()['wrap']==exported['bankBase']['living']+sum(exported['bankBase'][k] for k in ['family','travel','repay','condolence','other'])-90000
  page.locator('[data-action="add"]').first.click();page.locator('#operation').select_option('refund');page.locator('#title').fill('다른달 정산');page.locator('#amount').fill('10000');page.locator('#linkedID').select_option(str(purchase['id']))
  page.get_by_role('button',name='거래 저장',exact=True).click();assert '귀속 월' in page.locator('#form-error').inner_text()
  page.locator('#attributedMonth').fill('2026-10');page.get_by_role('button',name='거래 저장',exact=True).click()
  tab('home');page.locator('[data-modal="month"]').click();page.locator('#period').fill('2026-10');page.get_by_role('button',name='이 달 보기').click()
  assert summary()['outflow']==before-10000
  checks.append('카드 사용·다음달 납부 소비 분리, 다른 달 정산 귀속 월 저장 검증')
  # Plan editor live integration.
  tab('settings');page.locator('[data-action="add-plan"]').click();page.locator('#month').fill('2026-11');page.locator('#type').select_option('income');page.locator('#title').fill('11월 예상 월급');page.locator('#amount').fill('3000000')
  page.get_by_role('button',name='예상·계획 저장').click();assert summary()['income']==3000000 and summary()['actualIncome']==0
  checks.append('새 월 예상·계획은 예상 결산만 변경, 실제 수입·잔액과 분리')
  # Read-only source evidence; search results use original fields.
  tab('records');page.locator('[data-modal="archive"]').click();page.locator('#archive-all').check();page.locator('#archive-search').fill('두끼떡볶이')
  assert page.locator('#archive-results details').count()>0
  page.locator('#archive-results summary').first.click();assert 'source_row' not in page.locator('#archive-results dl').first.inner_text();assert '원본 행' in page.locator('#archive-results dl').first.inner_text()
  close();checks.append('원본 RAW 전체 이력 검색·펼치기, 읽기 전용')
  # Reset, confirming baseline after earlier tests; responsive screens.
  tab('settings');page.locator('[data-action="reset"]').first.click();page.locator('[data-action="confirm-reset"]').click()
  assert summary()['remaining']==1195895 and summary()['wrap']==8480881 and len(state()['transactions'])==45
  checks.append('확인 후 테스트 초기화, 원본 seed 결산으로 복원')
  # Genuine XLSX/XLS files are selected through the same file control as JSON.
  fixtures=ROOT.parent/'tests/fixtures';initial=state();tab('settings')
  original=ROOT.parent.parent.parent/'household-budget-history/source/ledger-handoff/sources/가계부/테트리스_가계부_2026_복구.xlsx'
  page.locator('#backup-file').set_input_files(original)
  page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
  assert page.get_by_role('dialog',name='엑셀 거래 가져오기').count()==1
  assert '2366건' in page.get_by_role('dialog').inner_text()
  assert page.locator('#confirm-excel').is_disabled();close();assert state()==initial
  page.locator('#backup-file').set_input_files(fixtures/'import.xlsx')
  page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
  assert page.locator('#confirm-excel').inner_text()=='3건 추가하기'
  close();assert state()==initial
  checks.append('실제 원본 XLSX 2,366건 중복 제외, 새 거래 검토 취소 시 원장 유지')
  page.locator('#backup-file').set_input_files(fixtures/'import.xlsx')
  page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
  page.locator('[data-excel-row="1"] summary').click()
  page.locator('[data-excel-index="1"][data-excel-field="planID"]').select_option('oct-fixed-0')
  assert page.locator('#excel-selected').inner_text().find('1건')>=0
  page.screenshot(path=str(ROOT/'screenshots/excel-review.png'))
  page.locator('#confirm-excel').click()
  assert len(state()['transactions'])==48 and len(state()['sourceArchive'])==2369
  assert summary()['wrap']==8470881 and summary()['remaining']==1176895
  assert summary()['monthlyFocus']['fixed']==1257055
  assert next(t for t in state()['transactions'] if t['title']=='엑셀 이체 테스트')['operation']=='pending'
  checks.append('XLSX 추가·원본 보관·고정비 예상 대체·이체 보류, 결산과 WRAP 재계산')
  changed=state();page.locator('#backup-file').set_input_files(fixtures/'import.xls')
  page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
  assert page.locator('#confirm-excel').is_disabled();assert '중복 제외' in page.get_by_role('dialog').inner_text()
  close();assert state()==changed
  page.locator('#backup-file').set_input_files(fixtures/'overlap.xlsx')
  page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
  assert page.locator('#confirm-excel').inner_text()=='1건 추가하기'
  page.locator('#confirm-excel').click();assert len(state()['transactions'])==49
  assert summary()['remaining']==1171895 and summary()['wrap']==8470881
  checks.append('순서가 다른 XLS 재입력 중복 방지, 기간이 겹치는 XLSX에서 새 행만 추가')
  changed=state();page.locator('#backup-file').set_input_files(fixtures/'invalid.xlsx')
  page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
  assert page.locator('#confirm-excel').is_disabled();assert '정수' in page.get_by_role('dialog').inner_text();close()
  page.locator('#backup-file').set_input_files(fixtures/'summary-only.xlsx')
  page.get_by_role('dialog',name='처리하지 못했어요').wait_for()
  assert '거래표를 찾지 못했습니다' in page.get_by_role('dialog').inner_text();close();assert state()==changed
  checks.append('잘못된 날짜·소수점 원 금액·거래표 없는 엑셀을 표시하고 저장 차단')
  for width in [440,390,320]:
   page.set_viewport_size({'width':width,'height':956})
   page.locator('#backup-file').set_input_files(fixtures/'overlap.xlsx')
   page.get_by_role('dialog',name='엑셀 거래 가져오기').wait_for()
   assert not page.get_by_role('dialog').evaluate('e=>e.scrollWidth>e.clientWidth')
   close()
  tab('settings');page.locator('[data-action="reset"]').first.click();page.locator('[data-action="confirm-reset"]').click()
  assert state()['transactions']==initial['transactions']
  shots=ROOT/'screenshots';shots.mkdir(exist_ok=True)
  for width,height in [(440,956),(390,844),(375,812),(320,568),(1440,1050)]:
   page.set_viewport_size({'width':width,'height':height})
   for t in ['home','records','reserve','balance','settings']:
    tab(t)
    assert not page.evaluate('document.documentElement.scrollWidth>innerWidth'),(width,t,'page')
    assert not page.locator('#screen').evaluate('e=>e.scrollWidth>e.clientWidth'),(width,t,'screen')
    if width==440:page.screenshot(path=str(shots/(t+'.png')))
    if t=='settings':
     # Scroll the card above the floating glass navigation before tapping.
     page.locator('[data-plan-type="fixed"] summary').evaluate('e=>e.scrollIntoView({block:"center"})')
     page.locator('[data-plan-type="fixed"] summary').click()
     assert not page.locator('#screen').evaluate('e=>e.scrollWidth>e.clientWidth'),(width,'expanded plans')
     page.locator('[data-plan-type="fixed"] summary').evaluate('e=>e.scrollIntoView({block:"center"})')
     page.locator('[data-plan-type="fixed"] summary').click()
   page.locator('[data-action="add"]').first.click()
   assert not page.get_by_role('dialog').evaluate('e=>e.scrollWidth>e.clientWidth'),(width,'editor')
   close()
   tab('home');page.locator('[data-focus="variable"]').click()
   assert not page.get_by_role('dialog').evaluate('e=>e.scrollWidth>e.clientWidth'),(width,'variable ledger')
   for row in page.get_by_role('dialog').locator('.variable-transaction').all():
    assert not row.evaluate('e=>e.scrollWidth>e.clientWidth'),(width,'variable row')
   close()
  checks.append('440/390/375/320px 및 데스크톱: 5탭과 입력 시트 가로 넘침 없음')
  page.set_viewport_size({'width':1440,'height':1050});tab('home');page.screenshot(path=str(shots/'overview.png'))
  assert not errors,errors
  assert not requests,requests
  report={'checks':checks,'passed':len(checks),'errors':errors,'network':'HTML injected with set_content; no navigation or external requests','storage':'Explicit localStorage test double: saved JSON restored in a fresh browser context; real browser disk persistence not verified','iOS_native_UI':'Not executed; this exercises the shared-core web companion'}
  (ROOT/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
  print(json.dumps(report,ensure_ascii=False))
 finally:
  if context:context.close()
  shutil.rmtree(profile,ignore_errors=True)
  for name in ['verification-backup.json','verification-bad.json']:
   file=ROOT/name
   if file.exists():file.unlink()
