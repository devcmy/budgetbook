"""Local HTML rendering without network. UI storage uses an explicit test double."""
import html, json, re
from pathlib import Path
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).parent
HTML=(ROOT/'index.html').read_text().replace('<link rel="stylesheet" href="style.css">','<style>'+(ROOT/'style.css').read_text()+'</style>')
for name in ['data','settings','focus','app']:
    HTML=HTML.replace(f'<script src="{name}.js"></script>','<script>'+(ROOT/f'{name}.js').read_text()+'</script>')
shots=ROOT/'screenshots';shots.mkdir(exist_ok=True)
checks,errors,storage=[],[],{}
def content(tab='home',sheet=None,stored=None):
    result=HTML.replace("const initial=new URLSearchParams(location.search).get('tab')",'const initial='+json.dumps(tab))
    for kind in ['input','cash']:
        result=result.replace(f"if(new URLSearchParams(location.search).get('sheet')==='{kind}')",'if('+str(sheet==kind).lower()+')')
    if stored is not None:
        script='<script>window.TEST_STORAGE='+json.dumps(stored)+';Object.defineProperty(window,"localStorage",{value:{getItem(k){return window.TEST_STORAGE[k]??null},setItem(k,v){window.TEST_STORAGE[k]=String(v)}}});</script>'
        result=result.replace('<script>window.SNAPSHOT',script+'<script>window.SNAPSHOT')
    return result
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path='/usr/bin/chromium',args=['--no-sandbox','--disable-crashpad-for-testing'])
    page=browser.new_page(viewport={'width':1440,'height':1050},locale='ko-KR')
    def load(tab='home',sheet=None,preserve=True):
        global page,storage
        viewport=page.viewport_size
        if preserve and page.evaluate('Boolean(window.TEST_STORAGE)'):storage=page.evaluate('window.TEST_STORAGE')
        page.close();page=browser.new_page(viewport=viewport,locale='ko-KR',timezone_id='Asia/Seoul')
        page.set_default_timeout(6000);page.on('pageerror',lambda e:errors.append(str(e)))
        page.set_content(content(tab,sheet,storage))
    load()
    assert page.locator('.focus-card').count()==4
    assert '1,195,895' in page.locator('.hero-amount').inner_text()
    assert page.locator('#tabbar button').count()==5
    checks.append('5개 탭 및 이번달 네 지출 메뉴')
    page.locator('[data-focus="cash"]').click()
    assert '11,900원' in page.get_by_role('dialog').inner_text()
    assert '결산에 추가할 금액' in page.get_by_role('dialog').inner_text()
    assert page.get_by_role('dialog').get_by_text('0원',exact=True).count()==1
    page.locator('[data-transaction="2333"]').click()
    assert '변동지출에 이미 포함' in page.get_by_role('dialog').inner_text()
    page.keyboard.press('Escape')
    page.locator('[data-focus="reserveUse"]').click()
    assert '9월에 사용' in page.get_by_role('dialog').inner_text()
    page.keyboard.press('Escape')
    page.locator('[data-action="classification"]').click()
    assert '네 카드의 금액을 모두 더하지' in page.get_by_role('dialog').inner_text()
    page.keyboard.press('Escape')
    checks.append('비예비비 발생 11,900 / 이미 집계 11,900 / 추가 0; 예비비 사용월 보존')
    page.locator('[data-homeday="6"]').click()
    assert '미입력' in page.locator('.day-detail').inner_text()
    page.locator('[data-homeday="2"]').click()
    assert '30,500원' in page.locator('.day-detail').inner_text()
    checks.append('일별 원 단위 금액; 미래 날짜는 0원이 아닌 미입력')
    page.locator('#tabbar [data-tab="settings"]').click()
    page.locator('[data-rename="tab.home"]').click()
    page.locator('#alias-name').fill('   ')
    page.get_by_role('button',name='이름 저장',exact=True).click()
    assert page.locator('#rename-error').inner_text()
    page.locator('#alias-name').fill('내 월간장부')
    page.get_by_role('button',name='이름 저장',exact=True).click()
    assert '내 월간장부' in page.locator('#tabbar').inner_text()
    page.locator('summary').filter(has_text='이번달 메뉴').click()
    page.locator('[data-rename="monthly.cash"]').click()
    page.locator('#alias-name').fill('계좌 현금지출')
    page.get_by_role('button',name='이름 저장',exact=True).click()
    page.locator('[data-preference="reduceTransparency"]').check()
    assert page.locator('body').evaluate("el=>el.classList.contains('reduce-transparency')")
    page.locator('[data-preference="compact"]').uncheck()
    assert page.locator('body').evaluate("el=>el.classList.contains('comfortable')")
    load()
    assert '내 월간장부' in page.locator('#tabbar').inner_text()
    assert '계좌 현금지출' in page.locator('[data-focus="cash"]').inner_text()
    assert '1,195,895' in page.locator('.hero-amount').inner_text()
    page.locator('[data-focus="cash"]').click()
    assert page.get_by_role('dialog').get_attribute('aria-label')=='계좌 현금지출'
    page.keyboard.press('Escape')
    checks.append('탭·메뉴 별칭 입력 검증, 저장 API와 저장값 복원; 금액·분류 ID 불변')
    load('settings')
    page.get_by_role('button',name='이름을 기본값으로 복원',exact=True).click()
    page.get_by_role('button',name='이름만 기본값으로 복원',exact=True).click()
    assert '이번달' in page.locator('#tabbar').inner_text()
    assert page.locator('[data-preference="reduceTransparency"]').is_checked()
    checks.append('이름 복원은 화면 설정을 유지')
    page.locator('[data-preference="reduceTransparency"]').uncheck()
    page.locator('[data-preference="compact"]').check()
    load('records')
    page.locator('#search').fill('우지커피');assert page.locator('.transaction').count()==1
    page.locator('#search').fill('')
    page.get_by_role('button',name='거래 입력').click()
    page.locator('#amount').fill('1.5');page.locator('#entry-title').fill('정산 입력 검증')
    page.get_by_role('button',name='시안에 저장').click();assert '정수' in page.locator('#form-error').inner_text()
    page.locator('#amount').fill('-2500');page.get_by_role('button',name='시안에 저장').click()
    page.locator('.transaction').filter(has_text='정산 입력 검증').click()
    page.get_by_role('button',name='수정하기').click();page.locator('#amount').fill('-3000')
    page.get_by_role('button',name='수정 저장').click();page.locator('.transaction').filter(has_text='정산 입력 검증').click()
    assert '-3,000' in page.get_by_role('dialog').inner_text()
    page.get_by_role('button',name='시안 기록 삭제').click()
    assert page.locator('.transaction').filter(has_text='정산 입력 검증').count()==0
    checks.append('검색·음수 금액·수정·삭제 회귀 검증')
    for width,height in [(440,956),(390,844),(375,812),(320,568)]:
        page.set_viewport_size({'width':width,'height':height})
        for tab in ['home','records','reserve','balance','settings']:
            load(tab)
            assert not page.evaluate('document.documentElement.scrollWidth>innerWidth'),(width,tab)
            assert not page.evaluate('screen.scrollWidth>screen.clientWidth'),(width,tab)
            if width==440:page.screenshot(path=str(shots/f'{tab}-v2.png'))
        load('home','cash')
        if width==440:
            page.locator('.close').evaluate('el=>el.blur()');page.screenshot(path=str(shots/'cash-v2.png'))
        assert not page.evaluate('document.querySelector(".sheet").scrollWidth>document.querySelector(".sheet").clientWidth')
    checks.append('Pro Max 440×956 및 390/375/320px: 5탭·상세 가로 넘침 없음')
    page.set_viewport_size({'width':1470,'height':1160});load()
    page.screenshot(path=str(shots/'03-overview-v2.png'))
    board=(ROOT/'board.html').read_text()
    def frame(match):
        src=match.group(1);tab=re.search(r'tab=([a-z]+)',src);sheet=re.search(r'sheet=([a-z]+)',src)
        return 'srcdoc="'+html.escape(content(tab.group(1) if tab else 'home',sheet.group(1) if sheet else None),quote=True)+'"'
    page.set_content(re.sub(r'src="(index.html[^"]*)"',frame,board));page.wait_for_timeout(200)
    for f in page.frames:f.evaluate('document.activeElement?.blur()')
    page.screenshot(path=str(shots/'04-design-board-v2.png'),full_page=True)
    assert not errors,errors
    browser.close()
report={'version':2,'prototype_checks':checks,'browser_errors':errors,'settings_storage_test':'API boundary/restoration tested with in-memory Web Storage double; browser disk and iOS UserDefaults unverified','native_build':'not_run_linux_no_xcode','accounting_engine':'not_implemented_design_phase'}
(ROOT/'verification-v2.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
print(json.dumps(report,ensure_ascii=False,indent=2))
