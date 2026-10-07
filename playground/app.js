/* Offline functional companion. Accounting is exclusively window.LedgerCore. */
(function () {
'use strict';
const seed = JSON.parse(document.getElementById('seed-data').textContent);
const labels = JSON.parse(document.getElementById('labels-data').textContent);
const KEY = 'tetris-budget-functional-v1';
const clone = x => JSON.parse(JSON.stringify(x));
const esc = x => String(x ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const won = x => Number(x).toLocaleString('ko-KR');
const operations = {expense:'지출',income:'수입',refund:'환급·정산',reserveAllocate:'예비비 실제적립',reserveUse:'예비비 사용',reserveReturn:'예비비 → 생활비 환입',jointAllocate:'공동계좌 납입',jointActivity:'공동계좌 활동',investmentAllocate:'투자금 적립',investmentBuy:'투자 매수',cardPayment:'카드대금 납부',transfer:'계좌 이동',pending:'대조 차이·보류'};
const locations = {card:'카드',bank:'일반 계좌·현금',living:'WRAP 생활비',joint:'공동계좌'};
const statuses = {confirmed:'확정',unconfirmed:'미확인',planned:'예상·계획'};
const icons = {
 home:'<rect x="3" y="3" width="7" height="7" rx="1.8"/><rect x="14" y="3" width="7" height="7" rx="1.8"/><rect x="3" y="14" width="7" height="7" rx="1.8"/><rect x="14" y="14" width="7" height="7" rx="1.8"/>',
 records:'<rect x="4" y="3" width="16" height="18" rx="3"/><path d="M8 8h8M8 12h8M8 16h5"/>',
 reserve:'<rect x="3" y="8" width="18" height="4" rx="1"/><path d="M5 12v9h14v-9M12 8v13"/><path d="M12 8C6 8 5 3 8 3s4 5 4 5c6 0 7-5 4-5s-4 5-4 5Z"/>',
 balance:'<rect x="3" y="5" width="18" height="15" rx="3"/><path d="M17 11h4v5h-4a2 2 0 0 1 0-5Z"/>',
 settings:'<path d="m9 3 1-1h4l1 1 1 3 3 1 2 3v4l-2 3-3 1-1 3-1 1h-4l-1-1-1-3-3-1-2-3v-4l2-3 3-1Z"/><circle cx="12" cy="12" r="3"/>',
 plus:'<path d="M12 5v14M5 12h14"/>',down:'<path d="m6 9 6 6 6-6"/>',close:'<path d="m6 6 12 12M18 6 6 18"/>',arrow:'<path d="m9 5 7 7-7 7"/>',search:'<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6"/>',info:'<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7v.1"/>',income:'<path d="M12 3v12m-5-5 5 5 5-5M4 17v4h16v-4"/>',expense:'<path d="m3 3 2 2 3 12h10l3-9H6"/><circle cx="9" cy="21" r="1"/><circle cx="18" cy="21" r="1"/>',transfer:'<path d="M3 7h17m-5-5 5 5-5 5M21 17H4m5-5-5 5 5 5"/>',calendar:'<rect x="3" y="5" width="18" height="16" rx="3"/><path d="M7 3v4M17 3v4M3 11h18"/>'
};
const svg = key => `<svg viewBox="0 0 24 24" aria-hidden="true">${icons[key]||icons.records}</svg>`;
let ledger, summary, storageMode = 'persistent', lastSaved = null;
let tab='home', month='2026-10', query='', calendarMode=false, selectedDay=5, showAllReserves=false;
let modal=null, pendingRestore=null, pendingExcel=null, returnFocus=null, toastTimer;
const screen=document.getElementById('screen'), overlay=document.getElementById('overlay');
function checkedSettings(value) {
 if(value === undefined) return {names:{},compact:true,reduceTransparency:false};
 if(!value || typeof value!=='object' || !value.names || typeof value.names!=='object' || typeof value.compact!=='boolean' || typeof value.reduceTransparency!=='boolean') throw Error('백업의 화면 설정 형식을 확인해주세요.');
 const names={};
 for(const [key,raw] of Object.entries(value.names)) {
  const label=labels.find(l=>l.id===key);
  if(!label || typeof raw!=='string' || !raw.trim() || [...raw.trim()].length>label.maxLength || /[\u0000-\u001f\u007f-\u009f]/.test(raw)) throw Error('백업의 메뉴 이름을 확인해주세요.');
  names[key]=raw.trim();
 }
 return {names,compact:value.compact,reduceTransparency:value.reduceTransparency};
}
function validateBackup(state) {
 const result=JSON.parse(window.ledgerDispatch(JSON.stringify({action:'validate',state})));
 if(!result.ok) throw Error(result.error);
 result.value.personalization=checkedSettings(result.value.personalization);
 return result.value;
}
function initialize() {
 try { lastSaved=localStorage.getItem(KEY); }
 catch (_) {storageMode='memory';}
 if(lastSaved !== null) {
  try {ledger=validateBackup(JSON.parse(lastSaved));}
  catch(error) {ledger=null;startupError(error.message);return;}
 } else {
  ledger=clone(seed);ledger.personalization=checkedSettings(undefined);
  if(storageMode==='persistent') {
   try {lastSaved=JSON.stringify(ledger);localStorage.setItem(KEY,lastSaved);}
   catch (_) {storageMode='memory';lastSaved=null;}
  }
 }
 render();
}
function startupError(message) {
 screen.innerHTML=`<div class="top"><h2>가계부를 열 수 없어요</h2></div><div class="error">${esc(message)}</div><div class="info">저장된 파일을 자동 초기화하지 않았어요. 먼저 원본을 내보내거나 정상 백업을 가져오세요.</div><div class="stack"><button class="primary" data-action="export">저장된 원본 내보내기</button><button class="secondary" data-action="import">백업 가져오기</button><button class="danger" data-action="reset">테스트 변경 초기화</button></div>`;
 document.getElementById('tabbar').innerHTML='';
}
function settings(){return ledger?.personalization||checkedSettings(undefined);}
function name(key){return settings().names[key] || labels.find(l=>l.id===key)?.defaultName || key;}
function toast(message){clearTimeout(toastTimer);const t=document.getElementById('toast');t.textContent=message;t.classList.add('show');toastTimer=setTimeout(()=>t.classList.remove('show'),2800);}
function errorMessage(error){const el=overlay.querySelector('#form-error');if(el)el.textContent=error.message||String(error);else show('error',{message:error.message||String(error)});}
function commit(candidate,{targetMonth=month,message='저장했어요.',targetTab=tab}={}) {
 const valid=validateBackup(candidate);
 // Validate and calculate before persistence or publishing any UI changes.
 window.LedgerCore.summarize(valid,targetMonth);
 const bytes=JSON.stringify(valid);
 if(new TextEncoder().encode(bytes).length>32*1024*1024)throw Error('가계부 파일이 32MB를 초과했습니다.');
 if(storageMode==='persistent') {
  try {
   if(localStorage.getItem(KEY)!==lastSaved) throw Error('다른 창에서 변경됐어요. 백업한 뒤 새로고침해주세요.');
   localStorage.setItem(KEY,bytes);
  } catch(error) {throw Error(error.message.includes('다른 창') ? error.message : '이 화면 저장소에 쓰지 못했습니다. 변경 내용은 적용하지 않았어요. 먼저 백업을 저장해주세요.');}
  lastSaved=bytes;
 }
 ledger=valid;month=targetMonth;tab=targetTab;hide();render();toast(storageMode==='memory'?'현재 테스트 화면에 적용했어요. 백업으로 보관할 수 있어요.':message);
}
function mutation(change,targetMonth=month) {commit(window.LedgerCore.mutate(ledger,change),{targetMonth,message:'거래를 저장했어요.'});}
function header(key,subtitle='') {return `<div class="top"><div><h2>${esc(name(key))}</h2><p class="sub">${esc(subtitle)}</p></div><div class="top-actions"><button class="round-btn glass-control toolbar-help" data-action="help" aria-label="테스트 예시">${svg('info')}</button><button class="round-btn glass-control" data-action="add" aria-label="거래 입력">${svg('plus')}</button></div></div>`;}
function row(title,amount,emphasis=false){return `<div class="row${emphasis?' emphasized':''}"><span>${esc(title)}</span><span class="value">${won(amount)}원</span></div>`;}
function badge(status){return `<span class="chip ${status==='confirmed'?'confirmed':status==='unconfirmed'?'unchecked':'expected'}">${esc(statuses[status]||status)}</span>`;}
function record(t) {
 const incoming=['income','refund'].includes(t.operation)||(t.operation==='jointActivity'&&t.sourceKind==='수입');
 const spending=['expense','reserveUse'].includes(t.operation)||(t.operation==='jointActivity'&&t.sourceKind==='지출');
 const amount=incoming?t.amount:-t.amount;
 return `<button class="transaction" data-record="${t.id}"><span class="icon ${incoming?'income':spending?'':'transfer'}">${svg(incoming?'income':spending?'expense':'transfer')}</span><span class="desc"><span class="title">${esc(t.title)}</span><span class="meta">${esc(t.category)} · ${esc(t.account)}</span>${t.status!=='confirmed'?`<span class="record-status">${esc(statuses[t.status])}</span>`:''}</span><span class="amount ${incoming?'income-text':spending?'':'transfer-text'}">${incoming&&amount>=0?'+':''}${won(amount)}원</span></button>`;
}
function cardList(entries,empty='확정 기록이 없어요.') {return entries.length?`<div class="card list-card">${entries.map(record).join('')}</div>`:`<div class="empty-state">${esc(empty)}</div>`;}
function variableList() {
 const ids=new Set(summary.variableRecordIDs),groups=new Map();
 const entries=ledger.transactions.filter(t=>ids.has(t.id)).sort((a,b)=>(b.date||'').localeCompare(a.date||'')||b.id-a.id);
 for(const t of entries){const day=t.date?.slice(0,10)||'';if(!groups.has(day))groups.set(day,[]);groups.get(day).push(t);}
 function transaction(t){
  const raw=t.sourceKey?ledger.sourceArchive.find(r=>r.importKey===t.sourceKey||(!r.importKey&&r.source_row===t.id&&t.sourceKey.endsWith(':'+r.source_row))):null;
  const originalCategory=raw&&(['변동비','비고정비','고정비'].includes(t.category)||t.category===raw['분류']);
  const category=originalCategory?(raw['분류']||t.category):t.category,subcategory=originalCategory?raw['소분류']:'';
  const originalNote=raw?[raw['추가입력'],raw['비고']].filter(Boolean).join(' · '):null;
  const memo=(raw&&t.note===originalNote?raw['추가입력']:t.note)||'';
  return `<button class="variable-transaction" data-record="${t.id}"><span class="variable-category"><span>${esc(category)}</span>${subcategory?`<small>${esc(subcategory)}</small>`:''}</span><span class="variable-description"><span class="variable-title">${esc(t.title)}${memo?` <small>${esc(memo)}</small>`:''}</span><span class="variable-account">${esc(t.account)}</span></span><span class="variable-amount ${t.amount<0?'credit':''}">${won(t.amount)}원</span></button>`;
 }
 return `<div class="variable-overview"><div><p>${esc(month.replace('-','년 '))}월 · 확정 지출</p><div class="amount-large">${won(summary.monthlyFocus.variable)}<small>원</small></div></div><span>${entries.length}건</span></div><div class="variable-ledger">${[...groups].map(([date,items])=>{
  const weekday=date?new Date(date+'T00:00:00Z').getUTCDay():null,total=items.reduce((sum,t)=>sum+t.amount,0);
  return `<section class="variable-day" data-date="${date}"><h4 class="variable-day-heading"><span class="variable-date">${date?`<strong>${Number(date.slice(8))}</strong><span class="variable-weekday ${weekday===0?'sunday':weekday===6?'saturday':''}">${['일','월','화','수','목','금','토'][weekday]}요일</span>`:'<strong class="unknown-day">날짜 미확인</strong>'}</span><span class="variable-day-total ${total<0?'credit':''}"><small>일 합계</small>${won(total)}원</span></h4>${items.map(transaction).join('')}</section>`;
 }).join('')||'<div class="empty-state">확정 기록이 없어요.</div>'}</div><div class="info">확정된 비고정비만 표시해요. 계좌 현금성 지출·예비비 사용이 포함된 경우 같은 거래를 다시 더하지 않아요. 거래를 누르면 전체 메모를 확인하고 수정할 수 있어요.</div>`;
}
function calendar(finance=false) {
 const [y,m]=month.split('-').map(Number),first=new Date(Date.UTC(y,m-1,1));
 const offset=(first.getUTCDay()+6)%7,count=new Date(Date.UTC(y,m,0)).getUTCDate();
 selectedDay=Math.min(selectedDay,count);
 return `<div class="card calendar-wrap"><div class="calendar-grid ${finance?'finance-calendar':''}">${['월','화','수','목','금','토','일'].map(d=>`<span class="weekday">${d}</span>`).join('')}${'<span class="empty"></span>'.repeat(offset)}${Array.from({length:count},(_,i)=>{const day=i+1,amount=summary.dailyAmounts.find(a=>a.day===day),has=summary.transactions.some(t=>t.date.startsWith(month+'-'+String(day).padStart(2,'0')));return `<button class="${selectedDay===day?'selected':''}" data-day="${day}" aria-label="${day}일"><span>${day}</span>${finance?`<small>${amount?won(amount.actual):'—'}</small>`:`<span class="dot" style="opacity:${has?1:0}"></span>`}</button>`;}).join('')}</div></div>`;
}
function home() {
 const f=summary.monthlyFocus;
 const has=summary.hasRecords||summary.hasForecast;
 let result=`<div class="top compact-top"><div class="eyebrow">${esc(name('tab.home'))} · 저장된 기록 기준</div><div class="top-actions"><button class="round-btn glass-control toolbar-help" data-action="help" aria-label="테스트 예시">${svg('info')}</button><button class="round-btn glass-control" data-action="add" aria-label="거래 입력">${svg('plus')}</button></div></div><div class="month-navigation"><button class="month-step month-prev glass-control" data-month-step="-1" aria-label="이전 달" ${month==='0001-01'?'disabled':''}>${svg('arrow')}</button><button class="month-title" data-modal="month" aria-label="조회월 선택">${Number(month.slice(0,4))}년 ${Number(month.slice(5))}월</button><button class="month-step glass-control" data-month-step="1" aria-label="다음 달" ${month==='9999-12'?'disabled':''}>${svg('arrow')}</button></div>`;
 if(summary.historicalPending)result+=`<div class="info">원본 RAW ${summary.archivedCount}건 보관 중 · 과거 결산은 미이관입니다. 표시 금액은 앱에 등록한 운영 거래 기준이에요. <button data-modal="archive" class="inline-link">원문 보기</button></div>`;
 if(!has)result+=`<div class="card"><h2>${summary.historicalPending?'운영 결산 미이관':'아직 미입력'}</h2><p class="muted">실제 거래와 예정 금액은 따로 관리해요.</p>${row('예비비 적립계획',summary.reserveMonthPlan)}<button class="primary" data-action="add">첫 기록 입력하기</button></div>`;
 else {
  result+=`<button class="hero compact-summary" data-modal="settlement"><div class="hero-main"><div class="hero-heading"><span class="hero-name">${esc(name('monthly.remaining'))}</span><span class="chip">${summary.hasForecast?'예상 포함':'확정 기록'}</span></div><div class="hero-amount" data-value="remaining">${won(summary.remaining)} <span class="unit">원</span></div></div><div class="compact-stats"><div class="stat"><small>${esc(name('monthly.income'))}</small><strong class="value">${won(summary.income)}원</strong></div><div class="stat"><small>${esc(name('monthly.outflow'))}</small><strong class="value">${won(summary.outflow)}원</strong></div></div></button>`;
  result+=`<div class="home-focus">${[
   ['variable',f.variable,'달력에 기록된 비고정비'],['fixed',f.fixed,'예상·미확인과 실제를 구분'],['cash',f.unplannedCash,'계좌에서 쓴 계획 외 현금'],['reserveUse',f.reserveUse,'목적자금의 이번달 실제 사용']
  ].map(([key,amount,hint])=>`<button data-modal="${key}" data-focus="${key}"><span class="focus-label">${esc(name('monthly.'+key))}${key==='fixed'&&summary.hasForecast?'<span class="chip expected">예상</span>':svg('arrow')}</span><strong>${won(amount)}원</strong><small>${hint}</small></button>`).join('')}</div>`;
  result+=`<div class="record-day-controls muted"><span>공동계좌 적립 ${won(f.jointAllocation)}</span><span>환급·할인 −${won(f.refund)}</span><button data-modal="classification" aria-label="집계 기준" class="toolbar-help">${svg('info')}</button></div>`;
  result+=`<div class="section-head"><h3>${esc(name('monthly.calendar'))}</h3><span class="muted">${esc(name('monthly.variable'))} · 원 단위</span></div>${calendar(true)}`;
  const dayEntries=ledger.transactions.filter(t=>summary.variableRecordIDs.includes(t.id)&&t.date?.startsWith(month+'-'+String(selectedDay).padStart(2,'0'))),daily=summary.dailyAmounts.find(d=>d.day===selectedDay);
  result+=`<div class="record-day-controls"><span>${Number(month.slice(5))}월 ${selectedDay}일</span><strong>${daily?won(daily.actual)+'원':'미입력'}</strong></div>${cardList(dayEntries.slice(0,3),'아직 실제 기록이 없는 날이에요.')}${dayEntries.length>3?`<button class="inline-link" data-action="day-records">이 날 ${dayEntries.length}건 전체 보기</button>`:''}`;
  result+=`<button class="tetris-card compact-tetris" data-modal="tetris"><span>${esc(name('monthly.tetris'))}</span><span><b>${summary.tetrisPending?'검토 대기':won(summary.tetrisRemaining)+'원'}</b><small>${summary.tetrisPending?'표시 규칙 확인 필요':'예상 · 차이 '+won(summary.remaining-summary.tetrisRemaining)+'원'}</small></span>${svg('arrow')}</button>`;
 }
 result+=`<div class="balance-links"><button data-tab="reserve"><small>${esc(name('tab.reserve'))} 잔액</small><strong>${won(summary.reserve)}원</strong></button><button data-tab="balance"><small>WRAP 잔액</small><strong>${won(summary.wrap)}원</strong></button></div><p class="preview-help">상단 ⓘ에서 테스트 예시를 입력할 수 있어요.</p>`;
 return result;
}
function filteredRecords() {return ledger.transactions.filter(t=>t.month===month&&(query===''||[t.title,t.account,t.category,t.note].some(v=>String(v).toLocaleLowerCase().includes(query.toLocaleLowerCase())))&&(!calendarMode||t.date?.startsWith(month+'-'+String(selectedDay).padStart(2,'0')))).sort((a,b)=>(b.date||'').localeCompare(a.date||'')||b.id-a.id);}
function moveMonth(delta) {
 const [year,current]=month.split('-').map(Number),index=year*12+current-1+delta,y=Math.floor(index/12),m=index%12+1;
 if(y<1||y>9999)return;
 const next=String(y).padStart(4,'0')+'-'+String(m).padStart(2,'0');
 window.LedgerCore.summarize(ledger,next);
 month=next;query='';
 const days=[31,y%4===0&&(y%100!==0||y%400===0)?29:28,31,30,31,30,31,31,30,31,30,31];
 selectedDay=Math.min(selectedDay,days[m-1]);render();screen.scrollTop=0;
}
function recordGroups() {
 const entries=filteredRecords(),days=[...new Set(entries.map(t=>t.date?t.date.slice(0,10):t.month+' · 날짜 미확인'))];
 const views=new Map(summary.transactions.map(t=>[t.id,t]));
 return !entries.length?'<div class="empty-state">찾는 기록이 없어요. 검색어나 날짜를 바꿔보세요.</div>':days.map(day=>{
  const items=entries.filter(t=>(t.date?t.date.slice(0,10):t.month+' · 날짜 미확인')===day),income=items.reduce((sum,t)=>sum+views.get(t.id).recordIncome,0),expense=items.reduce((sum,t)=>sum+views.get(t.id).recordExpense,0),unknown=day.includes('날짜 미확인'),weekday=unknown?'':['일','월','화','수','목','금','토'][new Date(day+'T00:00:00Z').getUTCDay()]+'요일';
  return `<section class="record-day" data-date="${esc(day)}"><div class="record-day-heading"><div><strong>${unknown?'날짜 미확인':Number(day.slice(8))}</strong> <small>${weekday} · ${items.length}건</small></div><div class="record-day-totals"><span class="income-text">수입 <b>${won(income)}원</b></span><span class="expense-text">지출 <b>${won(expense)}원</b></span></div></div>${cardList(items)}</section>`;
 }).join('');
}
function records() {return header('tab.records',`${month.replace('-','년 ')}월 · ${summary.transactions.length}건`)+`<button class="inline-link" data-modal="archive">${esc(name('common.archive'))} · 전체 이력</button><div class="segmented"><button data-mode="list" class="${calendarMode?'':'active'}">${esc(name('records.list'))}</button><button data-mode="calendar" class="${calendarMode?'active':''}">${esc(name('records.calendar'))}</button></div>${calendarMode?calendar():''}<label class="search">${svg('search')}<input id="search" aria-label="거래 검색" placeholder="내용, 분류, 결제수단 검색" value="${esc(query)}"></label><p class="record-totals-help">표시된 기록 기준 · 확정 개인 수입·소비의 일 합계입니다. 환급·적립·이체·카드 납부·공동 거래는 합계에서 제외하고 내역에 표시합니다.</p><div id="record-groups">${recordGroups()}</div>`;}
function reserve() {
 return header('tab.reserve','계획과 실제로 모인 돈을 따로')+`<div class="card"><div class="card-label">남은 금액 · 전체 운영 기록 기준</div><div class="card-amount" data-value="reserve">${won(summary.reserve)} 원</div><div class="reserve-stats"><div><small>누적 실제적립</small><strong>${won(summary.reserveAllocated)}원</strong></div><div><small>누적 목적자금 사용·환입</small><strong>${won(summary.reserveUsed)}원</strong></div></div></div><div class="card">${row(Number(month.slice(5))+'월 적립계획',summary.reserveMonthPlan)}${row('이번달 실제적립',summary.reserveMonthActual)}<p class="muted">계획 금액은 아직 잔액에 포함되지 않았어요.</p><div class="action-grid"><button class="secondary" data-case="reserveAllocate">실제적립 입력</button><button class="secondary" data-case="reserveUse">사용 입력</button></div></div><div class="section-head"><h3>${esc(name('reserve.purposes'))}</h3></div><div class="segmented"><button data-reserves="remaining" class="${showAllReserves?'':'active'}">${esc(name('reserve.remaining'))}</button><button data-reserves="all" class="${showAllReserves?'active':''}">${esc(name('reserve.all'))}</button></div>${summary.reserves.filter(r=>showAllReserves||r.balance>0).map(r=>`<button class="purpose" data-reserve="${r.id}"><div class="purpose-head"><h4>${esc(r.title)}</h4><strong>${won(r.balance)}원</strong></div><div class="progress"><i style="width:${r.allocated?Math.max(0,Math.min(100,r.balance/r.allocated*100)):0}%"></i></div><div class="purpose-group"><span>${esc(r.group)}</span><span>${Number(month.slice(5))}월 계획 ${won(r.plan)}원</span></div></button>`).join('')}<div class="info">남은 금액 = 실제적립 − 목적자금 사용·환입. 실제로 사용한 달의 결산에 한 번 반영해요.</div>`;
}
function balance() {
 const groups=[...new Set(summary.reserves.map(r=>r.group))], recent=summary.wrapRecordIDs.slice(0,5).map(id=>ledger.transactions.find(t=>t.id===id));
 return header('tab.balance','미래에셋 WRAP · 전체 운영 기록 기준')+`<div class="card"><div class="card-label">WRAP 총잔액</div><div class="card-amount" data-value="wrap">${won(summary.wrap)} 원</div><div class="balance-bar"><i style="width:${summary.wrap?summary.living/summary.wrap*100:0}%"></i></div><div class="legend"><span>생활비 ${summary.wrap?(summary.living/summary.wrap*100).toFixed(1):0}%</span><span>목적자금 ${summary.wrap?(summary.reserve/summary.wrap*100).toFixed(1):0}%</span></div></div><div class="section-head"><h3>${esc(name('balance.places'))}</h3><button class="text-btn" data-case="reserveAllocate">${esc(name('balance.movement'))} ${svg('arrow')}</button></div><p class="muted balance-explanation">같은 WRAP 계좌의 돈을 생활비와 예비비 목적자금으로 나눈 구성이에요. 별도 통장 잔액을 더한 값이 아닙니다.</p><button class="card row-button" data-modal="living">${row(name('balance.living'),summary.living,true)}<small class="muted">목적을 정하지 않은 생활 자금</small></button><div class="card">${row(name('balance.purposes'),summary.reserve,true)}${groups.map(group=>`<button class="row-button" data-purpose="${esc(group)}">${row(group,summary.reserves.filter(r=>r.group===group).reduce((s,r)=>s+r.balance,0))}</button>`).join('')}</div><div class="section-head"><h3>${esc(name('balance.recent'))}</h3></div>${cardList(recent)}<div class="info">생활비와 목적자금 사이 이동은 총잔액을 바꾸지 않아요. 일반 계좌·카드 잔액은 WRAP에 더하지 않아요.</div>`;
}
function nameRow(label){return `<button class="name-row" data-rename="${esc(label.id)}"><small>${esc(label.defaultName)}</small><span>${esc(name(label.id))} ›</span></button>`;}
function planOverview() {
 const plans=ledger.plans.filter(p=>p.month===month);
 const types={income:'수입 예상',fixed:'고정비 예상',reserve:'예비비 적립계획',familyContribution:'가족계좌 납입 · 중단'};
 return `<p class="muted">${esc(month.replace('-','년 '))}월 · 등록 계획액입니다. 실제 잔액은 바뀌지 않아요. 종류를 눌러 항목을 확인·수정하세요.</p><div class="plan-overview">${Object.entries(types).filter(([type])=>type!=='familyContribution'||plans.some(p=>p.type===type)).map(([type,title])=>{
  const items=plans.filter(p=>p.type===type),total=type==='familyContribution'?0:items.reduce((sum,p)=>sum+p.amount,0),unchecked=items.filter(p=>p.status==='미확인').length;
  return `<details class="card plan-group" data-plan-type="${type}"><summary><span><strong>${title}</strong><small>${items.length}개 항목${unchecked?' · 미확인 '+unchecked+'개':''}</small></span><span class="plan-total">${won(total)}<small>원 · 펼치기</small></span></summary><div class="plan-items">${items.map(p=>{
   const linked=ledger.transactions.filter(t=>t.planID===p.id),actual=linked.filter(t=>t.status==='confirmed').reduce((sum,t)=>sum+t.amount,0);
   return `<button class="plan-item" data-plan="${esc(p.id)}"><span><strong>${esc(p.title)}</strong><small>${esc(p.account)} · ${type==='familyContribution'?'중단':esc(p.status)}${linked.length?' · 실제 연결 '+won(actual)+'원':''}</small></span><span>${won(type==='familyContribution'?0:p.amount)}원 ${svg('arrow')}</span></button>`;
  }).join('')||'<p class="muted">등록한 항목이 없어요. 추가 버튼으로 등록하세요.</p>'}</div></details>`;
 }).join('')}</div>`;
}
function settingsScreen() {
 const groups=[...new Set(labels.filter(l=>!l.id.startsWith('tab.')).map(l=>l.group))];
 return header('tab.settings','내 가계부는, 내 이름으로.')+`<div class="card"><strong>나의 테트리스 가계부</strong><p class="muted">iPhone 17 Pro Max · 기능 테스트 화면</p><p class="storage-status ${storageMode==='memory'?'memory':''}">${storageMode==='memory'?'현재 화면에만 유지 · 백업으로 보관하세요.':'이 브라우저에 자동 저장됩니다.'}</p></div><div class="card"><h3 class="section-title">${esc(name('settings.backup'))}</h3><p class="muted">엑셀은 거래 추가 · JSON은 전체 백업 복원입니다.</p><div class="stack"><button class="secondary" data-action="export">${esc(name('settings.export'))}</button><button class="secondary" data-action="import">${esc(name('settings.import'))}</button></div><p class="muted">운영 거래 ${ledger.transactions.length}건 · 원본 RAW ${ledger.sourceArchive.length}건</p></div><div class="section-head"><h3>${esc(name('settings.plans'))}</h3><button class="text-btn" data-action="add-plan">추가 ${svg('plus')}</button></div>${planOverview()}<h3 class="section-title">탭 이름</h3><div class="card">${labels.filter(l=>l.id.startsWith('tab.')).map(nameRow).join('')}</div><h3 class="section-title">메뉴 이름</h3>${groups.map(group=>`<details class="card"><summary>${esc(group)}</summary>${labels.filter(l=>l.group===group).map(nameRow).join('')}</details>`).join('')}<div class="card"><label class="check"><input type="checkbox" data-pref="compact" ${settings().compact?'checked':''}> 촘촘한 보기</label><label class="check"><input type="checkbox" data-pref="reduceTransparency" ${settings().reduceTransparency?'checked':''}> 투명 효과 줄이기</label></div><div class="stack"><button class="secondary" data-action="reset-names">이름만 기본값으로 복원</button><button class="secondary" data-action="help">테스트 예시</button><button class="danger" data-action="reset">테스트 변경 초기화</button></div><p class="preview-help">웹에서 입력·계산·관리 기능을 확인하는 화면입니다. 기기 전용 Glass·키보드·파일 보호는 iPhone에서 확인합니다.</p>`;
}
function render() {
 if(!ledger)return;
 summary=window.LedgerCore.summarize(ledger,month);
 document.body.classList.toggle('comfortable',!settings().compact);document.body.classList.toggle('reduce-transparency',settings().reduceTransparency);
 document.querySelectorAll('.storage-status').forEach(el=>{el.textContent=storageMode==='memory'?'현재 화면에만 유지 · 백업으로 보관':'이 브라우저에 자동 저장';el.classList.toggle('memory',storageMode==='memory');});
 screen.innerHTML=({home,records,reserve,balance,settings:settingsScreen}[tab])();
 document.getElementById('tabbar').innerHTML=['home','records','reserve','balance','settings'].map(t=>`<button data-tab="${t}" class="${tab===t?'active':''}" aria-current="${tab===t?'page':'false'}">${svg(t)}<span class="tab-name">${esc(name('tab.'+t))}</span></button>`).join('');
}
function show(type,data={}) {returnFocus=document.activeElement;modal={type,...data};drawModal();}
function hide() {overlay.innerHTML='';modal=null;pendingRestore=null;if(returnFocus?.isConnected)returnFocus.focus();}
function field(label,html,wide=false) {return `<label class="field ${wide?'wide':''}">${esc(label)}${html}</label>`;}
function input(key,value,type='text',extra=''){return `<input name="${key}" id="${key}" type="${type}" value="${esc(value??'')}" ${extra}>`;}
function select(key,choices,value){return `<select name="${key}" id="${key}">${Object.entries(choices).map(([k,v])=>`<option value="${esc(k)}" ${String(value??'')===k?'selected':''}>${esc(v)}</option>`).join('')}</select>`;}
function freshEntry(overrides={}) {return Object.assign({id:0,date:month+'-06T12:00:00',month,title:'',account:'개인 계좌',category:'식비',amount:0,operation:'expense',status:'confirmed',funding:'bank',destination:'bank',expenseClass:'variable',reserveID:null,linkedID:null,attributedMonth:null,alreadyNet:false,displayAmount:null,planID:null,note:'',sourceKey:null,sourceKind:null},overrides);}
function editor(entry,ui) {
 const e=entry,u=ui||{},operation=u.operation??e.operation,klass=u.expenseClass??e.expenseClass,unknown=u.unknownDate??(e.date===null);
 const date=u.date??e.date?.slice(0,16)??e.month+'-06T12:00',m=unknown?(u.month??e.month):date.slice(0,7),reserveNeeded=['reserveAllocate','reserveUse','reserveReturn'].includes(operation);
 const plans=ledger.plans.filter(p=>p.month===m&&((operation==='income'&&p.type==='income')||(operation==='expense'&&klass==='fixed'&&p.type==='fixed')));
 const related=ledger.transactions.filter(t=>t.id!==e.id&&['expense','reserveUse'].includes(t.operation)&&(operation!=='cardPayment'||t.funding==='card')).sort((a,b)=>(b.date||'').localeCompare(a.date||''));
 return `<form id="entry-form" novalidate><div class="card">${field('거래 종류',select('operation',operations,operation))}${field('금액 (원)',input('amount',u.amount??(e.id===0&&e.amount===0?'':e.amount),'text','inputmode="decimal" class="amount-input" autocomplete="off" placeholder="0"'))}${field('내용',input('title',u.title??e.title,'text','maxlength="120" autocomplete="off"'))}${field('상태',select('status',statuses,u.status??e.status))}<p class="muted">확정 거래만 실제 금액과 잔액에 반영합니다.</p></div><div class="card"><label class="check"><input type="checkbox" name="unknownDate" id="unknownDate" ${unknown?'checked':''}> 날짜 미확인 · 월만 기록</label>${unknown?field('거래 월',input('month',u.month??e.month,'month')):field('실제 거래일',input('date',date,'datetime-local'))}${field('계좌·카드 이름',input('account',u.account??e.account,'text','maxlength="100"'))}${field('분류',input('category',u.category??e.category,'text','maxlength="100"'))}${!['reserveUse','reserveReturn','investmentBuy','pending'].includes(operation)?field(['income','refund'].includes(operation)?'입금 위치':'출금 위치',select('funding',locations,u.funding??e.funding)):''}${operation==='jointActivity'?field('공동계좌 입출금',select('sourceKind',{'지출':'지출','수입':'수입'},u.sourceKind??e.sourceKind??'지출')):''}${operation==='transfer'?field('도착 위치',select('destination',{bank:'일반 계좌',living:'WRAP 생활비'},u.destination??e.destination)):''}${reserveNeeded?field('예비비 목적',select('reserveID',{'':'선택해주세요',...Object.fromEntries(summary.reserves.map(r=>[r.id,r.title]))},u.reserveID??e.reserveID))+`<p class="muted">${operation==='reserveUse'?'목적자금에서 직접 지급한 금액을 기록해요.':'계획과 실제 이동은 별개입니다. 생활비에서 적립하면 WRAP 총잔액은 유지돼요.'}</p>`:''}</div>${['expense','reserveUse'].includes(operation)?`<div class="card">${field('지출 성격',select('expenseClass',{variable:name('monthly.variable'),fixed:name('monthly.fixed')},klass))}<label class="check"><input type="checkbox" name="alreadyNet" ${u.alreadyNet??e.alreadyNet?'checked':''}> 이미 정산을 뺀 본인부담 순액</label><p class="muted">순액이면 연결된 환급을 다시 차감하지 않아요.</p></div>`:''}${plans.length?`<div class="card">${field('대체할 예상 항목',select('planID',{'':'새 거래 · 예상과 별도',...Object.fromEntries(plans.map(p=>[p.id,p.title+' · '+won(p.amount)+'원']))},u.planID??e.planID))}<p class="muted">같은 예상 항목을 선택하면 실제 금액으로 대체해요.</p></div>`:''}${['refund','cardPayment'].includes(operation)?`<div class="card">${field('원래 지출',select('linkedID',{'':operation==='refund'?'월 환급·할인 (연결 없음)':'사용 내역 선택',...Object.fromEntries(related.map(t=>[t.id,t.month+' · '+t.title+' · '+won(t.amount)+'원']))},u.linkedID??e.linkedID))}${operation==='refund'?field('결산 귀속 월 · 다른 달 정산은 직접 선택',input('attributedMonth',u.attributedMonth??e.attributedMonth,'month')):'<p class="muted">소비를 다시 세지 않고 납부일의 계좌 출금만 반영해요.</p>'}</div>`:''}<div class="card">${field('메모',`<textarea name="note" id="note" maxlength="4000">${esc(u.note??e.note)}</textarea>`)}${field('테트리스 표시 금액 (선택)',input('displayAmount',u.displayAmount??e.displayAmount,'text','inputmode="numeric"'))}<p class="muted">임의로 반올림하지 않아요. 변경한 월 테트리스 결산은 검토 대기로 표시합니다.</p></div><div id="form-error" class="error" role="alert"></div><button type="submit" class="primary">${esc(name('entry.save'))}</button></form>`;
}
function readEditorUI(form) {const data=Object.fromEntries(new FormData(form));data.unknownDate=form.elements.unknownDate?.checked??false;data.alreadyNet=form.elements.alreadyNet?.checked??false;return data;}
function entryFromForm(form) {
 const u=readEditorUI(form),e=clone(modal.entry),number=String(u.amount??'').replace(/,/g,'').trim();
 if(!/^-?\d+$/.test(number)||!Number.isSafeInteger(Number(number))||Math.abs(Number(number))>9999999999)throw Error('원 단위 정수 금액을 입력해주세요.');
 const nullableInt=value=>value===''||value===undefined?null:Number(value);
 Object.assign(e,{title:String(u.title||'').trim(),amount:Number(number),operation:u.operation,status:u.status,account:u.account,category:u.category,note:u.note,funding:u.funding??e.funding,destination:u.destination??e.destination,expenseClass:u.expenseClass??e.expenseClass,alreadyNet:u.alreadyNet,reserveID:['reserveAllocate','reserveUse','reserveReturn'].includes(u.operation)?nullableInt(u.reserveID):null,planID:u.planID||null,linkedID:['refund','cardPayment'].includes(u.operation)?nullableInt(u.linkedID):null,attributedMonth:u.operation==='refund'?(u.attributedMonth||null):null,sourceKind:u.sourceKind??e.sourceKind});
 if(u.unknownDate){e.month=u.month;e.date=null;}else{e.month=String(u.date||'').slice(0,7);e.date=e.date?.slice(0,16)===u.date?e.date:u.date?.length===16?u.date+':00':u.date;}
 if(!u.displayAmount)e.displayAmount=null;else{if(!/^\d+$/.test(u.displayAmount)||!Number.isSafeInteger(Number(u.displayAmount)))throw Error('테트리스 표시 금액을 확인해주세요.');e.displayAmount=Number(u.displayAmount);}
 return e;
}
function openCase(operation,reserveID=22) {
 const values={expense:{title:'기능 테스트 · 계좌 지출',amount:10000,category:'식비',account:'미래에셋_WRAP형CMA',funding:'living'},reserveAllocate:{title:'여행비 실제적립',amount:100000,funding:'living',account:'미래에셋_WRAP형CMA',reserveID,category:'예비비 적립'},reserveUse:{title:'여행 숙박비',amount:40000,reserveID,category:'여행비'},income:{title:'실제 월급',amount:3000000,category:'급여',planID:ledger.plans.find(p=>p.month===month&&p.type==='income'&&p.id==='oct-salary')?.id??null}};
 show('entry',{entry:freshEntry({operation,...values[operation]})});
}
function details(entry) {return `<div class="detail-meta">${esc(entry.date?.replace('T',' ')||entry.month+' · 날짜 미확인')}</div><div class="amount-large">${won(['income','refund'].includes(entry.operation)||(entry.operation==='jointActivity'&&entry.sourceKind==='수입')?entry.amount:-entry.amount)} 원</div><div class="card">${rowText('내용',entry.title)}${rowText('거래 종류',operations[entry.operation])}${rowText('분류',entry.category)}${rowText('계좌·카드',entry.account)}${rowText('위치',locations[entry.funding]||entry.funding)}${rowText('상태',statuses[entry.status])}${entry.planID?rowText('대체한 예상 항목',ledger.plans.find(p=>p.id===entry.planID)?.title||entry.planID):''}${entry.displayAmount!==null?row('직접 지정 표시 금액',entry.displayAmount):''}</div>${entry.note?`<div class="info">${esc(entry.note)}</div>`:''}${entry.operation==='jointAllocate'?'<div class="info">개인 납입을 집계하고 이후 공동계좌 결제·수입은 다시 세지 않아요.</div>':''}${entry.operation==='expense'&&['living','bank'].includes(entry.funding)?'<div class="info">계좌 현금성 지출로 표시하지만 성격별 지출에 이미 포함해 추가 합산하지 않아요.</div>':''}<div class="stack"><button class="primary" data-edit="${entry.id}">${esc(name('entry.edit'))}</button><button class="danger" data-delete="${entry.id}">${esc(name('entry.delete'))}</button></div><p class="preview-help">원본 RAW와 수정·삭제 이력은 별도로 보관합니다.</p>`;}
function rowText(label,value){return `<div class="row"><span class="muted">${esc(label)}</span><span class="value" style="text-align:right;overflow-wrap:anywhere">${esc(value)}</span></div>`;}
function settlement() {const f=summary.monthlyFocus;return `<div class="info">남는 돈은 총수입 − 지출·적립이에요. 통장 잔액과 구분합니다. 미확인 고정비·예상 수입은 확정 합계에서 제외해요.</div><div class="card">${row('확정 기록 수입',summary.actualIncome)}${row('확정 지출·적립',summary.actualOutflow)}${row('확정 기록 남는 돈',summary.actualRemaining,true)}</div><div class="card">${row('총수입 · 예상 포함',summary.income)}${row(name('monthly.fixed'),f.fixed)}${row(name('monthly.variable'),f.variable)}${row('예비비 실제적립',f.reserveAllocation)}${row('공동계좌 적립',f.jointAllocation)}${row('투자금 적립',f.investmentAllocation)}${row('환급·할인 차감',-f.refund)}${row('지출·적립 합계',summary.outflow,true)}${row('남는 돈 · 예상 포함',summary.remaining,true)}</div><button class="primary" data-modal="fixed">${esc(name('monthly.fixed'))} 확인하기</button>`;}
function focusDetail(key) {
 const f=summary.monthlyFocus;
 if(key==='fixed')return `<p class="note">항목을 눌러 실제 결제로 바꾸세요. 같은 예상 항목을 대체해서 두 번 더하지 않아요.</p><div class="card">${summary.fixed.map((p,i)=>`<button class="row-button" data-fixed="${i}"><div class="row"><span>${esc(p.title)}<small>${esc(p.account)}</small></span><span class="value">${won(p.amount)}원<br>${badge(p.status==='확정'?'confirmed':p.status==='미확인'?'unconfirmed':'planned')}</span></div></button>`).join('')||'<p class="muted">이번달 고정비가 없어요.</p>'}</div>`;
 if(key==='cash')return `<p class="note">예비비 계획 밖의 계좌 현금성 지출</p><div class="amount-large">${won(f.unplannedCash)} 원</div>${cardList(summary.cashItems.map(c=>ledger.transactions.find(t=>t.id===c.sourceRow)))}<div class="card">${row('발생한 비예비비',f.unplannedCash)}${row('이미 고정·비고정비에 포함',f.cashAlreadyCounted)}${row('결산에 추가할 금액',f.cashAdditional,true)}</div><div class="info">카드 결제·예비비 사용·공동계좌 납입을 구분해요. 비예비비를 성격별 지출에 다시 더하지 않습니다.</div>`;
 if(key==='variable')return variableList();
 return `<p class="note">이번달 목적자금 실제 사용</p><div class="amount-large">${won(f.reserveUse)} 원</div>${cardList(ledger.transactions.filter(t=>t.month===month&&t.operation==='reserveUse'&&t.status==='confirmed'),'이번달 확정된 사용은 아직 없어요. 사용한 달에 결산을 반영합니다.')}<button class="primary" data-case="reserveUse">예비비 사용 입력</button>`;
}
function excelReview() {
 if(!pendingExcel)return '<p>엑셀 파일을 다시 선택해주세요.</p>';
 const {review,workbook,page,showExisting}=pendingExcel;
 const rows=review.rows.map((r,index)=>({...r,index})).filter(r=>showExisting||r.disposition!=='existing');
 const selected=review.rows.filter(r=>r.include&&r.disposition==='new');
 const selectField=(index,key,values,value)=>`<select data-excel-index="${index}" data-excel-field="${key}" aria-label="${esc(key)}">${Object.entries(values).map(([v,label])=>`<option value="${esc(v)}" ${String(value??'')===v?'selected':''}>${esc(label)}</option>`).join('')}</select>`;
 return `<p class="note">${esc(review.fileName)}<br>현재 가계부에 선택한 거래를 추가합니다.</p>${workbook.sheets.length>1?field('거래 시트',`<select id="excel-sheet">${workbook.sheets.map(s=>`<option ${s.name===review.sheet?'selected':''}>${esc(s.name)}</option>`).join('')}</select>`):''}<div class="card excel-counts">${rowText('읽은 거래',review.total+'건')}${rowText('기존 거래 · 중복 제외',review.existing+'건')}${rowText('오류 · 제외',review.invalid+'건')}<div id="excel-selected"></div></div><div class="info">고정비·수입은 아래에서 대체할 예상 항목을 선택해주세요. 이체·용도 미확인 거래는 보류로 저장하고, 가져온 뒤 기록 수정에서 처리합니다. 엑셀의 월별 결산·예비비·잔액 시트는 자동 교체하지 않습니다.</div><label class="check"><input id="excel-existing" type="checkbox" ${showExisting?'checked':''}> 중복 거래도 보기</label><div id="excel-rows">${rows.slice(page*20,(page+1)*20).map(r=>{
  if(r.disposition==='invalid')return `<div class="card error">${r.row}행 · ${esc(r.reason)}</div>`;
  const e=r.entry,plans=ledger.plans.filter(p=>p.month===e.month&&((p.type==='income'&&e.operation==='income')||(p.type==='fixed'&&e.operation==='expense'&&e.expenseClass==='fixed')));
  return `<details class="card excel-row" data-excel-row="${r.index}"><summary>${r.row}행 · ${esc(e.title)}<span>${won(e.amount)}원 · ${r.disposition==='existing'?'중복 제외':e.operation==='pending'?'처리 보류':statuses[e.status]}</span></summary><p class="muted">${esc(e.date||e.month+' · 날짜 미확인')}<br>${esc(e.account)} · ${esc(e.category)}<br>${esc(r.reason||'새 거래')}</p>${r.disposition==='new'?`<label class="check"><input type="checkbox" data-excel-index="${r.index}" data-excel-field="include" ${r.include?'checked':''}> 이 거래 가져오기</label>${field('반영 상태',selectField(r.index,'status',statuses,e.status))}${field('금액 (원)',`<input type="text" inputmode="numeric" value="${esc(e.amount)}" data-excel-index="${r.index}" data-excel-field="amount" aria-label="금액 (원)">`)}${e.operation==='expense'?field('지출 성격',selectField(r.index,'expenseClass',{fixed:name('monthly.fixed'),variable:name('monthly.variable')},e.expenseClass)):''}${['expense','income'].includes(e.operation)?field('계좌 위치',selectField(r.index,'funding',locations,e.funding)):''}${plans.length?field('대체할 예상 항목',selectField(r.index,'planID',{'':'새 거래 · 예상과 별도',...Object.fromEntries(plans.map(p=>[p.id,p.title+' · '+won(p.amount)+'원']))},e.planID)):''}`:''}</details>`;
 }).join('')||'<div class="info">추가할 새 거래가 없습니다. 기존 거래와 금액을 유지합니다.</div>'}</div>${rows.length>20?`<div class="action-grid"><button class="secondary" data-action="excel-prev" ${page===0?'disabled':''}>이전</button><button class="secondary" data-action="excel-next" ${(page+1)*20>=rows.length?'disabled':''}>다음</button></div><p class="muted">${page+1} / ${Math.ceil(rows.length/20)} 페이지</p>`:''}<div id="form-error" class="error" role="alert"></div><div class="stack"><button class="primary" id="confirm-excel" data-action="confirm-excel" ${!selected.length?'disabled':''}>${selected.length}건 추가하기</button><button class="secondary" data-action="close">취소</button></div>`;
}
function updateExcelCounts() {
 if(!pendingExcel||modal?.type!=='excel')return;
 const rows=pendingExcel.review.rows.filter(r=>r.include&&r.disposition==='new');
 const el=document.getElementById('excel-selected');if(el)el.innerHTML=rowText('선택한 새 거래',rows.length+'건')+rowText('보류 · 결산 미반영',rows.filter(r=>r.entry.operation==='pending').length+'건');
 const button=document.getElementById('confirm-excel');if(button){button.disabled=!rows.length;button.textContent=rows.length+'건 추가하기';}
}
function planForm(plan) {return `<form id="plan-form" novalidate><div class="card">${field('월',input('month',plan.month,'month'))}${field('계획 종류',select('type',{fixed:'고정비 예상',income:'수입 예상',reserve:'예비비 적립계획',...(plan.type==='familyContribution'?{familyContribution:'가족계좌 납입 · 중단'}:{})},plan.type))}${field('항목 이름',input('title',plan.title,'text','maxlength="120"'))}${field('예상 금액 (원)',input('amount',plan.amount,'text','inputmode="numeric"'))}${field('계좌·카드 이름',input('account',plan.account,'text','maxlength="100"'))}</div><div class="info">계획은 실제 잔액을 바꾸지 않아요. 실제 입력에서 연결하면 예상 금액을 대체해요. 취소한 계획은 금액을 0으로 기록해주세요.${plan.type==='familyContribution'?'<br>가족계좌 납입은 2026년 10월부터 적용 금액 0원으로 유지합니다.':''}</div><div id="form-error" class="error" role="alert"></div><button class="primary" type="submit">예상·계획 저장</button></form>`;}
function archiveContent() {return `<div class="info">원본 RAW · 읽기 전용<br>과거 분류와 본인부담을 그대로 보관합니다. 당시 공동계좌 50% 분담 등 원문 규칙은 현재 앱 계산에 적용하지 않아요.</div><label class="check"><input type="checkbox" id="archive-all"> 전체 기간 보기</label><label class="search">${svg('search')}<input id="archive-search" aria-label="원본 검색" placeholder="내용·계좌·분류 검색"></label><div id="archive-results"></div>`;}
function updateArchive() {
 const all=overlay.querySelector('#archive-all').checked,q=overlay.querySelector('#archive-search').value.toLocaleLowerCase();
 const entries=ledger.sourceArchive.filter(r=>(all||String(r['월']||r['일시']||'').startsWith(month))&&(!q||['내용','자산','분류','재분류','추가입력','비고'].some(k=>String(r[k]||'').toLocaleLowerCase().includes(q)))).sort((a,b)=>String(b['일시']||'').localeCompare(String(a['일시']||'')));
 overlay.querySelector('#archive-results').innerHTML=`<p class="muted">${entries.length}건</p>${entries.slice(0,100).map(r=>`<details class="card"><summary>${esc(r['내용'])}<span>${won(r['금액(KRW)'])}원</span></summary><dl class="raw-fields">${Object.entries(r).filter(([k])=>k!=='importKey').map(([k,v])=>`<dt>${esc(k==='source_row'?(r.import_row?'보관 ID':'원본 행'):({import_file:'가져온 파일',import_sheet:'엑셀 시트',import_row:'엑셀 행'}[k]||k))}</dt><dd>${esc(v??'—')}</dd>`).join('')}</dl></details>`).join('')}${entries.length>100?'<p class="muted">최근 100건 표시 · 검색하면 원하는 원문을 바로 찾을 수 있어요.</p>':''}`;
}
function drawModal() {
 if(!modal)return;
 let title='',body='';const f=summary?.monthlyFocus;
 switch(modal.type) {
  case 'entry':title=name(modal.entry.id?'entry.edit':'entry.add');body=editor(modal.entry,modal.ui);break;
  case 'record':title=name('records.detail');body=details(ledger.transactions.find(t=>t.id===modal.id));break;
  case 'delete':title='거래를 삭제할까요?';body=`<p class="note">${esc(ledger.transactions.find(t=>t.id===modal.id)?.title)}<br>결산·잔액을 다시 계산합니다. 원본 RAW와 이력은 보관해요.</p><div id="form-error" class="error" role="alert"></div><button class="danger" data-confirm-delete="${modal.id}">거래 삭제</button><button class="secondary" data-action="close">취소</button>`;break;
  case 'settlement':title=name('monthly.summary');body=settlement();break;
  case 'variable':case 'fixed':case 'cash':case 'reserveUse':title=name('monthly.'+modal.type);body=focusDetail(modal.type);break;
  case 'classification':title='이번달 집계 기준';body=`<div class="info">고정비·비고정비는 지출 성격, 비예비비·예비비 사용은 자금 사용을 살펴보는 메뉴예요. 네 금액을 모두 더하지 않아요.</div>${settlement()}`;break;
  case 'month':title=name('common.month');body=`<form id="month-form"><label class="field">조회 월${input('period',month,'month','required')}</label><button type="submit" class="primary">이 달 보기</button></form><p class="note">10월 45건부터 운영 결산을 연결했습니다. 이전 이력은 원본 보관함에서 볼 수 있어요.</p>`;break;
  case 'reserve': {
   const r=summary.reserves.find(r=>r.id===modal.id);title=r.title;body=`<div class="amount-large">${won(r.balance)} 원</div><div class="card">${row('누적 실제적립',r.allocated)}${row('누적 목적자금 사용·환입',r.used)}${row('이번달 적립계획',r.plan)}${rowText('예정일',r.date)}</div><div class="info">${esc(r.note)}</div><div class="action-grid"><button class="secondary" data-case="reserveAllocate" data-case-reserve="${r.id}">실제적립</button><button class="secondary" data-case="reserveUse" data-case-reserve="${r.id}">실제 사용</button></div>${cardList(ledger.transactions.filter(t=>t.reserveID===r.id))}`;break;
  }
  case 'purpose':title=modal.group;body=`<div class="card">${summary.reserves.filter(r=>r.group===modal.group).map(r=>`<button class="row-button" data-reserve="${r.id}">${row(r.title,r.balance)}</button>`).join('')}</div>`;break;
  case 'living':title=name('balance.living');body=`<div class="amount-large">${won(summary.living)} 원</div><div class="card">${row('생활비',summary.living)}${row('목적자금',summary.reserve)}${row('WRAP 총잔액',summary.wrap,true)}</div>${cardList(summary.wrapRecordIDs.map(id=>ledger.transactions.find(t=>t.id===id)).filter(t=>t.funding==='living'||t.operation==='reserveReturn'))}`;break;
  case 'tetris':title=name('monthly.tetris');body=summary.tetrisPending?`<div class="info">변경된 거래를 원 단위로 계산했어요. 월 테트리스 합계는 표시 규칙 확인 전까지 검토 대기로 표시합니다.</div><div class="card">${row('현재 원 단위 남는 돈',summary.remaining)}${ledger.transactions.filter(t=>t.month===month&&t.displayAmount!==null).map(t=>row(t.title+' · 직접 지정',t.displayAmount)).join('')}</div>`:`<div class="card">${row('원 단위 결산',summary.remaining)}${row('테트리스 결산 · 원본',summary.tetrisRemaining)}${row('차이',summary.remaining-summary.tetrisRemaining)}</div><div class="card">${summary.dailyAmounts.map(d=>rowText(d.day+'일',won(d.actual)+' → '+won(d.tetris)+'원')).join('')}</div>`;break;
  case 'rename': {
   const label=labels.find(l=>l.id===modal.key);title='이름 변경';body=`<form id="rename-form" novalidate><p class="note">기본 이름 · ${esc(label.defaultName)}<br>${label.maxLength}자 이내 · 표시 이름만 바뀌어요.</p>${field('표시할 이름',input('alias',name(modal.key),'text','autocomplete="off"'))}<div id="form-error" class="error" role="alert"></div><button class="primary" type="submit">이름 저장</button></form><button class="danger" data-name-default="${esc(modal.key)}">이 이름만 기본값으로</button>`;break;
  }
  case 'plan':title=name('entry.plan');body=planForm(modal.plan);break;
  case 'archive':title=name('common.archive');body=archiveContent();break;
  case 'excel':title='엑셀 거래 가져오기';body=excelReview();break;
  case 'restore':title='이 백업으로 복원할까요?';body=`<p class="note">운영 거래 <b>${modal.review.count}건</b><br>WRAP <b>${won(modal.review.wrap)}원</b><br>현재 테스트 데이터를 전체 교체합니다.</p><div class="info">먼저 현재 백업을 저장해주세요. iOS 앱 백업과 같은 형식을 사용합니다.</div><div id="form-error" class="error" role="alert"></div><div class="stack"><button class="secondary" data-action="export">현재 백업 저장</button><button class="primary" data-action="confirm-restore">이 백업으로 복원</button><button class="danger" data-action="close">취소</button></div>`;break;
  case 'reset':title='테스트 변경을 초기화할까요?';body=`<p class="note">이 화면의 거래·예상·이름 변경을 처음 상태로 되돌립니다. 변경한 내용은 먼저 JSON 백업으로 보관해주세요.</p><div id="form-error" class="error" role="alert"></div><div class="stack"><button class="secondary" data-action="export">현재 백업 저장</button><button class="danger" data-action="confirm-reset">테스트 변경 초기화</button><button class="secondary" data-action="close">취소</button></div>`;break;
  case 'help':title='직접 테스트해보세요';body=`<p class="note">예시를 선택하면 입력 화면이 채워져요. 거래 저장을 눌러 반영한 뒤 이번달·예비비·잔액을 살펴보세요.</p><div class="stack">${[['expense','계좌에서 1만원 지출'],['reserveAllocate','여행비 10만원 실제적립'],['reserveUse','여행비 4만원 사용'],['income','예상 월급을 실제 수입으로']].map(([op,text])=>`<button class="secondary" data-case="${op}">${text}</button>`).join('')}</div><div class="info">iOS와 같은 회계 모듈로 계산합니다. ${storageMode==='persistent'?'이 브라우저에 변경을 저장합니다.':'현재 화면의 저장소가 제한되어 변경은 이 화면에서만 유지됩니다. 백업으로 보관할 수 있어요.'}</div>`;break;
  case 'error':title='처리하지 못했어요';body=`<div class="error">${esc(modal.message)}</div><button class="secondary" data-action="close">확인</button>`;break;
 }
 overlay.innerHTML=`<div class="backdrop"><section class="sheet" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="handle"></div><div class="sheet-head"><h3>${esc(title)}</h3><button class="close" data-action="close" aria-label="닫기">${svg('close')}</button></div>${body}</section></div>`;
 if(modal.type==='archive')updateArchive();
 if(modal.type==='excel')updateExcelCounts();
 requestAnimationFrame(()=>overlay.querySelector('input:not([type=checkbox]),select,button')?.focus({preventScroll:true}));
}
function openFixed(index) {const fixed=summary.fixed[index],linked=ledger.transactions.find(t=>t.month===month&&(fixed.planID?t.planID===fixed.planID:t.title===fixed.title&&t.account===fixed.account&&t.expenseClass==='fixed'));show('entry',{entry:linked?clone(linked):freshEntry({title:fixed.title,amount:fixed.amount,account:fixed.account,expenseClass:'fixed',category:'고정비',funding:fixed.account==='CASH'?'bank':'card',planID:fixed.planID})});}
function exportBackup() {
 const bytes=ledger?JSON.stringify(ledger,null,2):lastSaved;
 if(!bytes)throw Error('백업할 저장본이 없습니다.');
 const url=URL.createObjectURL(new Blob([bytes],{type:'application/json'})),anchor=document.createElement('a');anchor.href=url;anchor.download='테트리스_가계부_기능테스트_백업.json';document.body.append(anchor);anchor.click();anchor.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);toast('백업 파일을 내보냈어요.');
}
async function reviewFile(file) {
 if(!file)return;
 if(file.size>32*1024*1024)throw Error('32MB 이하의 엑셀 또는 JSON 파일을 선택해주세요.');
 if(/\.(xlsx|xls)$/i.test(file.name)){
  if(!ledger)throw Error('먼저 정상 JSON 백업으로 가계부를 복원해주세요.');
  toast('엑셀을 읽고 있어요.');await new Promise(resolve=>requestAnimationFrame(resolve));
  const workbook=window.WorkbookImport.read(await file.arrayBuffer(),{fileName:file.name});
  pendingExcel={workbook,review:window.WorkbookImport.review(ledger,workbook),page:0,showExisting:false};show('excel');clearTimeout(toastTimer);document.getElementById('toast').classList.remove('show');return;
 }
 if(!/\.json$/i.test(file.name))throw Error('XLSX·XLS 엑셀 또는 JSON 백업을 선택해주세요.');
 const candidate=validateBackup(JSON.parse(await file.text())),s=window.LedgerCore.summarize(candidate,month);
 pendingRestore=candidate;show('restore',{review:{count:candidate.transactions.length,wrap:s.wrap}});
}
document.addEventListener('click',event=>{
 const button=event.target.closest('button');
 if(!button){if(event.target.classList.contains('backdrop'))hide();return;}
 try {
  const d=button.dataset;
  if(d.tab){tab=d.tab;query='';hide();render();screen.scrollTop=0;}
  else if(d.monthStep)moveMonth(Number(d.monthStep));
  else if(d.modal)show(d.modal);
  else if(d.record)show('record',{id:Number(d.record)});
  else if(d.edit)show('entry',{entry:clone(ledger.transactions.find(t=>t.id===Number(d.edit)))});
  else if(d.delete)show('delete',{id:Number(d.delete)});
  else if(d.confirmDelete){mutation({type:'delete',id:Number(d.confirmDelete)});toast('거래를 삭제했어요.');}
  else if(d.reserve)show('reserve',{id:Number(d.reserve)});
  else if(d.purpose)show('purpose',{group:d.purpose});
  else if(d.fixed!==undefined)openFixed(Number(d.fixed));
  else if(d.rename)show('rename',{key:d.rename});
  else if(d.nameDefault){const next=clone(ledger);delete next.personalization.names[d.nameDefault];commit(next,{message:'기본 이름으로 되돌렸어요.'});}
  else if(d.plan)show('plan',{plan:clone(ledger.plans.find(p=>p.id===d.plan))});
  else if(d.case)openCase(d.case,d.caseReserve?Number(d.caseReserve):22);
  else if(d.day){selectedDay=Number(d.day);render();}
  else if(d.mode){calendarMode=d.mode==='calendar';render();}
  else if(d.reserves){showAllReserves=d.reserves==='all';render();}
  else if(d.action) {
   switch(d.action){
    case 'add':show('entry',{entry:freshEntry()});break;
    case 'close':hide();break;
    case 'help':show('help');break;
    case 'day-records':tab='records';calendarMode=true;render();screen.scrollTop=0;break;
    case 'export':exportBackup();break;
    case 'import':document.getElementById('backup-file').click();break;
    case 'confirm-excel':{if(!pendingExcel)throw Error('엑셀을 다시 선택해주세요.');const selected=pendingExcel.review.rows.filter(r=>r.include&&r.disposition==='new');const candidate=window.WorkbookImport.apply(ledger,pendingExcel.review);commit(candidate,{targetTab:'records',targetMonth:selected[0]?.entry.month||month,message:selected.length+'건을 가져왔어요.'});pendingExcel=null;break;}
    case 'excel-next':pendingExcel.page++;drawModal();break;
    case 'excel-prev':pendingExcel.page=Math.max(0,pendingExcel.page-1);drawModal();break;
    case 'confirm-restore':if(!pendingRestore)throw Error('복원할 파일을 다시 선택해주세요.');commit(pendingRestore,{message:'백업을 복원했어요.'});break;
    case 'reset':show('reset');break;
    case 'confirm-reset':{const initial=clone(seed);initial.personalization=checkedSettings(undefined);commit(initial,{targetMonth:'2026-10',message:'테스트를 처음 상태로 되돌렸어요.'});tab='home';render();break;}
    case 'reset-names':{const next=clone(ledger);next.personalization.names={};commit(next,{message:'이름을 기본값으로 복원했어요.'});break;}
    case 'add-plan':show('plan',{plan:{id:crypto.randomUUID?crypto.randomUUID():'plan-'+Date.now()+'-'+Math.random().toString(16).slice(2),month,type:'fixed',title:'',account:'개인 계좌',status:'예상',amount:0}});break;
   }
  }
 }catch(error){errorMessage(error);}
});
document.addEventListener('submit',event=>{
 event.preventDefault();const form=event.target;
 try {
  if(form.id==='entry-form'){const e=entryFromForm(form);commit(window.LedgerCore.mutate(ledger,{type:'save',entry:e}),{targetMonth:e.month,targetTab:'records',message:'거래를 저장했어요.'});}
  else if(form.id==='month-form'){const m=form.elements.period.value;window.LedgerCore.summarize(ledger,m);month=m;hide();render();screen.scrollTop=0;}
  else if(form.id==='rename-form'){
   const label=labels.find(l=>l.id===modal.key),value=form.elements.alias.value.trim();
   if(!value||[...value].length>label.maxLength||/[\u0000-\u001f\u007f-\u009f]/.test(value))throw Error('1~'+label.maxLength+'자 이름을 입력해주세요.');
   const next=clone(ledger);next.personalization.names[modal.key]=value;commit(next,{message:'이름을 저장했어요.'});
  }else if(form.id==='plan-form'){
   const data=Object.fromEntries(new FormData(form)),number=data.amount.replace(/,/g,'').trim();
   if(!/^\d+$/.test(number))throw Error('0원 이상의 정수 예상 금액을 입력해주세요.');
   const plan=Object.assign({},modal.plan,data,{amount:Number(number),title:data.title.trim()});mutation({type:'savePlan',plan},plan.month);
  }
 }catch(error){errorMessage(error);}
});
document.addEventListener('input',event=>{
 if(event.target.id==='search'){query=event.target.value;document.getElementById('record-groups').innerHTML=recordGroups();}
 if(event.target.id==='archive-search')updateArchive();
});
document.addEventListener('change',event=>{
 const element=event.target;
 try {
  if(element.id==='backup-file'){const file=element.files[0];element.value='';reviewFile(file).catch(errorMessage);}
  else if(element.id==='excel-sheet'){pendingExcel.review=window.WorkbookImport.review(ledger,pendingExcel.workbook,element.value);pendingExcel.page=0;drawModal();}
  else if(element.id==='excel-existing'){pendingExcel.showExisting=element.checked;pendingExcel.page=0;drawModal();}
  else if(element.dataset.excelField){const row=pendingExcel.review.rows[Number(element.dataset.excelIndex)],key=element.dataset.excelField;if(key==='include')row.include=element.checked;else if(key==='amount'){const value=element.value.replace(/,/g,'').trim();row.entry.amount=/^-?\d+$/.test(value)?Number(value):null;}else if(key==='planID')row.entry.planID=element.value||null;else row.entry[key]=element.value;if(key==='expenseClass'){row.entry.planID=null;drawModal();}else updateExcelCounts();}
  else if(element.dataset.pref){const next=clone(ledger);next.personalization[element.dataset.pref]=element.checked;commit(next,{message:'화면 설정을 저장했어요.'});}
  else if(element.id==='archive-all')updateArchive();
  else if(element.closest('#entry-form')&&['operation','expenseClass','date','month','unknownDate'].includes(element.name)){
   const ui=readEditorUI(element.form);
   if(element.name==='operation'){
    ui.planID='';ui.linkedID='';ui.attributedMonth='';
    if(ui.operation==='jointActivity'){ui.funding='joint';ui.sourceKind='지출';}
    if(modal.entry.id===0)ui.category=operations[ui.operation];
    if(['reserveAllocate','reserveUse','reserveReturn'].includes(ui.operation)&&!ui.reserveID)ui.reserveID='22';
   }
   modal.ui=ui;drawModal();
  }
 }catch(error){errorMessage(error);}
});
document.addEventListener('keydown',event=>{
 if(!modal)return;
 if(event.key==='Escape'){event.preventDefault();hide();return;}
 if(event.key==='Tab'){
  const controls=[...overlay.querySelectorAll('button,input,select,textarea,summary,a[href]')].filter(e=>!e.disabled&&e.offsetParent!==null);
  const first=controls[0],last=controls.at(-1);
  if(event.shiftKey&&document.activeElement===first){event.preventDefault();last?.focus();}
  else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first?.focus();}
 }
});
window.BudgetPlayground={get state(){return clone(ledger);},get summary(){return clone(summary);},get storageMode(){return storageMode;},get month(){return month;},reviewFile};
initialize();
})();
