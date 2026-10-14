/* One accounting implementation, executed by JavaScriptCore on iOS and Node in tests.
   Integer KRW only. No network, system clock, rounding policy, or automatic reconciliation. */
(function (root) {
  'use strict';
  const SCHEMA = 1;
  const LIMIT = 9999999999;
  const buckets = ['living', 'family', 'travel', 'repay', 'condolence', 'other'];
  const operations = ['expense', 'income', 'refund', 'reserveAllocate', 'reserveUse', 'reserveReturn', 'jointAllocate', 'jointActivity', 'investmentAllocate', 'investmentBuy', 'cardPayment', 'transfer', 'pending'];
  const clone = x => JSON.parse(JSON.stringify(x));
  const check = (yes, message) => { if (!yes) throw new Error(message); };
  const money = x => Number.isSafeInteger(x) && Math.abs(x) <= LIMIT;
  const monthOK = x => typeof x === 'string' && /^\d{4}-(0[1-9]|1[0-2])$/.test(x);
  const string = (x, max) => typeof x === 'string' && x.length <= max;
  // Source-derived item + detail repair for the October import. Exact legacy
  // titles only: user renames, amounts, statuses, account labels and RAW stay intact.
  const legacyFixedTitles = {
  "oct-fixed-1": [
    "9월사용분(전달)",
    "관리비 · 9월사용분(전달)"
  ],
  "oct-fixed-2": [
    "9월사용분(전달)",
    "도시가스 · 9월사용분(전달)"
  ],
  "oct-fixed-3": [
    "이번달",
    "보험비 · 이번달"
  ],
  "oct-fixed-4": [
    "SKT",
    "무선통신비 · SKT"
  ],
  "oct-fixed-5": [
    "U모바일",
    "무선통신비 · U모바일"
  ],
  "oct-fixed-6": [
    "LG헬로비전",
    "유선통신비 · LG헬로비전"
  ],
  "oct-fixed-7": [
    "K패스",
    "교통비 · K패스"
  ],
  "oct-fixed-8": [
    "T우주 CU구독",
    "정기구독 · T우주 CU구독"
  ],
  "oct-fixed-9": [
    "신한은행",
    "전세대출이자 · 신한은행"
  ],
  "oct-fixed-10": [
    "Uber One",
    "정기구독 · Uber One"
  ],
  "oct-fixed-11": [
    "ChatGPT",
    "정기구독 · ChatGPT"
  ],
  "oct-fixed-12": [
    "로켓와우",
    "정기구독 · 로켓와우"
  ],
  "oct-fixed-13": [
    "카카오",
    "정기구독 · 카카오"
  ],
  "oct-fixed-14": [
    "T우주+YouTube",
    "정기구독 · T우주+YouTube"
  ],
  "oct-fixed-15": [
    "애플뮤직",
    "정기구독 · 애플뮤직"
  ],
  "oct-fixed-16": [
    "네이버+",
    "정기구독 · 네이버+"
  ],
  "oct-fixed-17": [
    "티빙디즈니",
    "정기구독 · 티빙디즈니"
  ],
  "oct-fixed-18": [
    "휴대폰",
    "할부 · 휴대폰"
  ],
  "oct-fixed-19": [
    "에어팟프로3",
    "할부 · 에어팟프로3"
  ],
  "oct-fixed-20": [
    "애플케어플러스",
    "할부 · 애플케어플러스"
  ]
};
  function repairFixedTitles(state) {
    const next = clone(state);
    const repair = item => {
      const pair = legacyFixedTitles[item.id];
      if (item.type === 'fixed' && item.month === '2026-10' && pair && item.title === pair[0]) item.title = pair[1];
    };
    next.plans.forEach(repair); next.originalPlans.forEach(repair);
    return next;
  }
  function dateOK(x, month) {
    if (x === null) return true; // month-only evidence stays month-only
    if (!string(x, 40) || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(x) || x.slice(0, 7) !== month) return false;
    const d = new Date(x + (/[Z+-]\d?/.test(x.slice(10)) ? '' : 'Z'));
    return !isNaN(d.getTime()) && d.toISOString().slice(0, 16) === x.slice(0, 16);
  }
  function validateEntry(t, state) {
    check(t && Number.isSafeInteger(t.id) && t.id > 0, '거래 ID가 올바르지 않습니다.');
    check(monthOK(t.month) && dateOK(t.date, t.month), '거래 날짜·월을 확인해주세요.');
    check(string(t.title, 120) && t.title.trim().length > 0, '내용을 입력해주세요.');
    check(string(t.account, 100) && string(t.category, 100) && string(t.note, 4000), '거래 문자열을 확인해주세요.');
    check(money(t.amount) && (t.operation === 'expense' || t.amount >= 0), '원 단위 정수 금액을 확인해주세요.');
    check(operations.includes(t.operation), '지원하지 않는 거래 종류입니다.');
    check(['confirmed', 'unconfirmed', 'planned'].includes(t.status), '확정 상태를 확인해주세요.');
    check(['card', 'bank', 'joint'].concat(buckets).includes(t.funding), '출금·입금 위치를 확인해주세요.');
    check(['bank', 'joint'].concat(buckets).includes(t.destination), '이체 목적지를 확인해주세요.');
    check(['fixed', 'variable'].includes(t.expenseClass), '고정비·비고정비 분류를 확인해주세요.');
    check(t.reserveID === null || state.reserves.some(r => r.id === t.reserveID), '예비비 목적을 확인해주세요.');
    check(!['reserveAllocate', 'reserveUse', 'reserveReturn'].includes(t.operation) || t.reserveID !== null, '예비비 목적을 선택해주세요.');
    check(typeof t.alreadyNet === 'boolean', '본인부담 순액 여부를 확인해주세요.');
    check(t.linkedID === null || Number.isSafeInteger(t.linkedID), '연결 거래를 확인해주세요.');
    check(t.attributedMonth === null || monthOK(t.attributedMonth), '정산 귀속 월을 확인해주세요.');
    check(t.displayAmount === null || (money(t.displayAmount) && t.displayAmount >= 0), '테트리스 표시 금액을 확인해주세요.');
    check(t.planID === null || state.plans.some(p => p.id === t.planID && p.month === t.month && ((p.type === 'income' && t.operation === 'income') || (p.type === 'fixed' && t.operation === 'expense' && t.expenseClass === 'fixed'))), '연결할 예상 항목과 거래 월·종류가 맞지 않습니다.');
    if (t.operation === 'transfer') {
      check(t.funding !== t.destination, '출발지와 도착지가 같습니다.');
      check(t.destination !== 'joint', '공동계좌 납입은 공동계좌 납입 종류로 입력해주세요.');
    }
    if (t.operation === 'reserveAllocate') check(['living','bank'].includes(t.funding), '예비비 실제적립은 생활비 또는 일반 계좌에서 이동해주세요.');
    if (t.operation === 'jointAllocate') check(['living','bank'].includes(t.funding), '공동계좌 납입은 개인 계좌에서 출금해주세요.');
    if (t.operation === 'income' || t.operation === 'refund') check(t.funding !== 'joint', '공동계좌 입금은 공동계좌 활동으로 기록해주세요.');
    if (t.operation === 'jointActivity') check(t.funding === 'joint', '공동계좌 활동의 위치를 확인해주세요.');
    // Purpose bucket balances must always be backed by a specific reserve item.
    if (!['reserveAllocate', 'reserveUse', 'reserveReturn', 'jointActivity', 'pending'].includes(t.operation)) {
      check(!buckets.slice(1).includes(t.funding) && !buckets.slice(1).includes(t.destination), '목적자금 이동은 예비비 적립·사용·생활비 환입으로 입력해주세요.');
    }
    check(t.operation !== 'expense' || t.funding !== 'joint', '공동계좌 소비는 공동계좌 활동으로 입력해주세요.');
  }
  function validate(state) {
    check(state && state.schemaVersion === SCHEMA, '지원하지 않는 백업 버전입니다. 원본 파일을 보관해주세요.');
    check(Array.isArray(state.transactions) && state.transactions.length <= 100000, '거래 목록을 확인해주세요.');
    check(Array.isArray(state.originals) && Array.isArray(state.reserves) && Array.isArray(state.plans) && Array.isArray(state.originalPlans) && Array.isArray(state.originalDaily) && Array.isArray(state.sourceArchive), '이관 기준 데이터가 없습니다.');
    check(string(state.asOf, 40) && monthOK(state.baseMonth) && money(state.originalTetrisOutflow) && money(state.originalTetrisRemaining), '기준일·표시 금액을 확인해주세요.');
    check(Number.isSafeInteger(state.nextID) && state.nextID > 0 && Array.isArray(state.audit) && state.audit.length <= 200000, '저장 상태를 확인해주세요.');
    buckets.forEach(b => check(state.bankBase && money(state.bankBase[b]), 'WRAP 기준 잔액을 확인해주세요.'));
    const displayDays = new Set();
    state.originalDaily.forEach(d=>{check(Number.isInteger(d.day)&&d.day>=1&&d.day<=31&&!displayDays.has(d.day)&&money(d.actual)&&money(d.tetris),'원본 일별 표시값을 확인해주세요.');displayDays.add(d.day);});
    const reserveIDs = new Set();
    state.reserves.forEach(r => {
      check(Number.isSafeInteger(r.id) && !reserveIDs.has(r.id) && buckets.slice(1).includes(r.bucket), '예비비 ID·계정을 확인해주세요.'); reserveIDs.add(r.id);
      ['plan', 'allocated', 'used', 'balance'].forEach(k => check(money(r[k]) && r[k] >= 0, '예비비 기준 금액을 확인해주세요.'));
      check(r.allocated - r.used === r.balance && string(r.group, 100) && string(r.title, 120) && string(r.date, 120) && string(r.note, 4000), '예비비 기준 데이터가 일치하지 않습니다.');
    });
    const planIDs = new Set();
    state.plans.forEach(p => {
      check(string(p.id, 100) && !planIDs.has(p.id) && monthOK(p.month) && ['fixed', 'income', 'reserve', 'familyContribution'].includes(p.type), '예상 항목을 확인해주세요.'); planIDs.add(p.id);
      check(money(p.amount) && p.amount >= 0 && string(p.title, 120) && string(p.account, 100) && ['예상', '미확인', '확정'].includes(p.status), '예상 금액·상태를 확인해주세요.');
    });
    const originalIDs = new Set();
    state.originals.forEach(t => { validateEntry(t, state); check(!originalIDs.has(t.id), '원본 ID가 중복됩니다.'); originalIDs.add(t.id); });
    const ids = new Set(), imports = new Set(), replacements = new Set();
    state.transactions.forEach(t => {
      validateEntry(t, state); check(!ids.has(t.id), '거래 ID가 중복됩니다.'); ids.add(t.id);
      check(t.sourceKey === null || string(t.sourceKey, 200), '원본 키를 확인해주세요.');
      if (t.sourceKey) { check(!imports.has(t.sourceKey), '원본 거래가 중복됩니다.'); imports.add(t.sourceKey); }
      if (t.planID && t.status !== 'planned') { check(!replacements.has(t.planID), '같은 예상 항목에 두 거래를 연결할 수 없습니다.'); replacements.add(t.planID); }
    });
    check(state.transactions.every(t => t.id < state.nextID), '다음 거래 ID를 확인해주세요.');
    state.transactions.forEach(t => {
      if (t.linkedID !== null) {
        const original = state.transactions.find(x => x.id === t.linkedID);
        check(original && original.id !== t.id, '연결된 거래가 없습니다. 먼저 연결을 해제해주세요.');
        check((t.operation === 'refund' || t.operation === 'cardPayment') && ['expense', 'reserveUse'].includes(original.operation), '정산·카드대금은 지출 거래에 연결해주세요.');
        if (t.operation === 'refund' && original.month !== t.month) check(t.attributedMonth !== null, '다른 달 정산은 귀속 월을 직접 선택해주세요.');
        if (t.operation === 'cardPayment') check(original.funding === 'card', '카드 사용 내역에 연결해주세요.');
      }
      if (t.operation === 'cardPayment') check(t.linkedID !== null && t.funding !== 'card', '카드대금의 출금 위치·원래 사용 내역을 선택해주세요.');
    });
    return state;
  }
  function effect(t, state) {
    const e = { income: 0, fixed: 0, variable: 0, refund: 0, reserveUse: 0, reserveAllocation: 0, jointAllocation: 0, investmentAllocation: 0, cash: 0, bank: {}, allocated: {}, used: {} };
    if (t.status !== 'confirmed' || t.operation === 'pending') return e;
    const a = t.amount, reserve = state.reserves.find(r => r.id === t.reserveID);
    const post = (b, v) => { if (buckets.includes(b)) e.bank[b] = (e.bank[b] || 0) + v; };
    const consume = () => { e[t.expenseClass] += a; };
    switch (t.operation) {
      case 'expense': consume(); post(t.funding, -a); if (t.funding === 'bank' || t.funding === 'living') e.cash = a; break;
      case 'income': e.income = a; post(t.funding, a); break;
      case 'refund': {
        const parent = state.transactions.find(x => x.id === t.linkedID);
        if (!parent || !parent.alreadyNet) e.refund = a;
        post(t.funding, a); break;
      }
      case 'reserveAllocate': e.reserveAllocation = a; post(t.funding, -a); post(reserve.bucket, a); e.allocated[reserve.id] = a; break;
      case 'reserveUse': consume(); e.reserveUse = a; post(reserve.bucket, -a); e.used[reserve.id] = a; break;
      case 'reserveReturn': post(reserve.bucket, -a); post('living', a); e.used[reserve.id] = a; break;
      case 'jointAllocate': e.jointAllocation = a; post(t.funding, -a); break;
      case 'investmentAllocate': e.investmentAllocation = a; post(t.funding, -a); break;
      case 'cardPayment': post(t.funding, -a); break;
      case 'transfer': post(t.funding, -a); post(t.destination, a); break;
      // Joint consumption, interest, investment buys are outside personal spending / WRAP.
      case 'jointActivity': case 'investmentBuy': break;
    }
    return e;
  }
  function activeForMonth(t, month) { return (t.operation === 'refund' ? t.attributedMonth || t.month : t.month) === month; }
  function totals(state, month, actualOnly) {
    let result = {income: 0, fixed: 0, variable: 0, refund: 0, reserveUse: 0, reserveAllocation: 0, jointAllocation: 0, investmentAllocation: 0, cash: 0};
    state.transactions.filter(t => activeForMonth(t, month)).forEach(t => {
      // Source unconfirmed fixed rows participate in the forecast, not in confirmed actuals.
      const entry = !actualOnly && t.status === 'unconfirmed' ? Object.assign({}, t, {status: 'confirmed'}) : t;
      const e = effect(entry, state); Object.keys(result).forEach(k => result[k] += e[k]);
    });
    if (!actualOnly) state.plans.filter(p => p.month === month).forEach(p => {
      const replaced = state.transactions.some(t => t.planID === p.id && t.status !== 'planned');
      if (!replaced && p.type === 'income') result.income += p.amount;
      if (!replaced && p.type === 'fixed') result.fixed += p.amount;
    });
    result.outflow = result.fixed + result.variable + result.reserveAllocation + result.jointAllocation + result.investmentAllocation - result.refund;
    result.remaining = result.income - result.outflow;
    Object.values(result).forEach(x => check(Number.isSafeInteger(x), '합계가 정수 허용 범위를 초과했습니다.'));
    return result;
  }
  function balances(state) {
    const bank = clone(state.bankBase), items = clone(state.reserves);
    const apply = (t, sign, context) => {
      const e = effect(t, context);
      Object.keys(e.bank).forEach(b => bank[b] += sign * e.bank[b]);
      items.forEach(r => { r.allocated += sign * (e.allocated[r.id] || 0); r.used += sign * (e.used[r.id] || 0); r.balance = r.allocated - r.used; });
    };
    const originalContext = Object.assign({}, state, {transactions: state.originals});
    state.originals.forEach(t => apply(t, -1, originalContext));
    state.transactions.forEach(t => apply(t, 1, state));
    items.forEach(r => check(r.allocated >= 0 && r.used >= 0 && r.balance >= 0 && money(r.balance), '예비비 사용·환입이 보유 금액을 초과합니다.'));
    buckets.forEach(b => { check(money(bank[b]) && bank[b] >= 0, 'WRAP 계정 잔액이 부족하거나 허용 범위를 초과합니다.'); if (b !== 'living') check(bank[b] === items.filter(r => r.bucket === b).reduce((s,r) => s+r.balance,0), '목적자금 원장이 일치하지 않습니다.'); });
    return {bank, items, wrap: buckets.reduce((s,b) => s+bank[b],0)};
  }
  function transactionView(t) {
    const incoming = ['income', 'refund'].includes(t.operation) || (t.operation === 'jointActivity' && t.sourceKind === '수입');
    const spending = ['expense', 'reserveUse'].includes(t.operation) || (t.operation === 'jointActivity' && t.sourceKind === '지출');
    return Object.assign({}, t, {date: t.date || t.month + ' · 날짜 미확인', amount: incoming ? t.amount : -t.amount, kind: incoming ? '수입' : spending ? '지출' : '이체', isPreview: false,
      recordIncome: t.status === 'confirmed' && t.operation === 'income' ? t.amount : 0,
      recordExpense: t.status === 'confirmed' && ['expense', 'reserveUse'].includes(t.operation) ? t.amount : 0});
  }
  function summarize(state, month) {
    validate(state); state = repairFixedTitles(state); check(monthOK(month), '조회 월을 확인해주세요.');
    const projected = totals(state, month, false), actual = totals(state, month, true), b = balances(state);
    const selected = state.transactions.filter(t => t.month === month).sort((a,b) => (b.date || '').localeCompare(a.date || '') || b.id-a.id);
    const confirmed = selected.filter(t => t.status === 'confirmed');
    const variable = confirmed.filter(t => ['expense','reserveUse'].includes(t.operation) && t.expenseClass === 'variable');
    const cashItems = confirmed.filter(t => t.operation === 'expense' && ['bank','living'].includes(t.funding)).map(t => ({sourceRow:t.id,title:t.title,date:t.date || '',amount:t.amount,alreadyCounted:t.amount,additional:0,note:'성격별 지출에 한 번 반영한 계좌 현금성 지출'}));
    const days = Array.from(new Set(variable.filter(t => t.date).map(t => Number(t.date.slice(8,10))))).sort((a,b)=>a-b);
    const unchanged = month === state.baseMonth && JSON.stringify(state.transactions) === JSON.stringify(state.originals) && JSON.stringify(state.plans) === JSON.stringify(state.originalPlans);
    const oldDaily = state.originalDaily || [];
    const daily = days.map(day => ({day, actual:variable.filter(t => t.date && Number(t.date.slice(8,10)) === day).reduce((s,t)=>s+t.amount,0), tetris: unchanged ? ((oldDaily.find(d=>d.day===day)||{}).tetris||0) : 0}));
    const fixed = state.plans.filter(p => p.type === 'fixed' && p.month === month).map(p => {
      const linked = selected.find(t=>t.planID===p.id && t.status!=='planned');
      return {title:p.title,account:p.account,amount:linked ? linked.amount : p.amount,status:linked ? (linked.status==='confirmed'?'확정':'미확인') : p.status,planID:p.id};
    });
    selected.filter(t=>t.operation==='expense'&&t.expenseClass==='fixed'&&!t.planID&&t.status!=='planned').forEach(t=>fixed.push({title:t.title,account:t.account,amount:t.amount,status:t.status==='confirmed'?'확정':'미확인',planID:null}));
    const familyPlan = state.plans.filter(p=>p.type==='familyContribution'&&p.month===month).reduce((s,p)=>s+(month >= '2026-10' ? 0 : p.amount),0);
    const plan = state.plans.filter(p=>p.type==='reserve'&&p.month===month).reduce((s,p)=>s+p.amount,0);
    return {
      asOf:state.asOf,month,archivedCount:state.sourceArchive.filter(r=>String(r['월']||r['일시']||'').slice(0,7)===month).length,historicalPending:month<state.baseMonth&&state.sourceArchive.some(r=>String(r['월']||r['일시']||'').slice(0,7)===month),income:projected.income,outflow:projected.outflow,remaining:projected.remaining,
      actualIncome:actual.income,actualOutflow:actual.outflow,actualRemaining:actual.remaining,
      hasRecords:selected.some(t=>t.status==='confirmed'),hasForecast:state.plans.some(p=>p.month===month&&['fixed','income'].includes(p.type)) || selected.some(t=>t.status==='unconfirmed' && t.operation!=='pending'),
      tetrisPending:!unchanged,tetrisOutflow:unchanged?state.originalTetrisOutflow:0,tetrisRemaining:unchanged?state.originalTetrisRemaining:0,
      wrap:b.wrap,living:b.bank.living,reserve:b.items.reduce((s,r)=>s+r.balance,0),reserveAllocated:b.items.reduce((s,r)=>s+r.allocated,0),reserveUsed:b.items.reduce((s,r)=>s+r.used,0),
      reserveMonthPlan:plan,reserveMonthActual:actual.reserveAllocation,familyPlan,
      transactions:selected.map(transactionView),reserves:b.items.map(r=>Object.assign({},r,{plan:month===state.baseMonth?r.plan:0})),fixed,
      monthlyFocus:{fixed:projected.fixed,variable:projected.variable,unplannedCash:actual.cash,cashAlreadyCounted:actual.cash,cashAdditional:0,reserveUse:actual.reserveUse,refund:projected.refund,jointAllocation:projected.jointAllocation,reserveAllocation:projected.reserveAllocation,investmentAllocation:projected.investmentAllocation},
      cashItems,dailyAmounts:daily,variableRecordIDs:variable.map(t=>t.id),wrapRecordIDs:state.transactions.filter(t=>Object.values(effect(t,state).bank).some(v=>v!==0)).sort((a,b)=>(b.date||'').localeCompare(a.date||'')||b.id-a.id).map(t=>t.id),
      reserveUseItems:confirmed.filter(t=>t.operation==='reserveUse').map(transactionView),sourceCount:state.sourceArchive.length
    };
  }
  function mutate(state, action) {
    validate(state); const next = repairFixedTitles(state);
    if (action.type === 'save') {
      const t = clone(action.entry), old = next.transactions.find(x=>x.id===t.id);
      if (!old) { t.id = next.nextID++; t.sourceKey = null; }
      else t.sourceKey = old.sourceKey; // Display edits never rewrite original provenance.
      const at = next.transactions.findIndex(x=>x.id===t.id);
      if (at<0) next.transactions.push(t); else next.transactions[at]=t;
      next.audit.push({action:old?'edit':'add',id:t.id,before:old||null,after:t});
    } else if (action.type === 'delete') {
      check(next.transactions.some(t=>t.id===action.id),'삭제할 거래가 없습니다.');
      check(!next.transactions.some(t=>t.linkedID===action.id),'연결된 정산·카드대금이 있습니다. 먼저 연결 거래를 수정·삭제해주세요.');
      const old=next.transactions.find(t=>t.id===action.id); next.transactions=next.transactions.filter(t=>t.id!==action.id); next.audit.push({action:'delete',id:action.id,before:old,after:null});
    } else if (action.type === 'savePlan') {
      const plan = clone(action.plan), at = next.plans.findIndex(p=>p.id===plan.id);
      const old = at<0 ? null : next.plans[at];
      if(old && next.originals.some(t=>t.planID===old.id)) check(plan.month===old.month && plan.type===old.type,'원본 거래에 연결된 예상 항목의 월·종류는 유지해주세요.');
      if(at<0) next.plans.push(plan); else next.plans[at]=plan;
      next.audit.push({action:'plan',id:plan.id,before:old,after:plan});
    } else if (action.type === 'import') {
      check(Array.isArray(action.entries), '정규화한 거래 목록이 필요합니다.');
      const archiveRows = action.archiveRows;
      check(archiveRows === undefined || (Array.isArray(archiveRows) && archiveRows.length === action.entries.length), '엑셀 원본 행을 확인해주세요.');
      let archiveID = next.sourceArchive.reduce((id,r)=>Math.max(id,Number.isSafeInteger(r.source_row)?r.source_row:0),0);
      action.entries.forEach((entry,index)=>{
        check(string(entry.sourceKey,200)&&entry.sourceKey.length>0,'이관 거래의 파일·행 원본 키가 필요합니다.');
        // Tombstones also prevent reimport resurrecting user deletions.
        if(next.originals.some(t=>t.sourceKey===entry.sourceKey)||next.transactions.some(t=>t.sourceKey===entry.sourceKey)||next.audit.some(a=>a.before&&a.before.sourceKey===entry.sourceKey)) return;
        const t=clone(entry);t.id=next.nextID++;next.transactions.push(t);next.audit.push({action:'import',id:t.id,before:null,after:t});
        if(archiveRows){
          const raw=clone(archiveRows[index]);
          check(raw && money(raw['금액(KRW)']) && string(raw['내용'],120) && string(raw['자산'],100) && (raw['본인부담']===null||money(raw['본인부담'])), '엑셀 원본 금액·내용을 확인해주세요.');
          raw.source_row=++archiveID;raw.importKey=t.sourceKey;next.sourceArchive.push(raw);
        }
      });
    } else throw new Error('지원하지 않는 작업입니다.');
    validate(next);balances(next); const months=new Set(next.transactions.flatMap(t=>[t.month,t.attributedMonth].filter(Boolean)));months.forEach(m=>totals(next,m,false));
    return next;
  }
  const api={schemaVersion:SCHEMA,validate,effect,totals,balances,summarize,mutate};
  root.LedgerCore=api;
  root.ledgerDispatch=function(input){try{const q=JSON.parse(input);let value;if(q.action==='summary')value=summarize(q.state,q.month);else if(q.action==='mutate')value=mutate(q.state,q.change);else if(q.action==='validate'){validate(q.state);q.state=repairFixedTitles(q.state);balances(q.state);q.state.transactions.forEach(t=>totals(q.state,t.month,false));value=q.state;}else throw new Error('지원하지 않는 요청입니다.');return JSON.stringify({ok:true,value});}catch(e){return JSON.stringify({ok:false,error:e.message||String(e)});}};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof globalThis!=='undefined'?globalThis:this);
