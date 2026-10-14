'use strict';
const test=require('node:test');const assert=require('node:assert/strict');
const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),vm=require('node:vm');
const core=require('../core/ledger-engine');const fixture=require('../ios/TetrisBudget/seed-ledger.json');
const fresh=()=>JSON.parse(JSON.stringify(fixture));
const entry=(overrides={})=>Object.assign({id:0,date:'2026-10-06T12:00:00',month:'2026-10',title:'테스트 거래',account:'개인 계좌',category:'식비',amount:10000,operation:'expense',status:'confirmed',funding:'bank',destination:'bank',expenseClass:'variable',reserveID:null,linkedID:null,attributedMonth:null,alreadyNet:false,displayAmount:null,planID:null,note:'',sourceKey:null},overrides);
const save=(s,t)=>core.mutate(s,{type:'save',entry:t});
const sum=(s,m='2026-10')=>core.summarize(s,m);
const delta=(s,field)=>sum(s)[field]-sum(fixture)[field];
test('Initial Excel forecast, complete 19 nonfixed transactions, balances reconcile',()=>{
 const x=sum(fresh());assert.deepEqual([x.income,x.outflow,x.remaining,x.wrap,x.living,x.reserve],[2780230,1584335,1195895,8480881,6330881,2150000]);
 assert.equal(x.variableRecordIDs.length,19);assert.deepEqual(x.dailyAmounts.map(d=>d.actual),[17300,30500,5000,112890,120890]);assert.equal(x.actualIncome,4180);assert.equal(x.actualOutflow,336280);
});
test('Fixed item names retain both Excel item/detail columns, including continued groups',()=>{
 const labels=require('../verification/fixed-item-labels.json');
 const fixed=sum(fresh()).fixed;
 assert.equal(labels.length,21);
 labels.forEach(label=>assert.equal(fixed.find(p=>p.planID===label.id).title,label.title));
 assert.equal(fixed[1].title,'관리비 · 9월사용분(전달)');
 assert.equal(fixed[2].title,'도시가스 · 9월사용분(전달)');
 assert.equal(fixed[3].title,'보험비 · 이번달');
 assert.equal(fixed[5].title,'무선통신비 · U모바일');
 assert.equal(fixed[12].title,'정기구독 · 로켓와우');
 assert.equal(fixed[19].title,'할부 · 에어팟프로3');
});
test('Legacy stored fixed titles and backups repair without overwriting edits or money',()=>{
 const labels=require('../verification/fixed-item-labels.json');
 let legacy=fresh();
 for(const list of [legacy.plans,legacy.originalPlans]){
  labels.forEach(label=>{list.find(p=>p.id===label.id).title=label.legacyTitle;});
 }
 const baseline=JSON.stringify(legacy);
 assert.equal(sum(legacy).fixed[1].title,'관리비 · 9월사용분(전달)');
 assert.equal(sum(legacy).tetrisPending,false);
 assert.equal(JSON.stringify(legacy),baseline);
 legacy.plans.find(p=>p.id==='oct-fixed-1').amount=175000;
 legacy.plans.find(p=>p.id==='oct-fixed-2').title='내 도시가스';
 legacy.personalization={names:{'tab.home':'내 월간장부'},compact:false,reduceTransparency:true};
 const parse=input=>JSON.parse(globalThis.ledgerDispatch(JSON.stringify({action:'validate',state:input}))).value;
 const restored=parse(legacy);
 assert.equal(restored.plans.find(p=>p.id==='oct-fixed-1').title,'관리비 · 9월사용분(전달)');
 assert.equal(restored.plans.find(p=>p.id==='oct-fixed-1').amount,175000);
 assert.equal(restored.plans.find(p=>p.id==='oct-fixed-2').title,'내 도시가스');
 assert.deepEqual(restored.transactions,legacy.transactions);assert.deepEqual(restored.originals,legacy.originals);
 assert.deepEqual(restored.sourceArchive,legacy.sourceArchive);assert.deepEqual(restored.personalization,legacy.personalization);
 assert.deepEqual(restored.audit,legacy.audit);assert.deepEqual(parse(restored),restored);
 assert.equal(sum(restored).outflow,sum(legacy).outflow);assert.equal(sum(restored).wrap,sum(legacy).wrap);
 const next=core.mutate(legacy,{type:'save',entry:entry({amount:5000})});
 assert.equal(next.plans.find(p=>p.id==='oct-fixed-1').title,'관리비 · 9월사용분(전달)');
 assert.equal(sum(next).outflow,sum(legacy).outflow+5000);
});
test('V01 Same normalized source imported twice; deletion tombstone respected',()=>{
 let s=fresh();const imported=entry({sourceKey:'new-file:row:1'});s=core.mutate(s,{type:'import',entries:[imported]});const after=JSON.stringify(s);
 s=core.mutate(s,{type:'import',entries:[imported]});assert.equal(JSON.stringify(s),after);
 const id=s.transactions.at(-1).id;s=core.mutate(s,{type:'delete',id});const deleted=JSON.stringify(s);s=core.mutate(s,{type:'import',entries:[imported]});assert.equal(JSON.stringify(s),deleted);
});
test('V02 Equal contents from distinct rows remain distinct',()=>{
 const s=core.mutate(fresh(),{type:'import',entries:[entry({sourceKey:'batch:1'}),entry({sourceKey:'batch:2'})]});assert.equal(s.transactions.length,47);assert.equal(delta(s,'outflow'),20000);
});
test('V03 Two-kki counted once, current WRAP includes it once',()=>{
 const x=sum(fresh());assert.equal(x.monthlyFocus.unplannedCash,11900);assert.equal(x.monthlyFocus.cashAdditional,0);assert.equal(core.effect(fixture.transactions.find(t=>t.id===2333),fixture).bank.living,-11900);
 const s=core.mutate(fresh(),{type:'delete',id:2333});assert.equal(delta(s,'wrap'),11900);assert.equal(delta(s,'outflow'),-11900);
});
test('V04 Joint allocation only; joint spending and interest excluded',()=>{
 let s=save(fresh(),entry({operation:'jointAllocate',funding:'living',amount:50000}));s=save(s,entry({operation:'jointActivity',funding:'joint',amount:30000}));s=save(s,entry({operation:'jointActivity',funding:'joint',amount:500}));assert.equal(delta(s,'wrap'),-50000);assert.equal(delta(s,'outflow'),50000);
});
test('V05 Actual reserve allocation is internal movement with invariant WRAP total',()=>{
 const s=save(fresh(),entry({operation:'reserveAllocate',reserveID:22,funding:'living',amount:100000}));assert.equal(delta(s,'wrap'),0);assert.equal(delta(s,'living'),-100000);assert.equal(delta(s,'reserve'),100000);assert.equal(sum(s).reserveMonthActual,100000);assert.equal(delta(s,'outflow'),100000);
});
test('V06 Hotel use reduces reserve/WRAP once and adds consumption once',()=>{
 const s=save(fresh(),entry({operation:'reserveUse',reserveID:22,amount:40000}));assert.equal(delta(s,'wrap'),-40000);assert.equal(delta(s,'reserve'),-40000);assert.equal(delta(s,'outflow'),40000);assert.equal(sum(s).monthlyFocus.reserveUse,40000);
});
test('V07 Past funding, current use belongs to actual usage month',()=>{
 const s=save(fresh(),entry({operation:'reserveUse',reserveID:22,amount:40000,date:'2026-11-03T12:00:00',month:'2026-11'}));assert.equal(delta(s,'outflow'),0);assert.equal(sum(s,'2026-11').actualOutflow,40000);
});
test('V08 Card purchase vs next-month payment have separate consumption and cash dates',()=>{
 let s=save(fresh(),entry({funding:'card',amount:50000}));const id=s.transactions.at(-1).id;
 s=save(s,entry({operation:'cardPayment',funding:'living',amount:50000,linkedID:id,month:'2026-11',date:'2026-11-02T12:00:00'}));assert.equal(delta(s,'outflow'),50000);assert.equal(sum(s,'2026-11').actualOutflow,0);assert.equal(delta(s,'wrap'),-50000);
});
test('V09 Linked refund reduces burden; cross-month attribution must be explicit',()=>{
 let s=save(fresh(),entry({amount:30000}));const id=s.transactions.at(-1).id;
 assert.throws(()=>save(s,entry({operation:'refund',amount:10000,linkedID:id,month:'2026-11',date:'2026-11-02T12:00:00'})),/귀속 월/);
 s=save(s,entry({operation:'refund',amount:10000,linkedID:id,month:'2026-11',date:'2026-11-02T12:00:00',attributedMonth:'2026-10'}));assert.equal(delta(s,'outflow'),20000);assert.equal(delta(s,'income'),0);
});
test('V10 Multiple settlements / already-net burden do not subtract twice',()=>{
 let s=save(fresh(),entry({amount:100000}));const id=s.transactions.at(-1).id;s=save(s,entry({operation:'refund',amount:30000,linkedID:id}));s=save(s,entry({operation:'refund',amount:20000,linkedID:id}));assert.equal(delta(s,'outflow'),50000);
 let n=save(fresh(),entry({amount:50000,alreadyNet:true}));const nid=n.transactions.at(-1).id;n=save(n,entry({operation:'refund',amount:50000,linkedID:nid}));assert.equal(delta(n,'outflow'),50000);
});
test('V11 Ordinary bank receipts preserve WRAP scope',()=>{
 let s=save(fresh(),entry({operation:'income',amount:1}));s=save(s,entry({operation:'refund',amount:100}));assert.equal(delta(s,'wrap'),0);assert.equal(delta(s,'income'),1);assert.equal(delta(s,'outflow'),-100);
});
test('V12 Approved interest 4178 is a single canonical transaction',()=>{
 let s=core.mutate(fresh(),{type:'delete',id:2367});assert.equal(sum(s).wrap,8476703);
 const original=fixture.originals.find(t=>t.id===2367);s=save(s,original);assert.equal(sum(s).wrap,8480881);assert.equal(sum(s).income,2780230);
});
test('V13 Unexplained discrepancy stays pending, no implicit interest',()=>{
 const s=save(fresh(),entry({operation:'pending',amount:4178,funding:'living'}));assert.equal(delta(s,'wrap'),0);assert.equal(delta(s,'income'),0);
});
test('V14 Future plan changes no actual balances; missing distinct from zero',()=>{
 const x=sum(fresh(),'2026-11');assert.equal(x.reserveMonthPlan,200000);assert.equal(x.hasRecords,false);assert.equal(x.reserveMonthActual,0);
 const s=save(fresh(),entry({operation:'reserveAllocate',reserveID:22,amount:100000,status:'planned',month:'2026-11',date:'2026-11-01T12:00:00'}));assert.equal(delta(s,'wrap'),0);assert.equal(delta(s,'reserve'),0);
});
test('V15 Integers and exact zero vs negative balance verdict',()=>{
 let s=save(fresh(),entry({operation:'income',amount:1234567,month:'2026-11',date:'2026-11-01T12:00:00'}));s=save(s,entry({amount:1234567,month:'2026-11',date:'2026-11-01T12:00:00'}));assert.equal(sum(s,'2026-11').remaining,0);assert.ok(sum(s,'2026-11').remaining>=0);
 s=save(s,entry({amount:1,month:'2026-11',date:'2026-11-01T12:00:00'}));assert.equal(sum(s,'2026-11').remaining,-1);assert.throws(()=>save(fresh(),entry({amount:1.1})),/정수/);
});
test('V16 Family contribution plans stopped October while originals preserved',()=>{
 const s=fresh();assert.equal(sum(s,'2026-11').familyPlan,0);assert.equal(sum(s,'2026-12').familyPlan,0);assert.equal(s.plans.find(p=>p.id==='family-2026-11').amount,200000);
});
test('V17 Investment allocation counted once, purchase does not repeat it',()=>{
 let s=save(fresh(),entry({operation:'investmentAllocate',amount:120000,funding:'living'}));s=save(s,entry({operation:'investmentBuy',amount:120000}));assert.equal(delta(s,'outflow'),120000);assert.equal(delta(s,'wrap'),-120000);assert.equal(sum(s).monthlyFocus.variable,286580);
});
test('V18 Unknown day remains nullable, not first day of month',()=>{
 const s=save(fresh(),entry({date:null,amount:10000}));assert.equal(s.transactions.at(-1).date,null);assert.equal(sum(s).dailyAmounts.reduce((n,d)=>n+d.actual,0),286580);assert.equal(delta(s,'outflow'),10000);
});
test('V19 Planned birthday date does not override actual September usage',()=>{
 const funded=fresh();const father=funded.reserves.find(r=>r.id===7);father.used-=200000;father.balance+=200000;funded.bankBase.family+=200000;const s=save(funded,entry({operation:'reserveUse',reserveID:7,amount:200000,month:'2026-09',date:'2026-09-20T12:00:00'}));assert.equal(sum(s,'2026-09').monthlyFocus.reserveUse,200000);assert.equal(sum(s).monthlyFocus.reserveUse,0);
});
test('V20 No invented ceil policy; source display preserved and modifications pending',()=>{
 const original=sum(fresh());assert.equal(original.tetrisOutflow,1595700);assert.equal(original.tetrisPending,false);
 const s=save(fresh(),entry({amount:1234,displayAmount:2000}));assert.equal(s.transactions.at(-1).displayAmount,2000);assert.equal(sum(s).tetrisPending,true);assert.equal(delta(s,'outflow'),1234);
});
test('Replacing fixed estimate and salary: no double-counting; deletion restores estimate',()=>{
 let s=save(fresh(),entry({amount:400000,expenseClass:'fixed',planID:'oct-fixed-0'}));assert.equal(delta(s,'outflow'),9000);
 s=save(s,entry({amount:3000000,operation:'income',planID:'oct-salary'}));assert.equal(sum(s).income,3004180);
 const id=s.transactions.find(t=>t.planID==='oct-fixed-0').id;s=core.mutate(s,{type:'delete',id});assert.equal(sum(s).monthlyFocus.fixed,1248055);
});
test('Edit/delete recompute all metrics; originals immutable',()=>{
 const original=JSON.stringify(fixture);let s=fresh();const t=s.transactions.find(t=>t.id===2333);s=save(s,Object.assign({},t,{amount:21900}));assert.equal(delta(s,'outflow'),10000);assert.equal(delta(s,'wrap'),-10000);assert.equal(sum(s).monthlyFocus.unplannedCash,21900);assert.equal(JSON.stringify(fixture),original);assert.equal(s.originals.find(t=>t.id===2333).amount,11900);
});
test('Linked deletion guard, duplicate plans, insufficient reserve rejected atomically',()=>{
 let s=save(fresh(),entry());const id=s.transactions.at(-1).id;s=save(s,entry({operation:'refund',linkedID:id,amount:100}));assert.throws(()=>core.mutate(s,{type:'delete',id}),/連|연결/);
 assert.throws(()=>save(fresh(),entry({operation:'reserveUse',reserveID:22,amount:1400001})),/초과/);assert.throws(()=>save(fresh(),entry({expenseClass:'fixed',planID:'oct-fixed-8'})),/두 거래/);
});
test('Backups: real filesystem round-trip, corrupt version/date/IDs/money rejected',()=>{
 const directory=fs.mkdtempSync(path.join(os.tmpdir(),'tetris-budget-'));try{
 let s=save(fresh(),entry({funding:'living'}));const p=path.join(directory,'ledger.json');fs.writeFileSync(p+'.tmp',JSON.stringify(s));fs.renameSync(p+'.tmp',p);
 const restored=JSON.parse(fs.readFileSync(p));core.validate(restored);assert.deepEqual(sum(restored),sum(s));
 const bad=fresh();bad.schemaVersion=2;assert.throws(()=>core.validate(bad),/버전/);
 const invalid=fresh();invalid.transactions[0].date='2026-10-32T12:00:00';assert.throws(()=>core.validate(invalid),/날짜/);
 const duplicate=fresh();duplicate.transactions.push(duplicate.transactions[0]);assert.throws(()=>core.validate(duplicate),/중복/);
 const fractional=fresh();fractional.bankBase.living=1.5;assert.throws(()=>core.validate(fractional),/잔액/);
 assert.throws(()=>JSON.parse('{'),SyntaxError);
 }finally{fs.rmSync(directory,{recursive:true});}
});
test('iOS dispatch executes same bundled engine in isolated JS global, structured errors',()=>{
 const code=fs.readFileSync(path.join(__dirname,'../ios/TetrisBudget/ledger-engine.js'),'utf8');assert.equal(code,fs.readFileSync(path.join(__dirname,'../core/ledger-engine.js'),'utf8'));
 const ctx=vm.createContext({});vm.runInContext(code,ctx);const result=JSON.parse(ctx.ledgerDispatch(JSON.stringify({action:'summary',state:fresh(),month:'2026-10'})));assert.equal(result.ok,true);assert.equal(result.value.wrap,8480881);
 const error=JSON.parse(ctx.ledgerDispatch('{'));assert.equal(error.ok,false);assert.equal(typeof error.error,'string');
});
test('Manual forecast plans update only projection; actual plan replacement and source provenance remain',()=>{
 let s=fresh();const original=JSON.stringify(s.originalPlans);
 s=core.mutate(s,{type:'savePlan',plan:{id:'nov-salary',month:'2026-11',type:'income',title:'급여',account:'은행',status:'예상',amount:3000000}});
 s=core.mutate(s,{type:'savePlan',plan:{id:'nov-rent',month:'2026-11',type:'fixed',title:'월세',account:'은행',status:'예상',amount:391000}});
 assert.equal(sum(s,'2026-11').remaining,2609000);assert.equal(sum(s,'2026-11').hasRecords,false);assert.equal(sum(s,'2026-11').actualIncome,0);assert.equal(delta(s,'wrap'),0);assert.equal(JSON.stringify(s.originalPlans),original);
 s=save(s,entry({operation:'income',amount:3100000,month:'2026-11',date:'2026-11-20T12:00:00',planID:'nov-salary'}));assert.equal(sum(s,'2026-11').income,3100000);
});
test('Personal joint funding guards and unsupported purpose transfers rejected',()=>{
 assert.throws(()=>save(fresh(),entry({operation:'transfer',funding:'living',destination:'joint'})),/공동계좌 납입/);
 assert.throws(()=>save(fresh(),entry({operation:'income',funding:'joint'})),/공동계좌 활동/);
 assert.throws(()=>save(fresh(),entry({operation:'reserveAllocate',reserveID:22,funding:'card'})),/생활비 또는 일반 계좌/);
});
test('Joint incoming source sign survives viewing; negative expense preserved as burden credit',()=>{
 const x=sum(fresh());const t=x.transactions.find(t=>t.id===2347);assert.equal(t.amount,50000);assert.equal(t.kind,'수입');
 const s=save(fresh(),entry({amount:-1000}));assert.equal(delta(s,'outflow'),-1000);assert.equal(s.transactions.at(-1).amount,-1000);
});
test('Changed forecast invalidates cached Tetris; malformed display cache refused during restore',()=>{
 const s=fresh();const plan=Object.assign({},s.plans.find(p=>p.id==='oct-salary'),{amount:3000000});const changed=core.mutate(s,{type:'savePlan',plan});assert.equal(sum(changed).tetrisPending,true);assert.equal(sum(changed).income,3004180);
 s.originalDaily[0].tetris=1.5;assert.throws(()=>core.validate(s),/일별 표시/);
});
test('Historical source evidence is preserved and clearly marked outside operational migration',()=>{
 const s=fresh(), september=sum(s,'2026-09');assert.equal(september.historicalPending,true);assert.ok(september.archivedCount>0);assert.equal(september.transactions.length,0);
 const october=sum(s);assert.equal(october.historicalPending,false);assert.equal(october.archivedCount,45);assert.equal(s.sourceArchive.length,2366);
});
test('Record day contributions include confirmed personal income and consumption, excluding movements and forecasts',()=>{
 let s=fresh();
 for(const overrides of [
  {operation:'income',amount:10000}, {amount:3000}, {amount:-500},
  {operation:'reserveUse',amount:1000,reserveID:22}, {amount:10000,status:'unconfirmed'},
  {operation:'refund',amount:500}, {operation:'cardPayment',amount:3000,linkedID:2323},
  {operation:'transfer',amount:700,destination:'living'},
  {operation:'jointActivity',amount:800,funding:'joint',sourceKind:'수입'},
  {operation:'reserveAllocate',amount:900,reserveID:22}
 ]) s=save(s,entry(overrides));
 const rows=sum(s).transactions.filter(t=>t.date.startsWith('2026-10-06'));
 assert.equal(rows.length,10);
 assert.equal(rows.reduce((n,t)=>n+t.recordIncome,0),10000);
 assert.equal(rows.reduce((n,t)=>n+t.recordExpense,0),3500);
 for(const t of rows.filter(t=>['refund','cardPayment','transfer','jointActivity','reserveAllocate'].includes(t.operation)||t.status!=='confirmed'))
  assert.deepEqual([t.recordIncome,t.recordExpense],[0,0]);
});
