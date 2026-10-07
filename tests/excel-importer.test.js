'use strict';
const test=require('node:test'),assert=require('node:assert/strict');
const fs=require('fs'),crypto=require('crypto'),vm=require('vm'),path=require('path');
global.XLSX=require('../core/vendor/xlsx.full.min.js');
global.LedgerCore=require('../core/ledger-engine.js');
const importer=require('../core/excel-importer.js'),core=global.LedgerCore;
const fixture=require('../ios/TetrisBudget/seed-ledger.json');
const fresh=()=>JSON.parse(JSON.stringify(fixture));
const headers=['일시','자산','내용','금액(KRW)','수입/지출','재분류','추가입력','상태'];
const row=(overrides={})=>Object.assign({date:'2026-10-07 12:30:00',account:'미래에셋_WRAP형CMA',title:'엑셀 커피',amount:10000,kind:'지출',category:'변동비',note:'메모 보존',status:'확정'},overrides);
function binary(rows,{format='xlsx',extraSheets=false,date1904=false,customHeaders=headers}={}){
 const book=XLSX.utils.book_new();if(extraSheets)XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['이번달','합계'],['결산',99999999]]),'10월');
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([customHeaders,...rows.map(r=>Array.isArray(r)?r:[r.date,r.account,r.title,r.amount,r.kind,r.category,r.note,r.status])]),'RAW');
 if(date1904)book.Workbook={WBProps:{date1904:true}};
 return XLSX.write(book,{type:'buffer',bookType:format});
}
const proposal=(state,rows,options)=>importer.review(state,importer.read(binary(rows,options),{fileName:'거래.xlsx'}));
const sum=state=>core.summarize(state,'2026-10');
test('Stable Unicode SHA-256 keys agree with cryptographic reference',()=>{
 for(const value of ['','abc','테트리스 가계부 · ☕ 😀',JSON.stringify(fixture)])assert.equal(importer.sha256(value),crypto.createHash('sha256').update(value).digest('hex'));
});
test('Actual original Tetris workbook has 2366 existing rows and no duplicate additions',t=>{
 const p=path.resolve(__dirname,'../../household-budget-history/source/ledger-handoff/sources/가계부/테트리스_가계부_2026_복구.xlsx');
 if(!fs.existsSync(p)){t.skip('Original handoff workbook is optional and is not distributed in the source ZIP.');return;}
 const book=importer.read(fs.readFileSync(p));assert.equal(book.preferred,'RAW');assert.deepEqual(book.sheets.map(s=>s.name),['RAW']);
 const state=fresh(),before=JSON.stringify(state),review=importer.review(state,book);
 assert.deepEqual([review.total,review.existing,review.invalid],[2366,2366,0]);assert.equal(review.rows.filter(r=>r.include).length,0);
 assert.throws(()=>importer.apply(state,review),/추가할 거래/);assert.equal(JSON.stringify(state),before);
});
test('XLSX import adds selected transactions and immutable RAW; WRAP and consumption agree',()=>{
 const state=fresh(),review=proposal(state,[row(),row({title:'추가 수입',amount:20000,kind:'수입'})],{extraSheets:true});
 const before=JSON.stringify(state),after=importer.apply(state,review);
 assert.equal(after.transactions.length,47);assert.equal(after.sourceArchive.length,2368);
 assert.equal(sum(after).wrap,sum(state).wrap+10000);assert.equal(sum(after).outflow,sum(state).outflow+10000);assert.equal(sum(after).income,sum(state).income+20000);
 assert.equal(sum(after).monthlyFocus.unplannedCash,21900);assert.equal(sum(after).monthlyFocus.cashAdditional,0);
 assert.equal(after.sourceArchive.at(-2)['추가입력'],'메모 보존');assert.equal(after.sourceArchive.at(-2).import_row,2);
 assert.deepEqual(after.originals,state.originals);assert.deepEqual(after.bankBase,state.bankBase);assert.equal(JSON.stringify(state),before);
});
test('XLS and renamed/reordered XLSX files deduplicate across file formats',()=>{
 const rows=[row(),row({title:'다른 거래',amount:4000})],state=importer.apply(fresh(),proposal(fresh(),rows));
 const book=importer.read(binary(rows.slice().reverse(),{format:'xls'}),{fileName:'이름을 바꾼 파일.xls'}),review=importer.review(state,book);
 assert.equal(review.existing,2);assert.equal(review.invalid,0);assert.ok(review.rows.every(r=>!r.include));
});
test('Identical source rows retain their multiplicity, overlap adds only extra occurrence',()=>{
 const base=fresh(),first=proposal(base,[row(),row()]);assert.equal(new Set(first.rows.map(r=>r.entry.sourceKey)).size,2);
 const state=importer.apply(base,first),review=proposal(state,[row(),row(),row()]);
 assert.equal(review.existing,2);assert.equal(review.rows.filter(r=>r.include).length,1);
 assert.equal(importer.apply(state,review).transactions.length,48);
});
test('User edits and deletions are not overwritten or resurrected by repeated Excel',()=>{
 let state=importer.apply(fresh(),proposal(fresh(),[row()]));const imported=state.transactions.at(-1);
 state=core.mutate(state,{type:'save',entry:{...imported,title:'내가 수정한 이름',amount:12000}});
 assert.equal(proposal(state,[row()]).existing,1);assert.equal(state.transactions.at(-1).title,'내가 수정한 이름');
 state=core.mutate(state,{type:'delete',id:imported.id});assert.equal(proposal(state,[row()]).existing,1);
 assert.ok(!state.transactions.some(t=>t.id===imported.id));
});
test('Explicit personal burden preserves gross RAW and consumes the burden once',()=>{
 const h=['날짜','계좌','내용','금액','구분','再분류','메모','상태','본인부담'];h[5]='재분류';
 const state=fresh(),review=proposal(state,[['2026-10-07','일반 계좌','식사','20,000','지출','변동비','공동 정산','확정',10000]],{customHeaders:h});
 const after=importer.apply(state,review);assert.equal(after.transactions.at(-1).amount,10000);assert.equal(after.transactions.at(-1).alreadyNet,true);
 assert.equal(after.sourceArchive.at(-1)['금액(KRW)'],20000);assert.equal(after.sourceArchive.at(-1)['본인부담'],10000);
 assert.equal(sum(after).outflow-sum(state).outflow,10000);
});
test('Transfers, ambiguous purpose spending and old WRAP movement remain pending',()=>{
 const state=fresh(),review=proposal(state,[row({kind:'이체출금',amount:30000}),row({category:'예비비',amount:40000}),row({date:'2026-10-04 12:00:00',title:'기준일 이전 이자',kind:'수입',amount:100})]);
 assert.ok(review.rows.every(r=>r.entry.operation==='pending'));
 const after=importer.apply(state,review);assert.equal(sum(after).wrap,sum(state).wrap);assert.equal(sum(after).outflow,sum(state).outflow);
});
test('Imported cashback and card payment require linking instead of duplicate income/consumption',()=>{
 const state=fresh(),review=proposal(state,[row({title:'카드 캐시백',kind:'수입',amount:1000}),row({title:'카드대금 납부',amount:30000})]);
 assert.ok(review.rows.every(r=>r.entry.operation==='pending'));const after=importer.apply(state,review);
 assert.equal(sum(after).income,sum(state).income);assert.equal(sum(after).outflow,sum(state).outflow);
});
test('Joint-account spending does not enter personal consumption or WRAP',()=>{
 const state=fresh(),review=proposal(state,[row({account:'20260606',amount:30000}),row({account:'20260606',kind:'수입',amount:200})]);
 assert.ok(review.rows.every(r=>r.entry.operation==='jointActivity'));
 const after=importer.apply(state,review);assert.equal(sum(after).wrap,sum(state).wrap);assert.equal(sum(after).outflow,sum(state).outflow);assert.equal(sum(after).income,sum(state).income);
});
test('Fixed actual can explicitly replace forecast without double counting',()=>{
 const state=fresh(),review=proposal(state,[row({title:'월세',amount:400000,category:'고정비',status:'미확인',account:'CASH'})]);
 review.rows[0].entry.planID='oct-fixed-0';review.rows[0].entry.status='confirmed';
 const after=importer.apply(state,review);assert.equal(sum(after).outflow-sum(state).outflow,9000);
});
test('Invalid dates, decimal KRW and missing amounts are visible excluded rows',()=>{
 const review=proposal(fresh(),[row(),row({date:'2026-02-30'}),row({amount:1.5}),row({amount:null})]);
 assert.equal(review.invalid,3);assert.equal(review.rows.filter(r=>r.include).length,1);
 assert.equal(importer.apply(fresh(),review).transactions.length,46);
});
test('Month-only evidence retains null day, 1904 Excel serial dates decode correctly',()=>{
 const state=fresh(),h=['월','자산','내용','금액','구분'];
 const r=proposal(state,[['2026-11','일반 계좌','날짜 미확인',1000,'지출']],{customHeaders:h});assert.equal(r.rows[0].entry.date,null);assert.equal(r.rows[0].entry.month,'2026-11');
 const serial=(Date.UTC(2026,10,7,12)-Date.UTC(1904,0,1))/86400000;
 const s=proposal(state,[row({date:serial})],{date1904:true});assert.equal(s.rows[0].entry.date,'2026-11-07T12:00:00.000');
});
test('Cancel/exclusion leaves state unchanged, stale proposals and invalid edits fail atomically',()=>{
 const state=fresh(),before=JSON.stringify(state),review=proposal(state,[row()]);review.rows[0].include=false;assert.throws(()=>importer.apply(state,review),/선택/);
 review.rows[0].include=true;review.rows[0].entry.amount=1.5;assert.throws(()=>importer.apply(state,review),/정수/);assert.equal(JSON.stringify(state),before);
 const valid=proposal(state,[row()]);const changed=core.mutate(state,{type:'savePlan',plan:{...state.plans[0],amount:500000}});assert.throws(()=>importer.apply(changed,valid),/변경/);
});
test('Changed source amount requires explicit selection and never silently overwrites',()=>{
 const state=importer.apply(fresh(),proposal(fresh(),[row()]));
 const review=proposal(state,[row({amount:12000})]);
 assert.equal(review.rows[0].include,false);assert.match(review.rows[0].reason,/금액이 달라/);
 assert.equal(state.transactions.at(-1).amount,10000);assert.throws(()=>importer.apply(state,review),/선택/);
});
test('Corrupt files, unrecognized summaries and currency mismatches fail clearly',()=>{
 assert.throws(()=>importer.read(Buffer.from('not excel')),/XLSX/);
 const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['결산','합계'],['이번달',10]]),'10월');
 assert.throws(()=>importer.read(XLSX.write(book,{bookType:'xlsx',type:'buffer'})),/거래표/);
 const r=proposal(fresh(),[['2026-10-07','계좌','USD',10,'지출','USD']],{customHeaders:['일시','자산','내용','금액','수입/지출','통화']});assert.equal(r.invalid,1);
});
test('Oversized inflated XLSX members are rejected before decompression',()=>{
 const bytes=Buffer.from(binary([row()])),central=bytes.indexOf(Buffer.from([0x50,0x4b,0x01,0x02]));assert.ok(central>=0);
 bytes.writeUInt32LE(65*1024*1024,central+24);assert.throws(()=>importer.read(bytes),/64MB/);
});
test('Vendor/parser/core dispatch run without Node, DOM, TextEncoder, atob or Web Crypto',()=>{
 const context=vm.createContext({});for(const f of ['core/ledger-engine.js','core/vendor/xlsx.full.min.js','core/excel-importer.js'])vm.runInContext(fs.readFileSync(path.resolve(__dirname,'..',f),'utf8'),context);
 const file=Buffer.from(binary([row()])).toString('base64');const response=JSON.parse(context.excelDispatch(JSON.stringify({action:'read',file,fileName:'기기 엑셀.xlsx'})));assert.equal(response.ok,true,response.error);assert.equal(response.value.sheets.length,1);
 const review=JSON.parse(context.excelDispatch(JSON.stringify({action:'review',state:fresh(),workbook:response.value})));assert.equal(review.ok,true);
 const result=JSON.parse(context.excelDispatch(JSON.stringify({action:'apply',state:fresh(),proposal:review.value})));assert.equal(result.ok,true);assert.equal(result.value.transactions.length,46);
});
