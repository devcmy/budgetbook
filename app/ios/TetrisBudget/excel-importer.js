/* Local-only spreadsheet import shared by browsers and iOS JavaScriptCore.
   Formula results are read from cached values; formulas/macros are never run. */
(function(root) {
 'use strict';
 const clone=x=>JSON.parse(JSON.stringify(x));
 const check=(value,message)=>{if(!value)throw Error(message);};
 const text=value=>value===undefined||value===null?'':String(value).trim();
 const normalize=value=>text(value).toLowerCase().replace(/[\s_]/g,'');
 const aliases={
  date:['일시','날짜','거래일시','거래일','date','datetime'],month:['월','거래월','month'],
  account:['자산','계좌','계좌/카드','결제수단','account','asset'],
  title:['내용','내역','거래내용','description','title'],amount:['금액(krw)','금액krw','krw','금액','amount'],
  kind:['수입/지출','수입지출','수입/지출구분','구분','유형','type','kind'],
  category:['분류','카테고리','category'],subcategory:['소분류','subcategory'],
  classification:['재분류','지출성격'],burden:['본인부담','본인부담금액'],method:['결제구분'],
  note:['추가입력','메모','memo','note'],remark:['비고'],status:['상태','status'],currency:['통화','currency']
 };
 // SHA-256, independent of Web Crypto and Node so JavaScriptCore uses the same keys.
 function sha256(input) {
  const bytes=[];
  for(let i=0;i<input.length;i++){
   let c=input.charCodeAt(i);
   if(c>=0xd800&&c<=0xdbff){const low=input.charCodeAt(++i);c=low>=0xdc00&&low<=0xdfff?0x10000+((c-0xd800)<<10)+(low-0xdc00):0xfffd;}
   if(c<128)bytes.push(c);else if(c<2048)bytes.push(192|(c>>6),128|(c&63));
   else if(c<65536)bytes.push(224|(c>>12),128|((c>>6)&63),128|(c&63));
   else bytes.push(240|(c>>18),128|((c>>12)&63),128|((c>>6)&63),128|(c&63));
  }
  const bitLength=bytes.length*8;bytes.push(128);while(bytes.length%64!==56)bytes.push(0);
  for(let i=7;i>=0;i--)bytes.push(i>=4?0:(bitLength/Math.pow(256,i))&255);
  const k=[0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  const h=[0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
  const rotate=(v,n)=>(v>>>n)|(v<<(32-n));
  for(let off=0;off<bytes.length;off+=64){
   const w=new Array(64);
   for(let i=0;i<16;i++)w[i]=(bytes[off+i*4]<<24)|(bytes[off+i*4+1]<<16)|(bytes[off+i*4+2]<<8)|bytes[off+i*4+3];
   for(let i=16;i<64;i++){const a=w[i-15],b=w[i-2];w[i]=((rotate(a,7)^rotate(a,18)^(a>>>3))+w[i-16]+(rotate(b,17)^rotate(b,19)^(b>>>10))+w[i-7])|0;}
   let [a,b,c,d,e,f,g,hh]=h;
   for(let i=0;i<64;i++){const t1=(hh+(rotate(e,6)^rotate(e,11)^rotate(e,25))+((e&f)^(~e&g))+k[i]+w[i])|0,t2=((rotate(a,2)^rotate(a,13)^rotate(a,22))+((a&b)^(a&c)^(b&c)))|0;hh=g;g=f;f=e;e=(d+t1)|0;d=c;c=b;b=a;a=(t1+t2)|0;}
   [a,b,c,d,e,f,g,hh].forEach((v,i)=>h[i]=(h[i]+v)|0);
  }
  return h.map(v=>(v>>>0).toString(16).padStart(8,'0')).join('');
 }
 function money(value,label='금액') {
  check(value!==null&&value!==undefined&&text(value)!=='',label+'이 비어 있거나 수식 결과가 없습니다.');
  let number=value;
  if(typeof value==='string'){
   let s=value.trim().replace(/^[₩￦]\s*/,'').replace(/\s*원$/,'');
   check(/^-?(?:\d+|\d{1,3}(?:,\d{3})+)$/.test(s),label+'은 원 단위 정수여야 합니다.');number=Number(s.replace(/,/g,''));
  }
  check(Number.isSafeInteger(number)&&Math.abs(number)<=9999999999,label+'은 원 단위 정수여야 합니다.');return number;
 }
 function monthOf(value,date1904=false) {
  if(typeof value==='number'){const d=root.XLSX.SSF.parse_date_code(value,{date1904});check(d,'거래 월을 확인해주세요.');return String(d.y).padStart(4,'0')+'-'+String(d.m).padStart(2,'0');}
  const match=text(value).match(/^(\d{4})[-./](\d{1,2})(?:[-./T\s]|$)/);
  check(match&&Number(match[2])>=1&&Number(match[2])<=12,'거래 월을 확인해주세요.');return match[1]+'-'+match[2].padStart(2,'0');
 }
 function dateOf(value,date1904=false) {
  if(value===null||value===undefined||text(value)==='')return null;
  let y,m,d,h=0,min=0,sec=0,ms=0;
  if(typeof value==='number'){
   const p=root.XLSX.SSF.parse_date_code(value,{date1904});check(p,'거래 날짜를 확인해주세요.');
   ({y,m,d}=p);h=p.H;min=p.M;sec=p.S;ms=Math.round((p.u||0)*1000);
   if(ms===1000){sec++;ms=0;}
  }else{
   const p=text(value).match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,6}))?)?)?$/);
   check(p,'날짜는 YYYY-MM-DD 또는 YYYY-MM-DD HH:mm:ss 형식이어야 합니다.');
   y=Number(p[1]);m=Number(p[2]);d=Number(p[3]);h=Number(p[4]||0);min=Number(p[5]||0);sec=Number(p[6]||0);ms=Number((p[7]||'').padEnd(3,'0').slice(0,3));
  }
  check(y>=1900&&y<=9999&&m>=1&&m<=12&&d>=1&&d<=31&&h<24&&min<60&&sec<=60,'거래 날짜가 올바르지 않습니다.');
  const day=new Date(Date.UTC(y,m-1,d));check(day.getUTCFullYear()===y&&day.getUTCMonth()===m-1&&day.getUTCDate()===d,'존재하지 않는 거래 날짜입니다.');
  return new Date(Date.UTC(y,m-1,d,h,min,sec,ms)).toISOString().replace(/Z$/,'');
 }
 function zipLimit(bytes) {
  if(bytes[0]!==80||bytes[1]!==75)return;
  const u16=p=>bytes[p]|(bytes[p+1]<<8),u32=p=>(bytes[p]|(bytes[p+1]<<8)|(bytes[p+2]<<16)|(bytes[p+3]<<24))>>>0;
  let end=-1;for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--){if(u32(p)===0x06054b50&&p+22+u16(p+20)===bytes.length){end=p;break;}}
  check(end>=0,'손상된 XLSX 압축 파일입니다.');
  const count=u16(end+10);let p=u32(end+16),total=0;
  check(count>0&&count<=10000&&p<end,'엑셀 압축 구조 또는 크기를 확인해주세요.');
  for(let i=0;i<count;i++){
   check(p+46<=end&&u32(p)===0x02014b50,'손상된 XLSX 파일입니다.');
   check(!(u16(p+8)&1),'암호화된 엑셀은 암호를 해제한 뒤 가져오세요.');
   total+=u32(p+24);check(total<=64*1024*1024,'압축을 푼 엑셀이 64MB를 초과합니다.');
   p+=46+u16(p+28)+u16(p+30)+u16(p+32);check(p<=end,'손상된 XLSX 파일입니다.');
  }
 }
 function base64Bytes(value) {
  check(typeof value==='string'&&value.length<=45*1024*1024&&/^[A-Za-z0-9+/]*={0,2}$/.test(value),'파일 인코딩을 확인해주세요.');
  const alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/',bytes=[];let bits=0,count=0;
  for(const c of value){if(c==='=')break;bits=(bits<<6)|alphabet.indexOf(c);count+=6;if(count>=8){count-=8;bytes.push((bits>>>count)&255);}}
  return new Uint8Array(bytes);
 }
 function read(input,options={}) {
  check(root.XLSX,'엑셀 읽기 모듈을 찾을 수 없습니다.');
  const bytes=options.type==='base64'?base64Bytes(input):new Uint8Array(input);
  check(bytes.length>0&&bytes.length<=32*1024*1024,'32MB 이하의 엑셀을 선택해주세요.');
  const zip=bytes[0]===80&&bytes[1]===75,ole=bytes[0]===208&&bytes[1]===207&&bytes[2]===17&&bytes[3]===224;
  check(zip||ole,'XLSX 또는 XLS 형식의 엑셀 파일을 선택해주세요.');zipLimit(bytes);
  let workbook;
  try{workbook=root.XLSX.read(bytes,{type:'array',cellDates:false,cellFormula:false,cellHTML:false,cellText:false,bookVBA:false,sheetRows:100002});}
  catch(_){throw Error('엑셀을 읽지 못했습니다. 파일 손상·암호를 확인해주세요.');}
  const date1904=Boolean(workbook.Workbook&&workbook.Workbook.WBProps&&workbook.Workbook.WBProps.date1904),sheets=[];
  check(workbook.SheetNames.length<=100,'엑셀 시트가 100개를 초과합니다.');
  for(const name of workbook.SheetNames){
   const sheet=workbook.Sheets[name];if(!sheet||!sheet['!ref'])continue;
   const range=root.XLSX.utils.decode_range(sheet['!fullref']||sheet['!ref']);
   check(range.e.r<100001&&range.e.c<256&&(range.e.r+1)*(range.e.c+1)<=3000000,'시트 '+name+'의 행·열 범위가 너무 큽니다. 거래표만 내보내 주세요.');
   let columns=null,headerRow=0;
   for(let r=0;r<=Math.min(range.e.r,29);r++){
    const found={};
    for(let c=0;c<=range.e.c;c++){
     const cell=sheet[root.XLSX.utils.encode_cell({r,c})],label=normalize(cell&&cell.v);
     for(const [key,values] of Object.entries(aliases))if(values.includes(label)&&found[key]===undefined)found[key]=c;
    }
    if(found.title!==undefined&&found.amount!==undefined&&found.account!==undefined&&found.kind!==undefined&&(found.date!==undefined||found.month!==undefined)){columns=found;headerRow=r;break;}
   }
   if(!columns)continue;
   const rows=[];
   for(let r=headerRow+1;r<=range.e.r;r++){
    const values={},missingFormula=[];
    for(const [key,c] of Object.entries(columns)){
     const cell=sheet[root.XLSX.utils.encode_cell({r,c})];values[key]=cell&&cell.v!==undefined?cell.v:null;
     if(cell&&cell.f&&cell.v===undefined)missingFormula.push(key);
    }
    if([values.title,values.amount,values.date,values.month].every(v=>v===null||text(v)===''))continue;
    rows.push({row:r+1,values,missingFormula});
   }
   sheets.push({name,headerRow:headerRow+1,date1904,hasBurden:columns.burden!==undefined,rows});
  }
  check(sheets.length,'거래표를 찾지 못했습니다. 일시(또는 월)·자산·내용·금액·수입/지출 열이 있는 RAW 또는 편한가계부 추출 엑셀을 선택해주세요.');
  return {fileName:text(options.fileName)||'가져온 엑셀',sheets,preferred:sheets.some(s=>normalize(s.name)==='raw')?sheets.find(s=>normalize(s.name)==='raw').name:sheets[0].name};
 }
 function signature(value) {
  return sha256(JSON.stringify([value.date||null,value.month,text(value.account),text(value.title),value.gross,text(value.kind)]));
 }
 function similarSignature(value) {
  return sha256(JSON.stringify([value.date||null,value.month,text(value.account),text(value.title),text(value.kind)]));
 }
 function rawValues(raw) {
  const date=dateOf(raw['일시']),month=date?date.slice(0,7):monthOf(raw['월']);
  return {date,month,account:text(raw['자산']),title:text(raw['내용']),gross:money(raw['금액(KRW)']),kind:text(raw['수입/지출'])};
 }
 function inventory(state) {
  const archive=new Map(),active=new Map(),similar=new Set(),metadata=new Map();
  for(const r of state.sourceArchive){try{const value=rawValues(r),key=signature(value);archive.set(key,(archive.get(key)||0)+1);similar.add(similarSignature(value));if(!metadata.has(key))metadata.set(key,r);}catch(_){} }
  const collect=t=>{
   if(t.sourceKey&&t.sourceKey.startsWith('excel:v1:'))return; // Imported RAW already represents it.
   const kind=t.sourceKind||({expense:'지출',income:'수입',refund:'지출'}[t.operation]);if(!kind)return;
   try{const value={date:t.date?dateOf(t.date):null,month:t.month,account:t.account,title:t.title,gross:t.amount,kind},key=signature(value);active.set(key,(active.get(key)||0)+1);similar.add(similarSignature(value));}catch(_){}
  };
  state.transactions.forEach(collect);
  // Originals/tombstones retain source evidence even after a user edits or deletes it.
  const original=new Map();
  state.originals.forEach(t=>{try{const key=signature({date:t.date?dateOf(t.date):null,month:t.month,account:t.account,title:t.title,gross:t.amount,kind:t.sourceKind||'지출'});original.set(key,(original.get(key)||0)+1);}catch(_){} });
  for(const [key,count] of original)archive.set(key,Math.max(archive.get(key)||0,count));
  for(const [key,count] of active)archive.set(key,Math.max(archive.get(key)||0,count));
  return {counts:archive,similar,metadata};
 }
 function normalizedRow(row,sheet,state,fileName) {
  const v=row.values,date=dateOf(v.date,sheet.date1904),month=date?date.slice(0,7):monthOf(v.month,sheet.date1904);
  if(date&&v.month!==null&&v.month!==undefined&&text(v.month))check(monthOf(v.month,sheet.date1904)===month,'일시와 거래 월이 다릅니다.');
  const title=text(v.title),account=text(v.account),gross=money(v.amount),kind=text(v.kind);
  check(title.length>0&&title.length<=120&&account.length>0&&account.length<=100,'내용·자산 이름을 확인해주세요.');
  const explicitKRW=text(v.currency);check(!explicitKRW||['KRW','원','₩','￦'].includes(explicitKRW.toUpperCase()),'원화 거래만 지원합니다. 금액(KRW) 열로 내보내세요.');
  let operation=kind==='수입'?'income':kind==='지출'?'expense':'pending',status='confirmed',amount=gross;
  check(operation==='expense'||gross>=0,'수입·이체 금액은 0원 이상이어야 합니다.');
  const category=text(v.classification)||text(v.category)||'미분류',note=[text(v.note),text(v.remark)].filter(Boolean).join(' · '),warnings=[];
  const fixed=category==='고정비',accountKey=normalize(account);
  const funding=accountKey==='20260606'?'joint':/wrap|랩형/.test(accountKey)?'living':/카드|card|loca|쿠팡와우|bcgoat/.test(accountKey)?'card':'bank';
  if(funding==='joint'&&['expense','income'].includes(operation))operation='jointActivity';
  if(fixed&&operation==='expense')status='unconfirmed';
  if(['확정','confirmed'].includes(text(v.status)))status='confirmed';
  else if(['미확인','unconfirmed'].includes(text(v.status)))status='unconfirmed';
  else if(['예상','계획','planned'].includes(text(v.status)))status='planned';
  else if(text(v.status))throw Error('상태는 확정·미확인·예상 중 하나여야 합니다.');
  let alreadyNet=false;
  if(operation==='expense'&&sheet.hasBurden){
   if(v.burden===null||text(v.burden)===''){status='unconfirmed';warnings.push('본인부담 값이 없어 금액을 미확인으로 기록');}
   else{amount=money(v.burden,'본인부담');alreadyNet=amount!==gross;if(alreadyNet)warnings.push('원금액 '+gross+'원 · 본인부담 '+amount+'원 사용');}
  }
  if(['expense','income'].includes(operation)&&(/예비비|상환비|투자|정산|환급|대조|계좌납입|공동통장/.test(category+' '+text(v.method))||/정산|환급|캐시백|카드대금|카드결제대금|카드납부/.test(title))){operation='pending';warnings.push('자금 용도·연결 거래 확인 필요');}
  if(funding==='living'&&month<=state.asOf.slice(0,7)&&(!date||date.slice(0,10)<=state.asOf.slice(0,10))){operation='pending';warnings.push('기준일 이전 WRAP 거래 · 기준 잔액과 중복 여부 확인');}
  if(operation==='pending')warnings.push('이체·미분류는 보류로 저장 · 결산 미반영');
  if(fixed&&status==='unconfirmed')warnings.push('고정비는 실제 여부와 대체할 예상 항목 확인');
  const entry={id:0,date,month,title,account,category,amount,operation,status,funding,destination:'bank',expenseClass:fixed?'fixed':'variable',reserveID:null,linkedID:null,attributedMonth:null,alreadyNet,displayAmount:null,planID:null,note:[note,...warnings].filter(Boolean).join(' · ').slice(0,4000),sourceKey:null,sourceKind:kind};
  const raw={'일시':date,'월':month,'자산':account,'분류':text(v.category)||null,'소분류':text(v.subcategory)||null,'내용':title,'금액(KRW)':gross,'수입/지출':kind,'추가입력':text(v.note)||null,'재분류':text(v.classification)||null,'본인부담':v.burden===null||v.burden===undefined||text(v.burden)===''?null:money(v.burden,'본인부담'),'결제구분':text(v.method)||null,'비고':text(v.remark)||null,import_file:fileName,import_sheet:sheet.name,import_row:row.row};
  return {entry,raw,key:signature({date,month,account,title,gross,kind}),reason:warnings.join(' · ')};
 }
 function review(state,workbook,sheetName) {
  root.LedgerCore.validate(state);
  const sheet=workbook.sheets.find(s=>s.name===(sheetName||workbook.preferred));check(sheet,'가져올 시트를 선택해주세요.');
  const existing=inventory(state),seen=new Map(),rows=[];
  for(const row of sheet.rows){
   try{
    const normalized=normalizedRow(row,sheet,state,workbook.fileName),occurrence=(seen.get(normalized.key)||0)+1;seen.set(normalized.key,occurrence);
    const key='excel:v1:'+normalized.key+':'+occurrence,known=occurrence<=(existing.counts.get(normalized.key)||0)||state.transactions.some(t=>t.sourceKey===key)||state.audit.some(a=>a.before&&a.before.sourceKey===key);
    const amountConflict=!known&&!existing.counts.has(normalized.key)&&existing.similar.has(similarSignature(rawValues(normalized.raw))),old=existing.metadata.get(normalized.key);
    const metadataConflict=known&&old&&((normalized.raw['재분류']&&old['재분류']&&normalized.raw['재분류']!==old['재분류'])||(normalized.raw['본인부담']!==null&&old['본인부담']!==null&&normalized.raw['본인부담']!==old['본인부담']));
    normalized.entry.sourceKey=key;
    rows.push({id:sheet.name+':'+row.row,row:row.row,disposition:known?'existing':'new',include:!known&&!amountConflict,reason:known?(metadataConflict?'같은 거래의 분류·본인부담이 다릅니다. 기존값을 유지했어요. 기록 수정에서 확인해주세요.':'이미 보관·입력한 거래 · 중복 제외'):amountConflict?'동일 일시·자산·내용의 기존 거래와 금액이 달라 선택을 해제했어요. 새 거래인지 확인한 뒤 선택해주세요.':normalized.reason,entry:normalized.entry,raw:normalized.raw});
   }catch(error){rows.push({id:sheet.name+':'+row.row,row:row.row,disposition:'invalid',include:false,reason:error.message,entry:null,raw:null});}
  }
  return {fileName:workbook.fileName,sheet:sheet.name,baseStamp:sha256(JSON.stringify(state)),total:rows.length,existing:rows.filter(r=>r.disposition==='existing').length,invalid:rows.filter(r=>r.disposition==='invalid').length,rows};
 }
 function apply(state,proposal) {
  check(proposal&&proposal.baseStamp===sha256(JSON.stringify(state)),'검토 중 가계부가 변경됐어요. 엑셀을 다시 선택해주세요.');
  const selected=proposal.rows.filter(r=>r.disposition==='new'&&r.include),entries=[],archiveRows=[];
  check(selected.length,'추가할 거래를 선택해주세요.');
  for(const row of selected){
   const e=clone(row.entry),key=signature(rawValues(row.raw));
   check(typeof e.sourceKey==='string'&&e.sourceKey.startsWith('excel:v1:'+key+':'),'엑셀 원본 키를 확인해주세요.');
   if(e.planID){const plan=state.plans.find(p=>p.id===e.planID);check(plan&&plan.month===e.month,'같은 월의 예상 항목을 선택해주세요.');}
   entries.push(e);archiveRows.push(clone(row.raw));
  }
  return root.LedgerCore.mutate(state,{type:'import',entries,archiveRows});
 }
 const api={read,review,apply,sha256};root.WorkbookImport=api;
 root.excelDispatch=function(input){try{const q=JSON.parse(input);let value;if(q.action==='read')value=read(q.file,{type:'base64',fileName:q.fileName});else if(q.action==='review')value=review(q.state,q.workbook,q.sheet);else if(q.action==='apply')value=apply(q.state,q.proposal);else throw Error('지원하지 않는 엑셀 작업입니다.');return JSON.stringify({ok:true,value});}catch(e){return JSON.stringify({ok:false,error:e.message||String(e)});}};
 if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof globalThis!=='undefined'?globalThis:this);
