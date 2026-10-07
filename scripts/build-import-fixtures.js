'use strict';
const fs=require('fs'),path=require('path');
const XLSX=require('../core/vendor/xlsx.full.min.js');
const target=path.resolve(__dirname,'../tests/fixtures');fs.mkdirSync(target,{recursive:true});
const columns=['일시','자산','내용','금액(KRW)','수입/지출','재분류','추가입력','상태'];
const rows=[
 ['2026-10-07 12:30:00','미래에셋_WRAP형CMA','엑셀 커피 테스트',10000,'지출','변동비','가져오기 테스트','확정'],
 ['2026-10-07 13:00:00','CASH','엑셀 월세 테스트',400000,'지출','고정비','예상 월세를 실제로','확정'],
 ['2026-10-07 14:00:00','미래에셋_WRAP형CMA','엑셀 이체 테스트',30000,'이체출금','이체',null,'확정']
];
function write(name,values,format='xlsx'){
 const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([columns,...values]),'RAW');
 fs.writeFileSync(path.join(target,name),XLSX.write(book,{bookType:format,type:'buffer'}));
}
write('import.xlsx',rows);write('import.xls',rows.slice().reverse(),'xls');
write('overlap.xlsx',[...rows,['2026-10-08 12:00:00','일반 계좌','추가 기간 거래',5000,'지출','변동비',null,'확정']]);
write('invalid.xlsx',[['2026-02-30','계좌','잘못된 날짜',1000,'지출'],['2026-10-07','계좌','소수점 금액',1.5,'지출']]);
const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['월','합계'],['10월',1000]]),'결산');fs.writeFileSync(path.join(target,'summary-only.xlsx'),XLSX.write(book,{bookType:'xlsx',type:'buffer'}));
console.log(JSON.stringify({fixtures:5,private_data:false}));
