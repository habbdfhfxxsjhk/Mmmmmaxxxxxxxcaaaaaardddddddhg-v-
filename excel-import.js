/* MAX CARD FINANCE PRO — Excel inventory workflow. Entirely local. */
(function(){
'use strict';
const KEY='maxcard_finance_v1',START='2026-11-17';
const $=id=>document.getElementById(id);
const WALLETS=['نقدي','شام كاش دولار','شام كاش ليرة','سيريتل كاش','MTN كاش','USDT','أخرى'];
const SCHEMA={
 sale:{name:'المبيعات',file:'MAX_CARD_SALES.xlsx',headers:['التاريخ','الخدمة','العميل','رقم الطلب','المحفظة','قيمة البيع USD','ملاحظات']},
 purchase:{name:'المشتريات',file:'MAX_CARD_PURCHASES.xlsx',headers:['التاريخ','الخدمة','المورد','رقم الطلب','المحفظة','تكلفة الشراء USD','ملاحظات']},
 balance:{name:'الأرصدة',file:'MAX_CARD_BALANCES.xlsx',headers:['التاريخ','المحفظة','الرصيد الفعلي USD','المرجع','ملاحظات']}
};
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const today=()=>{const d=new Date();return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');};
const amount=n=>(Number(n)||0).toFixed(2)+' $';
const norm=s=>String(s??'').replace(/\uFEFF/g,'').trim().replace(/\s+/g,' ');
function loadDB(){
 const x=JSON.parse(localStorage.getItem(KEY)||'null')||{version:1,initial:150,entries:[],closedDays:[],closedCycles:[]};
 if(!Array.isArray(x.entries)||!Array.isArray(x.closedDays)||!Array.isArray(x.closedCycles))throw Error('بيانات النظام غير سليمة. استعد نسخة احتياطية أولاً.');
 return x;
}
function dateValue(x){
 const s=norm(x);
 if(/^\d{4}-\d{2}-\d{2}$/.test(s)){
  const [y,m,d]=s.split('-').map(Number);
  if(new Date(Date.UTC(y,m-1,d)).toISOString().slice(0,10)===s)return s;
 }
 if(/^\d{5}(\.\d+)?$/.test(s)){const n=Number(s);if(n>=25000&&n<=90000)return new Date(Date.UTC(1899,11,30)+Math.floor(n)*86400000).toISOString().slice(0,10);}
 return '';
}
function walletValue(x){
 const s=norm(x);
 const aliases={'كاش':'نقدي','شام دولار':'شام كاش دولار','شام ليرة':'شام كاش ليرة','سيريتل':'سيريتل كاش','MTN':'MTN كاش','usdt':'USDT'};
 const w=aliases[s]||s;return WALLETS.includes(w)?w:'';
}
function signed(e){
 if(e.type==='transfer')return 0;
 if(['sale','deposit','adjustment','reconcile_in'].includes(e.type))return Number(e.amount)||0;
 return -(Number(e.amount)||0);
}
function getWalletBalances(db,maxDate,staged){
 const result=Object.fromEntries(WALLETS.map(w=>[w,w==='نقدي'?Number(db.initial||0):0]));
 for(const e of staged||db.entries){
  if(maxDate&&e.date>maxDate)continue;
  if(e.type==='transfer'){
   result[e.wallet]=(result[e.wallet]||0)-Number(e.amount||0);
   if(e.toWallet)result[e.toWallet]=(result[e.toWallet]||0)+Number(e.amount||0);
  }else result[e.wallet]=(result[e.wallet]||0)+signed(e);
 }
 return result;
}
function locked(db,date){
 return db.closedDays.includes(date)||db.closedCycles.some(c=>c&&date>=c.start&&date<=c.end);
}
function mapRows(kind,fileName,raw){
 const expected=SCHEMA[kind].headers.map(norm);
 if(!raw.length)throw Error('الملف فارغ');
 const head=raw[0].map(norm),columns=expected.map(x=>head.indexOf(x));
 if(columns.some(i=>i<0))throw Error('عناوين الأعمدة لا تطابق النموذج المختار. تأكد من اختيار نوع الجرد الصحيح وعدم تغيير العناوين.');
 const get=(r,i)=>r[columns[i]]??'';
 const results=[];
 raw.slice(1).forEach((r,i)=>{
  if(!r.some(x=>norm(x)))return;
  let item={kind,fileName,rowNumber:i+2,date:get(r,0),wallet:'',amount:'',note:'',service:'',party:'',order:''};
  if(kind==='balance'){item.wallet=get(r,1);item.amount=get(r,2);item.order=get(r,3);item.note=get(r,4);}
  else{item.service=get(r,1);item.party=get(r,2);item.order=get(r,3);item.wallet=get(r,4);item.amount=get(r,5);item.note=get(r,6);}
  results.push(item);
 });
 if(results.length>2000)throw Error('الحد الأقصى 2000 صف');
 return results;
}
function plan(rows,db){
 const imported=new Set(db.entries.filter(e=>e.importKey).map(e=>e.importKey));
 const withinFile=new Set(),snapshots=new Set(),staged=[...db.entries],results=[];
 const lastBalanceDate={};
 for(const r of rows){
  const date=dateValue(r.date),wallet=walletValue(r.wallet),value=Number(norm(r.amount).replace(/,/g,''));
  const order=norm(r.order).slice(0,120),service=norm(r.service).slice(0,120),party=norm(r.party).slice(0,120),note=norm(r.note).slice(0,400);
  const key=['excel',r.kind,r.fileName,r.rowNumber,date,wallet,value,service,party,order,note].join('|');
  let status='ready',reason='جاهز للإضافة',entry=null;
  if(!date||date<START||date>today()){status='error';reason='تاريخ غير صحيح أو خارج الفترة';}
  else if(!wallet){status='error';reason='المحفظة غير موجودة بالنظام';}
  else if(!Number.isFinite(value)||value<0||value>1e9||(r.kind!=='balance'&&value===0)){status='error';reason='مبلغ غير صحيح';}
  else if(Math.abs(value*100-Math.round(value*100))>.00001){status='error';reason='المبلغ يقبل منزلتين عشريتين فقط';}
  else if(locked(db,date)){status='error';reason='التاريخ مغلق مالياً';}
  else if(imported.has(key)||withinFile.has(key)){status='duplicate';reason='مكرر — سبق استيراد نفس السطر';}
  else{
   withinFile.add(key);
   const base={date,wallet,amount:Math.round(value*100)/100,customer:party,importKey:key,source:'excel',
    note:[service,order?'الطلب/المرجع: '+order:'',note].filter(Boolean).join(' • ')||'مستورد من Excel'};
   if(r.kind==='balance'){
    const unique=date+'|'+wallet;
    if(snapshots.has(unique)){status='error';reason='تكرر رصيد نفس المحفظة والتاريخ';}
    else if(lastBalanceDate[wallet]&&date<lastBalanceDate[wallet]){status='error';reason='رتب تواريخ أرصدة كل محفظة من الأقدم للأحدث';}
    else{
     snapshots.add(unique);lastBalanceDate[wallet]=date;
     const current=Math.round((getWalletBalances(db,date,staged)[wallet]||0)*100)/100;
     const diff=Math.round((base.amount-current)*100)/100;
     if(diff===0){status='unchanged';reason='الرصيد مطابق، لا حاجة لتسوية';}
     else{
      entry={...base,type:diff>0?'reconcile_in':'reconcile_out',amount:Math.abs(diff),
       customer:'تسوية جرد الرصيد',note:'فرق جرد: الفعلي '+base.amount.toFixed(2)+'$، المحسوب '+current.toFixed(2)+'$ • '+base.note};
      reason='تسوية فرق '+(diff>0?'+':'')+diff.toFixed(2)+' $';
     }
    }
   }else if(r.kind==='sale'||r.kind==='purchase'){
    entry={...base,type:r.kind};
   }else{status='error';reason='نوع الجرد غير معروف';}
   if(entry)staged.push(entry);
  }
  results.push({rowNumber:r.rowNumber,date:date||norm(r.date),wallet:wallet||norm(r.wallet),amount:Number.isFinite(value)?value:0,status,reason,entry});
 }
 return results;
}
let loaded=[];
const showStatus=(text,bad=false)=>{$('excelStatus').textContent=text;$('excelStatus').style.color=bad?'#ff9cab':'#d5c5ed';};
function clearPreview(){
 loaded=[];$('excelPreview').replaceChildren();$('excelKpis').replaceChildren();$('excelAdd').disabled=true;showStatus('');
}
function renderPreview(){
 const list=plan(loaded,loadDB()),ready=list.filter(r=>r.status==='ready'),errors=list.filter(r=>r.status==='error');
 const dupes=list.filter(r=>r.status==='duplicate'),unchanged=list.filter(r=>r.status==='unchanged');
 $('excelKpis').innerHTML='<span>✅ جاهز: '+ready.length+'</span><span>🔁 مكرر: '+dupes.length+'</span><span>🔒/⚠️ مرفوض: '+errors.length+'</span><span>✓ مطابق: '+unchanged.length+'</span>';
 $('excelAdd').disabled=ready.length===0;
 const names={ready:'جاهز',duplicate:'مكرر',error:'مرفوض',unchanged:'مطابق'};
 $('excelPreview').innerHTML=list.length?'<table><thead><tr><th>سطر</th><th>التاريخ</th><th>المحفظة</th><th>القيمة</th><th>النتيجة</th><th>التفاصيل</th></tr></thead><tbody>'+
  list.slice(0,200).map(r=>'<tr><td>'+r.rowNumber+'</td><td>'+escape(r.date)+'</td><td>'+escape(r.wallet)+'</td><td>'+amount(r.amount)+'</td><td class="'+(r.status==='ready'?'excel-success':r.status==='error'?'excel-error':'excel-skip')+'">'+names[r.status]+'</td><td>'+escape(r.reason)+'</td></tr>').join('')+'</tbody></table>':'<p class="muted">لا توجد أسطر بيانات في الملف.</p>';
 if(list.length>200)$('excelPreview').insertAdjacentHTML('beforeend','<p class="muted">عرض أول 200 سطر للمعاينة. سيتم فحص جميع السطور عند الإضافة.</p>');
 showStatus('تم فحص '+list.length+' سطر. سيُضاف '+ready.length+' فقط. تأكد من الأرقام قبل الاعتماد.');
}
async function preview(){
 clearPreview();
 const file=$('excelFile').files[0],kind=$('excelKind').value;
 if(!file){showStatus('اختر ملف Excel أولاً.',true);return;}
 $('excelCheck').disabled=true;showStatus('جارٍ قراءة الملف وفحص العمليات...');
 try{
  const raw=await window.MaxExcelReader.read(file);
  loaded=mapRows(kind,file.name,raw);
  renderPreview();
 }catch(e){clearPreview();showStatus('تعذرت القراءة: '+e.message,true);}
 finally{$('excelCheck').disabled=false;}
}
function add(){
 if(!loaded.length)return;
 try{
  const db=loadDB(),check=plan(loaded,db),ready=check.filter(r=>r.status==='ready').map(r=>r.entry);
  if(!ready.length){renderPreview();showStatus('لا توجد سطور جديدة قابلة للإضافة.',true);return;}
  if(db.entries.length+ready.length>50000)throw Error('تجاوزت الحد المسموح للسجلات');
  if(!confirm('إضافة '+ready.length+' حركة إلى السجل المالي وتحديث لوحة التحكم؟ يرجى التأكد أنك أخذت نسخة احتياطية.'))return;
  const at=new Date().toISOString();
  db.entries.push(...ready.map((e,i)=>({...e,id:typeof crypto.randomUUID==='function'?crypto.randomUUID():Date.now()+'-'+i+'-'+Math.random(),created:at})));
  localStorage.setItem(KEY,JSON.stringify(db));
  showStatus('✅ تم حفظ '+ready.length+' حركة بنجاح. يتم تحديث لوحة التحكم...');
  $('excelAdd').disabled=true;
  setTimeout(()=>location.reload(),650);
 }catch(e){showStatus('فشلت الإضافة: '+e.message,true);}
}
function renderWallets(){
 const target=$('walletKpis');if(!target)return;
 try{
  const wallet=getWalletBalances(loadDB());
  target.innerHTML=WALLETS.map(w=>'<div class="card"><span class="muted">'+escape(w)+'</span><strong>'+amount(wallet[w])+'</strong></div>').join('');
 }catch(e){target.textContent='تعذر حساب أرصدة المحافظ';}
}
/* A small dependency-free Excel/OpenXML template writer (stored ZIP). */
function templateBlob(kind){
 const schema=SCHEMA[kind],esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
 const cells=schema.headers.map((h,i)=>'<c r="'+String.fromCharCode(65+i)+'1" s="1" t="inlineStr"><is><t>'+esc(h)+'</t></is></c>').join('');
 const worksheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="7" width="24" customWidth="1"/></cols><sheetData><row r="1" ht="29" customHeight="1">'+cells+'</row></sheetData></worksheet>';
 const files={
 '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
 '_rels/.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
 'xl/workbook.xml':'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="'+esc(schema.name)+'" sheetId="1" r:id="rId1"/></sheets></workbook>',
 'xl/_rels/workbook.xml.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
 'xl/styles.xml':'<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="12"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF432160"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/></cellXfs></styleSheet>',
 'xl/worksheets/sheet1.xml':worksheet
 };
 const encoder=new TextEncoder(),parts=[],central=[];let offset=0;
 const crcTable=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let j=0;j<8;j++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;crcTable[n]=c>>>0;}
 const crc32=bytes=>{let c=0xffffffff;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0;};
 function header(size){return new Uint8Array(size);}
 const put16=(v,p,n)=>v.setUint16(p,n,true),put32=(v,p,n)=>v.setUint32(p,n>>>0,true);
 for(const [filename,content] of Object.entries(files)){
  const name=encoder.encode(filename),data=encoder.encode(content),crc=crc32(data);
  const local=header(30+name.length),lv=new DataView(local.buffer);
  put32(lv,0,0x04034b50);put16(lv,4,20);put16(lv,6,0x0800);put16(lv,8,0);
  put32(lv,14,crc);put32(lv,18,data.length);put32(lv,22,data.length);put16(lv,26,name.length);local.set(name,30);
  parts.push(local,data);
  const cent=header(46+name.length),cv=new DataView(cent.buffer);
  put32(cv,0,0x02014b50);put16(cv,4,20);put16(cv,6,20);put16(cv,8,0x0800);
  put32(cv,16,crc);put32(cv,20,data.length);put32(cv,24,data.length);put16(cv,28,name.length);put32(cv,42,offset);cent.set(name,46);
  central.push(cent);offset+=local.length+data.length;
 }
 const centralSize=central.reduce((n,x)=>n+x.length,0),end=header(22),ev=new DataView(end.buffer);
 put32(ev,0,0x06054b50);put16(ev,8,central.length);put16(ev,10,central.length);put32(ev,12,centralSize);put32(ev,16,offset);
 const total=offset+centralSize+22,output=new Uint8Array(total);let at=0;
 for(const part of [...parts,...central,end]){output.set(part,at);at+=part.length;}
 return new Blob([output],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function download(kind){
 const file=SCHEMA[kind].file,blob=templateBlob(kind),href=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=href;a.download=file;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(href),5000);
}
document.querySelectorAll('[data-template]').forEach(b=>b.addEventListener('click',()=>download(b.dataset.template)));
$('excelCheck').addEventListener('click',preview);
$('excelAdd').addEventListener('click',add);
$('excelClear').addEventListener('click',()=>{$('excelFile').value='';clearPreview();});
$('excelFile').addEventListener('change',clearPreview);
$('excelKind').addEventListener('change',()=>{$('excelFile').value='';clearPreview();});
renderWallets();
})();
