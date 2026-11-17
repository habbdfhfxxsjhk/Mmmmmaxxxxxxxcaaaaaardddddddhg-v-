/* MAX CARD Excel reader: local XLSX/CSV parsing, no data uploaded. */
(function(global){
  'use strict';
  const decoder=new TextDecoder('utf-8');
  const MAX_BYTES=5*1024*1024, MAX_UNPACKED=16*1024*1024;
  function csv(text){
    text=text.replace(/^\uFEFF/,'');const rows=[];let row=[],field='',quoted=false;
    for(let i=0;i<text.length;i++){const c=text[i];
      if(quoted){if(c==='"'&&text[i+1]==='"'){field+='"';i++;}else if(c==='"')quoted=false;else field+=c;}
      else if(c==='"'){quoted=true;}else if(c===','){row.push(field);field='';}
      else if(c==='\n'||c==='\r'){if(c==='\r'&&text[i+1]==='\n')i++;row.push(field);if(row.some(x=>x.trim()))rows.push(row);row=[];field='';}
      else field+=c;
    }row.push(field);if(row.some(x=>x.trim()))rows.push(row);
    if(rows.length>2001)throw Error('الملف يتجاوز 2000 سطر');return rows;
  }
  function numberAt(view,p){return view.getUint32(p,true)}
  function xml(text){const doc=new DOMParser().parseFromString(text,'application/xml');
    if(doc.getElementsByTagName('parsererror').length)throw Error('ملف Excel غير سليم (XML)');return doc;}
  async function xlsx(buffer){
    const b=new Uint8Array(buffer),v=new DataView(buffer);let end=-1;
    for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--){if(numberAt(v,i)===0x06054b50){end=i;break;}}
    if(end<0)throw Error('لم يتم التعرف على ملف Excel');
    const count=v.getUint16(end+10,true),dir=numberAt(v,end+16);if(count>1500)throw Error('ملف Excel غير مدعوم');
    let pos=dir;const directory=new Map();
    for(let i=0;i<count;i++){
      if(numberAt(v,pos)!==0x02014b50)throw Error('فهرس Excel تالف');
      const method=v.getUint16(pos+10,true),compressed=numberAt(v,pos+20),size=numberAt(v,pos+24);
      const nameSize=v.getUint16(pos+28,true),extra=v.getUint16(pos+30,true),comment=v.getUint16(pos+32,true);
      const local=numberAt(v,pos+42),name=decoder.decode(b.subarray(pos+46,pos+46+nameSize));
      if(local>=b.length||size>MAX_UNPACKED)throw Error('ملف Excel كبير أو غير صالح');
      directory.set(name,{method,compressed,size,local});pos+=46+nameSize+extra+comment;
    }
    async function read(name){
      const entry=directory.get(name);if(!entry)return null;
      const p=entry.local;if(numberAt(v,p)!==0x04034b50)throw Error('ملف Excel غير صالح');
      const start=p+30+v.getUint16(p+26,true)+v.getUint16(p+28,true);
      if(start+entry.compressed>b.length)throw Error('الملف غير مكتمل');
      const bytes=b.slice(start,start+entry.compressed);let output;
      if(entry.method===0)output=bytes;
      else if(entry.method===8){if(typeof DecompressionStream==='undefined')throw Error('متصفحك لا يدعم فك Excel: يمكنك حفظ الملف بصيغة CSV');
        const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        output=new Uint8Array(await new Response(stream).arrayBuffer());
      }else throw Error('نوع ضغط Excel غير مدعوم');
      if(output.length>MAX_UNPACKED)throw Error('ملف Excel كبير جداً');return decoder.decode(output);
    }
    let shared=[];const sharedXml=await read('xl/sharedStrings.xml');
    if(sharedXml){shared=Array.from(xml(sharedXml).getElementsByTagName('si'),si=>
      Array.from(si.getElementsByTagName('t')).map(x=>x.textContent).join(''));}
    const names=Array.from(directory.keys()).filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();
    const docText=await read(names[0]||'xl/worksheets/sheet1.xml');
    if(!docText)throw Error('ملف Excel لا يحتوي ورقة بيانات');
    const doc=xml(docText),result=[];
    for(const row of Array.from(doc.getElementsByTagName('row'))){
      const arr=[];for(const c of Array.from(row.getElementsByTagName('c'))){
        const match=/^([A-Z]+)/i.exec(c.getAttribute('r')||'');if(!match)continue;
        let col=0;for(const letter of match[1].toUpperCase())col=col*26+letter.charCodeAt(0)-64;col--;
        if(col>=32)continue;
        const type=c.getAttribute('t'),valNode=c.getElementsByTagName('v')[0];
        let value=valNode?valNode.textContent:'';
        if(type==='s')value=shared[Number(value)]??'';
        else if(type==='inlineStr')value=Array.from(c.getElementsByTagName('t')).map(x=>x.textContent).join('');
        else if(type==='b')value=value==='1'?'TRUE':'FALSE';
        arr[col]=value;
      }
      if(arr.some(x=>String(x||'').trim()))result.push(arr.map(x=>x??''));
      if(result.length>2001)throw Error('الملف يتجاوز 2000 سطر');
    }
    return result;
  }
  async function read(file){
    if(!file)throw Error('اختر الملف أولاً');
    if(file.size>MAX_BYTES)throw Error('الحد الأقصى للملف 5MB');
    if(/\.csv$/i.test(file.name))return csv(await file.text());
    if(/\.xlsx$/i.test(file.name))return xlsx(await file.arrayBuffer());
    throw Error('اختر ملف XLSX أو CSV فقط');
  }
  global.MaxExcelReader={read,csv,xlsx};
})(window);
