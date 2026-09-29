/* Çizelge dışa aktarma: Excel (.xlsx) ve PDF — tamamen çevrimdışı, kütüphanesiz */
const HEAD=["Tarih","Başlama Zamanı","Süresi","Ağrının Şiddeti (0-10)","Tetikleyiciler","Ağrı Kesici","Bulantı","Kusma","Işıktan Rahatsızlık","Sesten Rahatsızlık"];
const enc=new TextEncoder();
function rowsOf(y,m){
  const out=[],n=new Date(y,m+1,0).getDate();
  for(let d=1;d<=n;d++)for(const e of(data[key(y,m,d)]||[]))
    out.push([`${String(d).padStart(2,"0")}.${String(m+1).padStart(2,"0")}.${y}`,e.zaman||"",e.sure?`${e.sure} ${e.birim}`:"",String(e.sev),[e.tet,e.gida].filter(Boolean).join(" + "),e.kesici||"",e.bulanti?"X":"",e.kusma?"X":"",e.isik?"X":"",e.ses?"X":""]);
  return out;
}
function saveBlob(blob,name){
  const f=new File([blob],name,{type:blob.type});
  if(/Android/i.test(navigator.userAgent)&&navigator.canShare&&navigator.canShare({files:[f]})){navigator.share({files:[f],title:name}).catch(()=>{});return}
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();a.remove();
}
const fname=(y,m,ext)=>`bas-agrisi-gunlugu-${y}-${String(m+1).padStart(2,"0")}.${ext}`;

/* ---------- XLSX ---------- */
const crcT=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
const crc=b=>{let c=~0;for(const x of b)c=crcT[(c^x)&255]^(c>>>8);return ~c>>>0};
function zip(files){
  const parts=[],cd=[];let off=0;
  const u16=v=>[v&255,v>>8&255],u32=v=>[v&255,v>>8&255,v>>16&255,v>>24&255];
  for(const[name,txt]of files){
    const nb=enc.encode(name),d=enc.encode(txt),c=crc(d);
    const lh=new Uint8Array([0x50,0x4b,3,4,20,0,0,8,0,0,0,0,0x21,0,...u32(c),...u32(d.length),...u32(d.length),...u16(nb.length),0,0,...nb]);
    parts.push(lh,d);
    cd.push(new Uint8Array([0x50,0x4b,1,2,20,0,20,0,0,8,0,0,0,0,0x21,0,...u32(c),...u32(d.length),...u32(d.length),...u16(nb.length),0,0,0,0,0,0,0,0,0,0,0,0,...u32(off),...nb]));
    off+=lh.length+d.length;
  }
  const cs=cd.reduce((a,b)=>a+b.length,0);
  return new Blob([...parts,...cd,new Uint8Array([0x50,0x4b,5,6,0,0,0,0,...u16(files.length),...u16(files.length),...u32(cs),...u32(off),0,0])],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
const esc=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const colL=i=>String.fromCharCode(65+i);
function exportXlsx(){
  const y=cur.getFullYear(),m=cur.getMonth();
  const rows=rowsOf(y,m);while(rows.length<20)rows.push(HEAD.map(()=>""));
  const W=[13,16,12,14,44,20,9,9,14,14];
  const cell=(v,i,r,st)=>`<c r="${colL(i)}${r}" s="${st}"${v!==""&&!(i==3&&r>1)?' t="inlineStr"><is><t>'+esc(v)+"</t></is></c>":(v!==""?`><v>${v}</v></c>`:"/>")}`;
  let sd=`<row r="1" ht="40" customHeight="1">${HEAD.map((h,i)=>cell(h,i,1,1)).join("")}</row>`;
  rows.forEach((r,k)=>sd+=`<row r="${k+2}" ht="26" customHeight="1">${r.map((v,i)=>cell(v,i,k+2,2)).join("")}</row>`);
  const X='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const ns='xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const b=zip([
    ["[Content_Types].xml",X+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
    ["_rels/.rels",X+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ["xl/workbook.xml",X+`<workbook ${ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${AY[m]} ${y}" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels",X+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
    ["xl/styles.xml",X+`<styleSheet ${ns}><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" borderId="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf borderId="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs></styleSheet>`],
    ["xl/worksheets/sheet1.xml",X+`<worksheet ${ns}><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews><cols>${W.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${sd}</sheetData><pageSetup orientation="landscape" fitToHeight="0"/></worksheet>`]
  ]);
  saveBlob(b,fname(y,m,"xlsx"));
}

/* ---------- PDF (çizelge görüntüsü A4 yatay) ---------- */
function wrap(ctx,t,w){
  const words=String(t).split(" "),ls=[];let l="";
  for(const x of words){const n=l?l+" "+x:x;if(ctx.measureText(n).width>w&&l){ls.push(l);l=x}else l=n}
  if(l)ls.push(l);return ls;
}
function pages(y,m){
  const PW=1754,PH=1240,M=50,TW=PW-2*M,rows=rowsOf(y,m);
  const fr=[.085,.09,.075,.09,.235,.14,.06,.06,.085,.08];
  const cw=fr.map(f=>f*TW),RH=54,HH=90,TOP=150,per=Math.floor((PH-TOP-HH-70)/RH);
  const n=Math.max(1,Math.ceil(rows.length/per)),out=[];
  for(let p=0;p<n;p++){
    const c=document.createElement("canvas");c.width=PW;c.height=PH;const x=c.getContext("2d");
    x.fillStyle="#fff";x.fillRect(0,0,PW,PH);x.fillStyle="#000";x.strokeStyle="#000";
    x.font="bold 40px Arial,sans-serif";x.textAlign="center";x.fillText("BAŞ AĞRISI GÜNLÜĞÜ",PW/2,70);
    x.font="22px Arial,sans-serif";x.fillText(`${AY[m]} ${y}` + (n>1?`  (${p+1}/${n})`:""),PW/2,105);
    const slice=rows.slice(p*per,(p+1)*per);while(slice.length<per&&p==n-1&&slice.length<Math.min(per,18))slice.push(HEAD.map(()=>""));
    const H=HH+slice.length*RH;let cx=M;
    x.lineWidth=2;x.strokeRect(M,TOP,TW,H);x.lineWidth=1.5;
    x.beginPath();x.moveTo(M,TOP+HH);x.lineTo(M+TW,TOP+HH);
    slice.forEach((_,i)=>{const yy=TOP+HH+(i+1)*RH;x.moveTo(M,yy);x.lineTo(M+TW,yy)});
    cw.forEach((w,i)=>{cx+=w;if(i<cw.length-1){x.moveTo(cx,TOP);x.lineTo(cx,TOP+H)}});x.stroke();
    cx=M;
    cw.forEach((w,i)=>{x.font="bold 19px Arial,sans-serif";x.textAlign="center";
      const ls=wrap(x,HEAD[i],w-14);ls.forEach((l,k)=>x.fillText(l,cx+w/2,TOP+HH/2-(ls.length-1)*12+k*24+7));cx+=w});
    slice.forEach((r,ri)=>{cx=M;cw.forEach((w,i)=>{x.font="20px Arial,sans-serif";x.textAlign="center";
      const ls=wrap(x,r[i],w-14).slice(0,2);ls.forEach((l,k)=>x.fillText(l,cx+w/2,TOP+HH+ri*RH+RH/2-(ls.length-1)*11+k*22+7));cx+=w})});
    out.push(c.toDataURL("image/jpeg",.92));
  }
  return out;
}
function exportPdf(){
  const y=cur.getFullYear(),m=cur.getMonth(),imgs=pages(y,m),ch=[];let len=0;
  const add=v=>{const b=typeof v=="string"?enc.encode(v):v;ch.push(b);len+=b.length};
  const offs=[];const obj=(n,fn)=>{offs[n]=len;add(`${n} 0 obj\n`);fn();add("\nendobj\n")};
  add("%PDF-1.4\n");
  obj(1,()=>add("<</Type/Catalog/Pages 2 0 R>>"));
  obj(2,()=>add(`<</Type/Pages/Count ${imgs.length}/Kids[${imgs.map((_,i)=>`${3+i*3} 0 R`).join(" ")}]>>`));
  imgs.forEach((d,i)=>{
    const p=3+i*3,bin=Uint8Array.from(atob(d.split(",")[1]),c=>c.charCodeAt(0));
    obj(p,()=>add(`<</Type/Page/Parent 2 0 R/MediaBox[0 0 842 595]/Resources<</XObject<</Im0 ${p+2} 0 R>>>>/Contents ${p+1} 0 R>>`));
    const cs="q 842 0 0 595 0 0 cm /Im0 Do Q";
    obj(p+1,()=>add(`<</Length ${cs.length}>>\nstream\n${cs}\nendstream`));
    obj(p+2,()=>{add(`<</Type/XObject/Subtype/Image/Width 1754/Height 1240/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/DCTDecode/Length ${bin.length}>>\nstream\n`);add(bin);add("\nendstream")});
  });
  const N=3+imgs.length*3,xr=len;
  add(`xref\n0 ${N}\n0000000000 65535 f \n`+offs.slice(1,N).map(o=>String(o).padStart(10,"0")+" 00000 n \n").join("")+`trailer\n<</Size ${N}/Root 1 0 R>>\nstartxref\n${xr}\n%%EOF`);
  saveBlob(new Blob(ch,{type:"application/pdf"}),fname(y,m,"pdf"));
}
document.getElementById("xl").onclick=exportXlsx;
document.getElementById("pdf").onclick=exportPdf;
