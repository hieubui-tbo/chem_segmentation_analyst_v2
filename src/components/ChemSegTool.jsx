'use client';
import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import _ from "lodash";

// ═══════════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════════
const UNIT_CONVERSIONS = {
  TNE: 1, TONNE: 1, TONNES: 1, "METRIC TON": 1, "METRIC TONS": 1, MT: 1,
  "TONNE (METRIC TON)": 1,
  KGM: 0.001, GRM: 0.000001, LBS: 0.00045359237,
};
const AMBIGUOUS_UNITS = ["TON", "TONS"];
const INVALID_NAMES = [
  "NONE","OTHER","NOT FOUND","UNKNOWN","N/A","NA","NIL",
  "TO THE ORDER OF","TO ORDER","TO THE ORDER","NOT AVAILABLE",
  "NOT APPLICABLE","UNSPECIFIED","TBD","TBA",
];
const TEXT_PLACEHOLDERS = [...INVALID_NAMES, "-", "--"];
const SUFFIXES = [
  "COMPANY LIMITED","JOINT STOCK COMPANY","S.A. DE C.V.","SA DE CV",
  "CO.,LTD.","CO.,LTD","CO., LTD.","CO., LTD","CO LTD","CO.,JSC",
  "CORPORATION","CORP.","CORP","INC.","INC","LLC","LTD.","LTD",
  "LIMITED","PTE.","PTE","S.A.S","S.A.","SAS","SA","GMBH","AG",
  "B.V.","BV","N.V.","NV","PLC","S.R.L.","SRL","PTY.","PTY",
  "J.S.C.","J.S.C","JSC","OJSC","CJSC","TNHH","CP","CTCP",
  "CONG TY","TNHH TM","THUONG MAI","GROUP","HOLDING","HOLDINGS",
  "ENTERPRISE","ENTERPRISES","TRADING","INTERNATIONAL","INTL",
  "INT'L","MFG","MANUFACTURING","PVT","PRIVATE","P",
].sort((a,b)=>b.length-a.length);
const VN_PFX = [
  "CONG TY TNHH THUONG MAI","CONG TY TNHH TM","CONG TY TNHH MTV",
  "CONG TY TNHH","CONG TY CO PHAN","CONG TY CP","CONG TY",
].sort((a,b)=>b.length-a.length);

// ═══════════════════════════════════════════════════════════
// ENTITY RESOLUTION
// ═══════════════════════════════════════════════════════════
const diac = s => s.normalize("NFD").replace(/[\u0300-\u036f]/g,"");
function resolveEntity(raw){
  const flags=[];
  if(!raw||typeof raw!=="string"||!raw.trim()) return {key:"",display:"",flags:["Empty Name"]};
  let s=diac(raw.trim()).toUpperCase().replace(/[.,;:'"!?()[\]{}]/g," ").replace(/\s+/g," ").trim();
  if(INVALID_NAMES.includes(s)) return {key:s,display:raw.trim(),flags:["Invalid Company Name"]};
  for(const p of VN_PFX){if(s.startsWith(p+" ")||s===p){s=s.slice(p.length).trim();break;}}
  let go=true,pass=0;while(go&&pass<3){go=false;pass++;for(const sf of SUFFIXES){if(s.endsWith(" "+sf)||s===sf){s=s.slice(0,s.length-sf.length).replace(/[\s,;.]+$/,"").trim();go=true;break;}}}
  // fused suffix
  for(const fs of [{r:/SA DE CV$/,l:8},{r:/SA$/,l:2},{r:/LTD$/,l:3},{r:/LLC$/,l:3},{r:/SRL$/,l:3}]){
    if(fs.r.test(s)&&s.length>fs.l+2){const c=s.slice(0,s.length-fs.l);if(c&&!/\s$/.test(c)){s=c;flags.push("Fused Suffix");break;}}
  }
  s=s.replace(/[.,;:'"!?()[\]{}]/g," ").replace(/\s+/g," ").trim();
  const tk=s.split(" "),dd=[];for(let i=0;i<tk.length;i++){if(i>0&&tk[i]===tk[i-1]){flags.push("Dup Token: "+tk[i]);continue;}dd.push(tk[i]);}
  s=dd.join(" ");
  if(/\s[PSA]$/.test(s))s=s.slice(0,-2).trim();
  if(s.length<=2&&raw.trim().length>5)flags.push("Needs Official Name Review");
  const key=s||diac(raw.trim()).toUpperCase();
  const display=key.split(" ").map(w=>w.length<=2?w:w[0]+w.slice(1).toLowerCase()).join(" ");
  return {key,display,flags};
}
const compactK = s=>s.replace(/\s+/g,"").toUpperCase();

// ═══════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════
function classifyNum(val){
  if(val==null||val==="")return{s:"Missing",n:NaN};
  const sv=String(val).trim();
  if(TEXT_PLACEHOLDERS.includes(sv.toUpperCase()))return{s:"Text Placeholder",n:NaN};
  const num=parseFloat(sv);
  if(isNaN(num))return{s:"Invalid Numeric",n:NaN};
  if(num===0)return{s:"Zero",n:0};
  if(num<0)return{s:"Negative",n:num};
  return{s:"Valid",n:num};
}
function calcIQR(v){
  if(v.length<4)return null;
  const s=[...v].sort((a,b)=>a-b),n=s.length;
  const q1=s[Math.floor(n*.25)],q3=s[Math.floor(n*.75)],iqr=q3-q1;
  return{q1,q3,iqr,lo:q1-1.5*iqr,hi:q3+1.5*iqr,n};
}
function detectCols(h){
  const u=h.map(c=>(c||"").toString().toUpperCase().trim());
  const f=ps=>{for(const p of ps){const i=u.findIndex(x=>x.includes(p));if(i>=0)return h[i];}return null;};
  return{
    productDesc:f(["PRODUCT DESCRIPTION","PRODUCT DESC","DESCRIPTION","COMMODITY"]),
    supplier:f(["SUPPLIER","EXPORTER","SHIPPER","SELLER"]),
    purchaser:f(["PURCHASER","BUYER","IMPORTER","CONSIGNEE"]),
    countryOrigin:f(["COUNTRY OF ORIGIN","ORIGIN COUNTRY","ORIGIN","EXPORT COUNTRY"]),
    purchasingCountry:f(["PURCHASING COUNTRY","IMPORT COUNTRY","DESTINATION","DEST COUNTRY"]),
    unitPrice:f(["UNIT PRICE","PRICE PER UNIT","PRICE/UNIT","PRICE"]),
    totalValue:f(["TOTAL VALUE","VALUE","TOTAL AMOUNT","AMOUNT","FOB VALUE"]),
    quantity:f(["QUANTITY","VOLUME","QTY","WEIGHT"]),
    unit:f(["UNIT","UOM","UNIT OF MEASURE"]),
  };
}

// ═══════════════════════════════════════════════════════════
// UI COMPONENTS
// ═══════════════════════════════════════════════════════════
function StageNav({stage,setStage,maxStage}){
  const st=[
    {id:0,l:"Upload",ic:"📂"},{id:1,l:"Columns",ic:"🔗"},{id:2,l:"Quality",ic:"🔍"},
    {id:3,l:"Standardize",ic:"🏢"},{id:4,l:"Top 80%",ic:"🎯"},{id:5,l:"Industry",ic:"🏭"},
    {id:6,l:"Keywords",ic:"🏷"},{id:7,l:"Segment",ic:"⚙"},{id:8,l:"IQR",ic:"📐"},
    {id:9,l:"Export",ic:"📊"},
  ];
  return(<div style={S.stageNav}>{st.map(s=>(
    <button key={s.id} onClick={()=>s.id<=maxStage&&setStage(s.id)}
      style={{...S.stageBtn,...(s.id===stage?S.stageBtnActive:{}),...(s.id>maxStage?S.stageBtnDisabled:{})}}>
      <span style={{fontSize:14}}>{s.ic}</span><span>{s.l}</span>
    </button>
  ))}</div>);
}
function Upload({onFileRead,label,accept,compact}){
  const ref=useRef(null);const[drag,setDrag]=useState(false);const[fn,setFn]=useState("");
  const proc=useCallback(f=>{if(!f)return;setFn(f.name);const r=new FileReader();r.onload=e=>{try{onFileRead(XLSX.read(new Uint8Array(e.target.result),{type:"array"}),f.name)}catch(er){console.error(er)}};r.readAsArrayBuffer(f)},[onFileRead]);
  const cl=useCallback(e=>{e.stopPropagation();ref.current?.click()},[]);
  const ch=useCallback(e=>{proc(e.target.files[0]);e.target.value=""},[proc]);
  if(compact)return(<span style={{display:"inline-flex",alignItems:"center",gap:6}}>
    <button style={S.sec} onClick={cl}>{label||"Upload"}</button>
    {fn&&<span style={{fontSize:10,color:C.textMuted}}>{fn}</span>}
    <input ref={ref} type="file" accept={accept||".xlsx,.xls,.csv"} style={{display:"none"}} onChange={ch}/>
  </span>);
  return(<div style={{...S.drop,...(drag?S.dropAct:{})}} onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);proc(e.dataTransfer.files[0])}} onClick={cl}>
    <input ref={ref} type="file" accept={accept||".xlsx,.xls,.csv"} style={{display:"none"}} onChange={ch}/>
    <div style={S.dropIc}>⬆</div><div style={S.dropLbl}>{label||"Drop workbook or click to browse"}</div>
    <div style={S.dropH}>Supports .xlsx, .xls, .csv</div>
    {fn&&<div style={{marginTop:5,fontSize:10,color:C.accent}}>{fn}</div>}
  </div>);
}
function DT({data,max=50,title}){
  if(!data?.length)return<div style={S.emptyMsg}>No data</div>;
  const cols=Object.keys(data[0]),show=data.slice(0,max);
  return(<div style={S.tw}>
    {title&&<div style={S.tt}>{title} <span style={S.rc}>({data.length} rows)</span></div>}
    <div style={S.ts}><table style={S.tbl}>
      <thead><tr>{cols.map(c=><th key={c} style={S.th} title={c}>{c.length>24?c.slice(0,21)+"…":c}</th>)}</tr></thead>
      <tbody>{show.map((r,i)=><tr key={i} style={i%2?S.trA:{}}>{cols.map(c=>{const v=r[c]==null?"":String(r[c]);return<td key={c} style={S.td} title={v}>{v.length>36?v.slice(0,33)+"…":v}</td>})}</tr>)}</tbody>
    </table></div>
    {data.length>max&&<div style={S.mr}>Showing {max} of {data.length}</div>}
  </div>);
}
function M({label,value,sub}){return(<div style={S.mc}><div style={S.mv}>{value}</div><div style={S.ml}>{label}</div>{sub&&<div style={S.ms}>{sub}</div>}</div>)}
function SChart({data,year}){
  if(!data?.length)return null;
  const gs=_.groupBy(data,"Final_Segmentation");
  const en=Object.entries(gs).map(([n,r])=>({n,c:r.length,v:r.reduce((s,x)=>s+(parseFloat(x["TotalValue_After_Conversion"])||0),0)})).sort((a,b)=>b.v-a.v);
  const mx=Math.max(...en.map(e=>e.v),1);
  return(<div style={S.cb}><div style={S.ct}>{year}</div>{en.map(e=>(<div key={e.n} style={S.br}>
    <div style={S.bl} title={e.n}>{e.n.length>18?e.n.slice(0,15)+"…":e.n}</div>
    <div style={S.bt}><div style={{...S.bf,width:Math.max(e.v/mx*100,2)+"%"}}/></div>
    <div style={S.bv}>${(e.v/1e6).toFixed(2)}M ({e.c})</div>
  </div>))}</div>);
}

// ═══════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════
export default function ChemSegTool(){
  const[stage,setStage]=useState(0);
  const[maxStage,setMaxStage]=useState(0);
  const[processing,setP]=useState(false);
  const[statusMsg,setSM]=useState("");
  const[wbName,setWbName]=useState("");
  const[yearSheets,setYS]=useState({});
  const[rawH,setRawH]=useState([]);
  const[colMap,setCM]=useState({});
  const[colWarn,setColWarn]=useState([]);
  const[qReport,setQR]=useState({});
  const[stdData,setSD]=useState({});
  const[purList,setPL]=useState([]);
  const[compLog,setCL]=useState([]);
  const[dupes,setDupes]=useState([]);
  // V15: Top 80% purchaser ranking
  const[purRanking,setPurRanking]=useState({});
  const[indMaster,setIM]=useState(null);
  const[indLog,setIL]=useState([]);
  const[segs,setSegs]=useState([]);
  const[newSeg,setNewSeg]=useState("");
  const[kwConf,setKwConf]=useState([]);
  const[segData,setSegD]=useState({});
  const[iqrSum,setIqrSum]=useState([]);
  const[bef,setBef]=useState({});
  const[aft,setAft]=useState({});
  const[nonConv,setNC]=useState([]);

  const sdRef=useRef(stdData);useEffect(()=>{sdRef.current=stdData},[stdData]);
  const cmRef=useRef(colMap);useEffect(()=>{cmRef.current=colMap},[colMap]);
  const sgRef=useRef(segs);useEffect(()=>{sgRef.current=segs},[segs]);
  const prRef=useRef(purRanking);useEffect(()=>{prRef.current=purRanking},[purRanking]);

  const go=useCallback(s=>{setStage(s);setMaxStage(m=>Math.max(m,s))},[]);

  // ── S0: Upload ──
  const onUpload=useCallback((wb,name)=>{
    setWbName(name);const sh={},ah=new Set();
    wb.SheetNames.forEach(sn=>{const m=sn.trim().match(/\d{4}/);if(m){const d=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});if(d.length){sh[m[0]]=d;Object.keys(d[0]).forEach(h=>ah.add(h))}}});
    if(!Object.keys(sh).length)wb.SheetNames.forEach(sn=>{const d=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});if(d.length){sh[sn]=d;Object.keys(d[0]).forEach(h=>ah.add(h))}});
    const w=[];const yrs=Object.keys(sh);
    if(yrs.length>1){const rc=new Set(Object.keys(sh[yrs[0]]?.[0]||{}));yrs.slice(1).forEach(y=>{const yc=new Set(Object.keys(sh[y]?.[0]||{}));[...rc].filter(c=>!yc.has(c)).forEach(c=>w.push(y+" missing: "+c));[...yc].filter(c=>!rc.has(c)).forEach(c=>w.push(y+" extra: "+c))})}
    setColWarn(w);setYS(sh);const hd=[...ah];setRawH(hd);setCM(detectCols(hd));go(1);
  },[go]);

  // ── S2: Quality ──
  const runQuality=useCallback(()=>{
    setP(true);setSM("Checking quality...");
    setTimeout(()=>{
      const cm=cmRef.current,rp={};
      Object.entries(yearSheets).forEach(([yr,rows])=>{
        const n=rows.length,y={totalRows:n,fields:{}};
        ["unitPrice","totalValue","quantity","supplier","purchaser","countryOrigin","purchasingCountry"].forEach(f=>{
          const col=cm[f],isN=["unitPrice","totalValue","quantity"].includes(f);
          if(!col){y.fields[f]=isN?{missing:n,zero:0,negative:0,invalidNum:0,textPlaceholder:0,backfilled:0,unrecoverable:0,validPositive:0,totalIssues:n,issueRate:"100.0%",type:"numeric"}:{blank:n,none:0,other:0,totalIssues:n,issueRate:"100.0%",type:"text"};return}
          if(isN){let mi=0,z=0,ne=0,iv=0,tp=0,vp=0;rows.forEach(r=>{const c=classifyNum(r[col]);switch(c.s){case"Missing":mi++;break;case"Zero":z++;break;case"Negative":ne++;break;case"Invalid Numeric":iv++;break;case"Text Placeholder":tp++;break;default:vp++}});const iss=mi+z+ne+iv+tp;y.fields[f]={missing:mi,zero:z,negative:ne,invalidNum:iv,textPlaceholder:tp,backfilled:0,unrecoverable:0,validPositive:vp,totalIssues:iss,issueRate:((iss/n)*100).toFixed(1)+"%",type:"numeric"}}
          else{let bl=0,no=0,ot=0;rows.forEach(r=>{const v=r[col];if(v==null||v==="")bl++;else if(String(v).toUpperCase()==="NONE")no++;else if(String(v).toUpperCase()==="OTHER")ot++});const iss=bl+no+ot;y.fields[f]={blank:bl,none:no,other:ot,totalIssues:iss,issueRate:((iss/n)*100).toFixed(1)+"%",type:"text"}}
        });rp[yr]=y});
      setQR(rp);setP(false);setSM("");go(2);
    },20);
  },[yearSheets,go]);

  // ── S3: Standardization + Backfill ──
  const runStd=useCallback(()=>{
    setP(true);setSM("Entity resolution & backfill...");
    setTimeout(()=>{
      const cm=cmRef.current,ns={},ap=[],lg=[],ckMap=new Map(),bfc={};
      Object.entries(yearSheets).forEach(([yr,rows])=>{
        const bc={unitPrice:{b:0,u:0},totalValue:{b:0,u:0},quantity:{b:0,u:0}};
        ns[yr]=rows.map((row,idx)=>{
          const nr={...row};
          // Entity resolution
          const rs=resolveEntity(String(row[cm.supplier]||""));nr["Supplier_Standardize"]=rs.display;nr["Supplier_MatchingKey"]=rs.key;
          if(rs.flags.length)lg.push({Year:yr,Row:idx+1,Field:"Supplier",Raw:row[cm.supplier],Std:rs.display,Key:rs.key,Flags:rs.flags.join("; ")});
          const rp=resolveEntity(String(row[cm.purchaser]||""));nr["Purchaser_Standardize"]=rp.display;nr["Purchaser_MatchingKey"]=rp.key;
          if(rp.flags.length)lg.push({Year:yr,Row:idx+1,Field:"Purchaser",Raw:row[cm.purchaser],Std:rp.display,Key:rp.key,Flags:rp.flags.join("; ")});
          const ctry=String(row[cm.purchasingCountry]||"").toUpperCase().trim();
          const ck=compactK(rp.key)+"||"+ctry;if(!ckMap.has(ck))ckMap.set(ck,[]);ckMap.get(ck).push({raw:row[cm.purchaser],std:rp.display,key:rp.key,yr});

          // Numeric classify + backfill
          const pC=classifyNum(row[cm.unitPrice]),tC=classifyNum(row[cm.totalValue]),qC=classifyNum(row[cm.quantity]);
          const pV=pC.s==="Valid",tV=tC.s==="Valid",qV=qC.s==="Valid",vc=(pV?1:0)+(tV?1:0)+(qV?1:0);
          let upCl=pV?pC.n:null,tvCl=tV?tC.n:null,qtCl=qV?qC.n:null;
          let upSrc=pV?"Raw":pC.s,tvSrc=tV?"Raw":tC.s,qtSrc=qV?"Raw":qC.s,bfSt="No Backfill Needed";
          if(vc===2){
            if(!pV&&tV&&qV&&qtCl>0){upCl=tvCl/qtCl;upSrc="Calculated";bfSt="Backfilled Unit Price";bc.unitPrice.b++}
            else if(!tV&&pV&&qV){tvCl=upCl*qtCl;tvSrc="Calculated";bfSt="Backfilled Total Value";bc.totalValue.b++}
            else if(!qV&&pV&&tV&&upCl>0){qtCl=tvCl/upCl;qtSrc="Calculated";bfSt="Backfilled Quantity";bc.quantity.b++}
            else{bfSt="Insufficient Numeric Data";if(!pV)bc.unitPrice.u++;if(!tV)bc.totalValue.u++;if(!qV)bc.quantity.u++}
          }else if(vc<2){bfSt="Insufficient Numeric Data";if(!pV)bc.unitPrice.u++;if(!tV)bc.totalValue.u++;if(!qV)bc.quantity.u++}
          nr["UnitPrice_Clean"]=upCl!=null?upCl:"";nr["TotalValue_Clean"]=tvCl!=null?tvCl:"";nr["Quantity_Clean"]=qtCl!=null?qtCl:"";
          nr["UnitPrice_Source"]=upSrc;nr["TotalValue_Source"]=tvSrc;nr["Quantity_Source"]=qtSrc;nr["Numeric_Backfill_Status"]=bfSt;
          const allCl=upCl!=null&&tvCl!=null&&qtCl!=null;
          nr["Numeric_Quality_Flag"]=allCl?"OK":"Issue";
          // Pre-conversion value check
          if(upCl!=null&&qtCl!=null&&qtCl>0){
            const exp=upCl*qtCl;nr["Expected_TotalValue_PreConv"]=exp;
            if(tvCl!=null){const d=Math.abs(tvCl-exp);nr["PreConv_Value_Diff_USD"]=d;nr["Pre_Conversion_Value_Check"]=d>500?"Mismatch":"Matched"}
            else{nr["PreConv_Value_Diff_USD"]="";nr["Pre_Conversion_Value_Check"]="Not Calculable"}
          }else{nr["Expected_TotalValue_PreConv"]="";nr["PreConv_Value_Diff_USD"]="";nr["Pre_Conversion_Value_Check"]="Not Calculable"}

          // V15: Row_Ranking_Value (raw-basis for Top 80%)
          const rawTV=classifyNum(row[cm.totalValue]),rawUP=classifyNum(row[cm.unitPrice]),rawQT=classifyNum(row[cm.quantity]);
          if(rawTV.s==="Valid"&&rawTV.n>0){nr["Row_Ranking_Value"]=rawTV.n;nr["Ranking_Value_Source"]="Raw_TotalValue"}
          else if(rawUP.s==="Valid"&&rawUP.n>0&&rawQT.s==="Valid"&&rawQT.n>0){nr["Row_Ranking_Value"]=rawUP.n*rawQT.n;nr["Ranking_Value_Source"]="Raw_UnitPrice × Raw_Quantity"}
          else{nr["Row_Ranking_Value"]=null;nr["Ranking_Value_Source"]="Not Calculable";nr["Ranking_Value_Issue_Flag"]="Ranking_Value_Not_Calculable"}
          return nr;
        });
        bfc[yr]=bc;
        const pm=new Map();ns[yr].forEach(r=>{const k=r["Purchaser_MatchingKey"]+"||"+String(r[cmRef.current.purchasingCountry]||"").toUpperCase().trim();if(!pm.has(k))pm.set(k,{Raw_Purchaser_Name:r[cm.purchaser]||"",Purchaser_Standardize:r["Purchaser_Standardize"],Purchasing_Country:r[cm.purchasingCountry]||"",Year:yr})});pm.forEach(v=>ap.push(v));
      });
      const dp=[];ckMap.forEach((en,ck)=>{const uq=_.uniqBy(en,"key");if(uq.length>1)uq.forEach(e=>dp.push({CompactKey:ck.split("||")[0],Country:ck.split("||")[1],Raw:e.raw,Std:e.std,Key:e.key,Year:e.yr,Note:"Space variant"}))});
      setQR(prev=>{const u={...prev};Object.entries(bfc).forEach(([yr,bc])=>{if(!u[yr])return;const y={...u[yr],fields:{...u[yr].fields}};["unitPrice","totalValue","quantity"].forEach(f=>{if(y.fields[f]?.type==="numeric")y.fields[f]={...y.fields[f],backfilled:bc[f].b,unrecoverable:bc[f].u}});u[yr]=y});return u});
      setSD(ns);setPL(ap);setCL(lg);setDupes(dp);setP(false);setSM("");go(3);
    },20);
  },[yearSheets,go]);

  const exportPurList=useCallback(()=>{const wb=XLSX.utils.book_new();[...new Set(purList.map(p=>p.Year))].forEach(y=>{XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(purList.filter(p=>p.Year===y)),"Pur_"+y)});XLSX.writeFile(wb,"Purchaser_List.xlsx")},[purList]);

  // ── S4: Top 80% Purchaser Value-Scope Selection (V15) ──
  const runTop80=useCallback(()=>{
    setP(true);setSM("Calculating Top 80% purchaser scope...");
    setTimeout(()=>{
      const cm=cmRef.current,cur=sdRef.current,ranking={},updated={};
      Object.entries(cur).forEach(([yr,rows])=>{
        const grp={};
        rows.forEach(r=>{
          const k=(r["Purchaser_MatchingKey"]||r["Purchaser_Standardize"]||"").toUpperCase()+"||"+String(r[cm.purchasingCountry]||"").toUpperCase().trim();
          if(!grp[k])grp[k]={key:k,std:r["Purchaser_Standardize"],country:r[cm.purchasingCountry]||"",raw:r[cm.purchaser]||"",total:0,rows:0,calculable:0};
          grp[k].rows++;
          const rv=r["Row_Ranking_Value"];
          if(rv!=null&&!isNaN(rv)){grp[k].total+=rv;grp[k].calculable++}
        });
        const pArr=Object.values(grp).filter(p=>p.calculable>0).sort((a,b)=>b.total-a.total);
        const yearTotal=pArr.reduce((s,p)=>s+p.total,0);
        let cum=0;
        const ranked=pArr.map((p,i)=>{
          const share=yearTotal>0?p.total/yearTotal:0;
          const prevCum=cum;
          cum+=share;
          // Include this purchaser if cumulative BEFORE adding it was still < 80%
          const inScope=prevCum<0.8;
          return{...p,rank:i+1,share,cumShare:cum,scope:inScope?"Top_80_Value_Scope":"Long_Tail_20_Value_Scope"};
        });
        const ncPur=Object.values(grp).filter(p=>p.calculable===0).map(p=>({...p,rank:null,share:0,cumShare:null,scope:"Long_Tail_20_Value_Scope"}));
        ranking[yr]=[...ranked,...ncPur];
        // Create NEW array with scope annotations (don't mutate sdRef)
        const scopeMap=new Map();[...ranked,...ncPur].forEach(p=>scopeMap.set(p.key,p.scope));
        updated[yr]=rows.map(r=>{
          const k=(r["Purchaser_MatchingKey"]||r["Purchaser_Standardize"]||"").toUpperCase()+"||"+String(r[cm.purchasingCountry]||"").toUpperCase().trim();
          return{...r,Research_Scope_Flag:scopeMap.get(k)||"Long_Tail_20_Value_Scope"};
        });
      });
      setSD(updated);setPurRanking(ranking);setP(false);setSM("");go(4);
    },20);
  },[go]);

  const exportTop80=useCallback(()=>{
    const wb=XLSX.utils.book_new();
    Object.entries(purRanking).forEach(([yr,list])=>{
      const top=list.filter(p=>p.scope==="Top_80_Value_Scope").map(p=>({
        Year:yr,Raw_Purchaser:p.raw,Purchaser_Standardize:p.std,Purchasing_Country:p.country,
        Ranking_Value:p.total.toFixed(2),Year_Total:list.reduce((s,x)=>s+x.total,0).toFixed(2),
        Value_Share:(p.share*100).toFixed(2)+"%",Cum_Share:(p.cumShare*100).toFixed(2)+"%",
        Rank:p.rank,Research_Scope:p.scope
      }));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(top),"Top80_"+yr);
    });
    XLSX.writeFile(wb,"Top80_Purchaser_Research_List.xlsx");
  },[purRanking]);

  // ── S5: Industry Master ──
  const onIndustry=useCallback(wb=>{
    const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});setIM(data);
    const cm=cmRef.current,cur=sdRef.current,ml=[],ns={};
    Object.entries(cur).forEach(([yr,rows])=>{
      ns[yr]=rows.map(r=>{
        const nr={...r};
        const pk=(r["Purchaser_MatchingKey"]||r["Purchaser_Standardize"]||"").toUpperCase().trim();
        const pc=String(r[cm.purchasingCountry]||"").toUpperCase().trim();
        if(r["Research_Scope_Flag"]==="Top_80_Value_Scope"){
          const match=data.find(m=>{const ms=(m["Purchaser_Standardize"]||"").toUpperCase().trim();const mc=(m["Purchasing Country"]||m["Purchasing_Country"]||"").toUpperCase().trim();return(ms===pk||diac(ms)===pk)&&mc===pc});
          if(match){nr["Industry"]=match["Industry"]||"";nr["Industry Segment"]=match["Industry Segment"]||match["Industry_Segment"]||"";nr["Industry_Match_Status"]="Matched"}
          else{nr["Industry"]="";nr["Industry Segment"]="";nr["Industry_Match_Status"]="Not Matched"}
        }else{nr["Industry"]="";nr["Industry Segment"]="";nr["Industry_Match_Status"]="Long Tail - No Match Required"}
        ml.push({Year:yr,Purchaser:pk,Country:pc,Scope:r["Research_Scope_Flag"]||"",Status:nr["Industry_Match_Status"],Industry:nr["Industry"],Segment:nr["Industry Segment"]});
        return nr;
      });
    });
    setSD(ns);setIL(ml);go(5);
  },[go]);

  // ── S6: Keywords ──
  const addSeg=useCallback(()=>{const n=newSeg.trim();if(!n)return;setSegs(p=>[...p,{id:Date.now(),name:n,keywords:[]}]);setNewSeg("")},[newSeg]);
  const addKw=useCallback(sid=>{setSegs(p=>p.map(s=>s.id===sid?{...s,keywords:[...s.keywords,{keyword:"",note:""}]}:s))},[]);
  const updKw=useCallback((sid,ki,f,v)=>{setSegs(p=>p.map(s=>s.id===sid?{...s,keywords:s.keywords.map((k,i)=>i===ki?{...k,[f]:v}:k)}:s))},[]);
  const rmKw=useCallback((sid,ki)=>{setSegs(p=>p.map(s=>s.id===sid?{...s,keywords:s.keywords.filter((_,i)=>i!==ki)}:s))},[]);
  const rmSeg=useCallback(sid=>{setSegs(p=>p.filter(s=>s.id!==sid))},[]);
  const renSeg=useCallback((sid,n)=>{setSegs(p=>p.map(s=>s.id===sid?{...s,name:n}:s))},[]);
  const onKwImport=useCallback(wb=>{
    const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});if(!data.length)return;
    const cols=Object.keys(data[0]).filter(h=>{const u=h.toUpperCase().trim();return u!=="STT"&&u!=="NO"&&u!=="NO."&&u!=="#"&&u!=="INDEX"});
    setSegs(prev=>{const ns=prev.map(s=>({...s,keywords:[...s.keywords]}));cols.forEach(col=>{const sn=col.trim();if(!sn)return;let sg=ns.find(s=>s.name.toUpperCase()===sn.toUpperCase());if(!sg){sg={id:Date.now()+Math.random()*99999,name:sn,keywords:[]};ns.push(sg)}const ex=new Set(sg.keywords.map(k=>k.keyword.toUpperCase()));data.forEach(r=>{const kw=String(r[col]||"").trim();if(kw&&!ex.has(kw.toUpperCase())){sg.keywords.push({keyword:kw,note:""});ex.add(kw.toUpperCase())}})});return ns});
  },[]);
  const compConf=useCallback(()=>{const m=new Map();segs.forEach(s=>s.keywords.forEach(k=>{if(!k.keyword)return;const u=k.keyword.toUpperCase();if(!m.has(u))m.set(u,[]);m.get(u).push(s.name)}));const c=[];m.forEach((ss,kw)=>{if(ss.length>1)c.push({Keyword:kw,Segments:ss.join(", "),Count:ss.length})});setKwConf(c);return c},[segs]);

  // ── S7: Segmentation (V15 logic: KW first; Industry Segment fallback ONLY for Top 80%) ──
  const runSeg=useCallback(()=>{
    compConf();setP(true);setSM("Classifying rows...");
    setTimeout(()=>{
      const cur=sdRef.current,cm=cmRef.current,sg=sgRef.current,res={};
      Object.entries(cur).forEach(([yr,rows])=>{
        res[yr]=rows.map(r=>{
          const nr={...r},desc=String(r[cm.productDesc]||"").toUpperCase();
          const scope=r["Research_Scope_Flag"]||"Long_Tail_20_Value_Scope";
          let matched=null,matchedKw=null,allM=[];
          for(const s of sg){for(const k of s.keywords){if(k.keyword&&desc.includes(k.keyword.toUpperCase())){allM.push({seg:s.name,kw:k.keyword});if(!matched){matched=s.name;matchedKw=k.keyword}}}}
          nr["Matched_Segment"]=matched||"";nr["Matched_Keyword"]=matchedKw||"";nr["Keyword_Match_Status"]=matched?"Keyword Match":"No Match";
          const uSegs=[...new Set(allM.map(m=>m.seg))];
          nr["Keyword_Conflict"]=uSegs.length>1?"Multi: "+uSegs.join(", "):"";

          if(matched){
            nr["Final_Segmentation"]=matched;nr["Final_Segmentation_Source"]="Keyword";
            if(scope==="Top_80_Value_Scope"&&r["Industry Segment"]&&r["Industry Segment"]!==matched){
              nr["Segmentation_Note"]="KW: "+matchedKw+" [Industry Conflict: "+r["Industry Segment"]+"]";
            }else{nr["Segmentation_Note"]="KW: "+matchedKw}
          }else if(scope==="Top_80_Value_Scope"&&r["Industry Segment"]){
            nr["Final_Segmentation"]=r["Industry Segment"];nr["Final_Segmentation_Source"]="Industry Segment";nr["Segmentation_Note"]="Fallback (Top 80%)";
          }else{
            nr["Final_Segmentation"]="Unclassified";nr["Final_Segmentation_Source"]="None";
            nr["Segmentation_Note"]=scope==="Long_Tail_20_Value_Scope"?"Long Tail - KW only":"No match";
          }
          return nr;
        });
      });
      setSegD(res);setP(false);setSM("");go(7);
    },20);
  },[go,compConf]);

  // ── S8: Unit Conversion & IQR ──
  const runIQR=useCallback(()=>{
    setP(true);setSM("Converting & IQR...");
    setTimeout(()=>{
      const cm=cmRef.current,bf={},af={},iq=[],nc=[];
      Object.entries(segData).forEach(([yr,rows])=>{
        const proc=rows.map(r=>{
          const nr={...r},u=String(r[cm.unit]||"").toUpperCase().trim();
          nr["Original_Unit"]=u;
          if(AMBIGUOUS_UNITS.includes(u)){nr["Unit_Conversion_Status"]="Ambiguous (flagged)";nr["Quantity_MT"]="";nr["UnitPrice_per_MT"]="";nr["TotalValue_After_Conversion"]="";nr["Price_Eligible"]="No";nc.push({Year:yr,Unit:u,Note:"Ambiguous - confirm metric ton"});return nr}
          const fac=UNIT_CONVERSIONS[u];
          if(fac===undefined){nr["Unit_Conversion_Status"]="Non-Convertible";nr["Quantity_MT"]="";nr["UnitPrice_per_MT"]="";nr["TotalValue_After_Conversion"]="";nr["Price_Eligible"]="No";if(u)nc.push({Year:yr,Unit:u});return nr}
          nr["Unit_Conversion_Status"]="Converted";nr["MT_Factor"]=fac;
          if(r["Numeric_Quality_Flag"]==="Issue"){nr["Quantity_MT"]="";nr["UnitPrice_per_MT"]="Excluded";nr["TotalValue_After_Conversion"]="";nr["Price_Eligible"]="No";return nr}
          const tv=parseFloat(r["TotalValue_Clean"]),qt=parseFloat(r["Quantity_Clean"]),up=parseFloat(r["UnitPrice_Clean"]);
          nr["Quantity_MT"]=(!isNaN(qt)&&qt>0)?qt*fac:"";
          const qmt=parseFloat(nr["Quantity_MT"]);
          if(!isNaN(tv)&&tv>0&&!isNaN(qmt)&&qmt>0){nr["UnitPrice_per_MT"]=tv/qmt;nr["UnitPrice_per_MT_Src"]="TV/QMT"}
          else if(!isNaN(up)&&up>0&&fac>0){nr["UnitPrice_per_MT"]=up/fac;nr["UnitPrice_per_MT_Src"]="UP_Fallback"}
          else{nr["UnitPrice_per_MT"]="Invalid";nr["UnitPrice_per_MT_Src"]="N/A"}
          const pm=parseFloat(nr["UnitPrice_per_MT"]);
          nr["TotalValue_After_Conversion"]=(!isNaN(pm)&&!isNaN(qmt)&&qmt>0)?pm*qmt:"";
          nr["Price_Eligible"]=(!isNaN(pm)&&pm>0)?"Yes":"No";
          return nr;
        });
        // IQR
        const gs=_.groupBy(proc,"Final_Segmentation");
        Object.entries(gs).forEach(([seg,sr])=>{
          const ps=sr.filter(r=>r["Price_Eligible"]==="Yes").map(r=>parseFloat(r["UnitPrice_per_MT"])).filter(v=>!isNaN(v)&&v>0);
          const iqr=calcIQR(ps);
          if(!iqr){sr.forEach(r=>{r["IQR_Status"]="Insufficient Data";r["IQR_Note"]="<4 eligible"});return}
          const oc=ps.filter(v=>v<iqr.lo||v>iqr.hi).length;
          const wide=iqr.q1>0&&(iqr.hi-iqr.lo)/iqr.q1>10;
          iq.push({Year:yr,Segment:seg,Q1:iqr.q1.toFixed(2),Q3:iqr.q3.toFixed(2),IQR:iqr.iqr.toFixed(2),Lower:iqr.lo.toFixed(2),Upper:iqr.hi.toFixed(2),Sample:iqr.n,Outliers:oc,Rate:((oc/iqr.n)*100).toFixed(1)+"%",Warning:wide?"IQR_Range_Too_Wide":""});
          sr.forEach(r=>{const p=parseFloat(r["UnitPrice_per_MT"]);r["IQR_Q1"]=iqr.q1;r["IQR_Q3"]=iqr.q3;r["IQR_Lower"]=iqr.lo;r["IQR_Upper"]=iqr.hi;
            if(!isNaN(p)&&p>0){r["IQR_Status"]=(p<iqr.lo||p>iqr.hi)?"Outlier":"Normal";r["IQR_Note"]=r["IQR_Status"]==="Outlier"?p.toFixed(2)+" outside ["+iqr.lo.toFixed(2)+","+iqr.hi.toFixed(2)+"]":""}
            else{r["IQR_Status"]="Invalid Price";r["IQR_Note"]=""}
          });
        });
        bf[yr]=proc;af[yr]=proc.filter(r=>!(r["Pre_Conversion_Value_Check"]==="Mismatch"&&r["IQR_Status"]==="Outlier"));
      });
      setBef(bf);setAft(af);setIqrSum(iq);setNC(nc);setP(false);setSM("");go(8);
    },30);
  },[segData,go]);

  // ── S9: Export ──
  const exportAll=useCallback(()=>{
    const wb=XLSX.utils.book_new();
    // 00 Summary
    const sum=Object.entries(aft).flatMap(([y,rows])=>{const gs=_.groupBy(rows,"Final_Segmentation");return Object.entries(gs).map(([seg,sr])=>{const el=sr.filter(r=>r["Price_Eligible"]==="Yes");const tv=el.reduce((s,r)=>s+(parseFloat(r["TotalValue_After_Conversion"])||0),0);const qm=el.reduce((s,r)=>s+(parseFloat(r["Quantity_MT"])||0),0);return{Year:y,Segment:seg,Total_Rows:sr.length,Eligible:el.length,Total_Value:tv.toFixed(2),Total_MT:qm.toFixed(4),Avg_Price_MT:qm>0?(tv/qm).toFixed(2):"N/A"}})});
    if(sum.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sum),"00_Summary");
    // 01 Quality
    const dq=[];Object.entries(qReport).forEach(([y,yr])=>Object.entries(yr.fields).forEach(([f,s])=>{const b={Year:y,Field:f,Rows:yr.totalRows,Issues:s.totalIssues,Issue_Rate:s.issueRate};if(s.type==="numeric")dq.push({...b,Valid:s.validPositive,Missing:s.missing,Zero:s.zero,Negative:s.negative,Invalid:s.invalidNum,Text:s.textPlaceholder,Backfilled:s.backfilled||0,Unrecoverable:s.unrecoverable||0});else dq.push({...b,Blank:s.blank,None:s.none||0,Other:s.other||0})}));
    if(dq.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(dq),"01_Quality");
    if(purList.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(purList),"02_Purchaser_List");
    // 03 Top 80% Ranking
    const rkAll=Object.entries(purRanking).flatMap(([yr,list])=>list.map(p=>({Year:yr,Purchaser:p.std,Country:p.country,Value:p.total.toFixed(2),Share:(p.share*100).toFixed(2)+"%",Cum_Share:p.cumShare!=null?(p.cumShare*100).toFixed(2)+"%":"",Rank:p.rank||"",Scope:p.scope})));
    if(rkAll.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rkAll),"03_Top80_Ranking");
    if(indLog.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(_.uniqBy(indLog,r=>r.Purchaser+r.Country+r.Year)),"04_Industry_Match");
    if(compLog.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(compLog),"05_Company_Log");
    if(dupes.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(dupes),"05b_Duplicates");
    if(nonConv.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(_.uniqBy(nonConv,r=>r.Year+r.Unit)),"06_Non_Convertible");
    if(iqrSum.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(iqrSum),"07_IQR_Summary");
    if(kwConf.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(kwConf),"08_KW_Conflicts");
    Object.entries(bef).forEach(([y,r])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r),(y+"_Before").slice(0,31)));
    Object.entries(aft).forEach(([y,r])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r),(y+"_After").slice(0,31)));
    XLSX.writeFile(wb,"ChemSeg_Output.xlsx");
  },[aft,bef,qReport,purList,purRanking,indLog,compLog,dupes,nonConv,iqrSum,kwConf]);

  // ═══════════════════════════════════════════════════════════
  // RENDER
  // ═══════════════════════════════════════════════════════════
  const R=()=>{switch(stage){
    case 0:return(<div style={S.sc}>
      <div style={S.hero}><h2 style={S.heroT}>Chemical Segmentation Tool</h2><p style={S.heroS}>Upload multi-year trade workbook</p></div>
      <Upload onFileRead={onUpload}/>{wbName&&<div style={S.chip}>{wbName}</div>}
    </div>);
    case 1:return(<div style={S.sc}>
      <div style={S.mr2}>{Object.entries(yearSheets).map(([y,r])=><M key={y} label={"Year "+y} value={r.length} sub="rows"/>)}<M label="Columns" value={rawH.length}/></div>
      {colWarn.length>0&&<div style={S.warn}>{colWarn.map((w,i)=><div key={i}>{w}</div>)}</div>}
      <h3 style={S.secT}>Column Mapping</h3><p style={S.hint}>Auto-detected. Adjust if needed.</p>
      <div style={S.mapG}>{["productDesc:Product Description","supplier:Supplier","purchaser:Purchaser","countryOrigin:Country of Origin","purchasingCountry:Purchasing Country","unitPrice:Unit Price","totalValue:Total Value","quantity:Quantity","unit:Unit"].map(s=>{const[k,l]=s.split(":");return(<div key={k} style={S.mapR}><label style={S.mapL}>{l}</label><select value={colMap[k]||""} style={S.sel} onChange={e=>setCM(p=>({...p,[k]:e.target.value}))}><option value="">--</option>{rawH.map(h=><option key={h} value={h}>{h}</option>)}</select>{colMap[k]&&<span style={{color:C.success,marginLeft:4}}>✓</span>}</div>)})}</div>
      <button style={S.pri} onClick={runQuality}>Run Quality Check</button>
    </div>);
    case 2:return(<div style={S.sc}>
      <h3 style={S.secT}>Data Quality</h3>
      {Object.entries(qReport).map(([y,yr])=>{const n=yr.totalRows;const pct=v=>n>0?((v/n)*100).toFixed(1)+"%":"0%";return(<div key={y} style={S.yb}>
        <div style={S.yl}>Year {y} — {n} rows</div>
        <div style={S.qg}>{Object.entries(yr.fields).map(([f,s])=>{const ic=parseFloat(s.issueRate)>20?C.danger:parseFloat(s.issueRate)>5?C.warning:C.success;return(<div key={f} style={{...S.qc,borderLeft:"3px solid "+ic}}>
          <div style={S.qf}>{f}</div><div style={S.qr}>{s.totalIssues} issues ({s.issueRate})</div>
          {s.type==="numeric"?(<div style={S.qd}>
            <span style={s.validPositive===n?{color:C.success}:{}}>Valid:{s.validPositive}({pct(s.validPositive)})</span>
            {s.missing>0&&<span> Blank:{s.missing}({pct(s.missing)})</span>}
            {s.zero>0&&<span> Zero:{s.zero}</span>}{s.negative>0&&<span style={{color:C.danger}}> Neg:{s.negative}</span>}
            {s.invalidNum>0&&<span style={{color:C.danger}}> Inv:{s.invalidNum}</span>}
            {s.textPlaceholder>0&&<span style={{color:C.warning}}> Txt:{s.textPlaceholder}</span>}
            {(s.backfilled||0)>0&&<span style={{color:C.accent}}> BF:{s.backfilled}</span>}
            {(s.unrecoverable||0)>0&&<span style={{color:C.danger}}> Unrec:{s.unrecoverable}</span>}
          </div>):(<div style={S.qd}>{s.blank>0&&<span>Blank:{s.blank}({pct(s.blank)})</span>}{s.none>0&&<span> None:{s.none}</span>}{s.other>0&&<span> Other:{s.other}</span>}{s.totalIssues===0&&<span style={{color:C.success}}>All valid</span>}</div>)}
        </div>)})}</div>
      </div>)})}
      <button style={S.pri} onClick={runStd}>Run Standardization</button>
    </div>);
    case 3:{
      const tbf=Object.values(stdData).reduce((s,r)=>s+r.filter(x=>(x["Numeric_Backfill_Status"]||"").startsWith("Backfilled")).length,0);
      const tins=Object.values(stdData).reduce((s,r)=>s+r.filter(x=>x["Numeric_Backfill_Status"]==="Insufficient Numeric Data").length,0);
      const tmm=Object.values(stdData).reduce((s,r)=>s+r.filter(x=>x["Pre_Conversion_Value_Check"]==="Mismatch").length,0);
      return(<div style={S.sc}>
        <h3 style={S.secT}>Entity Resolution & Backfill</h3>
        <div style={S.mr2}><M label="Purchasers" value={purList.length}/><M label="Flagged" value={compLog.length} sub="review"/><M label="Dupes" value={dupes.length} sub="space variants"/><M label="Backfilled" value={tbf}/><M label="Insufficient" value={tins}/><M label="Mismatch" value={tmm} sub=">$500"/></div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><button style={S.sec} onClick={exportPurList}>Download All Purchasers</button></div>
        <DT data={purList.slice(0,50)} title="Purchaser List (preview)"/>
        {compLog.length>0&&<DT data={compLog.slice(0,30)} title="Entity Resolution Log"/>}
        <button style={S.pri} onClick={runTop80}>Calculate Top 80% Scope</button>
      </div>);
    }
    case 4:{
      const yrs=Object.keys(purRanking);
      return(<div style={S.sc}>
        <h3 style={S.secT}>Top 80% Purchaser Value Scope</h3>
        <p style={S.hint}>Only Top 80% purchasers need Industry Master research. Long-tail 20% segmented by keywords only.</p>
        {yrs.map(yr=>{const list=purRanking[yr]||[];const top=list.filter(p=>p.scope==="Top_80_Value_Scope");const tail=list.filter(p=>p.scope==="Long_Tail_20_Value_Scope");
          return(<div key={yr} style={S.yb}>
            <div style={S.yl}>Year {yr}</div>
            <div style={S.mr2}><M label="Top 80%" value={top.length} sub="research scope"/><M label="Long Tail 20%" value={tail.length} sub="KW only"/><M label="Total" value={list.length}/></div>
            <DT data={top.slice(0,30).map(p=>({Rank:p.rank,Purchaser:p.std,Country:p.country,Value:"$"+(p.total/1e6).toFixed(2)+"M",Share:(p.share*100).toFixed(1)+"%",Cum:(p.cumShare*100).toFixed(1)+"%"}))} title={"Top 80% Purchasers — "+yr}/>
          </div>);
        })}
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button style={S.pri} onClick={exportTop80}>Download Top 80% Research List</button>
        </div>
        <div style={S.dv}/>
        <h3 style={S.secT}>Upload Industry Master</h3>
        <p style={S.hint}>4 columns: Purchaser_Standardize, Purchasing Country, Industry, Industry Segment</p>
        <Upload onFileRead={onIndustry} label="Drop Industry Master"/>
        <button style={S.link} onClick={()=>go(6)}>Skip — Go to Keywords</button>
      </div>);
    }
    case 5:return(<div style={S.sc}>
      <h3 style={S.secT}>Industry Master Matched</h3>
      <div style={S.mr2}><M label="Records" value={indMaster?.length||0}/><M label="Matched" value={indLog.filter(l=>l.Status==="Matched").length}/><M label="Not Matched" value={indLog.filter(l=>l.Status==="Not Matched").length}/><M label="Long Tail Skip" value={indLog.filter(l=>l.Status.includes("Long Tail")).length}/></div>
      <DT data={_.uniqBy(indLog,r=>r.Purchaser+r.Country).slice(0,40)} title="Match Log"/>
      <button style={S.pri} onClick={()=>go(6)}>Proceed to Keywords</button>
    </div>);
    case 6:return(<div style={S.sc}>
      <h3 style={S.secT}>Segment Keywords</h3>
      <p style={S.hint}>STT | Segment1 | Segment2 | ... (columnar import)</p>
      <div style={S.addR}>
        <input style={S.addIn} placeholder="New segment name..." value={newSeg} onChange={e=>setNewSeg(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addSeg()}/>
        <button style={S.pri} onClick={addSeg} disabled={!newSeg.trim()}>+ Add</button>
        <Upload onFileRead={onKwImport} label="Import Keywords" accept=".xlsx,.csv" compact/>
      </div>
      {kwConf.length>0&&<div style={S.warn}><strong>Conflicts:</strong>{kwConf.map((c,i)=><div key={i}>{c.Keyword} → {c.Segments}</div>)}</div>}
      {!segs.length&&<div style={S.empty}>No segments yet</div>}
      {segs.map(sg=>(<div key={sg.id} style={S.sgC}><div style={S.sgH}>
        <input style={S.sgN} value={sg.name} onChange={e=>renSeg(sg.id,e.target.value)}/><span style={S.sgK}>{sg.keywords.length}kw</span><button style={S.danSm} onClick={()=>rmSeg(sg.id)}>✕</button>
      </div><div style={S.kwL}>{sg.keywords.map((k,ki)=>(<div key={ki} style={S.kwR}>
        <input style={S.kwI} value={k.keyword} placeholder="Keyword..." onChange={e=>updKw(sg.id,ki,"keyword",e.target.value)}/>
        <input style={S.kwN} value={k.note} placeholder="Note" onChange={e=>updKw(sg.id,ki,"note",e.target.value)}/>
        <button style={S.icB} onClick={()=>rmKw(sg.id,ki)}>✕</button>
      </div>))}<button style={S.ghost} onClick={()=>addKw(sg.id)}>+ Keyword</button></div></div>))}
      {segs.length>0&&segs.some(s=>s.keywords.length)&&<button style={S.pri} onClick={runSeg}>Run Segmentation</button>}
    </div>);
    case 7:{const yrs=Object.keys(segData),tot=yrs.reduce((s,y)=>s+(segData[y]?.length||0),0),kwM=yrs.reduce((s,y)=>s+(segData[y]?.filter(r=>r["Keyword_Match_Status"]==="Keyword Match").length||0),0),indM=yrs.reduce((s,y)=>s+(segData[y]?.filter(r=>r["Final_Segmentation_Source"]==="Industry Segment").length||0),0),unc=yrs.reduce((s,y)=>s+(segData[y]?.filter(r=>r["Final_Segmentation"]==="Unclassified").length||0),0);
      const pct=n=>tot>0?((n/tot)*100).toFixed(1)+"%":"0%";
      return(<div style={S.sc}>
        <h3 style={S.secT}>Segmentation Results</h3>
        <div style={S.mr2}><M label="Total" value={tot}/><M label="Keyword" value={kwM} sub={pct(kwM)}/><M label="Industry" value={indM} sub={pct(indM)}/><M label="Unclassified" value={unc} sub={pct(unc)}/></div>
        {yrs.map(y=><SChart key={y} data={segData[y]} year={y}/>)}
        <button style={S.pri} onClick={runIQR}>Run IQR</button>
      </div>);
    }
    case 8:return(<div style={S.sc}>
      <h3 style={S.secT}>Unit Conversion & IQR</h3>
      <div style={S.mr2}><M label="Non-Conv" value={_.uniqBy(nonConv,"Unit").length}/><M label="IQR Groups" value={iqrSum.length}/></div>
      {nonConv.length>0&&<DT data={_.uniqBy(nonConv,r=>r.Year+r.Unit)} title="Non-Convertible/Ambiguous"/>}
      <DT data={iqrSum} title="IQR Summary"/>
      {Object.entries(bef).map(([y,r])=>{const a=aft[y]||[];const rm=r.length-a.length;return(<div key={y} style={S.yb}><div style={S.yl}>Year {y}</div><div style={S.mr2}><M label="Before" value={r.length}/><M label="After" value={a.length}/><M label="Removed" value={rm} sub={r.length>0?((rm/r.length)*100).toFixed(1)+"%":"0%"}/></div></div>)})}
      <button style={S.pri} onClick={()=>go(9)}>View Results</button>
    </div>);
    case 9:{const yrs=Object.keys(aft);return(<div style={S.sc}>
      <h3 style={S.secT}>Final Results</h3>
      <div style={S.exBox}><button style={S.exBtn} onClick={exportAll}>Download Complete Workbook</button><p style={S.exH}>Summary, Quality, Purchaser List, Top 80% Ranking, Industry Match, Company Log, IQR, Before/After sheets</p></div>
      {yrs.map(y=>{const rows=aft[y]||[];const gs=_.groupBy(rows,"Final_Segmentation");
        const tbl=Object.entries(gs).map(([seg,sr])=>{const el=sr.filter(r=>r["Price_Eligible"]==="Yes");const tv=el.reduce((s,r)=>s+(parseFloat(r["TotalValue_After_Conversion"])||0),0);const qm=el.reduce((s,r)=>s+(parseFloat(r["Quantity_MT"])||0),0);return{Segment:seg,Rows:sr.length,Eligible:el.length,Value:"$"+(tv/1e6).toFixed(2)+"M",MT:qm.toFixed(1),Avg:qm>0?"$"+(tv/qm).toFixed(2):"N/A"}}).sort((a,b)=>b.Rows-a.Rows);
        return(<div key={y}><DT data={tbl} title={y+" — Clean Summary"}/><SChart data={rows} year={y+" (Clean)"}/></div>)
      })}
    </div>);
    }
    default:return null;
  }};

  return(<div style={S.app}>
    <header style={S.hdr}><div style={S.hdrL}><div style={S.logo}>⬡</div><div><div style={S.appT}>ChemSeg Analyst</div><div style={S.appS}>V15</div></div></div>{wbName&&<div style={S.hdrF}>{wbName}</div>}</header>
    <StageNav stage={stage} setStage={setStage} maxStage={maxStage}/>
    {processing&&<div style={S.pBar}><div style={S.spin}/><span>{statusMsg}</span></div>}
    <main style={S.main}>{R()}</main>
  </div>);
}

// ═══════════════════════════════════════════════════════════
// STYLES
// ═══════════════════════════════════════════════════════════
const C={bg:"#090E14",surface:"#111921",surface2:"#1A2530",border:"#253040",text:"#D0D8E0",textMuted:"#6B7F92",accent:"#00D4AA",accentDark:"#009E7E",accentGlow:"rgba(0,212,170,0.08)",danger:"#E05252",warning:"#E8A838",success:"#34C770",white:"#EEF2F6"};
const fn="'DM Sans',sans-serif",mo="'JetBrains Mono',monospace";
const S={
  app:{fontFamily:fn,background:C.bg,color:C.text,minHeight:"100vh",display:"flex",flexDirection:"column"},
  hdr:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 16px",borderBottom:"1px solid "+C.border,background:C.surface},
  hdrL:{display:"flex",alignItems:"center",gap:10},
  logo:{fontSize:20,color:C.accent,fontWeight:700,width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center",background:C.accentGlow,borderRadius:7,border:"1px solid "+C.accent+"30"},
  appT:{fontSize:15,fontWeight:700,color:C.white,letterSpacing:-.3},appS:{fontSize:9,color:C.textMuted},
  hdrF:{fontSize:10,color:C.textMuted,fontFamily:mo,background:C.surface2,padding:"2px 7px",borderRadius:3},
  stageNav:{display:"flex",gap:1,padding:"0 8px",background:C.surface,borderBottom:"1px solid "+C.border,overflowX:"auto"},
  stageBtn:{display:"flex",alignItems:"center",gap:4,padding:"8px 10px",background:"transparent",border:"none",color:C.textMuted,cursor:"pointer",fontSize:10,fontFamily:fn,whiteSpace:"nowrap",borderBottom:"2px solid transparent",transition:"all .15s"},
  stageBtnActive:{color:C.accent,borderBottomColor:C.accent,background:C.accentGlow},
  stageBtnDisabled:{opacity:.25,cursor:"not-allowed"},
  main:{flex:1,padding:16,maxWidth:1200,margin:"0 auto",width:"100%"},
  sc:{display:"flex",flexDirection:"column",gap:12},
  hero:{textAlign:"center",padding:"20px 0 6px"},heroT:{fontSize:22,fontWeight:700,color:C.white,letterSpacing:-.5},heroS:{fontSize:12,color:C.textMuted,marginTop:4},
  drop:{border:"2px dashed "+C.border,borderRadius:8,padding:"24px 14px",textAlign:"center",cursor:"pointer",background:C.surface,transition:"all .2s"},
  dropAct:{borderColor:C.accent,background:C.accentGlow},dropIc:{fontSize:26,color:C.accent,marginBottom:4},dropLbl:{fontSize:12,fontWeight:600,color:C.text},dropH:{fontSize:9,color:C.textMuted,marginTop:2},
  chip:{display:"inline-block",padding:"4px 8px",background:C.surface2,borderRadius:4,fontSize:10,color:C.accent,fontFamily:mo},
  secT:{fontSize:16,fontWeight:700,color:C.white,letterSpacing:-.2},hint:{fontSize:11,color:C.textMuted,margin:"1px 0 5px"},
  mr2:{display:"flex",gap:7,flexWrap:"wrap"},
  mc:{flex:"1 1 110px",background:C.surface,borderRadius:6,padding:"10px 12px",border:"1px solid "+C.border},
  mv:{fontSize:22,fontWeight:700,color:C.accent,fontFamily:mo},ml:{fontSize:9,color:C.textMuted,marginTop:1,textTransform:"uppercase",letterSpacing:.3},ms:{fontSize:9,color:C.textMuted},
  mapG:{display:"flex",flexDirection:"column",gap:6,marginBottom:12},mapR:{display:"flex",alignItems:"center",gap:7},mapL:{width:140,fontSize:11,fontWeight:600,color:C.text,flexShrink:0},
  sel:{flex:1,padding:"5px 8px",background:C.surface2,border:"1px solid "+C.border,borderRadius:4,color:C.text,fontSize:11,fontFamily:fn,outline:"none"},
  pri:{padding:"8px 18px",background:C.accent,color:C.bg,border:"none",borderRadius:5,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:fn,alignSelf:"flex-start"},
  sec:{padding:"6px 14px",background:"transparent",color:C.accent,border:"1px solid "+C.accent,borderRadius:4,fontSize:10,fontWeight:600,cursor:"pointer",fontFamily:fn,whiteSpace:"nowrap"},
  link:{background:"none",border:"none",color:C.textMuted,cursor:"pointer",fontSize:10,fontFamily:fn,textDecoration:"underline",padding:"4px 0",alignSelf:"flex-start"},
  ghost:{background:"none",border:"1px dashed "+C.border,color:C.textMuted,padding:"3px 9px",borderRadius:3,cursor:"pointer",fontSize:10,fontFamily:fn},
  danSm:{background:C.danger+"22",color:C.danger,border:"1px solid "+C.danger+"44",borderRadius:3,padding:"2px 6px",cursor:"pointer",fontSize:10,fontWeight:700,fontFamily:fn},
  icB:{background:"transparent",color:C.danger,border:"none",cursor:"pointer",fontSize:13,padding:"1px 4px",opacity:.7},
  exBtn:{padding:"10px 24px",background:"linear-gradient(135deg,"+C.accent+","+C.accentDark+")",color:C.bg,border:"none",borderRadius:6,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:fn},
  exBox:{background:C.surface,borderRadius:8,padding:16,textAlign:"center",border:"1px solid "+C.border},
  exH:{fontSize:9,color:C.textMuted,marginTop:5,maxWidth:480,margin:"5px auto 0"},
  yb:{background:C.surface,borderRadius:6,padding:10,border:"1px solid "+C.border},yl:{fontSize:13,fontWeight:700,color:C.accent,marginBottom:7},
  qg:{display:"flex",gap:6,flexWrap:"wrap"},qc:{flex:"1 1 150px",background:C.surface2,borderRadius:4,padding:"6px 9px"},qf:{fontSize:10,fontWeight:600,color:C.text},qr:{fontSize:14,fontWeight:700,color:C.warning,marginTop:1,fontFamily:mo},qd:{fontSize:8,color:C.textMuted,marginTop:1,lineHeight:1.4},
  dv:{borderTop:"1px solid "+C.border,margin:"5px 0"},
  warn:{background:C.warning+"18",border:"1px solid "+C.warning+"44",borderRadius:5,padding:"8px 12px",fontSize:11,color:C.warning,lineHeight:1.4},
  addR:{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap"},addIn:{flex:"1 1 160px",padding:"7px 10px",background:C.surface2,border:"1px solid "+C.border,borderRadius:4,color:C.text,fontSize:12,fontFamily:fn,outline:"none",minWidth:140},
  empty:{textAlign:"center",padding:"24px 14px",background:C.surface,borderRadius:8,border:"1px dashed "+C.border,color:C.textMuted,fontSize:12},
  sgC:{background:C.surface,borderRadius:6,border:"1px solid "+C.border,overflow:"hidden"},sgH:{display:"flex",alignItems:"center",gap:7,padding:"8px 10px",background:C.surface2,borderBottom:"1px solid "+C.border},
  sgN:{flex:1,background:"transparent",border:"none",color:C.accent,fontSize:13,fontWeight:700,fontFamily:fn,outline:"none"},sgK:{fontSize:9,color:C.textMuted},
  kwL:{padding:7,display:"flex",flexDirection:"column",gap:3},kwR:{display:"flex",gap:4,alignItems:"center"},
  kwI:{flex:3,padding:"4px 8px",background:C.surface2,border:"1px solid "+C.border,borderRadius:3,color:C.text,fontSize:11,fontFamily:fn,outline:"none",minWidth:80},
  kwN:{flex:2,padding:"4px 8px",background:C.surface2,border:"1px solid "+C.border,borderRadius:3,color:C.textMuted,fontSize:10,fontFamily:fn,outline:"none",minWidth:60},
  cb:{background:C.surface,borderRadius:6,padding:10,border:"1px solid "+C.border},ct:{fontSize:11,fontWeight:600,color:C.text,marginBottom:7},
  br:{display:"flex",alignItems:"center",gap:6,marginBottom:4},bl:{width:120,fontSize:9,color:C.textMuted,textAlign:"right",flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"},
  bt:{flex:1,height:16,background:C.surface2,borderRadius:3,overflow:"hidden"},bf:{height:"100%",background:"linear-gradient(90deg,"+C.accent+","+C.accentDark+")",borderRadius:3,minWidth:2,transition:"width .4s"},
  bv:{width:100,fontSize:8,color:C.textMuted,fontFamily:mo,flexShrink:0},
  tw:{background:C.surface,borderRadius:6,border:"1px solid "+C.border,overflow:"hidden"},tt:{fontSize:11,fontWeight:600,color:C.text,padding:"8px 10px",borderBottom:"1px solid "+C.border,background:C.surface2},rc:{fontSize:9,color:C.textMuted,fontWeight:400},
  ts:{overflowX:"auto",maxHeight:340},tbl:{width:"100%",borderCollapse:"collapse",fontSize:9,fontFamily:mo},
  th:{padding:"5px 8px",background:C.surface2,color:C.textMuted,textAlign:"left",fontWeight:600,fontSize:8,textTransform:"uppercase",letterSpacing:.2,borderBottom:"1px solid "+C.border,whiteSpace:"nowrap",position:"sticky",top:0,zIndex:1},
  td:{padding:"3px 8px",borderBottom:"1px solid "+C.border,whiteSpace:"nowrap",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",color:C.text},
  trA:{background:"rgba(0,0,0,.08)"},mr:{padding:"4px 10px",fontSize:8,color:C.textMuted,textAlign:"center"},emptyMsg:{padding:12,textAlign:"center",color:C.textMuted,fontSize:10},
  pBar:{display:"flex",alignItems:"center",gap:7,padding:"6px 16px",background:C.accentGlow,color:C.accent,fontSize:10,fontWeight:500,borderBottom:"1px solid "+C.accent+"30"},
  spin:{width:12,height:12,border:"2px solid "+C.border,borderTopColor:C.accent,borderRadius:"50%",animation:"spin .7s linear infinite",flexShrink:0},
};
