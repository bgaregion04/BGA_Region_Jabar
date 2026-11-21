/* FB Share Jabar dashboard — calc + render + exports */
(function(){
'use strict';
const SEED = window.__SEED__;
const ORDER = SEED.order;                       // 18 city names, table order
const UP = ORDER.map(s=>s.toUpperCase());
const K = ['tsel','xl','isat','three','sf','xlplus','ioh'];
const META_OP = {tsel:'Telkomsel',xl:'XLSmart',isat:'Indosat IM3',three:'Tri Indonesia',sf:'Smartfren',xlplus:'XL+',ioh:'IOH'};
const PATH_OPS = ['Telkomsel','XLSmart','Indosat IM3','Smartfren','Tri Indonesia'];
const PATH_LBL = {'Telkomsel':'Telkomsel','XLSmart':'XL','Indosat IM3':'Indosat IM3','Smartfren':'Smartfren','Tri Indonesia':'Tri Indonesia'};
const SHORT = {'Telkomsel':'TSEL','XLSmart':'XL','Indosat IM3':'IM3','Smartfren':'SF','Tri Indonesia':'Tri'};
const BRANCHES = ['BANDUNG','CIREBON','SOREANG','TASIK'];
const BR_NAME = {BANDUNG:'Bandung',CIREBON:'Cirebon',SOREANG:'Soreang',TASIK:'Tasikmalaya',JABAR:'Jabar'};
const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/* ---------- history (seed + saved weeks) ---------- */
const H = {ms:{}, ga:{}, jms:{}, jga:{}, jfull:{}, paths:{}, kms:{}, kga:{}};
for (const [d,v] of Object.entries(SEED.ms)) H.ms[+d]=v;
for (const [d,v] of Object.entries(SEED.ga)) H.ga[+d]=v;
for (const r of SEED.jms) H.jms[r[0]]={ioh:r[1],tsel:r[2],xlplus:r[3]};
for (const r of SEED.jga) H.jga[r[0]]={ioh:r[1],tsel:r[2],xlplus:r[3]};
for (const [d,v] of Object.entries(SEED.jfull)) H.jfull[+d]=v;
for (const [d,v] of Object.entries(SEED.paths||{})) H.paths[+d]=v;
for (const [d,v] of Object.entries(SEED.kms||{})) H.kms[+d]=v;
for (const [d,v] of Object.entries(SEED.kga||{})) H.kga[+d]=v;
function mergeWeek(d,p){ if(p.ms)H.ms[d]=p.ms; if(p.ga)H.ga[d]=p.ga; if(p.jms)H.jms[d]=p.jms; if(p.jga)H.jga[d]=p.jga; if(p.jfull)H.jfull[d]=p.jfull; if(p.paths)H.paths[d]=p.paths; if(p.kms)H.kms[d]=p.kms; if(p.kga)H.kga[d]=p.kga; }

/* ---------- date helpers (Excel serial) ---------- */
const EPOCH = Date.UTC(1899,11,30);
const toSerial = (y,m,d)=>Math.round((Date.UTC(y,m-1,d)-EPOCH)/864e5);
const fromSerial = s=>{const t=new Date(EPOCH+s*864e5);return {y:t.getUTCFullYear(),m:t.getUTCMonth()+1,d:t.getUTCDate()};};
const monthOf = s=>fromSerial(s).m;
const mPrev = F=> monthOf(F-28)===monthOf(F) ? F-35 : F-28;      // workbook rule
const yPrev = F=> monthOf(F-364)===monthOf(F) ? F-364 : F-371;
function isoWeek(s){const t=fromSerial(s);const dt=new Date(Date.UTC(t.y,t.m-1,t.d));const day=(dt.getUTCDay()+6)%7;dt.setUTCDate(dt.getUTCDate()-day+3);const fy=new Date(Date.UTC(dt.getUTCFullYear(),0,4));return 1+Math.round(((dt-fy)/864e5-3+((fy.getUTCDay()+6)%7))/7);}
const pad=n=>String(n).padStart(2,'0');
function labels(F){const t=fromSerial(F);const mon=MON[t.m-1];const yy=String(t.y).slice(2);
  return {mdy:`${t.m}/${t.d}/${t.y}`, kec:`${t.d}-${mon}`, ms:`${pad(t.d)} ${mon}-${yy}`, ga:`${pad(t.d)}-${mon}-${yy}`,
          wk:`W${Math.ceil(t.d/7)} ${mon}`, mtd:`MTD ${mon} ${t.y}`, iso:isoWeek(F), y:t.y, m:t.m, d:t.d, ymd:`${t.y}${pad(t.m)}${pad(t.d)}`, ym:`${t.y}${pad(t.m)}`};}

/* ---------- number helpers ---------- */
const r2 = x=>Math.round((x+(x>=0?1e-9:-1e-9))*100)/100;
const f2 = x=>(x==null||isNaN(x))?'–':r2(x).toFixed(2);
const fpp = x=>(x==null||isNaN(x))?'–':r2(x).toFixed(2)+'pp';
const fint = x=>(x==null||isNaN(x))?'–':Math.round(x).toLocaleString('en-US');
const fK0 = x=>{const k=Math.round(x/1000);return k===0?'0':k+'K';};
const fK1 = x=>(x>0?'+':'')+(x/1000).toFixed(1)+'K';
const sgnPct = x=>(x>0?'+':'')+r2(x).toFixed(2)+'%';

/* ---------- CSV parsing ---------- */
function parseCSV(text){
  text=text.replace(/^\uFEFF/,''); const rows=[]; let row=[],cell='',q=false;
  for(let i=0;i<text.length;i++){const c=text[i];
    if(q){ if(c==='"'){ if(text[i+1]==='"'){cell+='"';i++;} else q=false; } else cell+=c; }
    else if(c==='"') q=true; else if(c===','){row.push(cell);cell='';}
    else if(c==='\n'){row.push(cell);rows.push(row);row=[];cell='';} else if(c!=='\r') cell+=c; }
  if(cell!==''||row.length){row.push(cell);rows.push(row);} return rows;}
const num = s=>{ if(s==null)return null; s=String(s).trim(); if(s===''||s==='--'||s.toLowerCase()==='null')return null; const v=Number(s.replace(/,/g,'')); return isNaN(v)?null:v; };
function headerIndex(rows,first){ for(let i=0;i<Math.min(rows.length,6);i++) if(rows[i][0]&&rows[i][0].trim().startsWith(first)) return i; return -1; }
function detect(name,text){
  const rows=parseCSV(text);
  if(headerIndex(rows,'City')>=0) return {kind:'city',rows};
  if(headerIndex(rows,'Kecamatan')>=0) return {kind:'kec',rows};
  if(headerIndex(rows,'date')>=0 && /paths/i.test(name)){
    const up=name.toUpperCase(); let br=null;
    for(const b of ['JABAR','BANDUNG','CIREBON','SOREANG','TASIK']) if(up.startsWith(b)||up.includes(b+'PATHS')) {br=b;break;}
    return {kind:'paths',branch:br,rows};
  }
  return {kind:'unknown',rows};
}
function metaTable(rows,first){
  const h=headerIndex(rows,first); const hdr=rows[h]; const out=[];
  for(const r of rows.slice(h+1)){ if(!r[0]) continue; const o={}; hdr.forEach((k,i)=>{o[k]=r[i];}); out.push(o); }
  return {hdr,out};
}

/* ---------- state ---------- */
const S = {files:{}, city:null, kec:null, paths:{}, F:null, warn:[], metaMS:null, metaGA:null};
window.__S=S; window.__H=H;

function dateFromName(n){const m=n.match(/(20\d\d)-(\d\d)-(\d\d)/); return m?toSerial(+m[1],+m[2],+m[3]):null;}

function ingest(name,text){
  const d=detect(name,text);
  if(d.kind==='city'){ const {hdr,out}=metaTable(d.rows,'City'); S.city={name,hdr,rows:out}; const f=dateFromName(name); if(f){S.F=f;S.uploadF=f;} }
  else if(d.kind==='kec'){ const {hdr,out}=metaTable(d.rows,'Kecamatan'); S.kec={name,hdr,rows:out}; const f=dateFromName(name); if(f&&!S.city){S.F=f;S.uploadF=f;} }
  else if(d.kind==='paths' && d.branch){ const {out}=metaTable(d.rows,'date'); let pd=null;
    if(out[0] && /^\d{4}-\d\d-\d\d$/.test(out[0].date)){const [y,m,dd]=out[0].date.split('-').map(Number); pd=toSerial(y,m,dd);}
    S.paths[d.branch]={name,text,rows:out,date:pd}; if(!S.city&&!S.kec&&pd&&!S.uploadF) S.F=pd; }
  return d.kind==='paths'?(d.branch?'paths '+d.branch:'paths (branch not recognised)'):d.kind;
}

/* ---------- city values ---------- */
function cityVals(){
  const res={};
  if(!S.city){ const a=H.ms[S.F], g=H.ga[S.F]; if(!a) return res; UP.forEach((c,i)=>{ if(a[i]) res[c]={ms:a[i],ga:g&&g[i]?g[i]:K.map(()=>0),cust:null}; }); return res; }
  for(const r of S.city.rows){ const nm=r['City']; if(!nm||!/, Jawa Barat$/.test(nm)) continue; const c=nm.replace(/, Jawa Barat$/,'').trim().toUpperCase();
    const g=m=>K.map(k=>num(r[`${META_OP[k]}-Mobile ${m}`]));
    res[c]={ms:g('Market share'),ga:g('Gross adds'),cust:g('Total customers')}; }
  return res;
}
function jabarDerived(cv){
  const i5=[0,2,3,1,4]; let W=0; const acc=K.map(()=>0);
  for(const c of UP){ const v=cv[c]; if(!v||!v.cust) continue; let cs=0,ss=0; for(const i of [0,1,2,3,4]){cs+=v.cust[i]||0; ss+=v.ms[i]||0;} const w=cs/ss; W+=w; K.forEach((k,i)=>acc[i]+=(v.ms[i]||0)*w); }
  return W?acc.map(a=>r2(a/W)):null;
}

/* ---------- main compute ---------- */
function compute(){
  const C={}; if(!S.F) S.F=Math.max(...Object.keys(H.ms).map(Number)); if(!S.city&&!H.ms[S.F]) return null;
  C.fromHistory=!S.city;
  const F=S.F, X=F-7, P=mPrev(F), Y=yPrev(F); C.F=F;C.X=X;C.P=P;C.Y=Y;C.L=labels(F);
  const cv=cityVals(); C.cv=cv;
  C.missingCities=UP.filter(c=>!cv[c]);
  const idx={tsel:0,ioh:6,xlplus:5};
  const hv=(d,i,k)=>{const a=H.ms[d];return a&&a[i]?a[i][idx[k]]:null;};
  const gv=(d,i,k)=>{const a=H.ga[d];return a&&a[i]?a[i][idx[k]]:null;};
  // ---- MS table
  C.ms=ORDER.map((nm,i)=>{const v=cv[UP[i]]; if(!v) return null;
    const cur={tsel:v.ms[0],ioh:v.ms[6],xlplus:v.ms[5]};
    const d=(dd)=>({tsel:cur.tsel-hv(dd,i,'tsel'),ioh:cur.ioh-hv(dd,i,'ioh'),xlplus:cur.xlplus-hv(dd,i,'xlplus')});
    return rowMS(nm,cur,hv(X,i,'tsel')==null?null:d(X),hv(P,i,'tsel')==null?null:d(P),hv(Y,i,'tsel')==null?null:d(Y));});
  // Jabar
  const derived=jabarDerived(cv); const hj=H.jfull[F];
  C.jDerived=derived||hj||K.map(()=>0);
  const jfull=S.metaMS||(C.fromHistory&&hj?hj:null)||derived||hj; C.jfull=jfull; C.jFromMeta=!!S.metaMS||(C.fromHistory&&!!hj);
  const jcur={tsel:jfull[0],ioh:jfull[6],xlplus:jfull[5]};
  const jd=dd=>{const p=H.jms[dd];return p?{tsel:jcur.tsel-p.tsel,ioh:jcur.ioh-p.ioh,xlplus:jcur.xlplus-p.xlplus}:null;};
  C.msJ=rowMS('Jabar',jcur,jd(X),jd(P),jd(Y));
  const rows=C.ms.filter(Boolean);
  C.nWin=rows.filter(r=>r.win).length;
  const wins=rows.filter(r=>r.win), loses=rows.filter(r=>!r.win);
  C.minWin=wins.length?wins.reduce((a,b)=>b.gap<a.gap?b:a):null;
  C.minLose=loses.length?loses.reduce((a,b)=>b.gap>a.gap?b:a):null;
  C.nNeg=rows.filter(r=>r.wow&&r.mom&&r2(r.wow.tsel)<0&&r2(r.mom.tsel)<0).length;
  const lst=(key,asc)=>{const a=C.ms.map((r,j)=>({r,j})).filter(o=>o.r&&o.r[key]);
    a.sort((p,q)=>asc?(r2(p.r[key].tsel)-r2(q.r[key].tsel))||(q.j-p.j):(r2(q.r[key].tsel)-r2(p.r[key].tsel))||(p.j-q.j));
    return a.slice(0,3).map(o=>[o.r.name,r2(o.r[key].tsel)]);};
  C.msTopW=lst('wow',0);C.msBotW=lst('wow',1);C.msTopM=lst('mom',0);C.msBotM=lst('mom',1);
  // ---- GA table
  C.ga=ORDER.map((nm,i)=>{const v=cv[UP[i]]; if(!v) return null; const cur=[v.ga[0]||0,v.ga[6]||0,v.ga[5]||0];
    const w=gv(X,i,'tsel')==null?null:[cur[0]-gv(X,i,'tsel'),cur[1]-gv(X,i,'ioh'),cur[2]-gv(X,i,'xlplus')];
    const m=gv(P,i,'tsel')==null?null:[cur[0]-gv(P,i,'tsel'),cur[1]-gv(P,i,'ioh'),cur[2]-gv(P,i,'xlplus')];
    return rowGA(nm,cur,w,m);});
  const gaSrc=S.metaGA||(H.jga[F]?[H.jga[F].tsel,H.jga[F].ioh,H.jga[F].xlplus]:null); C.gaFromTrend=!S.metaGA&&!!gaSrc;
  if(gaSrc){const cur=gaSrc; const pw=H.jga[X], pm=H.jga[P];
    C.gaJ=rowGA('JABAR',cur,pw?[cur[0]-pw.tsel,cur[1]-pw.ioh,cur[2]-pw.xlplus]:null,pm?[cur[0]-pm.tsel,cur[1]-pm.ioh,cur[2]-pm.xlplus]:null);} else C.gaJ=null;
  const gr=C.ga.filter(Boolean); C.gaHi={TSEL:0,IOH:0,'XL+':0}; gr.forEach(r=>C.gaHi[r.hi]++);
  const glst=(key,asc)=>{const a=C.ga.map((r,j)=>({r,j})).filter(o=>o.r&&o.r[key]);
    a.sort((p,q)=>asc?(p.r[key][0]-q.r[key][0])||(q.j-p.j):(q.r[key][0]-p.r[key][0])||(p.j-q.j));
    return a.slice(0,3).map(o=>[o.r.name,o.r[key][0]]);};
  C.gaTopW=glst('wow',0);C.gaBotW=glst('wow',1);C.gaTopM=glst('mom',0);C.gaBotM=glst('mom',1);
  // ---- Paths
  S.warn=[]; const PT={}; C.pathSrc={};
  for(const b of ['JABAR',...BRANCHES]){ const up=S.paths[b];
    if(up&&up.date===F){PT[b]={text:up.text,rows:up.rows}; C.pathSrc[b]='upload';}
    else if(H.paths[F]&&H.paths[F][b]){const t=H.paths[F][b]; PT[b]={text:t,rows:metaTable(parseCSV(t),'date').out}; C.pathSrc[b]='saved';}
    else if(up) S.warn.push(`Paths ${BR_NAME[b]} is for ${up.date?fromSerial(up.date).d+' '+MON[fromSerial(up.date).m-1]:'another week'}, not this week. Upload this week’s file.`); }
  C.pathText={}; for(const b in PT) C.pathText[b]=PT[b].text;
  C.paths={}; for(const b in PT) C.paths[b]=matrix(PT[b].rows);
  const ks=Object.keys(PT); C.dupPaths=new Set();
  for(let a=0;a<ks.length;a++)for(let b=a+1;b<ks.length;b++) if(PT[ks[a]].text===PT[ks[b]].text){C.dupPaths.add(ks[a]);C.dupPaths.add(ks[b]); S.warn.push(`Paths ${BR_NAME[ks[a]]} and ${BR_NAME[ks[b]]} are identical files. One branch was exported twice; re-export the correct one before using it.`);}
  if(!gaSrc) S.warn.push('Gross add Jabar for this week is not in the trend yet. Enter TSEL, IOH and XL+ from the Meta Gross adds screenshot in Update data.');
  if(!C.jFromMeta) S.warn.push('Jabar share is derived from city data (customer-weighted). Enter the Meta screenshot values in Update data to use the official figures.');
  if(C.missingCities.length) S.warn.push('Missing Jabar cities in City CSV: '+C.missingCities.join(', '));
  if(!H.ms[X]) S.warn.push('No history for the previous week; WoW cannot be calculated.');
  C.caption=caption(C);
  return C;
}
function rowMS(name,cur,wow,mom,yoy){
  const t=cur.tsel,i=cur.ioh,x=cur.xlplus,mx=Math.max(t,i,x);
  const hi=t===mx?'TSEL':(i===mx?'IOH':'XL+'); const comp=i>=x?'IOH':'XL+';
  const hg=d=>{if(!d)return null;const a=[r2(d.tsel),r2(d.ioh),r2(d.xlplus)];const m=Math.max(...a);return a[0]===m?'TSEL':(a[1]===m?'IOH':'XL+');};
  return {name,cur,wow,mom,yoy,win:t===mx,gap:r2(t-Math.max(i,x)),to:comp,hi,rank:1+(i>t)+(x>t),hiComp:comp,hgW:hg(wow),hgM:hg(mom)};
}
function rowGA(name,cur,wow,mom){const mx=Math.max(...cur);
  return {name,cur,wow,mom,win:cur[0]===mx,gap:cur[0]-Math.max(cur[1],cur[2]),hi:cur[0]===mx?'TSEL':(cur[1]===mx?'IOH':'XL+'),rank:1+(cur[1]>cur[0])+(cur[2]>cur[0])};}
function matrix(rows){
  const m={}; let lt=0;
  for(const r of rows){ if(!PATH_OPS.includes(r.from)||!PATH_OPS.includes(r.to)||r.from===r.to) continue;
    const v=num(r.amount); if(v==null) lt++; m[r.from+'|'+r.to]=v==null?0:v; }
  const out={},inn={},io={}; let tot=0;
  for(const a of PATH_OPS){ out[a]=0; inn[a]=0; }
  for(const a of PATH_OPS)for(const b of PATH_OPS){ if(a===b)continue; const v=m[a+'|'+b]||0; out[a]+=v; inn[b]+=v; tot+=v; }
  for(const a of PATH_OPS) io[a]=inn[a]-out[a];
  const others=PATH_OPS.filter(o=>o!=='Telkomsel');
  const tOut=others.reduce((a,b)=>(m['Telkomsel|'+b]||0)>(m['Telkomsel|'+a]||0)?b:a);
  const tIn=others.reduce((a,b)=>(m[b+'|Telkomsel']||0)>(m[a+'|Telkomsel']||0)?b:a);
  const gainMax=PATH_OPS.reduce((a,b)=>io[b]>io[a]?b:a);
  return {m,out,inn,io,tot,tOut,tIn,gainMax,lt};
}

/* ---------- kecamatan ---------- */
function kecCurrent(F){ // returns {ms:[433][3], ga:[433][3]} for week F (tsel, ioh, xlplus)
  if(S.kec&&S.F===F){ const kv=kecVals(); const pick=(v,f)=>v?[v[f][0],v[f][6],v[f][5]]:null;
    return {ms:SEED.kecmeta.map(m=>pick(kv[m[1]],'ms')), ga:SEED.kecmeta.map(m=>pick(kv[m[1]],'ga'))}; }
  return {ms:H.kms[F]||null, ga:H.kga[F]||null}; }
function kecTable(C,cityName){
  const F=C.F, cur=kecCurrent(F); const up=cityName.toUpperCase(); const idx=SEED.kecmeta.map((m,i)=>[m,i]).filter(([m])=>m[2]===up);
  const g=(arr,i,k)=>arr&&arr[i]?arr[i][k]:null;
  const ms=[],ga=[];
  for(const [m,i] of idx){ const c=cur.ms&&cur.ms[i]; if(c&&c[0]!=null){ const v={tsel:c[0],ioh:c[1]||0,xlplus:c[2]||0};
      const d=dd=>{const p=H.kms[dd]&&H.kms[dd][i]; return p&&p[0]!=null?{tsel:v.tsel-p[0],ioh:v.ioh-(p[1]||0),xlplus:v.xlplus-(p[2]||0)}:null;};
      ms.push(Object.assign(rowMS(m[0],v,d(C.X),d(C.P),d(C.Y)),{kid:m[1]})); }
    const q=cur.ga&&cur.ga[i]; if(q){ const v=[q[0]||0,q[1]||0,q[2]||0];
      const d=dd=>{const p=H.kga[dd]&&H.kga[dd][i]; return p?[v[0]-(p[0]||0),v[1]-(p[1]||0),v[2]-(p[2]||0)]:null;};
      ga.push(Object.assign(rowGA(m[0],v,d(C.X),d(C.P)),{kid:m[1]})); } }
  return {ms,ga,hasCur:!!cur.ms};
}
/* ---------- caption ---------- */
function caption(C){
  const L=C.L, J=C.msJ, lines=[]; const S1=' ';
  const list=(a,fmt)=>a.map((x,i)=>`${i+1}. ${x[0]} ${fmt(x[1])}`).join('\n');
  const ppf=v=>r2(v).toFixed(2)+'pp'; const gaf=v=>Math.round(v).toLocaleString('en-US');
  const comp=J.to;
  let s=`Assalamualaikum wr wb,
Semangat pagi Bang @Andoz, Pak @Danny dan kaka-kaka JABAR..
Berikut terlampir update untuk *Mobile FB share* terbaru pekan ini dengan informasi Before & After XL+.
${S1}
*FB Share Performance ${L.mtd} (${L.mdy})*
1. FB Share ${L.ms} (IOH & after XL+) Jabar ${f2(J.cur.tsel)}% WoW ${J.wow?ppf(J.wow.tsel):'–'}, MoM ${J.mom?ppf(J.mom.tsel):'–'}, Gap to ${comp} ${sgnPct(J.gap)}. TSEL di rank ${J.rank}, dengan total ${C.nWin} City Win dari 18.
2. Gap WIN terkecil di ${C.minWin?C.minWin.name:'–'} Gap to ${C.minWin?C.minWin.to:'–'} ${C.minWin?sgnPct(C.minWin.gap):'–'} dan Gap LOSE terkecil ada di ${C.minLose?C.minLose.name:'–'} dengan Gap to ${C.minLose?C.minLose.to:'–'} ${C.minLose?sgnPct(C.minLose.gap):'–'}.
3. ${C.nNeg} dari 18 kota mengalami growth WoW dan MoM Market Share negative.
${S1}
*Top 3 WoW*
${list(C.msTopW,ppf)}
*Bottom 3 WoW*
${list(C.msBotW,ppf)}
${S1}
*Top 3 MoM*
${list(C.msTopM,ppf)}
*Bottom 3 MoM*
${list(C.msBotM,ppf)}
${S1}
*Gross Add Performance ${L.mtd} (${L.mdy})*
`;
  if(C.gaJ){const g=C.gaJ; const lead=g.hi==='TSEL'?null:g.hi;
    const kw=v=>v===0?'0':fK0(v);
    s+=`1. Gross add TSEL sebesar ${fK0(g.cur[0])} di rank ${g.rank} dengan gap ${fK0(g.gap)} ke ${lead||'kompetitor'} di rank 1, WoW ${g.wow?kw(g.wow[0]):'–'} dan MoM ${g.mom?kw(g.mom[0]):'–'}.\n`;
  } else s+=`1. ⟨isi Gross Add Jabar dari screenshot Meta⟩\n`;
  s+=`2. Gross Add TSEL tertinggi di ${C.gaHi.TSEL} Kota, XL+ di ${C.gaHi['XL+']} Kota, IOH di ${C.gaHi.IOH} Kota.\n`;
  if(C.gaJ&&C.gaJ.wow&&C.gaJ.mom){const g=C.gaJ;
    const tw=v=>v===0?'flat':(v>0?'tumbuh positive':'tumbuh negative'); const cw=v=>v===0?'flat':(v>0?'naik':'turun');
    const ts=tw(g.wow[0])===tw(g.mom[0])?`${tw(g.wow[0])} secara WoW dan MoM`:`${tw(g.wow[0])} secara WoW dan ${tw(g.mom[0])} secara MoM`;
    const cs=(n,w,m)=>cw(w)===cw(m)?`${n} ${cw(w)} WoW dan MoM`:`${n} ${cw(w)} WoW dan ${cw(m)} MoM`;
    s+=`3. Gross add TSEL ${ts}. ${cs('IOH',g.wow[1],g.mom[1])}, ${cs('XL+',g.wow[2],g.mom[2])}.\n`;
  } else s+=`3. ⟨isi Gross Add Jabar dari screenshot Meta⟩\n`;
  s+=`${S1}
*Top 3 WoW*
${list(C.gaTopW,gaf)}
*Bottom 3 WoW*
${list(C.gaBotW,gaf)}
${S1}
*Top 3 MoM*
${list(C.gaTopM,gaf)}
*Bottom 3 MoM*
${list(C.gaBotM,gaf)}
${S1}
`;
  const PJ=C.paths.JABAR && !C.dupPaths.has('JABAR')?C.paths.JABAR:null;
  if(PJ){const t=PJ.io.Telkomsel; const lead=SHORT[PJ.gainMax];
    const word=t<0?'negative':'positive';
    const mean=t<0?'(artinya lebih banyak yang berpindah dari Tsel dibandingkan yang masuk dari kompetitor)':'(artinya lebih banyak yang masuk ke Tsel dibandingkan yang berpindah ke kompetitor)';
    s+=`Secara movement *churn Telkomsel Jabar ${word} ${fK1(t)} ${mean}.* Overall ${lead} memimpin pergerakan baik secara movement maupun net add yang meningkat hingga ${fK1(PJ.io[PJ.gainMax])} pada ${lead}.\n${S1}\n`;
  } else s+=`⟨paragraf churn — butuh Paths Jabar⟩\n${S1}\n`;
  const okB=BRANCHES.filter(b=>C.paths[b]&&!C.dupPaths.has(b));
  if(okB.length===4){
    const M=b=>C.paths[b]; const gm=M(okB[0]).gainMax;
    const allSame=okB.every(b=>M(b).gainMax===gm); const allPos=okB.every(b=>M(b).io[gm]>0); const allTneg=okB.every(b=>M(b).io.Telkomsel<0);
    const bestB=okB.reduce((a,b)=>M(b).io[gm]>M(a).io[gm]?b:a); const worstB=okB.reduce((a,b)=>M(b).io.Telkomsel<M(a).io.Telkomsel?b:a);
    if(allSame&&allPos&&allTneg)
      s+=`Pergerakan kompetisi menunjukkan ${SHORT[gm]} mencatat gain subs positif di all branch, dengan ${SHORT[gm]} terbesar di Branch ${BR_NAME[bestB]}. Sementara, *TSEL mengalami gain subs negatif di all branch, terdalam di Branch ${BR_NAME[worstB]} dengan delta gain ${fK1(M(worstB).io.Telkomsel)}.*\n`;
    else s+=`⟨paragraf branch: arah tidak seragam di semua branch — sesuaikan kalimat⟩\n`;
  } else s+=`⟨paragraf branch — butuh Paths 4 branch yang valid⟩\n`;
  s+=`${S1}\n*Per Branch:*\n`;
  for(const b of BRANCHES){ const M=C.paths[b];
    if(!M||C.dupPaths.has(b)){ s+=`- ${BR_NAME[b]}: ⟨butuh Paths ${BR_NAME[b]}⟩\n`; continue; }
    const o=PATH_LBL[M.tOut]==='Telkomsel'?'':SHORT[M.tOut], i=SHORT[M.tIn];
    const mv=M.tOut===M.tIn?`Move out TSEL terbanyak ke ${o} dan move in juga dari ${i}.`:`Move out TSEL terbanyak ke ${o} dan move in terbanyak dari ${i}.`;
    s+=`- ${BR_NAME[b]}: ${mv} Delta gain terbesar di ${SHORT[M.gainMax]} ${fK1(M.io[M.gainMax])}\n`; }
  s+=`${S1}
Detail Data Level Kecamatan per City terbaru updated di *Web Nobar Jabar*
http://10.54.68.99/optimus/nobar#
intip.in/srmjabar
${S1}
Terima kasih🙏🏻
`;
  return s;
}

/* ---------- exports: nobar CSV ---------- */
function csvLine(a){return a.map(v=>{v=v==null?'':String(v);return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;}).join(',');}
function nobarCity(C){const L=C.L; const out=[['city','city_id','cluster','branch','region','tsel','xl','isat','three','smartfren','xlplus','ioh','event_date','week','month','year','yearweek','yearmonth']];
  for(const m of SEED.citymeta){const v=C.cv[m[0]]; out.push([...m,...K.map((k,i)=>v&&v.ms[i]!=null?v.ms[i].toFixed(2):'0.00'),L.mdy,L.iso,L.m,L.y,`${L.y}${L.iso}`,L.ym]);}
  return out.map(csvLine).join('\r\n')+'\r\n';}
function nobarRegion(C){const L=C.L; const out=[['region','tsel','xl','isat','three','smartfren','xlplus','ioh','level','event_date','week','month','year','yearweek','yearmonth']];
  for(const [r,lvl] of SEED.regmeta){const v=r==='Jabar'?C.jfull.map(x=>x.toFixed(2)):K.map(()=>''); out.push([r,...v,lvl,L.mdy,L.iso,L.m,L.y,`${L.y}${L.iso}`,L.ym]);}
  return out.map(csvLine).join('\r\n')+'\r\n';}
function kecVals(){const res={}; if(!S.kec) return res;
  for(const r of S.kec.rows){const nm=r['Kecamatan 2022']; if(!nm) continue; const id=nm.split('_').pop();
    const g=m=>K.map(k=>num(r[`${META_OP[k]}-Mobile ${m}`])); res[id]={ms:g('Market share'),ga:g('Gross adds'),cust:g('Total customers'),name:nm};}
  return res;}
function nobarKec(C){const L=C.L, kv=kecVals(); const out=[['kecamatan','kec_id','city','city_id','cluster','branch','region','region_id','provinsi','prov_id','area','area_id','tsel','xl','isat','three','smartfren','xlplus','ioh','event_date','week','month','year','yearweek','yearmonth']];
  let miss=0; for(const m of SEED.kecmeta){const v=kv[m[1]]; if(!v)miss++; out.push([...m,...K.map((k,i)=>v&&v.ms[i]!=null?v.ms[i].toFixed(2):'0.00'),L.kec,L.iso,L.m,L.y,`${L.y}${L.iso}`,L.ym]);}
  C.kecMissing=miss; return out.map(csvLine).join('\r\n')+'\r\n';}

/* ---------- exports: xlsx (minimal writer on JSZip) ---------- */
function colName(n){let s='';n++;while(n){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
const xesc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function sheetXML(aoa,dateCols){
  let x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
  aoa.forEach((row,ri)=>{x+=`<row r="${ri+1}">`; row.forEach((v,ci)=>{ if(v==null||v==='')return; const ref=colName(ci)+(ri+1);
    if(typeof v==='number') x+=`<c r="${ref}"${ri>0&&dateCols&&dateCols.includes(ci)?' s="1"':''}><v>${v}</v></c>`;
    else x+=`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;}); x+='</row>';});
  return x+'</sheetData></worksheet>';}
async function buildXlsx(sheets){
  const z=new JSZip(); const n=sheets.length;
  z.file('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  z.file('_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file('xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s,i)=>`<sheet name="${xesc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`);
  z.file('xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  z.file('xl/styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="m/d/yyyy"/></numFmts><fonts count="1"><font><sz val="10"/><name val="Poppins"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>');
  sheets.forEach((s,i)=>z.file(`xl/worksheets/sheet${i+1}.xml`,sheetXML(s.aoa,s.dateCols)));
  return z.generateAsync({type:(typeof Blob!=='undefined'?'blob':'nodebuffer'),mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function excelSheets(C){
  const F=C.F, L=C.L, kv=kecVals(); const sh=[];
  const dd=r=>r?[r.tsel,r.ioh,r.xlplus].map(r2):['','',''];
  sh.push({name:'MS_City',aoa:[['Kabupaten','Tsel','IOH','XL+','Win/Lose','GAP','Highest','Rank TSEL','WoW Tsel','WoW IOH','WoW XL+','MoM Tsel','MoM IOH','MoM XL+','YoY Tsel','YoY IOH','YoY XL+','Highest Competitor','Highest Growth WoW','Highest Growth MoM'],
    ...[...C.ms.filter(Boolean),C.msJ].map(r=>[r.name,r.cur.tsel,r.cur.ioh,r.cur.xlplus,r.win?'Win':'Lose',r.gap,r.hi,r.rank,...dd(r.wow),...dd(r.mom),...dd(r.yoy),r.hiComp,r.hgW||'',r.hgM||''])]});
  sh.push({name:'GA_City',aoa:[['Kabupaten','Tsel','IOH','XL+','Win/Lose','GAP','Highest','Rank TSEL','WoW Tsel','WoW IOH','WoW XL+','MoM Tsel','MoM IOH','MoM XL+'],
    ...[...C.ga.filter(Boolean),...(C.gaJ?[C.gaJ]:[])].map(r=>[r.name,...r.cur,r.win?'Win':'Lose',r.gap,r.hi,r.rank,...(r.wow||['','','']),...(r.mom||['','',''])])]});
  const pm=[]; for(const b of ['JABAR',...BRANCHES]){const M=C.paths[b]; if(!M) continue;
    pm.push([`${BR_NAME[b]}${C.dupPaths.has(b)?' (DUPLICATE FILE – do not use)':''}`]); pm.push(['From \\ To',...PATH_OPS.map(o=>PATH_LBL[o]),'Move to','In-Out']);
    for(const a of PATH_OPS) pm.push([PATH_LBL[a],...PATH_OPS.map(o=>o===a?'':(M.m[a+'|'+o]||0)),M.out[a],M.io[a]]);
    pm.push(['Move In',...PATH_OPS.map(o=>M.inn[o]),M.tot,'']); pm.push([]);}
  sh.push({name:'Paths_Matrix',aoa:pm});
  // paste-ready RAW rows (same column order as master workbooks)
  const cityId={}; SEED.citymeta.forEach(m=>cityId[m[0]]=+m[1]);
  sh.push({name:'RAW_fb_city',dateCols:[9],aoa:[['city','city_id','tsel','xl','isat','three','smartfren','xlplus','ioh','event_date'],...SEED.citymeta.map(m=>[m[0],+m[1],...(C.cv[m[0]]?C.cv[m[0]].ms:K.map(()=>'')).map(v=>v==null?'NULL':v),F])]});
  if(S.city) sh.push({name:'RAW_abs_city',dateCols:[8],aoa:[['city','tsel','xl','isat','three','smartfren','xlplus','ioh','event_date'],...SEED.citymeta.map(m=>[m[0],...(C.cv[m[0]]&&C.cv[m[0]].cust?C.cv[m[0]].cust:K.map(()=>'')).map(v=>v==null?0:v),F])]});
  sh.push({name:'RAW_ga_city',dateCols:[9],aoa:[['city','tsel','xl','isat','three','smartfren','xlplus','ioh','level','event_date'],...SEED.citymeta.map(m=>[m[0],...(C.cv[m[0]]?C.cv[m[0]].ga:K.map(()=>0)).map(v=>v==null?0:v),'City',F])]});
  const rep=[['event_date','city','MS tsel','MS xl','MS isat','MS three','MS smartfren','MS xlplus','MS ioh','GA tsel','GA xl','GA isat','GA three','GA smartfren','GA xlplus','GA ioh','CUST tsel','CUST xl','CUST isat','CUST three','CUST smartfren','CUST xlplus','CUST ioh','CHURN tsel','CHURN xl','CHURN isat','CHURN three','CHURN smartfren','CHURN xlplus','CHURN ioh','NET tsel','NET xl','NET isat','NET three','NET smartfren','NET xlplus','NET ioh']];
  if(S.city) for(const r of S.city.rows){const nm=r['City']; if(!nm) continue; const row=[F,nm.replace(/,[^,]*$/,'').trim().toUpperCase()];
    for(const g of ['Market share','Gross adds','Total customers','Churn','Net adds']) for(const k of K){const v=num(r[`${META_OP[k]}-Mobile ${g}`]); row.push(v==null?'--':v);} rep.push(row);}
  if(S.city) sh.push({name:'RAW_report',dateCols:[0],aoa:rep});
  const kh=['area','area_id','region','region_id','provinsi','prov_id','city','city_id','kecamatan','kec_id','tsel','xl','isat','three','smartfren','xlplus','ioh','event_date'];
  const kr=f=>SEED.kecmeta.map(m=>{const v=kv[m[1]]; return [m[10],+m[11],m[6],+m[7],m[8],+m[9],m[2],+m[3],m[0],+m[1],...K.map((k,i)=>v&&v[f][i]!=null?v[f][i]:0),F];});
  if(S.kec){ sh.push({name:'RAW_kec_fb',dateCols:[17],aoa:[kh,...kr('ms')]}); sh.push({name:'RAW_kec_abs',dateCols:[17],aoa:[kh,...kr('cust')]}); sh.push({name:'RAW_kec_grossadd',dateCols:[17],aoa:[kh,...kr('ga')]}); }
  sh.push({name:'RAW_regional',dateCols:[9],aoa:[['region','tsel','xl','isat','three','smartfren','xlplus','ioh','level','event_date'],...SEED.regmeta.map(([r,l])=>[r,...(r==='Jabar'?C.jfull:K.map(()=>'')),l,F])]});
  sh.push({name:'GA_regional_col',dateCols:[],aoa:[['GA (Fb_share_regional sheet GA, new column)',''],['date',L.mdy],['IOH',C.gaJ?C.gaJ.cur[1]:''],['Telkomsel',C.gaJ?C.gaJ.cur[0]:''],['XL+',C.gaJ?C.gaJ.cur[2]:'']]});
  const pr=[['branch','date','description','from','to','percentage','amount']];
  for(const b of ['JABAR',...BRANCHES]){const t=C.pathText[b]; if(!t) continue; const P={rows:metaTable(parseCSV(t),'date').out}; for(const r of P.rows){ if(!r.date) continue; const [y,mo,d]=r.date.split('-').map(Number); const a=num(r.amount);
    pr.push([b,toSerial(y,mo,d),r.description,r.from,r.to,num(r.percentage),a==null?r.amount:a]);}}
  sh.push({name:'Paths_raw',dateCols:[1],aoa:pr});
  return sh;
}

/* ---------- exports: pptx ---------- */
const COL={red:'E4002B',navy:'1F2A44',green:'00B050',yellow:'FFC000',blue:'0070C0',gray:'7F7F7F',ice:'EEF3F8',white:'FFFFFF',neg:'FF0000'};
const HI_FILL={TSEL:['C00000','FFFFFF'],IOH:['FFFF00','000000'],'XL+':['0070C0','FFFFFF']};
const OP_FILL={'Telkomsel':['FF0000','FFFFFF'],'XLSmart':['0070C0','FFFFFF'],'Indosat IM3':['FFC000','000000'],'Smartfren':['5BC0EB','000000'],'Tri Indonesia':['000000','FFFFFF']};
function header(pptx,sl,title,headline,L){
  sl.background={color:'FFFFFF'};
  sl.addShape(pptx.ShapeType.rect,{x:0,y:0,w:13.333,h:1.0,fill:{color:COL.red},line:{color:COL.red}});
  sl.addText(title,{x:0.25,y:0.05,w:9.5,h:0.36,fontFace:'Poppins',fontSize:18,bold:true,color:'FFFFFF',margin:0});
  sl.addText(headline,{x:0.25,y:0.42,w:10.6,h:0.55,fontFace:'Poppins',fontSize:11,color:'FFFFFF',margin:0,valign:'top'});
  sl.addText('Update : '+L.wk,{x:11.0,y:0.55,w:2.1,h:0.35,fontFace:'Poppins',fontSize:12,bold:true,color:'FFFFFF',align:'right',margin:0});
}
const cell=(t,o={})=>({text:String(t),options:Object.assign({fontFace:'Poppins',fontSize:6.5,align:'center',valign:'middle',margin:0.01},o)});
const ncell=(v,fmt,o={})=>cell(v==null?'–':fmt(v),Object.assign({color:v!=null&&r2(v)<0?COL.neg:'000000'},o));
async function buildPptx(C){
  const pptx=new PptxGenJS(); pptx.layout='LAYOUT_WIDE'; pptx.theme={headFontFace:'Poppins',bodyFontFace:'Poppins'};
  const L=C.L, J=C.msJ;
  // slide 1 MS
  let sl=pptx.addSlide();
  header(pptx,sl,'Market Share Performance',`FB Share ${L.ms} (IOH & after XL+) Jabar ${f2(J.cur.tsel)}% WoW ${J.wow?fpp(J.wow.tsel):'–'}, MoM ${J.mom?fpp(J.mom.tsel):'–'}, Gap to ${J.to} ${r2(J.gap).toFixed(2)} TSEL di rank ${J.rank}, dengan total ${C.nWin} City Win dari 18`,L);
  const JMS=Object.assign({},H.jms,{[C.F]:{ioh:J.cur.ioh,tsel:J.cur.tsel,xlplus:J.cur.xlplus}});
  const jd=Object.keys(JMS).map(Number).sort((a,b)=>a-b).filter(d=>d<=C.F).slice(-95);
  const lab=jd.map(d=>{const t=fromSerial(d);return `${MON[t.m-1]}-${String(t.y).slice(2)}`;});
  sl.addChart(pptx.ChartType.line,[{name:'IOH',labels:lab,values:jd.map(d=>JMS[d].ioh)},{name:'Telkomsel',labels:lab,values:jd.map(d=>JMS[d].tsel)},{name:'XL+',labels:lab,values:jd.map(d=>JMS[d].xlplus)}],
    {x:0.2,y:1.1,w:8.7,h:1.85,chartColors:['FFC000','E4002B','0070C0'],lineSize:1.5,lineDataSymbol:'none',showLegend:true,legendPos:'r',legendFontFace:'Poppins',legendFontSize:8,
     catAxisLabelFontFace:'Poppins',catAxisLabelFontSize:7,valAxisLabelFontFace:'Poppins',valAxisLabelFontSize:7,valAxisHidden:true,valGridLine:{style:'none'},catAxisLabelFrequency:9,catAxisLabelRotate:0,
     showTitle:true,title:'Market Share Trend',titleFontFace:'Poppins',titleFontSize:9});
  sl.addText([
    {text:`TSEL (IOH & After XL+) memiliki ${C.nWin} kota WIN.`,options:{bullet:true,breakLine:true}},
    {text:`Gap WIN terkecil di ${C.minWin?C.minWin.name:'–'} Gap to ${C.minWin?C.minWin.to:'–'} ${C.minWin?sgnPct(C.minWin.gap):'–'} dan Gap LOSE terkecil ada di ${C.minLose?C.minLose.name:'–'} dengan Gap to ${C.minLose?C.minLose.to:'–'} ${C.minLose?sgnPct(C.minLose.gap):'–'}.`,options:{bullet:true,breakLine:true}},
    {text:`${C.nNeg} dari 18 kota mengalami growth WoW dan MoM Market Share negative.`,options:{bullet:true}}],
    {x:9.1,y:1.15,w:4.0,h:1.8,fontFace:'Poppins',fontSize:9.5,color:'000000',valign:'top',paraSpaceAfter:4});
  const H1=['Kabupaten','Tsel','IOH','XL+','Win/Lose','GAP','Highest','Rank TSEL','Tsel','IOH','XL+','Tsel','IOH','XL+','Tsel','IOH','XL+','Highest Competitor','WoW','MoM'];
  const hdrO={bold:true,color:'FFFFFF',fill:{color:COL.red}};
  const top=[cell('Kabupaten',Object.assign({rowspan:2},hdrO)),cell(`FB Share ${L.ms} (IOH & after XL+)`,Object.assign({colspan:7},hdrO)),cell('WOW',Object.assign({colspan:3},hdrO)),cell('MOM',Object.assign({colspan:3},hdrO)),cell('YOY',Object.assign({colspan:3},hdrO)),cell('Highest Competitor',Object.assign({rowspan:2},hdrO)),cell('Highest Growth',Object.assign({colspan:2},hdrO))];
  const sub=H1.slice(1,17).map(t=>cell(t,{bold:true,fill:{color:t==='Tsel'?'C00000':t==='IOH'?'FFFF00':t==='XL+'?'0070C0':'D9D9D9'},color:t==='IOH'?'000000':(t==='Tsel'||t==='XL+'?'FFFFFF':'000000')}));
  sub.push(cell('WoW',{bold:true,fill:{color:'D9D9D9'}}),cell('MoM',{bold:true,fill:{color:'D9D9D9'}}));
  const body=[...C.ms.filter(Boolean),J].map(r=>{const isJ=r===J; const b=isJ?{bold:true}:{};
    const mx=Math.max(r.cur.tsel,r.cur.ioh,r.cur.xlplus); const sf=v=>v===mx?{fill:{color:COL.green}}:{};
    const dd=d=>d?[d.tsel,d.ioh,d.xlplus].map(v=>ncell(v,fpp,b)):[cell('–'),cell('–'),cell('–')];
    const hf=h=>h?{fill:{color:HI_FILL[h][0]},color:HI_FILL[h][1],bold:true}:{};
    return [cell(r.name,Object.assign({align:'left'},b)),cell(f2(r.cur.tsel)+'%',Object.assign(sf(r.cur.tsel),b)),cell(f2(r.cur.ioh)+'%',Object.assign(sf(r.cur.ioh),b)),cell(f2(r.cur.xlplus)+'%',Object.assign(sf(r.cur.xlplus),b)),
      cell(r.win?'Win':'Lose',{bold:true,color:r.win?'000000':COL.neg}),ncell(r.gap,v=>r2(v).toFixed(2)+'%',b),cell(r.hi,hf(r.hi)),cell(r.rank,b),...dd(r.wow),...dd(r.mom),...dd(r.yoy),cell(r.hiComp,hf(r.hiComp)),cell(r.hgW||'–',hf(r.hgW)),cell(r.hgM||'–',hf(r.hgM))];});
  const cw=[1.25,0.58,0.58,0.58,0.55,0.6,0.6,0.55,0.62,0.62,0.62,0.62,0.62,0.62,0.62,0.62,0.62,0.75,0.6,0.6];
  sl.addTable([top,sub,...body],{x:0.2,y:3.05,w:cw.reduce((a,b)=>a+b),colW:cw,rowH:0.19,border:{type:'solid',pt:0.5,color:'BFBFBF'},fontFace:'Poppins'});
  // slide 2 GA
  sl=pptx.addSlide(); const g=C.gaJ;
  const gh=g?`Gross add TSEL sebesar ${fK0(g.cur[0])} di rank ${g.rank} dengan gap ${fK0(g.gap)} ke ${g.hi} di rank 1 dengan WoW ${g.wow?fK0(g.wow[0]):'–'} dan MoM ${g.mom?fK0(g.mom[0]):'–'}. Gross Add TSEL tertinggi di ${C.gaHi.TSEL} Kota, XL+ di ${C.gaHi['XL+']} Kota, IOH di ${C.gaHi.IOH} Kota.`:`Gross add Jabar menunggu angka Meta. Gross Add TSEL tertinggi di ${C.gaHi.TSEL} Kota, XL+ di ${C.gaHi['XL+']} Kota, IOH di ${C.gaHi.IOH} Kota.`;
  header(pptx,sl,'Gross Add Performance',gh,L);
  const JGA=Object.assign({},H.jga,g?{[C.F]:{tsel:g.cur[0],ioh:g.cur[1],xlplus:g.cur[2]}}:{});
  const gd=Object.keys(JGA).map(Number).sort((a,b)=>a-b).filter(d=>d<=C.F).slice(-95);
  const gl=gd.map(d=>{const t=fromSerial(d);return `${MON[t.m-1]}-${String(t.y).slice(2)}`;});
  sl.addChart(pptx.ChartType.line,[{name:'IOH',labels:gl,values:gd.map(d=>JGA[d].ioh)},{name:'Telkomsel',labels:gl,values:gd.map(d=>JGA[d].tsel)},{name:'XL+',labels:gl,values:gd.map(d=>JGA[d].xlplus)}],
    {x:0.2,y:1.1,w:8.7,h:1.85,chartColors:['FFC000','E4002B','0070C0'],lineSize:1.5,lineDataSymbol:'none',showLegend:true,legendPos:'r',legendFontFace:'Poppins',legendFontSize:8,
     catAxisLabelFontFace:'Poppins',catAxisLabelFontSize:7,valAxisLabelFontFace:'Poppins',valAxisLabelFontSize:7,valAxisHidden:true,valGridLine:{style:'none'},catAxisLabelFrequency:9,catAxisLabelRotate:0,showTitle:true,title:'Gross Add Share Trend',titleFontFace:'Poppins',titleFontSize:9});
  const gb=[];
  if(g){const ord=[['XL+',g.cur[2]],['TSEL',g.cur[0]],['IOH',g.cur[1]]].sort((a,b)=>b[1]-a[1]);
    gb.push(`Gross add terbesar dimiliki ${ord[0][0]} ${fint(ord[0][1])}, TSEL ${fint(g.cur[0])} di peringkat ${g.rank} dengan gap ${fint(g.gap)}.`);
    if(g.wow&&g.mom){const cw2=v=>v===0?'flat':(v>0?'naik':'turun'); gb.push(`TSEL ${cw2(g.wow[0])} secara WoW dan ${cw2(g.mom[0])} secara MoM. IOH ${cw2(g.wow[1])} WoW dan ${cw2(g.mom[1])} MoM, XL+ ${cw2(g.wow[2])} WoW dan ${cw2(g.mom[2])} MoM.`);}}
  else gb.push('Baris JABAR menunggu angka Gross adds dari screenshot Meta. Tabel kota sudah terupdate.');
  sl.addText(gb.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<gb.length-1}})),{x:9.1,y:1.15,w:4.0,h:1.8,fontFace:'Poppins',fontSize:9.5,valign:'top',paraSpaceAfter:4});
  const gtop=[cell('Kabupaten',Object.assign({rowspan:2},hdrO)),cell(`Gross Add ${L.ga}`,Object.assign({colspan:7},hdrO)),cell('WOW',Object.assign({colspan:3},hdrO)),cell('MOM',Object.assign({colspan:3},hdrO))];
  const gsub=['Tsel','IOH','XL+','Win/Lose','GAP','Highest','Rank TSEL','Tsel','IOH','XL+','Tsel','IOH','XL+'].map(t=>cell(t,{bold:true,fill:{color:t==='Tsel'?'C00000':t==='IOH'?'FFFF00':t==='XL+'?'0070C0':'D9D9D9'},color:t==='IOH'?'000000':(t==='Tsel'||t==='XL+'?'FFFFFF':'000000')}));
  const grows=[...C.ga.filter(Boolean)]; if(g) grows.push(g);
  const gbody=grows.map(r=>{const b=r===g?{bold:true}:{}; const mx=Math.max(...r.cur); const sf=v=>v===mx?{fill:{color:COL.green}}:{};
    const dd=d=>d?d.map(v=>ncell(v,fint,b)):[cell('–'),cell('–'),cell('–')]; const hf={fill:{color:HI_FILL[r.hi][0]},color:HI_FILL[r.hi][1],bold:true};
    return [cell(r.name,Object.assign({align:'left'},b)),...r.cur.map(v=>cell(fint(v),Object.assign(sf(v),b))),cell(r.win?'Win':'Lose',{bold:true,color:r.win?'000000':COL.neg}),ncell(r.gap,fint,b),cell(r.hi,hf),cell(r.rank,b),...dd(r.wow),...dd(r.mom)];});
  if(!g) gbody.push([cell('JABAR',{bold:true,align:'left'}),...Array(13).fill(0).map(()=>cell('–',{bold:true}))]);
  const gcw=[1.6,0.85,0.85,0.85,0.75,0.85,0.75,0.75,0.85,0.85,0.85,0.85,0.85,0.85];
  sl.addTable([gtop,gsub,...gbody],{x:0.2,y:3.05,w:gcw.reduce((a,b)=>a+b),colW:gcw,rowH:0.19,border:{type:'solid',pt:0.5,color:'BFBFBF'},fontFace:'Poppins'});
  // slide 3/4 paths
  const capLines=C.caption.split('\n');
  const churnLine=(capLines.find(l=>l.startsWith('Secara movement'))||'').replace(/\*/g,'');
  const branchLine=(capLines.find(l=>l.startsWith('Pergerakan kompetisi'))||'Paths branch belum lengkap.').replace(/\*/g,'');
  const mtx=(sl,M,x,y,w,fs,dup)=>{ if(!M){sl.addText('Paths belum diupload',{x,y,w,h:0.5,fontFace:'Poppins',fontSize:10,color:COL.gray});return;}
    const hdr=[cell('',{fill:{color:'FFFFFF'}}),cell('',{fill:{color:'FFFFFF'}}),...PATH_OPS.map(o=>cell(PATH_LBL[o],{bold:true,fill:{color:OP_FILL[o][0]},color:OP_FILL[o][1],fontSize:fs})),cell('Move to',{bold:true,fontSize:fs}),cell('In- Out',{bold:true,fontSize:fs})];
    const rows=[[cell('',{}),cell('',{}),cell('To',{colspan:5,bold:true,fill:{color:'7F7F7F'},color:'FFFFFF',fontSize:fs}),cell('',{}),cell('',{})],hdr];
    PATH_OPS.forEach((a,ai)=>{const r=[]; if(ai===0) r.push(cell('From',{rowspan:5,fill:{color:'7F7F7F'},color:'FFFFFF',fontSize:fs}));
      r.push(cell(PATH_LBL[a],{fill:{color:OP_FILL[a][0]},color:OP_FILL[a][1],fontSize:fs}));
      PATH_OPS.forEach(o=>r.push(cell(o===a?'':fint(M.m[a+'|'+o]||0),{fontSize:fs})));
      r.push(cell(fint(M.out[a]),{bold:true,fontSize:fs}),cell(fint(M.io[a]),{bold:true,fontSize:fs,color:M.io[a]<0?COL.neg:'00A050'})); rows.push(r);});
    rows.push([cell('',{}),cell('Move In',{bold:true,fontSize:fs}),...PATH_OPS.map(o=>cell(fint(M.inn[o]),{bold:true,fontSize:fs})),cell(fint(M.tot),{bold:true,fontSize:fs}),cell('',{})]);
    const cws=[0.9,1.3,0.95,0.95,0.95,0.95,0.95,0.95,0.95].map(v=>v*w/9.85);
    sl.addTable(rows,{x,y,w,colW:cws,rowH:fs>8?0.42:0.24,border:{type:'solid',pt:0.5,color:'BFBFBF'},fontFace:'Poppins'});
    if(dup) sl.addText('File Paths identik dengan branch lain – jangan dipakai',{x,y:y-0.3,w,h:0.25,fontFace:'Poppins',fontSize:9,bold:true,color:COL.neg});};
  sl=pptx.addSlide(); header(pptx,sl,'Churn Movement (Paths)',churnLine||'Paths Jabar belum diupload.',L);
  sl.addText('Regional Jabar',{x:0.3,y:1.15,w:4,h:0.4,fontFace:'Poppins',fontSize:16,bold:true,color:COL.red});
  mtx(sl,C.paths.JABAR,1.6,1.8,10.2,12,C.dupPaths.has('JABAR'));
  sl=pptx.addSlide(); header(pptx,sl,'Churn Movement per Branch',branchLine,L);
  const pos=[[0.2,1.15],[6.8,1.15],[0.2,4.25],[6.8,4.25]];
  BRANCHES.forEach((b,i)=>{const M=C.paths[b]; const [x,y]=pos[i];
    sl.addText('Branch '+BR_NAME[b],{x,y,w:6.2,h:0.3,fontFace:'Poppins',fontSize:12,bold:true,color:COL.red,margin:0});
    const bl=capLines.find(l=>l.startsWith('- '+BR_NAME[b]+':'))||''; sl.addText(bl.replace(/^- [^:]+: /,''),{x,y:y+0.3,w:6.2,h:0.4,fontFace:'Poppins',fontSize:8.5,margin:0,valign:'top'});
    mtx(sl,M,x,y+0.78,6.2,6.5,C.dupPaths.has(b));});
  return pptx.write({outputType:'blob'});
}

window.FB={kecTable,kecCurrent,num,toSerial,csvRows:t=>parseCSV(t).slice(2).filter(r=>r[0]&&r[0]!=='date'),compute,ingest,caption,nobarCity,nobarRegion,nobarKec,excelSheets,buildXlsx,buildPptx,mergeWeek,labels,H,S,K,ORDER,UP,PATH_OPS,PATH_LBL,BR_NAME,BRANCHES,fromSerial,mPrev,r2,f2,fpp,fint,fK1,sgnPct,MON};
})();

;
(function(){
'use strict';
const FB=window.FB, $=s=>document.querySelector(s), r2=FB.r2;
const el=(t,a={},...kids)=>{const e=document.createElement(t);for(const[k,v]of Object.entries(a)){if(k==='class')e.className=v;else if(k==='text')e.textContent=v;else e.setAttribute(k,v);}kids.flat().forEach(k=>e.append(k instanceof Node?k:document.createTextNode(String(k))));return e;};
const NS='http://www.w3.org/2000/svg'; const mk=(t,a)=>{const e=document.createElementNS(NS,t);for(const[k,v]of Object.entries(a))e.setAttribute(k,v);return e;};
const fmtD=s=>{const t=FB.fromSerial(s);return `${t.d} ${FB.MON[t.m-1]}`;};
const sc=v=>v==null?'':(r2(v)<0?'neg':(r2(v)>0?'pos':''));
let C=null, selCity=0;

/* ---------- series helpers (history + current week) ---------- */
function jSeries(){const o={}; for(const [d,v] of Object.entries(FB.H.jms)) if(+d<C.F) o[d]=v; o[C.F]={tsel:C.msJ.cur.tsel,ioh:C.msJ.cur.ioh,xlplus:C.msJ.cur.xlplus}; return o;}
function cSeries(i){const o={}; for(const [d,a] of Object.entries(FB.H.ms)){ if(+d>=C.F||!a[i]) continue; o[d]={tsel:a[i][0],ioh:a[i][6],xlplus:a[i][5]}; }
  const r=C.ms[i]; if(r) o[C.F]={tsel:r.cur.tsel,ioh:r.cur.ioh,xlplus:r.cur.xlplus}; return o;}
function gSeries(i){const o={}; const z=v=>v?v:null; for(const [d,a] of Object.entries(FB.H.ga)){ if(+d>=C.F||!a[i]) continue; o[d]={tsel:z(a[i][0]),ioh:z(a[i][6]),xlplus:z(a[i][5])}; }
  const r=C.ga[i]; if(r) o[C.F]={tsel:r.cur[0],ioh:r.cur[1],xlplus:r.cur[2]}; return o;}
function jgSeries(){const o={}; for(const [d,v] of Object.entries(FB.H.jga)) if(+d<C.F) o[d]=v; if(C.gaJ) o[C.F]={tsel:C.gaJ.cur[0],ioh:C.gaJ.cur[1],xlplus:C.gaJ.cur[2]}; return o;}
const dates=s=>Object.keys(s).map(Number).sort((a,b)=>a-b);
function wowList(s,k){const d=dates(s); const out=[]; for(let i=1;i<d.length;i++) if(d[i]-d[i-1]===7) out.push([d[i],r2(s[d[i]][k]-s[d[i-1]][k])]); return out;}
function streak(w){let n=0; for(let i=w.length-1;i>=0&&w[i][1]<0;i--) n++; return n;}
const gapOf=v=>r2(v.tsel-Math.max(v.ioh,v.xlplus));
const winOf=v=>v.tsel>=Math.max(v.ioh,v.xlplus);

/* ---------- insights ---------- */
function insights(){
  const out=[]; const js=jSeries(); const w=wowList(js,'tsel'); const last8=w.slice(-8); const neg8=last8.filter(x=>x[1]<0).length;
  const J=C.msJ; const st=streak(w);
  out.push({k:'share',t:`Telkomsel share ${J.wow&&r2(J.wow.tsel)<0?'fell':'rose'} ${Math.abs(r2(J.wow?J.wow.tsel:0)).toFixed(2)}pp WoW to ${FB.f2(J.cur.tsel)}%. It went down in ${neg8} of the last ${last8.length} weeks${st>=2?`, and ${st} weeks in a row now`:''}.`,cls:neg8>=5?'neg':''});
  const d=dates(js); const d4=d.filter(x=>x<=C.F-28).pop();
  if(d4!=null){const g0=gapOf(js[d4]), g1=gapOf(js[C.F]); out.push({k:'gap',t:`Gap to ${J.to} is ${FB.sgnPct(g1)}, ${Math.abs(r2(g1-g0)).toFixed(2)}pp ${g1<g0?'wider':'narrower'} than four weeks ago (${FB.sgnPct(g0)} on ${fmtD(d4)}).`,cls:g1<g0?'neg':'pos'});}
  const prev=FB.H.ms[C.X]; if(prev){ const won=[],lost=[]; C.ms.forEach((r,i)=>{ if(!r||!prev[i]) return; const pw=prev[i][0]>=Math.max(prev[i][6],prev[i][5]); if(r.win&&!pw) won.push(r.name); if(!r.win&&pw) lost.push(r.name);});
    const pn=prev.filter(a=>a&&a[0]>=Math.max(a[6],a[5])).length;
    out.push({k:'win',t:`${C.nWin} of 18 cities won, against ${pn} last week.${won.length?' New wins: '+won.join(', ')+'.':''}${lost.length?' Lost: '+lost.join(', ')+'.':''}${!won.length&&!lost.length?' No city changed side.':''}`,cls:C.nWin<pn?'neg':(C.nWin>pn?'pos':'')}); }
  const P4=C.F-28; const mv=C.ms.map((r,i)=>{const a=FB.H.ms[P4]; return r&&a&&a[i]?{n:r.name,v:r2(r.cur.tsel-a[i][0])}:null;}).filter(Boolean).sort((a,b)=>a.v-b.v);
  if(mv.length){const lo=mv[0], hi=mv[mv.length-1]; out.push({k:'move',t:`Over four weeks the largest Telkomsel drop is ${lo.n} (${FB.fpp(lo.v)}) and the largest gain is ${hi.n} (${FB.fpp(hi.v)}). ${mv.filter(x=>x.v<0).length} of ${mv.length} cities are down.`,cls:''});}
  if(C.gaJ){const g=C.gaJ.cur, s3=g[0]/(g[0]+g[1]+g[2])*100, b3=J.cur.tsel/(J.cur.tsel+J.cur.ioh+J.cur.xlplus)*100;
    out.push({k:'ga',t:`Telkomsel takes ${s3.toFixed(1)}% of gross adds among the three, below its ${b3.toFixed(1)}% share of the base. If that holds, base share keeps sliding. This is a pattern, not proof of cause.`,cls:s3<b3?'neg':'pos'});}
  const PJ=C.paths.JABAR; if(PJ&&!C.dupPaths.has('JABAR')){ const comp=FB.PATH_OPS.filter(o=>o!=='Telkomsel').map(o=>({o,net:(PJ.m[o+'|Telkomsel']||0)-(PJ.m['Telkomsel|'+o]||0)})).sort((a,b)=>a.net-b.net);
    out.push({k:'churn',t:`Net churn for Telkomsel is ${FB.fK1(PJ.io.Telkomsel)}. The biggest net loss is to ${FB.PATH_LBL[comp[0].o]} (${FB.fK1(comp[0].net)}); ${FB.PATH_LBL[PJ.gainMax]} gains most overall (${FB.fK1(PJ.io[PJ.gainMax])}).`,cls:PJ.io.Telkomsel<0?'neg':'pos'}); }
  return out;
}
function renderInsights(){ const box=$('#insights'); box.textContent=''; box.append(el('h2',{},'What changed this week'));
  const ul=el('ul',{}); insights().forEach(x=>ul.append(el('li',{class:'ins'+(x.cls?' ins-'+x.cls:''),'data-k':x.k},x.t))); box.append(ul); }

/* ---------- gap bars (market share tab) ---------- */
function gapBars(){ const host=$('#gap-bars'); host.textContent=''; const prev=FB.H.ms[C.X];
  const rows=C.ms.map((r,i)=>r?{n:r.name,g:r.gap,to:r.to,p:prev&&prev[i]?r2(prev[i][0]-Math.max(prev[i][6],prev[i][5])):null}:null).filter(Boolean).sort((a,b)=>b.g-a.g);
  const W=Math.max(320,host.clientWidth||600), rh=20, m={l:128,r:56,t:30,b:8}, Hh=m.t+rows.length*rh+m.b;
  const ext=Math.max(...rows.map(r=>Math.abs(r.g)),...rows.map(r=>Math.abs(r.p||0)))*1.05;
  const x=v=>m.l+(W-m.l-m.r)*(v+ext)/(2*ext);
  const svg=mk('svg',{viewBox:`0 0 ${W} ${Hh}`,width:'100%',role:'img','aria-label':'Gap to leading competitor by city'});
  const t=mk('text',{x:8,y:16,class:'ct'}); t.textContent='Gap to the leading competitor by city (dot = last week)'; svg.append(t);
  svg.append(mk('line',{x1:x(0),x2:x(0),y1:m.t-4,y2:Hh-m.b,stroke:'var(--ink2)'}));
  rows.forEach((r,i)=>{const y=m.t+i*rh; const lb=mk('text',{x:m.l-8,y:y+14,class:'ax','text-anchor':'end'}); lb.textContent=r.n; svg.append(lb);
    svg.append(mk('rect',{x:Math.min(x(0),x(r.g)),y:y+4,width:Math.abs(x(r.g)-x(0)),height:rh-8,fill:r.g>=0?'var(--pos)':'var(--neg)',opacity:.85,rx:2}));
    if(r.p!=null) svg.append(mk('circle',{cx:x(r.p),cy:y+rh/2,r:3.2,fill:'var(--surface)',stroke:'var(--ink)','stroke-width':1.4}));
    const v=mk('text',{x:r.g>=0?x(r.g)+4:x(r.g)-4,y:y+14,class:'ax','text-anchor':r.g>=0?'start':'end'}); v.textContent=FB.sgnPct(r.g)+' '+r.to; svg.append(v);});
  host.append(svg);}

/* ---------- heatmap ---------- */
function heatmap(){ const host=$('#heat'); host.textContent=''; const N=13;
  const all=[...Array(N).keys()].map(k=>C.F-7*(N-1-k));
  const t=el('table',{class:'grid heat'}); t.append(el('thead',{},el('tr',{},el('th',{class:'l'},'City'),...all.map(d=>el('th',{},fmtD(d))),el('th',{},'Down streak'),el('th',{},'13-week change'))));
  const tb=el('tbody'); const rows=[['Jabar',jSeries()],...FB.ORDER.map((n,i)=>[n,cSeries(i)])];
  for(const [n,s] of rows){ const tr=el('tr',{class:n==='Jabar'?'total':''},el('td',{class:'l'},n));
    all.forEach(d=>{const a=s[d], b=s[d-7]; const v=a&&b?r2(a.tsel-b.tsel):null; const c=el('td',{class:'hc'},v==null?'–':v.toFixed(2));
      if(v!=null){const k=Math.min(1,Math.abs(v)/0.3); c.style.background=v<0?`rgba(208,2,27,${0.12+0.6*k})`:(v>0?`rgba(14,138,75,${0.12+0.6*k})`:'transparent'); if(k>0.6)c.style.color='#fff';}
      c.title=`${n}, week to ${fmtD(d)}: ${v==null?'no data':v.toFixed(2)+'pp'}`; tr.append(c);});
    const w=wowList(s,'tsel'); const st=streak(w); const d0=s[all[0]-7]||s[all[0]], dn=s[C.F]; const ch=d0&&dn?r2(dn.tsel-d0.tsel):null;
    tr.append(el('td',{class:st>=3?'neg b':''},st?st+' wk':'–'),el('td',{class:sc(ch)+' b'},ch==null?'–':FB.fpp(ch))); tb.append(tr);}
  t.append(tb); host.append(el('div',{class:'scroll'},t));}

/* ---------- city explorer ---------- */
function cityView(){ const sel=$('#city-sel'); if(!sel.options.length){ FB.ORDER.forEach((n,i)=>sel.append(el('option',{value:i},n))); sel.addEventListener('change',()=>{selCity=+sel.value;cityView(); if(window.__kecTables) window.__kecTables();}); }
  sel.value=selCity; const i=selCity, r=C.ms[i], g=C.ga[i]; const s=cSeries(i); const w=wowList(s,'tsel'); const box=$('#city-kpi'); box.textContent='';
  if(!r){box.append(el('p',{class:'muted'},'This city is missing from the City CSV.'));return;}
  const ds=dates(s); const d13=ds.filter(d=>d<=C.F-91).pop(); const ch13=d13?r2(r.cur.tsel-s[d13].tsel):null;
  const k=(t,v,sub,cls)=>el('div',{class:'kpi'},el('div',{class:'kpi-t'},t),el('div',{class:'kpi-v '+(cls||'')},v),el('div',{class:'kpi-s'},sub));
  box.append(k('Telkomsel share',FB.f2(r.cur.tsel)+'%',`WoW ${r.wow?FB.fpp(r.wow.tsel):'–'}  ·  MoM ${r.mom?FB.fpp(r.mom.tsel):'–'}  ·  YoY ${r.yoy?FB.fpp(r.yoy.tsel):'–'}`));
  box.append(k(`Gap to ${r.to}`,FB.sgnPct(r.gap),`${r.win?'Win':'Lose'}  ·  rank ${r.rank} of 3`,sc(r.gap)));
  box.append(k('13-week change',ch13==null?'–':FB.fpp(ch13),`Down ${w.slice(-13).filter(x=>x[1]<0).length} of the last 13 weeks; streak ${streak(w)} wk`,sc(ch13)));
  if(g) box.append(k('Gross adds, Telkomsel',FB.fint(g.cur[0]),`${(g.cur[0]/(g.cur[0]+g.cur[1]+g.cur[2])*100).toFixed(1)}% of the three  ·  highest: ${g.hi}`));
  window.__lineChart($('#city-ms'),s,`${FB.ORDER[i]}: market share (%)`,v=>v.toFixed(2)+'%');
  window.__lineChart($('#city-ga'),gSeries(i),`${FB.ORDER[i]}: gross adds`,v=>Math.round(v).toLocaleString('en-US'));}

/* ---------- acquisition vs base (gross adds tab) ---------- */
function acqScatter(){ const host=$('#acq'); host.textContent='';
  const pts=C.ms.map((r,i)=>{const g=C.ga[i]; if(!r||!g) return null; const gs=g.cur[0]/(g.cur[0]+g.cur[1]+g.cur[2])*100, bs=r.cur.tsel/(r.cur.tsel+r.cur.ioh+r.cur.xlplus)*100; return {n:r.name,x:r.gap,y:r2(gs-bs),gs,bs};}).filter(Boolean);
  const W=Math.max(320,host.clientWidth||600), Hh=340, m={l:48,r:16,t:30,b:36};
  const xe=Math.max(...pts.map(p=>Math.abs(p.x)))*1.1, ye=Math.max(...pts.map(p=>Math.abs(p.y)))*1.15||1;
  const x=v=>m.l+(W-m.l-m.r)*(v+xe)/(2*xe), y=v=>m.t+(Hh-m.t-m.b)*(1-(v+ye)/(2*ye));
  const svg=mk('svg',{viewBox:`0 0 ${W} ${Hh}`,width:'100%',role:'img','aria-label':'Acquisition share versus base share by city'});
  const t=mk('text',{x:8,y:16,class:'ct'}); t.textContent='Acquisition vs base: gross-add share minus base share (pp), against gap to leader'; svg.append(t);
  svg.append(mk('line',{x1:x(0),x2:x(0),y1:m.t,y2:Hh-m.b,class:'gl'}),mk('line',{x1:m.l,x2:W-m.r,y1:y(0),y2:y(0),class:'gl'}));
  const q=(tx,xx,yy,a)=>{const e=mk('text',{x:xx,y:yy,class:'ax','text-anchor':a}); e.textContent=tx; svg.append(e);};
  q('Losing, but acquiring above base',m.l+4,m.t+12,'start'); q('Winning and acquiring above base',W-m.r-4,m.t+12,'end');
  q('Losing and acquiring below base',m.l+4,Hh-m.b-6,'start'); q('Winning, but acquiring below base',W-m.r-4,Hh-m.b-6,'end');
  const xl=mk('text',{x:(m.l+W-m.r)/2,y:Hh-8,class:'ax','text-anchor':'middle'}); xl.textContent='Gap to leading competitor, market share (pp)'; svg.append(xl);
  const tip=el('div',{class:'tip',hidden:''}); host.style.position='relative';
  pts.forEach(p=>{const c=mk('circle',{cx:x(p.x),cy:y(p.y),r:6,fill:p.y>=0?'var(--pos)':'var(--neg)',opacity:.8,stroke:'var(--surface)','stroke-width':1.5,tabindex:0});
    const show=()=>{tip.hidden=false; tip.textContent=''; tip.append(el('b',{},p.n),el('div',{},`Gap ${FB.sgnPct(p.x)}`),el('div',{},`Gross-add share ${p.gs.toFixed(1)}%`),el('div',{},`Base share ${p.bs.toFixed(1)}%`)); const rr=svg.getBoundingClientRect(); tip.style.left=Math.min(rr.width-150,x(p.x)*rr.width/W+10)+'px'; tip.style.top=(y(p.y)*rr.width/W)+'px';};
    c.addEventListener('pointerenter',show); c.addEventListener('focus',show); c.addEventListener('pointerleave',()=>tip.hidden=true); c.addEventListener('blur',()=>tip.hidden=true); svg.append(c);
    if(Math.abs(p.y)>ye*0.45||Math.abs(p.x)>xe*0.6){const l=mk('text',{x:x(p.x)+8,y:y(p.y)+4,class:'ax'}); l.textContent=p.n; svg.append(l);}});
  host.append(svg,tip,el('p',{class:'note'},'Shares are among Telkomsel, IOH and XL+ only. A city above zero is winning a bigger slice of new customers than of the base. Correlational: it flags where to look, it does not say why.'));}

/* ---------- churn flows ---------- */
function flows(){ const host=$('#flows'); host.textContent='';
  for(const b of ['JABAR',...FB.BRANCHES]){ const M=C.paths[b]; if(!M||C.dupPaths.has(b)) continue;
    const comp=FB.PATH_OPS.filter(o=>o!=='Telkomsel'); const mx=Math.max(...comp.map(o=>Math.max(M.m['Telkomsel|'+o]||0,M.m[o+'|Telkomsel']||0)))||1;
    const card=el('div',{class:'flow'},el('h3',{},b==='JABAR'?'Jabar':FB.BR_NAME[b]));
    comp.forEach(o=>{const out=M.m['Telkomsel|'+o]||0, inn=M.m[o+'|Telkomsel']||0, net=inn-out;
      card.append(el('div',{class:'frow'},el('span',{class:'fl'},FB.PATH_LBL[o]),
        el('span',{class:'fbar out'},el('i',{style:`width:${out/mx*100}%`})),el('span',{class:'fbar in'},el('i',{style:`width:${inn/mx*100}%`})),el('b',{class:sc(net)},FB.fK1(net))));});
    card.append(el('p',{class:'note'},`Left bar: Telkomsel customers lost to them. Right bar: won from them. Net ${FB.fK1(M.io.Telkomsel)}.`)); host.append(card);}
  if(!host.children.length) host.append(el('p',{class:'muted'},'Upload valid Paths files to see flows.'));}

window.__perfRender=function(c){C=c; renderInsights(); gapBars(); heatmap(); cityView(); acqScatter(); flows();};
window.__perfResize=function(){ if(!C) return; gapBars(); cityView(); acqScatter(); };
})();

;
(function(){
'use strict';
const FB=window.FB, $=s=>document.querySelector(s), r2=FB.r2, MAP=window.__MAP__;
const el=(t,a={},...kids)=>{const e=document.createElement(t);for(const[k,v]of Object.entries(a)){if(k==='class')e.className=v;else if(k==='text')e.textContent=v;else e.setAttribute(k,v);}kids.flat().forEach(k=>e.append(k instanceof Node?k:document.createTextNode(String(k))));return e;};
const NS='http://www.w3.org/2000/svg'; const mk=(t,a)=>{const e=document.createElementNS(NS,t);for(const[k,v]of Object.entries(a))e.setAttribute(k,v);return e;};
const sc=v=>v==null?'':(r2(v)<0?'neg':(r2(v)>0?'pos':''));
let C=null;
/* ---------- map ---------- */
const imgs={}; function loadImg(k,src){return imgs[k]||(imgs[k]=new Promise(res=>{const i=new Image(); i.onload=()=>res(i); i.src='data:image/png;base64,'+src;}));}
let maskData=null;
let mapGen=0;
async function drawMap(){
  const host=$('#map'); if(!host) return; const gen=++mapGen;
  const W=MAP.w,Hh=MAP.h; const [base,mask]=await Promise.all([loadImg('b',MAP.base),loadImg('m',MAP.mask)]);
  if(gen!==mapGen) return; /* a newer draw started while images loaded: let it win */
  host.textContent='';
  const dw=host.clientWidth||W; const S=Math.min(4,Math.max(2,Math.ceil(dw*(window.devicePixelRatio||1)/W))); const cv=el('canvas',{width:W*S,height:Hh*S,role:'img','aria-label':'Map of Jawa Barat coloured by the operator with the highest FB share in each city'}); const ctx=cv.getContext('2d');
  ctx.imageSmoothingEnabled=true; ctx.imageSmoothingQuality='high'; ctx.drawImage(base,0,0,W*S,Hh*S);
  const mc=document.createElement('canvas'); mc.width=W*S; mc.height=Hh*S; const mx=mc.getContext('2d'); mx.imageSmoothingEnabled=false; mx.drawImage(mask,0,0,W*S,Hh*S);
  const md=mx.getImageData(0,0,W*S,Hh*S).data; maskData={d:md,w:W*S,h:Hh*S};
  const id=ctx.getImageData(0,0,W*S,Hh*S); const p=id.data; const pal=MAP.pal;
  const win=C.ms.map(r=>r?r.hi:null);
  for(let i=0;i<md.length;i+=4){ const c=md[i]; if(!c||md[i+3]<128) continue; const w=win[c-1]; if(!w) continue; const col=pal[w]; const f=md[i+1]/128;
    p[i]=Math.min(255,col[0]*f); p[i+1]=Math.min(255,col[1]*f); p[i+2]=Math.min(255,col[2]*f); }
  ctx.putImageData(id,0,0); cv.style.width='100%'; cv.style.height='auto'; cv.style.display='block';
  const tip=el('div',{class:'tip',hidden:''}); host.style.position='relative'; host.append(cv,tip);
  const cnt={TSEL:0,IOH:0,'XL+':0}; win.forEach(w=>{if(w)cnt[w]++;});
  const rgb=a=>`rgb(${a.map(Math.round).join(',')})`;
  host.append(el('div',{class:'legend'},...[['TSEL','Telkomsel'],['XL+','XL+'],['IOH','IOH']].map(([k,l])=>el('span',{},el('i',{style:`background:${rgb(pal[k])};height:10px;width:14px`}),`${l} highest: ${cnt[k]} cities`))));
  const at=ev=>{const r=cv.getBoundingClientRect(); const x=Math.floor((ev.clientX-r.left)/r.width*maskData.w), y=Math.floor((ev.clientY-r.top)/r.height*maskData.h); if(x<0||y<0||x>=maskData.w||y>=maskData.h) return 0; const k=(y*maskData.w+x)*4; return maskData.d[k+3]>128?maskData.d[k]:0;};
  cv.addEventListener('pointermove',ev=>{const c=at(ev); if(!c){tip.hidden=true;cv.style.cursor='default';return;} const row=C.ms[c-1]; if(!row){tip.hidden=true;return;} cv.style.cursor='pointer';
    tip.hidden=false; tip.textContent=''; tip.append(el('b',{},row.name),el('div',{},`Telkomsel ${FB.f2(row.cur.tsel)}%  ·  IOH ${FB.f2(row.cur.ioh)}%  ·  XL+ ${FB.f2(row.cur.xlplus)}%`),el('div',{class:sc(row.gap)},`${row.win?'Win':'Lose'}, gap to ${row.to} ${FB.sgnPct(row.gap)}`),el('div',{class:'muted'},'Click to open the city'));
    const r=host.getBoundingClientRect(); tip.style.left=Math.min(r.width-220,ev.clientX-r.left+12)+'px'; tip.style.top=(ev.clientY-r.top+12)+'px';});
  cv.addEventListener('pointerleave',()=>tip.hidden=true);
  if(window.__renderOverview) requestAnimationFrame(()=>window.__renderOverview());
  cv.addEventListener('click',ev=>{const c=at(ev); if(!c) return; const sel=$('#city-sel'); if(sel){ sel.value=c-1; sel.dispatchEvent(new Event('change')); } location.hash='#city';});
}
/* ---------- kecamatan tables ---------- */
const HIC={TSEL:'hi-tsel',IOH:'hi-ioh','XL+':'hi-xl'};
const th=(t,c)=>el('th',{class:c||''},t), td=(t,c)=>el('td',{class:c||''},t);
function kecTables(){
  const i=+($('#city-sel').value||0); const name=FB.ORDER[i]; const K=FB.kecTable(C,name); const box=$('#city-kec'); box.textContent='';
  if(!K.hasCur){ box.append(el('p',{class:'muted'},'Kecamatan data for this week is not loaded yet. Upload Kec_Latest in Update data.')); return; }
  const city=C.ms[i], cga=C.ga[i];
  box.append(el('h3',{class:'sub'},`Kecamatan market share · ${name} (${K.ms.length} kecamatan)`));
  const t=el('table',{class:'grid'}); t.append(el('thead',{},el('tr',{},th('Kecamatan','l'),th('Tsel','op-t'),th('IOH','op-i'),th('XL+','op-x'),th('Win/Lose'),th('Gap'),th('Highest'),th('Rank'),th('WoW Tsel'),th('WoW IOH'),th('WoW XL+'),th('MoM Tsel'),th('MoM IOH'),th('MoM XL+'),th('YoY Tsel'),th('YoY IOH'),th('YoY XL+'))));
  const tb=el('tbody'); const rows=[...K.ms].sort((a,b)=>a.name.localeCompare(b.name)); if(city) rows.push(Object.assign({},city,{name:name+' (city)',isTotal:true}));
  for(const r of rows){ const mx=Math.max(r.cur.tsel,r.cur.ioh,r.cur.xlplus); const s=v=>td(FB.f2(v)+'%',v===mx?'max':''); const d=x=>x?[x.tsel,x.ioh,x.xlplus].map(v=>td(FB.fpp(v),sc(v))):[td('–'),td('–'),td('–')];
    tb.append(el('tr',{class:r.isTotal?'total':''},td(r.name,'l'),s(r.cur.tsel),s(r.cur.ioh),s(r.cur.xlplus),td(r.win?'Win':'Lose',r.win?'win':'lose'),td(r2(r.gap).toFixed(2)+'%',sc(r.gap)),td(r.hi,HIC[r.hi]),td(r.rank),...d(r.wow),...d(r.mom),...d(r.yoy))); }
  t.append(tb); box.append(el('div',{class:'scroll'},t));
  const nw=K.ms.filter(r=>r.win).length; box.append(el('p',{class:'note'},`Telkomsel wins ${nw} of ${K.ms.length} kecamatan in ${name}.`));
  box.append(el('h3',{class:'sub'},`Kecamatan gross adds · ${name}`));
  const g=el('table',{class:'grid'}); g.append(el('thead',{},el('tr',{},th('Kecamatan','l'),th('Tsel','op-t'),th('IOH','op-i'),th('XL+','op-x'),th('Win/Lose'),th('Gap'),th('Highest'),th('Rank'),th('WoW Tsel'),th('WoW IOH'),th('WoW XL+'),th('MoM Tsel'),th('MoM IOH'),th('MoM XL+'))));
  const gb=el('tbody'); const grs=[...K.ga].sort((a,b)=>a.name.localeCompare(b.name)); if(cga) grs.push(Object.assign({},cga,{name:name+' (city)',isTotal:true}));
  for(const r of grs){ const mx=Math.max(...r.cur); const d=x=>x?x.map(v=>td(FB.fint(v),sc(v))):[td('–'),td('–'),td('–')];
    gb.append(el('tr',{class:r.isTotal?'total':''},td(r.name,'l'),...r.cur.map(v=>td(FB.fint(v),v===mx&&mx>0?'max':'')),td(r.win?'Win':'Lose',r.win?'win':'lose'),td(FB.fint(r.gap),sc(r.gap)),td(r.hi,HIC[r.hi]),td(r.rank),...d(r.wow),...d(r.mom))); }
  g.append(gb); box.append(el('div',{class:'scroll'},g));
  box.append(el('p',{class:'note'},'Meta rounds small kecamatan gross adds; values under 100 often show as 0, so ties and zero rows are common here.'));
}
/* ---------- sankey ---------- */
const OPS=['Telkomsel','XLSmart','Indosat IM3','Tri Indonesia','Smartfren'];
const LBL={'Telkomsel':'TSEL','XLSmart':'XLSmart','Indosat IM3':'IM3','Tri Indonesia':'Tri','Smartfren':'Smartfren'};
const COLR={'Telkomsel':'#E4002B','XLSmart':'#2F6DB5','Indosat IM3':'#E8B923','Tri Indonesia':'#2E9E4F','Smartfren':'#D9458A'};
const rb=v=>(v/1000).toLocaleString('id-ID',{minimumFractionDigits:1,maximumFractionDigits:1});
let sankeyBranch='JABAR';
function sankey(){
  const host=$('#sankey'); if(!host) return; host.textContent='';
  const avail=['JABAR',...FB.BRANCHES].filter(b=>C.paths[b]&&!C.dupPaths.has(b));
  if(!avail.length){host.append(el('p',{class:'muted'},'Upload the Paths CSVs to see the flow.'));return;}
  if(!avail.includes(sankeyBranch)) sankeyBranch=avail[0];
  const sw=el('div',{class:'seg',role:'tablist'}); avail.forEach(b=>{const bt=el('button',{type:'button',class:'sec'+(b===sankeyBranch?' on':''),'aria-pressed':String(b===sankeyBranch)},b==='JABAR'?'Jabar':FB.BR_NAME[b]); bt.addEventListener('click',()=>{sankeyBranch=b;sankey();}); sw.append(bt);});
  const M=C.paths[sankeyBranch]; const m=(a,b)=>M.m[a+'|'+b]||0;
  const W=Math.max(340,host.clientWidth||700), Hh=420, top=60, bot=60, nodeW=14, gap=14, padL=W<520?86:150, padR=W<520?86:150;
  const tot=OPS.reduce((s,o)=>s+M.out[o],0); const scale=(Hh-top-bot-gap*(OPS.length-1))/tot;
  const L={},R={}; let y=top; for(const o of OPS){ L[o]={y,h:M.out[o]*scale}; y+=M.out[o]*scale+gap; }
  y=top; for(const o of OPS){ R[o]={y,h:M.inn[o]*scale}; y+=M.inn[o]*scale+gap; }
  const svg=mk('svg',{viewBox:`0 0 ${W} ${Hh}`,width:'100%',role:'img','aria-label':'Sankey of customers moving between operators'});
  const t1=mk('text',{x:14,y:26,class:'sk-title'}); t1.textContent=`SANKEY PERPINDAHAN PELANGGAN (rb) · ${sankeyBranch==='JABAR'?'JABAR':FB.BR_NAME[sankeyBranch].toUpperCase()}`; svg.append(t1);
  const t2=mk('text',{x:W-14,y:26,class:'sk-date','text-anchor':'end'}); t2.textContent=`Meta ${C.L.d} ${['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'][C.L.m-1]}`; svg.append(t2);
  const h1=mk('text',{x:padL+nodeW/2,y:top-10,class:'sk-hd','text-anchor':'middle'}); h1.textContent='KELUAR (churn)'; svg.append(h1);
  const h2=mk('text',{x:W-padR-nodeW/2,y:top-10,class:'sk-hd','text-anchor':'middle'}); h2.textContent='MASUK (port-in)'; svg.append(h2);
  const lo={},ro={}; OPS.forEach(o=>{lo[o]=L[o].y; ro[o]=R[o].y;});
  const x0=padL+nodeW, x1=W-padR-nodeW; const links=[]; const labels=[];
  for(const a of OPS) for(const b of OPS){ if(a===b) continue; const v=m(a,b); if(!v) continue; const h=v*scale; links.push({a,b,v,y0:lo[a],y1:ro[b],h}); lo[a]+=h; ro[b]+=h; }
  const g=mk('g',{}); svg.append(g);
  const topLinks=new Set(links.filter(l=>l.a==='Telkomsel'||l.b==='Telkomsel').sort((p,q)=>q.v-p.v).slice(0,4));
  links.sort((p,q)=>(p.a==='Telkomsel'||p.b==='Telkomsel')-(q.a==='Telkomsel'||q.b==='Telkomsel'));
  for(const l of links){ const mid=(x0+x1)/2; const d=`M${x0},${l.y0} C${mid},${l.y0} ${mid},${l.y1} ${x1},${l.y1} L${x1},${l.y1+l.h} C${mid},${l.y1+l.h} ${mid},${l.y0+l.h} ${x0},${l.y0+l.h} Z`;
    const tsel=l.a==='Telkomsel'||l.b==='Telkomsel'; const p=mk('path',{d,fill:COLR[l.a],'fill-opacity':tsel?0.55:0.18,class:'sk-link'});
    const tt=mk('title',{}); tt.textContent=`${LBL[l.a]} → ${LBL[l.b]}: ${rb(l.v)} rb`; p.append(tt); g.append(p);
    if(topLinks.has(l)&&l.h>12){ const tx=mk('text',{x:l.a==='Telkomsel'?x1-6:x0+6,y:(l.a==='Telkomsel'?l.y1:l.y0)+l.h/2+4,class:'sk-val','text-anchor':l.a==='Telkomsel'?'end':'start'}); tx.textContent=rb(l.v); labels.push(tx);} }
  for(const o of OPS){ svg.append(mk('rect',{x:padL,y:L[o].y,width:nodeW,height:Math.max(2,L[o].h),fill:COLR[o],rx:2}));
    svg.append(mk('rect',{x:W-padR-nodeW,y:R[o].y,width:nodeW,height:Math.max(2,R[o].h),fill:COLR[o],rx:2}));
    const a=mk('text',{x:padL-8,y:L[o].y+L[o].h/2+5,class:'sk-node','text-anchor':'end'}); a.append(Object.assign(document.createElementNS(NS,'tspan'),{textContent:LBL[o]+' '})); const s1=document.createElementNS(NS,'tspan'); s1.setAttribute('class','sk-num'); s1.textContent=rb(M.out[o]); a.append(s1); svg.append(a);
    const b=mk('text',{x:W-padR+8,y:R[o].y+R[o].h/2+5,class:'sk-node'}); b.append(Object.assign(document.createElementNS(NS,'tspan'),{textContent:LBL[o]+' '})); const s2=document.createElementNS(NS,'tspan'); s2.setAttribute('class','sk-num'); s2.textContent=rb(M.inn[o]); b.append(s2); svg.append(b); }
  labels.forEach(t=>svg.append(t));
  svg.append(mk('line',{x1:14,x2:W-14,y1:Hh-42,y2:Hh-42,class:'sk-rule'}));
  const net=M.io.Telkomsel, wr=M.inn.Telkomsel/(M.inn.Telkomsel+M.out.Telkomsel)*100, best=M.gainMax;
  const f=(x,anchor,parts)=>{const t=mk('text',{x,y:Hh-16,class:'sk-foot','text-anchor':anchor}); parts.forEach(([s,c])=>{const sp=document.createElementNS(NS,'tspan'); if(c) sp.setAttribute('class',c); sp.textContent=s; t.append(sp);}); svg.append(t);};
  const sg=v=>(v>0?'+':v<0?'−':'')+rb(Math.abs(v))+' rb';
  f(14,'start',[['Net TSEL ',''],[sg(net),net<0?'sk-neg':'sk-pos']]); f(W/2,'middle',[['Win rate ',''],[wr.toLocaleString('id-ID',{maximumFractionDigits:1})+'%','sk-b']]);
  f(W-14,'end',[['Net terbesar ',''],[`${LBL[best]} ${sg(M.io[best])}`,M.io[best]>=0?'sk-pos':'sk-neg']]);
  host.append(sw,el('div',{class:'sk-card'},svg),el('p',{class:'note'},'Win rate = customers won by Telkomsel ÷ (won + lost). Bands are coloured by the operator customers leave; bands touching Telkomsel are emphasised. Values in thousands (rb).'));
}
window.__vizRender=function(c){C=c; drawMap(); sankey(); if($('#city-sel')&&$('#city-sel').options.length) kecTables();};
let lastMW=0; window.__vizResize=function(){ if(C){ sankey(); const w=($('#map')||{}).clientWidth||0; if(Math.abs(w-lastMW)>40){lastMW=w; drawMap();} } };
window.__kecTables=()=>{ if(C) kecTables(); };
})();

;
/* Master workbook updater: runs in the browser on last week's workbooks + this week's CSVs. */
(function(){
'use strict';
const FB=window.FB;
const te=new TextEncoder(), td=new TextDecoder();
/* ---------- xlsb record helpers ---------- */
function rd(b,i){let t=b[i++]; if(t&0x80){t=(t&0x7f)|(b[i++]<<7);} let sz=0,sh=0; for(let k=0;k<4;k++){const c=b[i++]; sz|=(c&0x7f)<<sh; sh+=7; if(!(c&0x80))break;} return [t,sz>>>0,i];}
function* recs(b){let i=0; const n=b.length; while(i<n){const [t,sz,j]=rd(b,i); yield [t,j,j+sz,i]; i=j+sz;}}
function enc(t,p){const h=[]; if(t<128)h.push(t); else h.push((t&0x7f)|0x80,t>>7); let n=p.length; for(;;){const c=n&0x7f; n>>>=7; if(n){h.push(c|0x80);} else {h.push(c);break;}} const o=new Uint8Array(h.length+p.length); o.set(h,0); o.set(p,h.length); return o;}
const dv=(b,o)=>new DataView(b.buffer,b.byteOffset+o);
const u32=(b,o)=>(b[o]|(b[o+1]<<8)|(b[o+2]<<16)|(b[o+3]<<24))>>>0;
const f64buf=new DataView(new ArrayBuffer(8)); const f64=(b,o)=>{for(let k=0;k<8;k++) f64buf.setUint8(k,b[o+k]); return f64buf.getFloat64(0,true);};
function rkVal(v){let r; if(v&2){r=(v|0)>>2;} else {f64buf.setUint32(4,(v&0xFFFFFFFC)>>>0,true); f64buf.setUint32(0,0,true); r=f64buf.getFloat64(0,true);} if(v&1)r/=100; return r;}
function cellNum(b,t,s){ if(t===2) return rkVal(u32(b,s+8)); if(t===5||t===9) return f64(b,s+8); return null; }
function cellReal(col,style,v){const p=new Uint8Array(16); const d=new DataView(p.buffer); d.setUint32(0,col,true); d.setUint32(4,style,true); d.setFloat64(8,v,true); return enc(5,p);}
function cellRkInt(col,style,n){const p=new Uint8Array(12); const d=new DataView(p.buffer); d.setUint32(0,col,true); d.setUint32(4,style,true); d.setUint32(8,((n<<2)|2)>>>0,true); return enc(2,p);}
function cellIsst(col,style,i){const p=new Uint8Array(12); const d=new DataView(p.buffer); d.setUint32(0,col,true); d.setUint32(4,style,true); d.setUint32(8,i,true); return enc(7,p);}
function cellStr(col,style,s){const p=new Uint8Array(12+2*s.length); const d=new DataView(p.buffer); d.setUint32(0,col,true); d.setUint32(4,style,true); d.setUint32(8,s.length,true); for(let i=0;i<s.length;i++) d.setUint16(12+2*i,s.charCodeAt(i),true); return enc(6,p);}
const numCell=(col,style,v)=>Number.isInteger(v)&&Math.abs(v)<2**29?cellRkInt(col,style,v):cellReal(col,style,v);
const td16=new TextDecoder('utf-16le');
function wstr(b,o){const n=u32(b,o); return [td16.decode(b.subarray(o+4,o+4+2*n)),o+4+2*n];}
function concat(parts){let n=0; for(const p of parts) n+=p.length; const o=new Uint8Array(n); let k=0; for(const p of parts){o.set(p,k); k+=p.length;} return o;}
async function xlsbBook(zip){
  const wb=await zip.file('xl/workbook.bin').async('uint8array'); const rels=await zip.file('xl/_rels/workbook.bin.rels').async('string');
  const rid={}; rels.replace(/<Relationship [^>]*>/g,m=>{const id=(m.match(/Id="([^"]+)"/)||[])[1], tg=(m.match(/Target="([^"]+)"/)||[])[1]; if(id)rid[id]=tg; return m;});
  const sheets={}; for(const [t,s,e] of recs(wb)) if(t===156){let o=s+8; const [r,o2]=wstr(wb,o); const [nm]=wstr(wb,o2); sheets[nm]='xl/'+rid[r].replace(/^\/?xl\//,'');}
  const sst=[]; const sf=zip.file('xl/sharedStrings.bin'); if(sf){const b=await sf.async('uint8array'); for(const [t,s] of recs(b)) if(t===19) sst.push(wstr(b,s+1)[0]);}
  return {wb,sheets,sst,sstIdx:new Map(sst.map((s,i)=>[s,i]).reverse())};
}
function setFullCalc(wb){const parts=[]; for(const [t,s,e,h] of recs(wb)){ if(t===157){const p=wb.slice(s,e); const d=new DataView(p.buffer); d.setUint16(24,d.getUint16(24,true)|1,true); parts.push(enc(t,p));} else parts.push(wb.subarray(h,e)); } return concat(parts);}
async function dropBinIndex(zip,sheetPath){
  const n=sheetPath.match(/sheet(\d+)\.bin$/)[1]; const relp=`xl/worksheets/_rels/sheet${n}.bin.rels`; const rf=zip.file(relp); if(!rf) return;
  const r=await rf.async('string'); const m=r.match(/<Relationship [^>]*binaryIndex[^>]*\/>/); if(!m) return;
  const tg=(m[0].match(/Target="([^"]+)"/)||[])[1]; const r2=r.replace(m[0],'');
  if(/<Relationship /.test(r2)) zip.file(relp,r2); else zip.remove(relp);
  const bi='xl/worksheets/'+tg.replace(/^.*\//,''); zip.remove(bi);
  let ct=await zip.file('[Content_Types].xml').async('string'); ct=ct.replace(new RegExp(`<Override PartName="/${bi.replace(/[.]/g,'\\.')}"[^>]*/>`),''); zip.file('[Content_Types].xml',ct);
}
/* scan a sheet: rows whose date column == max date */
function scanDates(b,dateCol){ let r=-1,max=-Infinity,rows=[],lastRow=-1;
  for(const [t,s,e] of recs(b)){ if(t===0){r=u32(b,s); lastRow=r;} else if((t===2||t===5)&&u32(b,s)===dateCol){const v=cellNum(b,t,s); if(v>max){max=v;rows=[r];} else if(v===max) rows.push(r);} }
  return {max,rows,lastRow}; }
function collectRows(b,rowSet,maxCol){const m=new Map(); let r=-1; const hdr=new Map();
  for(const [t,s,e,h] of recs(b)){ if(t===0){r=u32(b,s); if(rowSet.has(r)){hdr.set(r,b.slice(s,e)); m.set(r,[]);}} else if(m.has(r)&&t>=2&&t<=11){ if(maxCol==null||u32(b,s)<=maxCol) m.get(r).push([t,b.slice(s,e)]); } }
  return {cells:m,hdr};}
/* insert value cells into rows that already exist (with formula cells) */
function insertIntoRows(b,map){ const parts=[]; let last=0, r=-1, done=0;
  for(const [t,s,e,h] of recs(b)){ if(t===0){ r=u32(b,s); if(map.has(r)){ parts.push(b.subarray(last,e)); parts.push(map.get(r)); last=e; done++; } } }
  parts.push(b.subarray(last)); if(done!==map.size) throw new Error(`Expected ${map.size} target rows, found ${done}.`); return concat(parts);}
function appendBeforeEnd(b,newRecs,newLast){ const parts=[]; let last=0;
  for(const [t,s,e,h] of recs(b)){ if(t===148){ parts.push(b.subarray(last,s)); const p=b.slice(s,e); const d=new DataView(p.buffer); d.setUint32(4,Math.max(d.getUint32(4,true),newLast),true); parts.push(p); last=e; }
    else if(t===146){ parts.push(b.subarray(last,h)); parts.push(newRecs); last=h; } }
  parts.push(b.subarray(last)); return concat(parts);}

/* ---------- pivot (xlsb) ---------- */
const dayMs=864e5; const toDate=s=>new Date(Date.UTC(1899,11,30)+s*dayMs); const serialOf=d=>Math.round((d-Date.UTC(1899,11,30))/dayMs);
function targetSet(vis,F){ // vis: sorted serials
  if(!vis.length) return null; const L=vis[vis.length-1];
  if(vis.length>=3&&vis.every((v,i)=>i===0||v-vis[i-1]===7)){const n=vis.length; return new Set([...Array(n).keys()].map(k=>F-7*k));}
  const P=FB.mPrev, Y=d=>{const m=FB.fromSerial(d-364).m===FB.fromSerial(d).m; return m?d-364:d-371;};
  const pats=[d=>[P(d),d],d=>[P(d),d-7,d],d=>[Y(d),P(d),d-7,d],d=>[d-7,d]];
  const vs=new Set(vis); for(const fn of pats){const a=fn(L); if(a.length===vs.size&&a.every(x=>vs.has(x))) return new Set(fn(F));}
  return null;}
function cacheDatesBin(b){const out={}; let fi=-1; for(const [t,s,e] of recs(b)){ if(t===183) fi++; if(t===191&&fi>=0){const d=dv(b,s); if(d.getUint16(0,true)===0x20){const n=d.getUint32(2,true); const a=[]; for(let k=0;k<n;k++){const y=d.getUint16(6+8*k,true),m=d.getUint16(8+8*k,true),dd=d.getUint16(10+8*k,true); a.push(serialOf(Date.UTC(y,m-1,dd)));} out[fi]=a;}}} return out;}
async function fixPivotsXlsb(zip,F,log,label){
  const names=Object.keys(zip.files).filter(n=>/^xl\/pivotTables\/pivotTable\d+\.bin$/.test(n));
  for(const n of names){ const rel=zip.file(n.replace('pivotTables/','pivotTables/_rels/')+'.rels'); if(!rel) continue;
    const tg=((await rel.async('string')).match(/Target="([^"]+)"/)||[])[1]; if(!tg) continue; const cp='xl/pivotCache/'+tg.replace(/^.*\//,'');
    const cd=cacheDatesBin(await zip.file(cp).async('uint8array')); const b=await zip.file(n).async('uint8array');
    const fields=[]; let cur=null; for(const [t,s,e] of recs(b)){ if(t===285){cur=[];fields.push(cur);} else if(t===282&&cur){cur.push({off:s+1,type:b[s],flags:b[s+1],ic:dv(b,s+3).getInt32(0,true)});} else if(t===286) cur=null; }
    let changed=false; const pn=n.match(/(\d+)\.bin/)[1];
    for(const [fi,ds] of Object.entries(cd)){ const it=(fields[fi]||[]).filter(i=>i.type===0&&i.ic>=0&&i.ic<ds.length); if(!it.length||!it.some(i=>i.flags&1)) continue;
      const vis=it.filter(i=>!(i.flags&1)).map(i=>ds[i.ic]).sort((a,b)=>a-b); const mx=Math.max(...ds);
      if(mx<F-7){log.push(`${label} pivot ${pn}: data source stops at ${fmt(mx)}; left as is.`);continue;}
      const tgs=targetSet(vis,F); if(!tgs){log.push(`${label} pivot ${pn}: date filter pattern not recognised; left as is.`);continue;}
      const visS=new Set(vis); const same=[...tgs].filter(x=>ds.includes(x)).every(x=>visS.has(x))&&vis.every(x=>tgs.has(x)); if(same){log.push(`${label} pivot ${pn}: already on ${fmt(F)}.`);continue;}
      for(const i of it){ const want=tgs.has(ds[i.ic]); b[i.off]=(i.flags&~1)|(want?0:1); }
      changed=true; const inCache=tgs.has(F)&&ds.includes(F);
      log.push(`${label} pivot ${pn}: ${vis.length} dates ending ${fmt(vis[vis.length-1])} → ${tgs.size} ending ${fmt(F)}.${inCache?'':' The new week enters this pivot when Excel refreshes it on open; if it shows unticked, tick it once and turn on “Include new items in manual filter” for the field.'}`); }
    if(changed) zip.file(n,b); }
}
function gen(zip){ for(const n of Object.keys(zip.files)) if(zip.files[n].dir) delete zip.files[n]; return zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}}); }
const fmt=s=>{const t=FB.fromSerial(s); return `${t.d} ${FB.MON[t.m-1]} ${t.y}`;};

/* ---------- fb_city ---------- */
const K=FB.K, OPN={tsel:'Telkomsel',xl:'XLSmart',isat:'Indosat IM3',three:'Tri Indonesia',sf:'Smartfren',xlplus:'XL+',ioh:'IOH'};
async function updateCity(file,C,log,progress){
  const zip=await JSZip.loadAsync(file); const B=await xlsbBook(zip); const F=C.F;
  { const b0=await zip.file(B.sheets['RAW_fb']).async('uint8array'); const mx=scanDates(b0,9).max;
    if(mx>F) throw new Error(`fb_city already has ${fmt(mx)}, newer than the dashboard week ${fmt(F)}.`);
    if(mx===F){ log.push(`fb_city already has ${fmt(F)}; data kept, only pivots and refresh settings updated.`); zip.file('xl/workbook.bin',setFullCalc(B.wb)); await fixPivotsXlsb(zip,F,log,'fb_city'); progress('fb_city: compressing'); return gen(zip); }
    if(!FB.S.city) throw new Error(`fb_city needs this week’s City CSV (${fmt(F)}); upload it above first.`); }
  const csv=new Map(); for(const r of FB.S.city.rows){const nm=r['City']; if(nm&&/, Jawa Barat$/.test(nm)) csv.set(nm.replace(/, Jawa Barat$/,'').trim().toUpperCase(),r);}
  const val=(c,k,m)=>FB.num(csv.get(c)[`${OPN[k]}-Mobile ${m}`]);
  for(const [sheet,dateCol,metric] of [['RAW_fb',9,'Market share'],['RAW_abs',8,'Total customers']]){
    progress(`fb_city: ${sheet}`); const path=B.sheets[sheet]; if(!path) throw new Error(`fb_city has no sheet ${sheet}.`);
    const b=await zip.file(path).async('uint8array'); const sc=scanDates(b,dateCol);
    if(sc.max>=F) throw new Error(`fb_city ${sheet} already has data for ${fmt(sc.max)}.`);
    if(sc.max!==F-7) log.push(`fb_city ${sheet}: last week in file is ${fmt(sc.max)}, not ${fmt(F-7)}; appended anyway.`);
    const src=sc.rows.sort((a,b)=>a-b); const {cells}=collectRows(b,new Set(src),dateCol); const base=src[src.length-1]+1; const ins=new Map();
    src.forEach((sr,i)=>{ const cs=cells.get(sr); const nameRec=cs.find(([t,p])=>u32(p,0)===0); const nm=nameRec&&nameRec[0]===7?B.sst[u32(nameRec[1],8)]:null;
      if(!nm||!csv.has(nm)) throw new Error(`fb_city ${sheet}: city ${nm} not in City CSV.`);
      const out=[]; for(const [t,p] of cs){ const col=u32(p,0), st=u32(p,4);
        if(col===dateCol) out.push(cellRkInt(col,st,F));
        else if(sheet==='RAW_fb'&&col>=2&&col<=8){const v=val(nm,K[col-2],metric); out.push(v==null?cellIsst(col,st,B.sstIdx.get('NULL')??0):cellReal(col,st,v));}
        else if(sheet==='RAW_abs'&&col>=1&&col<=7){const v=val(nm,K[col-1],metric); out.push(numCell(col,st,v==null?0:v));}
        else out.push(enc(t,p)); }
      ins.set(base+i,concat(out)); });
    zip.file(path,insertIntoRows(b,ins)); await dropBinIndex(zip,path); log.push(`fb_city ${sheet}: +${src.length} rows for ${fmt(F)}.`);
  }
  progress('fb_city: RAW_report'); { const path=B.sheets['RAW_report']; const b=await zip.file(path).async('uint8array');
    const rows=FB.S.city.rows.filter(r=>r['City']); const sty={},num={},dash={}; let r=-1;
    for(const [t,s,e] of recs(b)){ if(t===0) r=u32(b,s); else if(r>=3&&(t===2||t===5||t===7||t===6)){const c=u32(b,s),st=u32(b,s+4); if(r===3&&!(c in sty))sty[c]=st; if((t===2||t===5)&&!(c in num))num[c]=st; if(t===7&&!(c in dash))dash[c]=st;} }
    const groups=['Market share','Gross adds','Total customers','Churn','Net adds'], order=['tsel','xl','isat','three','sf','xlplus','ioh'];
    const parts=[]; let last=0, skip=false; r=-1; let done=0;
    for(const [t,s,e,h] of recs(b)){
      if(t===0){ r=u32(b,s); parts.push(b.subarray(last,e)); last=e; skip=false;
        if(r>=3&&r-3<rows.length){ const rec=rows[r-3]; const nm=rec['City'].replace(/,[^,]*$/,'').trim().toUpperCase(); const out=[cellRkInt(0,sty[0]||0,F), B.sstIdx.has(nm)?cellIsst(1,sty[1]||0,B.sstIdx.get(nm)):cellStr(1,sty[1]||0,nm)]; let c=2;
          for(const g of groups) for(const k of order){ const v=FB.num(rec[`${OPN[k]}-Mobile ${g}`]); out.push(v==null?cellIsst(c,dash[c]??sty[c]??0,B.sstIdx.get('--')??0):numCell(c,num[c]??sty[c]??0,v)); c++; }
          parts.push(concat(out)); skip=true; done++; } continue; }
      if(skip&&t>=2&&t<=11){ parts.push(b.subarray(last,h)); last=e; continue; }
    }
    parts.push(b.subarray(last)); zip.file(path,concat(parts)); await dropBinIndex(zip,path);
    if(done!==rows.length) log.push(`fb_city RAW_report: ${rows.length} cities in CSV but ${done} rows replaced; check the sheet.`); else log.push(`fb_city RAW_report: replaced with ${done} cities.`); }
  zip.file('xl/workbook.bin',setFullCalc(B.wb)); await fixPivotsXlsb(zip,F,log,'fb_city');
  progress('fb_city: compressing'); return gen(zip);
}
/* ---------- fb_kec ---------- */
async function updateKec(file,C,log,progress){
  const zip=await JSZip.loadAsync(file); const B=await xlsbBook(zip); const F=C.F;
  { const b0=await zip.file(B.sheets['RAW_fb']).async('uint8array'); const mx=scanDates(b0,17).max;
    if(mx>F) throw new Error(`fb_kec already has ${fmt(mx)}, newer than the dashboard week ${fmt(F)}.`);
    if(mx===F){ log.push(`fb_kec already has ${fmt(F)}; data kept, only pivots and refresh settings updated.`); zip.file('xl/workbook.bin',setFullCalc(B.wb)); await fixPivotsXlsb(zip,F,log,'fb_kec'); progress('fb_kec: compressing'); return gen(zip); }
    if(!FB.S.kec) throw new Error(`fb_kec needs this week’s Kec CSV (${fmt(F)}); upload it above first.`); }
  const KM=new Map(); for(const r of FB.S.kec.rows){const nm=r['Kecamatan 2022']; if(nm) KM.set(+nm.split('_').pop(),r);}
  const map={10:'tsel',11:'xl',12:'isat',13:'three',14:'sf',15:'xlplus',16:'ioh'};
  for(const [sheet,metric] of [['RAW_fb','Market share'],['RAW_abs','Total customers'],['RAW_grossadd','Gross adds']]){
    progress(`fb_kec: ${sheet}`); const path=B.sheets[sheet]; const b=await zip.file(path).async('uint8array'); const sc=scanDates(b,17);
    if(sc.max>=F) throw new Error(`fb_kec ${sheet} already has data for ${fmt(sc.max)}.`);
    const src=sc.rows.sort((a,b)=>a-b); const {cells,hdr}=collectRows(b,new Set(src),null); let nr=sc.lastRow; const out=[]; let miss=0;
    for(const sr of src){ nr++; const h=hdr.get(sr).slice(); new DataView(h.buffer).setUint32(0,nr,true); out.push(enc(0,h));
      const cs=cells.get(sr); const idc=cs.find(([t,p])=>u32(p,0)===9); const kid=Math.round(cellNum(idc[1],idc[0],0)); const rec=KM.get(kid); if(!rec) miss++;
      for(const [t,p] of cs){ const col=u32(p,0), st=u32(p,4);
        if(col===17) out.push(cellRkInt(col,st,F)); else if(col>=10&&col<=16){const v=rec?FB.num(rec[`${OPN[map[col]]}-Mobile ${metric}`]):null; out.push(numCell(col,st,v==null?0:v));} else out.push(enc(t,p)); } }
    zip.file(path,appendBeforeEnd(b,concat(out),nr)); await dropBinIndex(zip,path); log.push(`fb_kec ${sheet}: +${src.length} rows${miss?`, ${miss} kecamatan missing in CSV (written as 0)`:''}.`);
  }
  zip.file('xl/workbook.bin',setFullCalc(B.wb)); await fixPivotsXlsb(zip,F,log,'fb_kec');
  progress('fb_kec: compressing'); return gen(zip);
}
/* ---------- xlsx helpers ---------- */
async function xlsxSheets(zip){const wb=await zip.file('xl/workbook.xml').async('string'); const rels=await zip.file('xl/_rels/workbook.xml.rels').async('string'); const rid={};
  rels.replace(/<Relationship [^>]*>/g,m=>{const id=(m.match(/Id="([^"]+)"/)||[])[1], tg=(m.match(/Target="([^"]+)"/)||[])[1]; rid[id]=tg; return m;});
  const sh={}; wb.replace(/<sheet [^>]*>/g,m=>{const nm=(m.match(/name="([^"]+)"/)||[])[1], r=(m.match(/r:id="([^"]+)"/)||[])[1]; sh[nm.replace(/&amp;/g,'&')]='xl/'+rid[r].replace(/^\/?xl\//,''); return m;}); return {wb,sh,rid};}
const colNum=s=>{let n=0; for(const ch of s) n=n*26+ch.charCodeAt(0)-64; return n;};
const colName=n=>{let s=''; while(n){const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s;};
function rowBounds(s,r,from){const i=s.indexOf(`<row r="${r}"`,from||0); if(i<0) return null; const gt=s.indexOf('>',i); if(s[gt-1]==='/') return [i,gt+1]; return [i,s.indexOf('</row>',gt)+6];}
function cellEnd(x,k){const gt=x.indexOf('>',k); return x[gt-1]==='/'?gt+1:x.indexOf('</c>',gt)+4;}
function putCellXml(rowXml,ref,cx){ // replace or insert a complete cell xml in column order
  const k=rowXml.indexOf(`<c r="${ref}"`); if(k>=0) return rowXml.slice(0,k)+cx+rowXml.slice(cellEnd(rowXml,k));
  const cn=colNum(ref.replace(/\d+/g,'')); let ins=-1;
  for(const mm of rowXml.matchAll(/<c r="([A-Z]+)\d+"/g)){ if(colNum(mm[1])<cn) ins=cellEnd(rowXml,mm.index); else break; }
  if(ins<0){ const gt=rowXml.indexOf('>'); if(rowXml[gt-1]==='/') return rowXml.slice(0,gt-1)+'>'+cx+'</row>'; ins=gt+1; }
  return rowXml.slice(0,ins)+cx+rowXml.slice(ins);}
function styleOf(rowXml,ref){const mm=rowXml.match(new RegExp(`<c r="${ref}"(?: s="(\\d+)")?`)); return mm&&mm[1]?` s="${mm[1]}"`:'';}
function setCell(rowXml,ref,val,styleFrom){ const sty=styleOf(rowXml,ref)||(styleFrom?styleOf(rowXml,styleFrom):''); return putCellXml(rowXml,ref,val==null?`<c r="${ref}"${sty}/>`:`<c r="${ref}"${sty}><v>${val}</v></c>`);}
function fixPivotsXlsx(files,F,log,label){ // files: {pivotTableXml, cacheXml} pairs already as strings; returns modified pivot xml
  const out={};
  for(const [n,{t,cd}] of Object.entries(files)){
    const cfs=[...cd.matchAll(/<cacheField [^>]*>[\s\S]*?<\/cacheField>|<cacheField [^>]*\/>/g)].map(m=>m[0]);
    const pfs=[...t.matchAll(/<pivotField\b[^>]*?\/>|<pivotField\b[^>]*?[^/]>[\s\S]*?<\/pivotField>/g)];
    let nt=t, ch=false;
    for(let fi=0;fi<cfs.length&&fi<pfs.length;fi++){ const cf=cfs[fi]; if(!/containsDate="1"/.test(cf)) continue;
      const sh=[...cf.matchAll(/<(d|m|s|n)(?: v="([^"]*)")?\/>/g)].map(m=>m[1]==='d'?serialOf(Date.UTC(+m[2].slice(0,4),+m[2].slice(5,7)-1,+m[2].slice(8,10))):null);
      const p=pfs[fi][0]; const items=[...p.matchAll(/<item ([^>]*)\/>/g)].map(m=>m[1]).filter(a=>/x="\d+"/.test(a)&&!/t=/.test(a));
      if(!items.some(a=>/h="1"/.test(a))) continue;
      const vis=items.filter(a=>!/h="1"/.test(a)).map(a=>sh[+a.match(/x="(\d+)"/)[1]]).filter(v=>v!=null).sort((a,b)=>a-b);
      const tgs=targetSet(vis,F); if(!tgs){log.push(`${label} ${n}: date filter pattern not recognised; left as is.`);continue;}
      let np=p.replace(/<item ([^>]*)\/>/g,(m0,a)=>{const x=a.match(/x="(\d+)"/); if(!x||/t=/.test(a)) return m0; const a2=a.replace(/\s*h="1"/,'').trim(); return '<item '+(tgs.has(sh[+x[1]])?a2:'h="1" '+a2)+'/>';});
      np=np.replace(/^<pivotField\b/,m=>m).replace(/^<pivotField((?:(?!>)[\s\S])*?)(\/?)>/,(m0,a,sl)=>`<pivotField${/includeNewItemsInFilter=/.test(a)?a.replace(/includeNewItemsInFilter="\d"/,'includeNewItemsInFilter="1"'):a+' includeNewItemsInFilter="1"'}${sl}>`);
      nt=nt.replace(p,np); ch=true; log.push(`${label} ${n}: ${vis.length} dates ending ${fmt(vis[vis.length-1])} → ${tgs.size} ending ${fmt(F)}; new dates are included when Excel refreshes.`); }
    if(ch) out[n]=nt; }
  return out;}
async function pivotPairs(zip){ const wb=await zip.file('xl/workbook.xml').async('string'); const rels=await zip.file('xl/_rels/workbook.xml.rels').async('string'); const rid={};
  rels.replace(/<Relationship [^>]*>/g,m=>{rid[(m.match(/Id="([^"]+)"/)||[])[1]]=(m.match(/Target="([^"]+)"/)||[])[1]; return m;});
  const caches={}; wb.replace(/<pivotCache cacheId="(\d+)" r:id="([^"]+)"\/>/g,(m,c,r)=>{caches[c]='xl/'+rid[r]; return m;});
  const pairs={}; for(const n of Object.keys(zip.files).filter(n=>/^xl\/pivotTables\/pivotTable\d+\.xml$/.test(n))){const t=await zip.file(n).async('string'); const c=t.match(/cacheId="(\d+)"/)[1]; pairs[n]={t,cd:await zip.file(caches[c]).async('string'),cp:caches[c]};}
  return pairs;}
/* ---------- Fb_share_regional ---------- */
async function toPngBlob(img){ if(img.type==='image/png') return img; const bmp=await createImageBitmap(img); const cv=document.createElement('canvas'); cv.width=bmp.width; cv.height=bmp.height; cv.getContext('2d').drawImage(bmp,0,0); return await new Promise(r=>cv.toBlob(r,'image/png')); }
async function updateRegional(file,C,png,log,progress){
  const zip=await JSZip.loadAsync(file); const {wb,sh}=await xlsxSheets(zip); const F=C.F, L=C.L;
  let newDefs=wb; let mxR=-1; { const s0=await zip.file(sh['RAW']).async('string'); for(const q of s0.matchAll(/<c r="J\d+"[^>]*><v>(\d+)<\/v>/g)) if(+q[1]>mxR) mxR=+q[1]; }
  if(mxR>F) throw new Error(`Fb_share_regional already has ${fmt(mxR)}, newer than the dashboard week ${fmt(F)}.`);
  const dataDone=mxR===F; if(dataDone) log.push(`Fb_share_regional already has ${fmt(F)}; data kept, only pivots and refresh settings updated.`);
  if(!dataDone){
  progress('regional: RAW'); { const p=sh['RAW']; let s=await zip.file(p).async('string');
    const rows=[...s.matchAll(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g)]; let max=-1; for(const m of rows){const j=m[0].match(/<c r="J\d+"[^>]*><v>(\d+)<\/v>/); if(j&&+j[1]>max) max=+j[1];}
    if(max>=F) throw new Error(`Fb_share_regional RAW already has ${fmt(max)}.`);
    const blk=rows.filter(m=>new RegExp(`<c r="J\\d+"[^>]*><v>${max}</v>`).test(m[0])); const n=blk.length; const lastR=+blk[n-1][1];
    const cmap={B:0,C:1,D:2,E:3,F:4,G:5,H:6}; let add='';
    for(const m of blk){ const r=+m[1], nr=r+n; let x=m[0].replace(`<row r="${r}"`,`<row r="${nr}"`).replace(new RegExp(`r="([A-Z]+)${r}"`,'g'),(q,c)=>`r="${c}${nr}"`).replace(`<v>${max}</v>`,`<v>${F}</v>`);
      if(/<v>\d+<\/v>/.test(x) && x.includes('>Jabar<')===false){ /* Jabar detected via shared string index below */ }
      add+=x; }
    // Jabar row: find by shared string 'Jabar'
    const ss=await zip.file('xl/sharedStrings.xml').async('string'); const si=[...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>m[1].replace(/<[^>]+>/g,'')); const jIdx=si.indexOf('Jabar');
    add=add.replace(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g,row=>{ if(!new RegExp(`<c r="A\\d+"[^>]*t="s"[^>]*><v>${jIdx}</v>`).test(row)) return row; const rn=row.match(/<row r="(\d+)"/)[1]; let x=row; for(const [c,i] of Object.entries(cmap)) x=setCell(x,c+rn,C.jfull[i],c+rn); return x;});
    const end=blk[n-1].index+blk[n-1][0].length; s=s.slice(0,end)+add+s.slice(end); s=s.replace(/<dimension ref="A1:([A-Z]+)\d+"\/>/,(m,c)=>`<dimension ref="A1:${c}${lastR+n}"/>`); zip.file(p,s); log.push(`Fb_share_regional RAW: +${n} rows; Jabar ${C.jFromMeta?'from Meta':'derived from cities'}.`); }
  } // end RAW part
  const hasF=(s,row)=>{const b=rowBounds(s,row); if(!b) return false; return new RegExp(`<v>${F}</v>`).test(s.slice(b[0],b[1]));};
  const nextCol=(s,row)=>{const b=rowBounds(s,row); const rx=s.slice(b[0],b[1]); let best=0; for(const m of rx.matchAll(/<c r="([A-Z]+)\d+"[^>]*><v>(\d+)<\/v>/g)) if(+m[2]>=40000&&colNum(m[1])>best) best=colNum(m[1]); return best;};
  progress('regional: PPT and GA'); { const p=sh['PPT']; let s=await zip.file(p).async('string'); const lc=nextCol(s,1); const nc=colName(lc+1), pc=colName(lc);
    const put=(r,v)=>{const b=rowBounds(s,r); if(!b) return; const rx=s.slice(b[0],b[1]); s=s.slice(0,b[0])+setCell(rx,nc+r,v,pc+r)+s.slice(b[1]);};
    if(!hasF(s,1)){ put(1,F); put(3,C.jfull[6]); put(4,C.jfull[0]); put(5,C.jfull[5]); newDefs=newDefs.replace(new RegExp(`PPT!\\$A\\$1:\\$${pc}\\$5`,'g'),`PPT!$A$1:$${nc}$5`); log.push(`Fb_share_regional PPT: share trend column ${nc} = ${fmt(F)}.`); }
    else log.push(`Fb_share_regional PPT: share trend already has ${fmt(F)}.`);
    const glc=nextCol(s,9); const gnc=colName(glc+1), gpc=colName(glc);
    const put2=(r,v)=>{const b=rowBounds(s,r); if(!b) return; const rx=s.slice(b[0],b[1]); s=s.slice(0,b[0])+setCell(rx,gnc+r,v,gpc+r)+s.slice(b[1]);};
    if(!C.gaJ) log.push('Fb_share_regional PPT: gross-add trend not updated (Gross adds Jabar not entered).');
    else if(hasF(s,9)) log.push(`Fb_share_regional PPT: gross-add trend already has ${fmt(F)}.`);
    else { put2(9,F); put2(11,C.gaJ.cur[1]); put2(12,C.gaJ.cur[0]); put2(13,C.gaJ.cur[2]); newDefs=newDefs.replace(new RegExp(`PPT!\\$A\\$9:\\$${gpc}\\$13`,'g'),`PPT!$A$9:$${gnc}$13`); log.push(`Fb_share_regional PPT: gross-add trend column ${gnc} = ${fmt(F)}.`); }
    s=s.replace(/(<row r="(?:1|3|4|5|9|11|12|13)" spans="1:)(\d+)"/g,(m,a,n)=>a+Math.max(+n,lc+1)+'"'); zip.file(p,s); }
  if(!C.gaJ) log.push('Fb_share_regional GA: column not added (Gross adds Jabar not entered).');
  else if(hasF(await zip.file(sh['GA']).async('string'),1)) log.push(`Fb_share_regional GA: already has ${fmt(F)}.`);
  else { const p=sh['GA']; let s=await zip.file(p).async('string'); const lc=nextCol(s,1); const nc=colName(lc+1), pc=colName(lc);
    const put=(r,v)=>{const b=rowBounds(s,r); const rx=s.slice(b[0],b[1]); s=s.slice(0,b[0])+setCell(rx,nc+r,v,pc+r)+s.slice(b[1]);};
    put(1,F); put(3,C.gaJ.cur[1]); put(4,C.gaJ.cur[0]); put(5,C.gaJ.cur[2]);
    s=s.replace(/(<row r="(?:1|3|4|5)" spans="1:)(\d+)"/g,(m,a,n)=>a+Math.max(+n,lc+1)+'"').replace(/<dimension ref="A1:([A-Z]+)(\d+)"\/>/,(m,c,r)=>`<dimension ref="A1:${colNum(c)<lc+1?nc:c}${r}"/>`); zip.file(p,s); log.push(`Fb_share_regional GA: column ${nc} = ${fmt(F)}.`); }
  progress('regional: Paths'); { const p=sh['Paths']; let s=await zip.file(p).async('string'); const BL={JABAR:'A',BANDUNG:'K',CIREBON:'U',SOREANG:'AE',TASIK:'AO'};
    for(const [b,c0] of Object.entries(BL)){ const t=C.pathText[b]; if(!t||C.dupPaths.has(b)){log.push(`Fb_share_regional Paths ${b}: not updated (${t?'duplicate file':'no file'}).`);continue;}
      const data=FB.csvRows(t).filter(r=>r.length>=6); const cols=[...Array(6).keys()].map(i=>colName(colNum(c0)+i));
      for(let i=0;i<23;i++){ const r=3+i; const b2=rowBounds(s,r); if(!b2) continue; let rx=s.slice(b2[0],b2[1]);
        const d=data[i]; const vals=d?[(()=>{const [y,m,dd]=d[0].split('-').map(Number); return FB.toSerial(y,m,dd);})(),d[1],d[2],d[3],Number(d[4]),/^-?[\d,]+$/.test(d[5])?Number(d[5].replace(/,/g,'')):d[5]]:[null,null,null,null,null,null];
        cols.forEach((c,k)=>{ const v=vals[k]; const ref=c+r; const sty=styleOf(rx,ref);
          const cx=v==null?`<c r="${ref}"${sty}/>`:(typeof v==='number'?`<c r="${ref}"${sty}><v>${v}</v></c>`:`<c r="${ref}"${sty} t="inlineStr"><is><t>${String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;')}</t></is></c>`);
          if(v!=null||rx.includes(`<c r="${ref}"`)) rx=putCellXml(rx,ref,cx); });
        s=s.slice(0,b2[0])+rx+s.slice(b2[1]); } log.push(`Fb_share_regional Paths ${b}: ${data.length} rows.`); }
    zip.file(p,s); }
  if(png){ progress('regional: screenshot'); png=await toPngBlob(png); const media=Object.keys(zip.files).filter(n=>/^xl\/media\/image\d+\./.test(n)).map(n=>+n.match(/image(\d+)/)[1]); const k=Math.max(0,...media)+1;
    const gaPath=sh['GA']; const srel=gaPath.replace('worksheets/','worksheets/_rels/')+'.rels'; const sr=await zip.file(srel).async('string'); const dtg=sr.match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/);
    if(dtg){ const dp='xl/drawings/'+dtg[1], drp='xl/drawings/_rels/'+dtg[1]+'.rels'; let d=await zip.file(dp).async('string'), dr=await zip.file(drp).async('string');
      const ids=[...dr.matchAll(/Id="rId(\d+)"/g)].map(m=>+m[1]); const nid='rId'+(Math.max(0,...ids)+1);
      dr=dr.replace('</Relationships>',`<Relationship Id="${nid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${k}.png"/></Relationships>`);
      const anchors=[...d.matchAll(/<xdr:twoCellAnchor[^>]*>[\s\S]*?<\/xdr:twoCellAnchor>/g)]; const last=anchors[anchors.length-1][0]; const mx=Math.max(...[...d.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(m=>+m[1]));
      const bmp=await createImageBitmap(png); const cx=+(last.match(/<a:ext cx="(\d+)"/)||[0,0])[1]; const cy=Math.round(cx*bmp.height/bmp.width);
      const na=last.replace(/r:embed="[^"]+"/,`r:embed="${nid}"`).replace(/<xdr:cNvPr id="\d+" name="[^"]+"/,`<xdr:cNvPr id="${mx+1}" name="Picture ${mx}"`).replace(/<a16:creationId [^>]*\/>/g,'').replace(/<a:ext cx="\d+" cy="\d+"\/>/,`<a:ext cx="${cx}" cy="${cy}"/>`);
      zip.file(dp,d.replace('</xdr:wsDr>',na+'</xdr:wsDr>')); zip.file(drp,dr); zip.file(`xl/media/image${k}.png`,await png.arrayBuffer());
      let ct=await zip.file('[Content_Types].xml').async('string'); if(!/Extension="png"/.test(ct)) ct=ct.replace('<Default ','<Default Extension="png" ContentType="image/png"/><Default '); zip.file('[Content_Types].xml',ct); log.push('Fb_share_regional GA: screenshot pasted.'); } }
  const pairs=await pivotPairs(zip); const fx=fixPivotsXlsx(pairs,F,log,'Fb_share_regional'); for(const [n,t] of Object.entries(fx)) zip.file(n,t);
  for(const {cd,cp} of Object.values(pairs)) if(!/refreshOnLoad="1"/.test(cd)) zip.file(cp,cd.replace('<pivotCacheDefinition ','<pivotCacheDefinition refreshOnLoad="1" '));
  newDefs=newDefs.replace(/<calcPr([^>]*?)\/>/,(m,a)=>/fullCalcOnLoad/.test(a)?m:`<calcPr${a} fullCalcOnLoad="1"/>`); zip.file('xl/workbook.xml',newDefs);
  progress('regional: compressing'); return gen(zip);
}
const S_GA_OK=C=>!!C.gaJ;
/* ---------- fb_gross_add ---------- */
async function updateGA(file,C,log,progress){
  const zip=await JSZip.loadAsync(file); const {wb,sh}=await xlsxSheets(zip); const F=C.F;
  progress('gross add: RAW'); const p=sh['RAW']; let s=await zip.file(p).async('string');
  const ss=await zip.file('xl/sharedStrings.xml').async('string'); const si=[...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>m[1].replace(/<[^>]+>/g,''));
  let max=-1; for(const m of s.matchAll(/<c r="J(\d+)"[^>]*><v>(\d+)<\/v>/g)) if(+m[2]>max) max=+m[2];
  if(max>F) throw new Error(`fb_gross_add already has ${fmt(max)}, newer than the dashboard week ${fmt(F)}.`);
  if(max===F){ s=null; log.push(`fb_gross_add already has ${fmt(F)}; data kept, only refresh settings updated.`);
    let w=wb.replace(/<calcPr([^>]*?)\/>/,(q,a)=>/fullCalcOnLoad/.test(a)?q:`<calcPr${a} fullCalcOnLoad="1"/>`); zip.file('xl/workbook.xml',w);
    const pr=await pivotPairs(zip); for(const {cd,cp} of Object.values(pr)) if(!/refreshOnLoad="1"/.test(cd)) zip.file(cp,cd.replace('<pivotCacheDefinition ','<pivotCacheDefinition refreshOnLoad="1" '));
    progress('gross add: compressing'); return gen(zip); }
  if(!FB.S.city) throw new Error(`fb_gross_add needs this week’s City CSV (${fmt(F)}); upload it above first.`);
  const rowsMax=[]; for(const m of s.matchAll(new RegExp(`<c r="J(\\d+)"[^>]*><v>${max}</v>`,'g'))) rowsMax.push(+m[1]); rowsMax.sort((a,b)=>a-b);
  const n=rowsMax.length, OP={B:'tsel',C:'xl',D:'isat',E:'three',F:'sf',G:'xlplus',H:'ioh'};
  const csv=new Map(); for(const r of FB.S.city.rows){const nm=r['City']; if(nm&&/, Jawa Barat$/.test(nm)) csv.set(nm.replace(/, Jawa Barat$/,'').trim().toUpperCase(),r);}
  const pieces=[]; let cursor=0; const repl=[];
  for(let i=0;i<n;i++){ const sr=rowsMax[i], nr=rowsMax[n-1]+1+i; const sb=rowBounds(s,sr), tb=rowBounds(s,nr); if(!sb||!tb) throw new Error('fb_gross_add RAW: expected empty rows after the last week.');
    let x=s.slice(sb[0],sb[1]); const ai=x.match(/<c r="A\d+" t="s"><v>(\d+)<\/v>/); const nm=si[+ai[1]]; const rec=csv.get(nm); if(!rec) throw new Error(`fb_gross_add: city ${nm} not in City CSV.`);
    x=x.replace(`<row r="${sr}"`,`<row r="${nr}"`).replace(new RegExp(`r="([A-Z]+)${sr}"`,'g'),(q,c)=>`r="${c}${nr}"`).replace(new RegExp(`\\(A${sr},`,'g'),`(A${nr},`).replace(`<v>${max}</v>`,`<v>${F}</v>`);
    for(const [c,k] of Object.entries(OP)){ const v=FB.num(rec[`${OPN[k]}-Mobile Gross adds`])||0; x=x.replace(new RegExp(`(<c r="${c}${nr}" s="\\d+"><v>)[^<]*(</v>)`),`$1${v}$2`); }
    repl.push([tb[0],tb[1],x]); }
  repl.sort((a,b)=>a[0]-b[0]); for(const [a,b,x] of repl){ pieces.push(s.slice(cursor,a),x); cursor=b; } pieces.push(s.slice(cursor)); s=null;
  zip.file(p,pieces.join('')); log.push(`fb_gross_add RAW: +${n} rows for ${fmt(F)}.`);
  let w=wb.replace(/<calcPr([^>]*?)\/>/,(m,a)=>/fullCalcOnLoad/.test(a)?m:`<calcPr${a} fullCalcOnLoad="1"/>`); zip.file('xl/workbook.xml',w);
  if(zip.file('xl/calcChain.xml')){ zip.remove('xl/calcChain.xml'); let r=await zip.file('xl/_rels/workbook.xml.rels').async('string'); r=r.replace(/<Relationship [^>]*calcChain[^>]*\/>/,''); zip.file('xl/_rels/workbook.xml.rels',r);
    let ct=await zip.file('[Content_Types].xml').async('string'); ct=ct.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/,''); zip.file('[Content_Types].xml',ct); }
  const pairs=await pivotPairs(zip); for(const {cd,cp} of Object.values(pairs)) if(!/refreshOnLoad="1"/.test(cd)) zip.file(cp,cd.replace('<pivotCacheDefinition ','<pivotCacheDefinition refreshOnLoad="1" '));
  progress('gross add: compressing'); return gen(zip);
}
window.MASTER={updateCity,updateKec,updateRegional,updateGA,targetSet};
})();

;
(function(){
'use strict';
const FB=window.FB; const WEEKS={}; let EMB=0;
try{ const el=document.getElementById('saved-weeks'); const j=el?JSON.parse(el.textContent||'{}'):{}; for(const [d,p] of Object.entries(j)){ try{ FB.mergeWeek(+d,JSON.parse(p)); WEEKS[d]=p; EMB++; }catch(e){} } }catch(e){}
const $=s=>document.querySelector(s);
let C=null, db=null, downloads=null, dirtyMeta=false, OWNER=false;
function applyRole(){ document.body.classList.toggle('is-owner',OWNER); }
const el=(t,a={},...kids)=>{const e=document.createElement(t);for(const[k,v]of Object.entries(a)){if(k==='class')e.className=v;else if(k==='text')e.textContent=v;else e.setAttribute(k,v);}kids.flat().forEach(k=>e.append(k instanceof Node?k:document.createTextNode(String(k))));return e;};
const OPS7=[['tsel','Telkomsel'],['xl','XLSmart'],['isat','Indosat IM3'],['three','Tri'],['sf','Smartfren'],['xlplus','XL+'],['ioh','IOH']];

/* ---------- capabilities ---------- */
(async()=>{ if(!window.claude||!window.claude.use) return;
  try{ const usr=await window.claude.use('user'); OWNER=usr?!!(await usr.isOwner()):false; }catch(e){ OWNER=false; }
  applyRole(); renderStatus(); show();
  try{ downloads=await window.claude.use('downloads'); }catch(e){}
  try{ db=await window.claude.use('db'); }catch(e){}
  if(db) loadHistory();
  renderStatus();
})();
async function loadHistory(){
  const t=setTimeout(()=>{setHist('History: seed only (saved weeks did not load in time).');},10000);
  try{ const snap=await db.collection('weeks').get(); clearTimeout(t); let n=0;
    snap.docs.forEach(d=>{const v=d.data(); if(!v||!v.payload) return; try{FB.mergeWeek(+d.id,JSON.parse(v.payload)); WEEKS[d.id]=v.payload; n++;}catch(e){}});
    const sd=(()=>{const t=FB.fromSerial(Math.max(...Object.keys(window.__SEED__.ms).map(Number)));return `${t.d} ${FB.MON[t.m-1]} ${t.y}`;})(); const lastW=(()=>{const t=FB.fromSerial(Math.max(...Object.keys(FB.H.ms).map(Number)));return `${t.d} ${FB.MON[t.m-1]} ${t.y}`;})();
    const k=Object.keys(WEEKS).length; setHist(k?`History: built-in to ${sd}, plus ${k} saved week${k>1?'s':''} (latest ${lastW}).`:`History: built-in to ${sd}. No saved weeks yet.`);
    if(!FB.S.city&&!FB.S.uploadF) FB.S.F=null; run();
  }catch(e){clearTimeout(t); setHist('History: seed only. Saved weeks could not be read ('+(e.code||'error')+').');}
}
function setHist(t){$('#hist').textContent=t;}

/* ---------- file loading ---------- */
const need=[['city','City CSV'],['kec','Kec CSV'],['JABAR','Paths Jabar'],['BANDUNG','Paths Bandung'],['CIREBON','Paths Cirebon'],['SOREANG','Paths Soreang'],['TASIK','Paths Tasikmalaya']];
function renderStatus(){
  const box=$('#need'); box.textContent='';
  const dup=C?C.dupPaths:new Set();
  for(const [k,l] of need){ const have=k==='city'?!!FB.S.city:k==='kec'?!!FB.S.kec:!!FB.S.paths[k];
    const saved=!have&&C&&((k==='city'&&C.fromHistory)||(C.pathSrc&&C.pathSrc[k]==='saved'));
    const st=dup.has(k)?'dup':(have||saved?'ok':'miss');
    box.append(el('li',{class:'chip '+st},el('span',{class:'dot','aria-hidden':'true'}),l+(st==='dup'?' · duplicate':(saved?' · saved week':''))));}
  $('#btn-save').disabled=!(C&&OWNER); $('#dl-note').hidden=!!downloads||!window.claude;
}
async function handleFiles(list){
  const log=[];
  for(const f of list){ if(!/\.csv$/i.test(f.name)){log.push(`${f.name}: skipped (not CSV)`);continue;}
    const txt=await f.text(); const k=FB.ingest(f.name,txt); log.push(`${f.name}: ${k}`); }
  if((FB.S.city||FB.S.kec)&&!FB.S.uploadF){
    const pd=Object.values(FB.S.paths).map(p=>p.date).filter(Boolean).sort((a,b)=>b-a)[0];
    if(pd){ FB.S.F=pd; FB.S.uploadF=pd; const t=FB.fromSerial(pd); log.push(`Week date taken from the Paths files: ${t.d} ${FB.MON[t.m-1]} ${t.y} (the City/Kec file names have no date).`); }
    else { FB.S.city=null; FB.S.kec=null; FB.S.F=null; log.push('⚠ City/Kec not used: the file name has no week date. Rename it like City_Latest_2026-11-08.csv / Kec_Latest_2026-11-08.csv, or add the Paths files in the same upload.'); }
  }
  $('#log').textContent=log.join('\n'); run(true);
}
$('#file').addEventListener('change',e=>handleFiles([...e.target.files]));
const dz=$('#drop');
['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('over');}));
['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('over');}));
dz.addEventListener('drop',e=>handleFiles([...e.dataTransfer.files]));

/* ---------- meta inputs ---------- */
function buildMeta(){
  const ms=$('#meta-ms'); OPS7.forEach(([k,l],i)=>ms.append(el('label',{class:'fld'},el('span',{},l),el('input',{type:'number',step:'0.01',id:'ms-'+k,inputmode:'decimal'}))));
  const ga=$('#meta-ga'); [['tsel','Telkomsel'],['ioh','IOH'],['xlplus','XL+']].forEach(([k,l])=>ga.append(el('label',{class:'fld'},el('span',{},l),el('input',{type:'number',step:'1000',id:'ga-'+k,inputmode:'numeric',placeholder:'e.g. 77000'}))));
  document.querySelectorAll('#meta input').forEach(i=>i.addEventListener('change',()=>{readMeta();run();}));
}
function readMeta(){
  const ms=OPS7.map(([k])=>parseFloat($('#ms-'+k).value)); FB.S.metaMS=ms.every(v=>!isNaN(v))&&touched('#meta-ms')?ms:null;
  const ga=['tsel','ioh','xlplus'].map(k=>parseFloat($('#ga-'+k).value)); FB.S.metaGA=ga.every(v=>!isNaN(v))&&touched('#meta-ga')?ga:null;
}
const touched=sel=>[...document.querySelectorAll(sel+' input')].some(i=>i.dataset.touched==='1');
function dirtyMetaCheck(){ return touched('#meta-ms'); }
document.addEventListener('input',e=>{if(e.target.closest&&e.target.closest('#meta'))e.target.dataset.touched='1';});

/* ---------- run & render ---------- */
let lastF=null, lastUp=false;
function run(fresh){
  C=FB.compute();
  if(C&&C.F!==lastF){
    const keep=lastF!==null&&!lastUp&&!C.fromHistory&&(touched('#meta-ms')||touched('#meta-ga'));
    if(keep){ readMeta(); const t=FB.fromSerial(C.F); const lg=$('#log'); lg.textContent=(lg.textContent?lg.textContent+'\n':'')+`Meta values typed before the upload are used for ${t.d} ${FB.MON[t.m-1]} ${t.y}. Check them against this week’s screenshots.`; }
    else { /* a different uploaded week: forget the previous week's Meta inputs */ document.querySelectorAll('#meta input').forEach(i=>{i.dataset.touched='';i.value='';}); FB.S.metaMS=null; FB.S.metaGA=null; }
    C=FB.compute(); lastF=C?C.F:null; lastUp=C?!C.fromHistory:false; }
  renderStatus();
  if(!C){ $('#nodata').hidden=false; $('#asof').textContent='The uploaded files are for a week with no city data yet. Add City_Latest for that week.'; return; }
  $('#nodata').hidden=true;
  OPS7.forEach(([k],i)=>{const inp=$('#ms-'+k); if(inp.dataset.touched!=='1') inp.value=(C.jfull||C.jDerived)[i].toFixed(2);});
  [['tsel',0],['ioh',1],['xlplus',2]].forEach(([k,i])=>{const inp=$('#ga-'+k); if(inp.dataset.touched!=='1') inp.value=C.gaJ?C.gaJ.cur[i]:'';});
  $('#wk').textContent=`Update : ${C.L.wk}`;
  $('#asof').textContent=`Data week ending ${C.L.d} ${FB.MON[C.L.m-1]} ${C.L.y}  ·  WoW vs ${fmtD(C.X)}  ·  MoM vs ${fmtD(C.P)}  ·  ${C.fromHistory?'from saved history':'from uploaded files'}`;
  $('#src-ms').textContent=C.jFromMeta?'Jabar share: Meta screenshot values.':'Jabar share: derived from city data (customer-weighted) until Meta values are entered.';
  $('#src-ga').textContent=S_GA(); if(typeof mwRender==='function') mwRender();
  const w=$('#warn'); w.textContent=''; FB.S.warn.forEach(t=>w.append(el('li',{},t))); w.hidden=!FB.S.warn.length;
  renderOverview(); renderMS(); renderGA(); renderPaths(); $('#caption').value=C.caption; if(window.__perfRender) window.__perfRender(C); if(window.__vizRender) window.__vizRender(C);
}
const fmtD=s=>{const t=FB.fromSerial(s);return `${t.d} ${FB.MON[t.m-1]}`;};
const jmsNow=()=>{const o=Object.assign({},FB.H.jms); if(C) o[C.F]={tsel:C.msJ.cur.tsel,ioh:C.msJ.cur.ioh,xlplus:C.msJ.cur.xlplus}; return o;};
const jgaNow=()=>{const o=Object.assign({},FB.H.jga); if(C&&C.gaJ) o[C.F]={tsel:C.gaJ.cur[0],ioh:C.gaJ.cur[1],xlplus:C.gaJ.cur[2]}; return o;};
function S_GA(){ if(FB.S.metaGA) return 'Gross adds Jabar: values entered here.'; if(C.gaFromTrend) return 'Gross adds Jabar: taken from the Jabar gross-add trend for this week (not a sum of cities).'; return 'Gross adds Jabar: not in the trend for this week yet. Enter the Meta screenshot values.'; }
const sgnCls=v=>v==null?'':(FB.r2(v)<0?'neg':(FB.r2(v)>0?'pos':''));
function kpi(title,val,sub,cls){return el('div',{class:'kpi'},el('div',{class:'kpi-t'},title),el('div',{class:'kpi-v '+(cls||'')},val),el('div',{class:'kpi-s'},...sub));}
const span=(t,c)=>el('span',{class:c||''},t);
window.__renderOverview=()=>{ if(C) renderOverview(); };
function renderOverview(){
  const J=C.msJ, g=C.gaJ, PJ=C.paths.JABAR; const box=$('#kpis'); box.textContent='';
  box.append(kpi('Telkomsel share, Jabar',FB.f2(J.cur.tsel)+'%',[span('WoW '),span(J.wow?FB.fpp(J.wow.tsel):'–',sgnCls(J.wow&&J.wow.tsel)),span('  MoM '),span(J.mom?FB.fpp(J.mom.tsel):'–',sgnCls(J.mom&&J.mom.tsel))]));
  box.append(kpi(`Gap to ${J.to}`,FB.sgnPct(J.gap),[span(`Rank ${J.rank} of 3  ·  ${J.to} ${FB.f2(J.to==='IOH'?J.cur.ioh:J.cur.xlplus)}%`)],sgnCls(J.gap)));
  box.append(kpi('Cities won',`${C.nWin} of 18`,[span(`${C.nNeg} cities fell both WoW and MoM`)]));
  box.append(kpi('Gross adds, Telkomsel',g?FB.fint(g.cur[0]):'Needs Meta value',g?[span('WoW '),span(g.wow?FB.fint(g.wow[0]):'–',sgnCls(g.wow&&g.wow[0])),span('  MoM '),span(g.mom?FB.fint(g.mom[0]):'–',sgnCls(g.mom&&g.mom[0]))]:[span('Enter the Jabar values from the Meta Gross adds screenshot')],g?'':'muted'));
  box.append(kpi('Churn In-Out, Telkomsel',PJ?FB.fK1(PJ.io.Telkomsel):'Needs Paths',PJ?[span(`${FB.PATH_LBL[PJ.gainMax]} leads at ${FB.fK1(PJ.io[PJ.gainMax])}`)]:[span('Upload JABARPaths CSV')],PJ?sgnCls(PJ.io.Telkomsel):'muted'));
  const mh=document.querySelector('.map-card'); lineChart($('#ov-chart'),jmsNow(),'Market share trend, Jabar (%)',v=>v.toFixed(2)+'%',mh&&mh.clientHeight>260&&innerWidth>900?Math.max(260,mh.clientHeight-60):260);
}
function th(t,c){return el('th',{class:c||''},t);}
function td(t,c){return el('td',{class:c||''},t);}
const HIC={TSEL:'hi-tsel',IOH:'hi-ioh','XL+':'hi-xl'};
function renderMS(){
  lineChart($('#ms-chart'),jmsNow(),'Market share trend, Jabar (%)',v=>v.toFixed(2)+'%');
  const t=$('#ms-table'); t.textContent='';
  t.append(el('thead',{},el('tr',{},th('City','l'),th('Tsel','op-t'),th('IOH','op-i'),th('XL+','op-x'),th('Win/Lose'),th('Gap'),th('Highest'),th('Rank'),th('WoW Tsel'),th('WoW IOH'),th('WoW XL+'),th('MoM Tsel'),th('MoM IOH'),th('MoM XL+'),th('YoY Tsel'),th('YoY IOH'),th('YoY XL+'))));
  const tb=el('tbody'); const all=[...C.ms.filter(Boolean),C.msJ];
  for(const r of all){ const mx=Math.max(r.cur.tsel,r.cur.ioh,r.cur.xlplus); const s=v=>td(FB.f2(v)+'%',v===mx?'max':'');
    const d=x=>x?[x.tsel,x.ioh,x.xlplus].map(v=>td(FB.fpp(v),sgnCls(v))):[td('–'),td('–'),td('–')];
    tb.append(el('tr',{class:r===C.msJ?'total':''},td(r.name,'l'),s(r.cur.tsel),s(r.cur.ioh),s(r.cur.xlplus),td(r.win?'Win':'Lose',r.win?'win':'lose'),td(FB.r2(r.gap).toFixed(2)+'%',sgnCls(r.gap)),td(r.hi,HIC[r.hi]),td(r.rank),...d(r.wow),...d(r.mom),...d(r.yoy)));}
  t.append(tb);
  const L=(a,f)=>el('ol',{},...a.map(x=>el('li',{},el('span',{},x[0]),el('b',{class:sgnCls(x[1])},f(x[1])))));
  const tl=$('#ms-lists'); tl.textContent='';
  tl.append(el('div',{},el('h4',{},'Top 3 WoW'),L(C.msTopW,FB.fpp)),el('div',{},el('h4',{},'Bottom 3 WoW'),L(C.msBotW,FB.fpp)),el('div',{},el('h4',{},'Top 3 MoM'),L(C.msTopM,FB.fpp)),el('div',{},el('h4',{},'Bottom 3 MoM'),L(C.msBotM,FB.fpp)));
}
function renderGA(){
  lineChart($('#ga-chart'),jgaNow(),'Gross adds trend, Jabar',v=>Math.round(v).toLocaleString('en-US'));
  const t=$('#ga-table'); t.textContent='';
  t.append(el('thead',{},el('tr',{},th('City','l'),th('Tsel','op-t'),th('IOH','op-i'),th('XL+','op-x'),th('Win/Lose'),th('Gap'),th('Highest'),th('Rank'),th('WoW Tsel'),th('WoW IOH'),th('WoW XL+'),th('MoM Tsel'),th('MoM IOH'),th('MoM XL+'))));
  const tb=el('tbody'); const all=[...C.ga.filter(Boolean)]; if(C.gaJ) all.push(C.gaJ);
  for(const r of all){ const mx=Math.max(...r.cur); const d=x=>x?x.map(v=>td(FB.fint(v),sgnCls(v))):[td('–'),td('–'),td('–')];
    tb.append(el('tr',{class:r===C.gaJ?'total':''},td(r.name,'l'),...r.cur.map(v=>td(FB.fint(v),v===mx?'max':'')),td(r.win?'Win':'Lose',r.win?'win':'lose'),td(FB.fint(r.gap),sgnCls(r.gap)),td(r.hi,HIC[r.hi]),td(r.rank),...d(r.wow),...d(r.mom)));}
  if(!C.gaJ) tb.append(el('tr',{class:'total'},td('JABAR','l'),...Array(13).fill(0).map(()=>td('–'))));
  t.append(tb);
  const L=(a)=>el('ol',{},...a.map(x=>el('li',{},el('span',{},x[0]),el('b',{class:sgnCls(x[1])},FB.fint(x[1])))));
  const tl=$('#ga-lists'); tl.textContent='';
  tl.append(el('div',{},el('h4',{},'Top 3 WoW'),L(C.gaTopW)),el('div',{},el('h4',{},'Bottom 3 WoW'),L(C.gaBotW)),el('div',{},el('h4',{},'Top 3 MoM'),L(C.gaTopM)),el('div',{},el('h4',{},'Bottom 3 MoM'),L(C.gaBotM)));
  $('#ga-hi').textContent=`Highest gross adds by city: XL+ ${C.gaHi['XL+']}, IOH ${C.gaHi.IOH}, Telkomsel ${C.gaHi.TSEL}. Ties go to Telkomsel first, as in the workbook.`;
}
const OPC={'Telkomsel':'op-t','XLSmart':'op-xl','Indosat IM3':'op-im3','Smartfren':'op-sf','Tri Indonesia':'op-tri'};
function matrixEl(b){ const M=C.paths[b]; const wrap=el('div',{class:'mx'});
  wrap.append(el('h3',{},b==='JABAR'?'Regional Jabar':'Branch '+FB.BR_NAME[b]));
  if(!M){wrap.append(el('p',{class:'muted'},`Upload ${b==='TASIK'?'TASIK':b}Paths CSV.`));return wrap;}
  if(C.dupPaths.has(b)) wrap.append(el('p',{class:'alert'},'This file is identical to another branch’s file. Re-export it before use.'));
  const t=el('table',{class:'grid'}); const O=FB.PATH_OPS;
  const lbl=o=>b==='JABAR'?FB.PATH_LBL[o]:({'Indosat IM3':'IM3','Tri Indonesia':'Tri','Smartfren':'SF','XLSmart':'XL','Telkomsel':'Tsel'})[o];
  t.append(el('thead',{},el('tr',{},th('From / To','l'),...O.map(o=>th(lbl(o),OPC[o])),th('Move to'),th('In-Out'))));
  const tb=el('tbody');
  for(const a of O) tb.append(el('tr',{},td(lbl(a),'l '+OPC[a]),...O.map(o=>td(o===a?'':FB.fint(M.m[a+'|'+o]||0))),td(FB.fint(M.out[a]),'b'),td(FB.fint(M.io[a]),'b '+sgnCls(M.io[a]))));
  tb.append(el('tr',{class:'total'},td('Move in','l'),...O.map(o=>td(FB.fint(M.inn[o]))),td(FB.fint(M.tot)),td('')));
  t.append(tb); wrap.append(el('div',{class:'scroll'},t));
  const line=C.caption.split('\n').find(l=>l.startsWith('- '+FB.BR_NAME[b]+':'));
  if(b!=='JABAR'&&line) wrap.append(el('p',{class:'note'},line.replace(/^- [^:]+: /,'')));
  return wrap;}
function renderPaths(){ const box=$('#paths'); box.textContent='';
  const cl=C.caption.split('\n'); const churn=cl.find(l=>l.startsWith('Secara movement')), br=cl.find(l=>l.startsWith('Pergerakan kompetisi'));
  $('#paths-lede').textContent=(churn||'Upload the Paths CSVs to see churn movement.').replace(/\*/g,'');
  box.append(matrixEl('JABAR')); const g=el('div',{class:'mx-grid'}); FB.BRANCHES.forEach(b=>g.append(matrixEl(b))); box.append(el('p',{class:'lede'},(br||'').replace(/\*/g,'')),g);}

/* ---------- SVG line chart ---------- */
window.__lineChart=lineChart;
function lineChart(host,series,title,fmt,HH){
  host.textContent=''; const ds=Object.keys(series).map(Number).sort((a,b)=>a-b).filter(d=>!C||d<=C.F).slice(-80);
  if(!ds.length){host.append(el('p',{class:'muted'},'No trend data.'));return;}
  const W=Math.max(300,host.clientWidth||600), Hh=HH||220, m={l:44,r:12,t:28,b:26};
  const keys=[['tsel','Telkomsel','var(--tsel)'],['ioh','IOH','var(--ioh)'],['xlplus','XL+','var(--xlp)']];
  let lo=Infinity,hi=-Infinity; ds.forEach(d=>keys.forEach(([k])=>{const v=series[d][k]; if(v!=null){lo=Math.min(lo,v);hi=Math.max(hi,v);}}));
  const pad=(hi-lo)*0.08||1; lo-=pad; hi+=pad;
  const x=i=>m.l+(W-m.l-m.r)*i/(ds.length-1||1), y=v=>m.t+(Hh-m.t-m.b)*(1-(v-lo)/(hi-lo));
  const NS='http://www.w3.org/2000/svg'; const svg=document.createElementNS(NS,'svg'); svg.setAttribute('viewBox',`0 0 ${W} ${Hh}`); svg.setAttribute('width','100%'); svg.setAttribute('role','img'); svg.setAttribute('aria-label',title);
  const mk=(t,a)=>{const e=document.createElementNS(NS,t);for(const[k,v]of Object.entries(a))e.setAttribute(k,v);return e;};
  const tx=mk('text',{x:m.l,y:16,class:'ct'}); tx.textContent=title; svg.append(tx);
  for(let i=0;i<=3;i++){const v=lo+(hi-lo)*i/3; svg.append(mk('line',{x1:m.l,x2:W-m.r,y1:y(v),y2:y(v),class:'gl'})); const t=mk('text',{x:m.l-6,y:y(v)+3,class:'ax','text-anchor':'end'}); t.textContent=fmt(v).replace('.00%','%'); svg.append(t);}
  let lastM=-1; ds.forEach((d,i)=>{const t=FB.fromSerial(d); if(t.m!==lastM&&(t.m===1||t.m===4||t.m===7||t.m===10)&&i>2&&i<ds.length-3){const e=mk('text',{x:x(i),y:Hh-8,class:'ax','text-anchor':'middle'}); e.textContent=`${FB.MON[t.m-1]} ${String(t.y).slice(2)}`; svg.append(e);} lastM=t.m;});
  keys.forEach(([k,l,c])=>{let p=''; ds.forEach((d,i)=>{const v=series[d][k]; if(v==null)return; p+=(p?'L':'M')+x(i).toFixed(1)+','+y(v).toFixed(1);}); svg.append(mk('path',{d:p,fill:'none',stroke:c,'stroke-width':2,'stroke-linejoin':'round'}));
    const v=series[ds[ds.length-1]][k]; if(v!=null){svg.append(mk('circle',{cx:x(ds.length-1),cy:y(v),r:3,fill:c}));}});
  const guide=mk('line',{y1:m.t,y2:Hh-m.b,class:'guide',visibility:'hidden'}); svg.append(guide);
  const tip=el('div',{class:'tip',hidden:''}); host.style.position='relative'; host.append(svg,tip);
  const leg=el('div',{class:'legend'},...keys.map(([k,l,c])=>el('span',{},el('i',{style:`background:${c}`}),`${l} ${fmt(series[ds[ds.length-1]][k]??NaN)}`))); host.append(leg);
  const move=ev=>{const r=svg.getBoundingClientRect(); const px=(ev.touches?ev.touches[0].clientX:ev.clientX)-r.left; const sx=px*W/r.width;
    const i=Math.max(0,Math.min(ds.length-1,Math.round((sx-m.l)/((W-m.l-m.r)/(ds.length-1||1))))); const d=ds[i];
    guide.setAttribute('x1',x(i));guide.setAttribute('x2',x(i));guide.setAttribute('visibility','visible');
    tip.hidden=false; tip.textContent=''; tip.append(el('b',{},fmtD(d)+' '+FB.fromSerial(d).y),...keys.map(([k,l])=>el('div',{},`${l}: ${series[d][k]==null?'–':fmt(series[d][k])}`)));
    const left=Math.min(r.width-150,Math.max(0,px+12)); tip.style.left=left+'px'; tip.style.top='30px';};
  svg.addEventListener('pointermove',move); svg.addEventListener('touchstart',move,{passive:true}); svg.addEventListener('pointerleave',()=>{tip.hidden=true;guide.setAttribute('visibility','hidden');});
}
let rz; window.addEventListener('resize',()=>{clearTimeout(rz);rz=setTimeout(()=>{if(C){renderOverview();renderMS();renderGA(); if(window.__perfResize) window.__perfResize(); if(window.__vizResize) window.__vizResize();}},150);});

/* ---------- tabs ---------- */
const VIEWS=['overview','share','city','heatmap','grossadd','churn','update','export'];
function show(){const h=location.hash.slice(1); let v=VIEWS.includes(h)?h:'overview'; if((v==='update'||v==='export')&&!OWNER) v='overview';
  VIEWS.forEach(n=>{$('#v-'+n).hidden=n!==v; const b=$(`#tabs a[href="#${n}"]`); b.setAttribute('aria-current',n===v?'page':'false');});
  if(C) {renderOverview();renderMS();renderGA(); if(window.__perfResize) window.__perfResize(); if(window.__vizResize) window.__vizResize();}}
window.addEventListener('hashchange',show);

/* ---------- exports ---------- */
function msg(t,bad){const m=$('#ex-msg'); m.textContent=t; m.className=bad?'alert':'ok';}
async function save(filename,data){
  if(!downloads){msg('Downloads are not available in this view. Open the page in claude.ai to save files.',true);return;}
  try{await downloads.save({filename,data}); msg(`Saved ${filename}.`);}catch(e){ msg(e.code==='declined'?`${filename} was not saved.`:`Could not save ${filename} (${e.code||'error'}).`, e.code!=='declined');}
}
$('#ex-ppt').addEventListener('click',async()=>{ if(!C)return; if(typeof PptxGenJS==='undefined'){msg('The PowerPoint library did not load. Reload the page and try again.',true);return;}
  msg('Building the deck…'); try{ const b=await FB.buildPptx(C); await save('market_share_'+C.L.ymd+'.pptx',b);}catch(e){msg('Deck build failed: '+e.message,true);} });
$('#ex-xlsx').addEventListener('click',async()=>{ if(!C)return; if(typeof JSZip==='undefined'){msg('The Excel library did not load. Reload the page and try again.',true);return;}
  msg('Building the workbook…'); try{ const b=await FB.buildXlsx(FB.excelSheets(C)); await save('FB_share_update_'+C.L.ymd+'.xlsx',b);}catch(e){msg('Workbook build failed: '+e.message,true);} });
$('#ex-city').addEventListener('click',()=>C&&save(`nobar_ms_city_update_${C.L.ymd}.csv`,FB.nobarCity(C)));
$('#ex-region').addEventListener('click',()=>C&&save(`nobar_ms_region_update_${C.L.ymd}.csv`,FB.nobarRegion(C)));
$('#ex-kec').addEventListener('click',()=>{ if(!C)return; if(!FB.S.kec){msg('Upload the Kec CSV first.',true);return;} save(`nobar_ms_kec_update_${C.L.ymd}.csv`,FB.nobarKec(C));});
$('#ex-cap').addEventListener('click',async()=>{ const ta=$('#caption'); try{await navigator.clipboard.writeText(ta.value); msg('Caption copied.');}catch(e){ ta.focus(); ta.select(); const ok=document.execCommand&&document.execCommand('copy'); msg(ok?'Caption copied.':'Select the caption and copy it manually.',!ok);} });
$('#ex-captxt').addEventListener('click',()=>C&&save(`caption_FB_share_${C.L.ymd}.txt`,C.caption));
$('#btn-save').addEventListener('click',async()=>{ if(!C||!OWNER)return;
  const p={ms:FB.UP.map(c=>C.cv[c]?C.cv[c].ms:null),ga:FB.UP.map(c=>C.cv[c]?C.cv[c].ga.map(v=>v||0):null),jms:{ioh:C.jfull[6],tsel:C.jfull[0],xlplus:C.jfull[5]},jfull:C.jfull,paths:{}};
  for(const b in C.pathText) if(!C.dupPaths.has(b)) p.paths[b]=C.pathText[b];
  if(FB.S.kec&&FB.S.F===C.F){ const kc=FB.kecCurrent(C.F); p.kms=kc.ms; p.kga=kc.ga; }
  if(C.gaJ) p.jga={tsel:C.gaJ.cur[0],ioh:C.gaJ.cur[1],xlplus:C.gaJ.cur[2]};
  if(p.ms.some(v=>!v)){msg('Some Jabar cities are missing from the City CSV, so this week was not saved.',true);return;}
  const btn=$('#btn-save'); btn.disabled=true; const pay=JSON.stringify(p); const note=C.jFromMeta?'':' Jabar share is still derived from city data.';
  if(db){ try{ await db.doc('weeks/'+C.F).set({payload:pay,savedAt:new Date().toISOString()}); }catch(e){ /* the page republish below is what public viewers see */ } }
  FB.mergeWeek(C.F,p); WEEKS[String(C.F)]=pay;
  let art=null; try{ art=await window.claude.use('artifact'); }catch(e){}
  if(!art){ msg(`Week ${C.L.wk} saved for signed-in viewers only: this view cannot update the public page.`+note,true); btn.disabled=false; return; }
  msg(`Publishing week ${C.L.wk} for everyone… the page reloads when done.`);
  const html=window.__INDEX_PRE+'<script id="saved-weeks" type="application/json">'+JSON.stringify(WEEKS).replace(/</g,'\\u003c')+'</'+'script>'+window.__INDEX_POST;
  try{ await art.publish(html); }catch(e){ const c=e&&e.code||'error'; msg(c==='conflict'?'Someone published a newer version first; the page is reloading. Save again after it loads.':`Week ${C.L.wk} saved for signed-in viewers, but the public page was not updated (${c}).`+note,true); btn.disabled=false; } });

/* ---------- master workbooks ---------- */
const MW={city:null,kec:null,reg:null,ga:null,png:null};
function mwKind(n){n=n.toLowerCase().replace(/[\s\-]+/g,'_');
  if(/\.(png|jpe?g)$/.test(n))return 'png'; if(/\.zip$/.test(n))return 'zip';
  if(/\.xlsb$/.test(n)){ if(/fb_?city/.test(n))return 'city'; if(/fb_?kec/.test(n))return 'kec'; }
  if(/\.xlsx$/.test(n)){ if(/regional/.test(n))return 'reg'; if(/gross_?add/.test(n))return 'ga'; }
  return null;}
async function mwUnzip(f){ const z=await JSZip.loadAsync(f); const out=[];
  for(const n of Object.keys(z.files)){ if(z.files[n].dir||/__MACOSX|\/\./.test(n)) continue; const base=n.split('/').pop(); if(!/\.(xlsb|xlsx|png)$/i.test(base)) continue;
    out.push(new File([await z.files[n].async('blob')],base)); }
  return out; }
function mwRender(){ const ul=$('#mw-list'); ul.textContent=''; const lab={city:'fb_city',kec:'fb_kec',reg:'Fb_share_regional',ga:'fb_gross_add',png:'Gross adds screenshot'};
  for(const k of Object.keys(lab)) ul.append(el('li',{class:MW[k]?'ok':'muted'},`${lab[k]}: ${MW[k]?MW[k].name+' ('+(MW[k].size/1e6).toFixed(1)+' MB)':'not added'}`));
  const any=!!(MW.city||MW.kec||MW.reg||MW.ga); $('#mw-run').disabled=!(C&&any);
  $('#mw-prog').textContent=!any?'Add the workbooks to start.':(C.fromHistory?`Week ${C.L.d} ${FB.MON[C.L.m-1]} from saved history: workbooks that already contain it get their pivots moved and are exported; workbooks from an older week need this week’s CSVs above.`:`Ready to update to ${C.L.d} ${FB.MON[C.L.m-1]} ${C.L.y}.`); }
$('#mw-file').addEventListener('change',async e=>{ const notes=[]; const files=[...e.target.files];
  for(const f0 of files){ let list=[f0];
    if(mwKind(f0.name)==='zip'){ try{ $('#mw-prog').textContent='Opening '+f0.name+'…'; list=await mwUnzip(f0); if(!list.length) notes.push(`${f0.name}: no workbook inside.`); else notes.push(`${f0.name}: opened, using ${list.map(x=>x.name).join(', ')}.`);}catch(err){ notes.push(`${f0.name}: could not open this zip (${err.message}).`); continue; } }
    for(const f of list){ const k=mwKind(f.name); if(k&&k!=='zip') MW[k]=f; else notes.push(`${f.name}: not recognised. Expected fb_city….xlsb, fb_kec….xlsb, Fb_share_regional….xlsx, fb_gross_add….xlsx or a .png.`); } }
  e.target.value=''; mwRender(); const lg=$('#mw-log'); lg.textContent=''; notes.forEach(t=>lg.append(el('li',{class:/not recognised|could not|no workbook/.test(t)?'alert':''},t))); });
async function saveBlob(name,blob){ (window.__saved=window.__saved||{})[name]=blob; if(!downloads) return 'nodl';
  let fname=name, data=blob;
  if(/\.xlsb$/i.test(name)){ /* claude.ai does not allow .xlsb downloads: deliver the workbook inside a .zip */ const z=new JSZip(); z.file(name,blob,{binary:true,createFolders:false}); data=await z.generateAsync({type:'blob',compression:'STORE'}); fname=name.replace(/\.xlsb$/i,'.zip'); }
  try{ await downloads.save({filename:fname,data}); return fname===name?'saved':'savedzip'; }catch(e){ return e.code||'error'; } }
function dlButton(name){ return el('button',{class:'sec',type:'button','data-f':name},'Download '+name); }
document.addEventListener('click',async e=>{const b=e.target.closest&&e.target.closest('button[data-f]'); if(!b) return; const n=b.dataset.f; const st=await saveBlob(n,window.__saved[n]); b.textContent=st==='saved'?'Saved '+n:st==='savedzip'?'Saved '+n.replace(/\.xlsb$/i,'.zip'):(st==='declined'?'Download '+n:`Could not save (${st})`);});
$('#mw-run').addEventListener('click',async()=>{ if(!C) return; const btn=$('#mw-run'); btn.disabled=true; const logUl=$('#mw-log'); logUl.textContent='';
  const prog=t=>{$('#mw-prog').textContent=t+'…';}; const L=C.L; const ym=L.ym;
  const jobs=[['city',`fb_city_${ym}.xlsb`,(f,log)=>MASTER.updateCity(f,C,log,prog)],['kec',`fb_kec_${ym}.xlsb`,(f,log)=>MASTER.updateKec(f,C,log,prog)],
    ['reg','Fb_share_regional.xlsx',(f,log)=>MASTER.updateRegional(f,C,MW.png,log,prog)],['ga','fb_gross_add.xlsx',(f,log)=>MASTER.updateGA(f,C,log,prog)]];
  for(const [k,name,fn] of jobs){ if(!MW[k]) continue; const log=[]; const t0=performance.now();
    try{
      const blob=await fn(MW[k],log); prog(`saving ${name}`); const st=await saveBlob(name,blob);
      log.forEach(t=>logUl.append(el('li',{},t)));
      const li=el('li',{class:st==='saved'||st==='savedzip'?'ok':'alert'},`${name}: ${st==='saved'?'saved':st==='savedzip'?'saved as '+name.replace(/\.xlsb$/i,'.zip')+' (extract it to get the .xlsb)':st==='nodl'?'ready, but downloads are not available in this view (open the page in claude.ai)':st==='declined'?'download was declined':'not saved ('+st+')'} · ${((performance.now()-t0)/1000).toFixed(0)} s · ${(blob.size/1e6).toFixed(1)} MB `);
      if(st!=='saved'&&st!=='savedzip') li.append(dlButton(name)); logUl.append(li);
    }catch(e){ log.forEach(t=>logUl.append(el('li',{},t))); logUl.append(el('li',{class:'alert'},`${name}: ${e.message}`)); } }
  $('#mw-prog').textContent='Done. Each file was offered for download; if a prompt was missed, use its Download button below. Open each file in Excel once so the pivots refresh, then save it for next week.'; btn.disabled=false; });
window.__mwSet=(k,f)=>{MW[k]=f; mwRender();};

buildMeta(); applyRole(); show(); run();
})();
