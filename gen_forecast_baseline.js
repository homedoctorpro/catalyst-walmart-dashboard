// One-time generator: freeze the base/downside/upside monthly projections as the
// plan-of-record, so future weeks can be tracked against it (Plan vs Actual).
// Re-run only when intentionally re-baselining. Reads the live dashboard.html data.
const fs = require("fs");

const html = fs.readFileSync("dashboard.html", "utf8");
const start = html.indexOf("const DATA = ");
let i = html.indexOf("{", start), depth = 0, end = -1;
for (let j = i; j < html.length; j++) {
  const ch = html[j];
  if (ch === "{") depth++;
  else if (ch === "}") { depth--; if (depth === 0) { end = j + 1; break; } }
}
const DATA = JSON.parse(html.slice(i, end));

const SKU_SHORT = { "CATALYST15ORIG":"15O","CATALYST15UNSCEN":"15U","CATALYST34LBORIGINAL":"34O","CATALYSTPET34LBUNSCE":"34U" };
const SKUS = DATA.skus;
const MONTH = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const FC_ASP = { "15":9.18, "34":12.98 };
const WPM = 52/12;
const HORIZON = 24;

function fridays(y,mi){let c=0;const d=new Date(Date.UTC(y,mi,1));const last=new Date(Date.UTC(y,mi+1,0)).getUTCDate();for(let day=1;day<=last;day++){d.setUTCDate(day);if(d.getUTCDay()===5)c++;}return c;}
function size(s){return /15/.test(s)?"15":"34";}
function nextKey(k){const[y,mm]=k.split("-").map(Number);return mm===12?`${y+1}-01`:`${y}-${String(mm+1).padStart(2,"0")}`;}
function mlabel(k){const[y,mm]=k.split("-").map(Number);return `${MONTH[mm-1]} ${String(y).slice(2)}`;}

function monthlyRollup(){
  const wk=DATA.week_dates||{};const groups={};
  DATA.weeks.forEach(w=>{const iso=wk[w];if(!iso)return;const[y,mm]=iso.split("-").map(Number);const key=`${y}-${String(mm).padStart(2,"0")}`;(groups[key]=groups[key]||[]).push(w);});
  const months=Object.keys(groups).sort();const byMonth={};
  months.forEach(key=>{const[ys,ms]=key.split("-");const year=+ys,mi=+ms-1;const weeks=groups[key];const exp=fridays(year,mi);const isPartial=weeks.length<exp;const series={};
    SKUS.forEach(s=>{let q=0,has=false;weeks.forEach(w=>{const m=(DATA.metrics[w]||{})[s];if(m&&m.pos_qty!=null){q+=m.pos_qty;has=true;}});series[s]={pos_qty:has?q:null};});
    byMonth[key]={isPartial,series};});
  return {months,byMonth};
}
function cumTxns(p){const c=p.customers;const n=p.total_customers.length;const o=[];for(let i=0;i<n;i++)o.push((c["1"][i]||0)+(c["2"][i]||0)*2+(c["3"][i]||0)*3+(c["4"][i]||0)*4+(c["5"][i]||0)*5+(c["6+"][i]||0)*6);return o;}
function wnr(p){const cum=cumTxns(p);const tc=p.total_customers;const nT=[],rT=[];for(let i=0;i<tc.length;i++){const nc=Math.max(0,i===0?(tc[i]||0):((tc[i]||0)-(tc[i-1]||0)));const tt=Math.max(0,i===0?(cum[i]||0):((cum[i]||0)-(cum[i-1]||0)));nT.push(nc);rT.push(Math.max(0,tt-nc));}return{nT,rT};}
function shares(){const tr=DATA.trial_repeat;if(!tr||!tr.weeks)return{};const bm={};["15O","15U","34O","34U"].forEach(sku=>{const p=tr.products&&tr.products[sku];if(!p)return;const nr=wnr(p);tr.weeks.forEach((iso,i)=>{const[y,mm]=iso.split("-");const k=`${y}-${mm}`;bm[k]=bm[k]||{};bm[k][sku]=bm[k][sku]||{n:0,r:0};bm[k][sku].n+=nr.nT[i]||0;bm[k][sku].r+=nr.rT[i]||0;});});const sh={};Object.keys(bm).forEach(k=>{sh[k]={};Object.keys(bm[k]).forEach(sku=>{const d=bm[k][sku];const t=d.n+d.r;sh[k][sku]=t>0?d.n/t:null;});});return sh;}
function buildHist(){const monthly=monthlyRollup();const sh=shares();const latest={};Object.keys(sh).sort().forEach(k=>["15O","15U","34O","34U"].forEach(s=>{if(sh[k][s]!=null)latest[s]=sh[k][s];}));const months=[];const byMonth={};
  monthly.months.forEach(key=>{const mo=monthly.byMonth[key];if(mo.isPartial)return;months.push(key);byMonth[key]={};SKUS.forEach(ls=>{const ss=SKU_SHORT[ls];const tot=(mo.series[ls]&&mo.series[ls].pos_qty)||0;let s=(sh[key]||{})[ss];if(s==null)s=latest[ss];if(s==null)s=1;const nu=tot*s;byMonth[key][ls]={newUnits:nu,repeatUnits:tot-nu,total:tot};});});
  return {months,byMonth};
}
function placements(){const lw=DATA.store_weeks[DATA.store_weeks.length-1];const st=DATA.weekly_stores[lw]||{};const by=Object.fromEntries(SKUS.map(s=>[s,0]));Object.values(st).forEach(sd=>Object.keys(sd.skus||{}).forEach(sku=>{if(sku in by)by[sku]++;}));return {by,total:SKUS.reduce((a,s)=>a+by[s],0),week:lw};}

const hist=buildHist();
const PL=placements();

function forecast(a){const months=hist.months.slice();const newBy={},repBy={},totBy={};SKUS.forEach(s=>{newBy[s]=[];repBy[s]=[];totBy[s]=[];});
  hist.months.forEach(k=>SKUS.forEach(sku=>{const d=hist.byMonth[k][sku];newBy[sku].push(d.newUnits);repBy[sku].push(d.repeatUnits);totBy[sku].push(d.total);}));
  const floor={};SKUS.forEach(sku=>{floor[sku]=(newBy[sku][newBy[sku].length-1]||0)*a.new_floor;});
  let cursor=months[months.length-1];
  for(let f=0;f<HORIZON;f++){cursor=nextKey(cursor);months.push(cursor);
    SKUS.forEach(sku=>{const arr=newBy[sku];const prev=arr[arr.length-1]||0;const nv=floor[sku]+(prev-floor[sku])*(1-a.d_new);arr.push(nv);
      const idx=arr.length-1;let rep=0;for(let k=1;;k++){const ci=idx-2*k;if(ci<0)break;rep+=(arr[ci]||0)*a.r_initial*Math.pow(1-a.r_decay,k-1);}repBy[sku].push(rep);totBy[sku].push(nv+rep);});}
  // monthly totals
  const out=months.map((key,idx)=>{let n=0,r=0,t=0,w=0;SKUS.forEach(sku=>{const tt=totBy[sku][idx]||0;n+=newBy[sku][idx]||0;r+=repBy[sku][idx]||0;t+=tt;w+=tt*FC_ASP[size(sku)];});
    return {key,label:mlabel(key),units:Math.round(t),new:Math.round(n),repeat:Math.round(r),wholesale:Math.round(w),usw:+(t/PL.total/WPM).toFixed(4)};});
  return {historicalCount:hist.months.length,months:out};
}

const PRESETS={downside:{decay:3,floor:30,repeat:50,rdecay:7.5},base:{decay:2,floor:50,repeat:50,rdecay:5},upside:{decay:1,floor:85,repeat:60,rdecay:4}};
function toA(p){return {d_new:p.decay/100,new_floor:p.floor/100,r_initial:p.repeat/100,r_decay:p.rdecay/100};}

const fcBase=forecast(toA(PRESETS.base));
const hc=fcBase.historicalCount;
const lastHist=hist.months[hist.months.length-1];
const scenarios={};
for(const name of ["downside","base","upside"]){
  const fc=forecast(toA(PRESETS[name]));
  scenarios[name]={preset:PRESETS[name],months:fc.months.slice(hc)};  // forecast months only
}
const baseline={
  meta:{
    frozen_as_of:new Date().toISOString().slice(0,10),
    generated_from_week:DATA.store_weeks[DATA.store_weeks.length-1],
    last_complete_month:lastHist,
    last_complete_month_label:mlabel(lastHist),
    placements:{total:PL.total,week:PL.week},
    weeks_per_month:WPM,
    horizon_months:HORIZON,
    repeat_basis:"manual cohort presets (downside/base/upside)",
    note:"Plan-of-record frozen scenarios. Compare live monthly actuals against these lines."
  },
  scenarios
};
fs.writeFileSync("forecast_baseline.json",JSON.stringify(baseline,null,2));
console.log("Wrote forecast_baseline.json");
console.log("frozen_as_of",baseline.meta.frozen_as_of,"| from week",baseline.meta.generated_from_week,"| last complete",baseline.meta.last_complete_month_label,"| placements",PL.total);
["downside","base","upside"].forEach(n=>{const m=scenarios[n].months;console.log(`${n}: ${m[0].label} units=${m[0].units} wholesale=$${m[0].wholesale} usw=${m[0].usw}  ...  ${m[11].label} units=${m[11].units}`);});
