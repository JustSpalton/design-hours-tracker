let breakdownState={records:[]};
let breakdownView='status';
let breakdownMetric='jobs';

function breakdownInitials(name){return String(name||'').trim().split(/\s+/).filter(Boolean).map(x=>x[0]).join('').toUpperCase()}
function resolveBreakdownDesigner(imported){
  const exact=state.designers.find(d=>norm(d.name)===norm(imported));if(exact)return exact.name;
  const ini=breakdownInitials(imported);const matches=state.designers.filter(d=>breakdownInitials(d.name)===ini);
  return matches.length===1?matches[0].name:null;
}
function normalProduct(value){const v=String(value||'').replace(/\s+/g,' ').trim();if(!v)return '';if(/^posi\s*[- ]?\s*joist$/i.test(v))return 'Posi Joist';if(/^i\s*[- ]?\s*joist$/i.test(v))return 'I Joist';return v}
function normalStatus(value){const v=String(value||'').replace(/\s+/g,' ').trim();if(!v)return '';if(/^budget(?:\s+quote)?$/i.test(v))return 'Budget';if(/^otp$/i.test(v))return 'OTP';if(/^amendment$/i.test(v))return 'Amendment';return v}
function addCount(obj,key){if(key)obj[key]=(obj[key]||0)+1}
function addHours(obj,key,value){if(key&&Number.isFinite(value))obj[key]=Math.round(((obj[key]||0)+Number(value))*100)/100}
function breakdownColumns(ws,range){
  let productCol=null,statusCol=null;
  const report=getReportColumns(ws,range);
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+20);r++)for(let c=range.s.c;c<=range.e.c;c++){
    const t=cellText(ws[XLSX.utils.encode_cell({r,c})]);
    if(productCol==null&&/^product$/i.test(t))productCol=c;
    if(statusCol==null&&/^status$/i.test(t))statusCol=c;
  }
  return {productCol:productCol??3,estimatorCol:report.estimatorCol,statusCol:statusCol??10,hoursCol:report.hoursCol};
}
function extractBreakdowns(wb,week,sourceFile){
  const combined=new Map();
  for(const sheetName of wb.SheetNames){
    const ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])continue;const range=XLSX.utils.decode_range(ws['!ref']);
    const {productCol,estimatorCol,statusCol,hoursCol}=breakdownColumns(ws,range);const starts=[];
    for(let r=range.s.r;r<=range.e.r;r++){const name=sectionNameAtRow(ws,r,range);if(name)starts.push({r,name})}
    for(let i=0;i<starts.length;i++){
      const start=starts[i],end=i+1<starts.length?starts[i+1].r-1:range.e.r;const designer=resolveBreakdownDesigner(start.name);if(!designer)continue;
      const key=`${norm(designer)}|${week}`;
      if(!combined.has(key))combined.set(key,{designer,week,products:{},statuses:{},product_hours:{},status_hours:{},source_file:sourceFile});
      const rec=combined.get(key);
      for(let r=start.r+1;r<=end;r++){
        const rowEstimator=cellText(ws[XLSX.utils.encode_cell({r,c:estimatorCol})]);if(norm(rowEstimator)!==norm(start.name))continue;
        const product=normalProduct(cellText(ws[XLSX.utils.encode_cell({r,c:productCol})]));
        const status=normalStatus(cellText(ws[XLSX.utils.encode_cell({r,c:statusCol})]));
        const hours=numericHours(ws[XLSX.utils.encode_cell({r,c:hoursCol})]);
        addCount(rec.products,product);addCount(rec.statuses,status);
        if(hours!=null&&hours>=0&&hours<=1000){addHours(rec.product_hours,product,hours);addHours(rec.status_hours,status,hours)}
      }
    }
  }
  return [...combined.values()];
}
async function loadBreakdowns(){try{breakdownState=await api('/api/breakdowns')}catch(_){breakdownState={records:[]}}}
async function saveBreakdownsFromFiles(files){
  const excel=[...files].filter(isExcelFile);if(!excel.length)return;const records=[];
  for(const file of excel){try{const d=parseWeekFilename(file.name);if(!d)continue;const week=isoDate(d);const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellFormula:true,cellNF:true,cellText:true});records.push(...extractBreakdowns(wb,week,file.name))}catch(e){console.warn('Breakdown import skipped',file.name,e)}}
  if(!records.length)return;
  try{
    const res=await fetch('/api/breakdowns',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({records})});
    if(res.ok){
      breakdownState=await res.json();renderDetail();
      if(typeof renderProductTotals==='function')renderProductTotals();
      if(typeof refreshReportPeriods==='function'&&!document.getElementById('managerReportModal')?.hidden)refreshReportPeriods();
    }
  }catch(e){console.warn('Breakdown save failed',e)}
}
function breakdownWeeks(){return typeof rangeWeeks==='function'?rangeWeeks():(selectedWeek?[selectedWeek]:[])}
function breakdownTotals(designer,field){
  const weeks=new Set(breakdownWeeks()),out={};
  for(const rec of breakdownState.records||[]){if(norm(rec.designer)!==norm(designer)||!weeks.has(rec.week))continue;for(const [k,v] of Object.entries(rec[field]||{}))out[k]=(out[k]||0)+Number(v||0)}
  return out;
}
function breakdownHourCoverage(designer,field){
  const weeks=new Set(breakdownWeeks()),hourField=field==='products'?'product_hours':'status_hours';
  const rows=(breakdownState.records||[]).filter(rec=>norm(rec.designer)===norm(designer)&&weeks.has(rec.week)&&Object.keys(rec[field]||{}).length);
  return {rows:rows.length,withHours:rows.filter(rec=>Object.prototype.hasOwnProperty.call(rec,hourField)).length};
}
function barRows(items,metric='jobs',coverage=null){
  if(metric==='hours'&&coverage&&coverage.rows&&!coverage.withHours)return '<div class="breakdown-empty">Hours were not stored for this period. Re-import the weekly estimator spreadsheet to add hours by Product and Job Type.</div>';
  if(!items.length||!items.some(x=>Number(x[1])>0))return '<div class="breakdown-empty">No breakdown data for this period. Re-import the weekly estimator spreadsheet to add it.</div>';
  const visible=items.filter(([,v])=>Number(v)>0),max=Math.max(1,...visible.map(x=>Number(x[1])));
  return `<div class="breakdown-bars">${visible.map(([label,value])=>`<div class="breakdown-row"><div class="breakdown-label">${escapeHtml(label)}</div><div class="breakdown-track"><div class="breakdown-fill" style="width:${Math.max(4,Number(value)/max*100)}%"></div></div><div class="breakdown-count">${metric==='hours'?`${Number(value).toFixed(1)}h`:Number(value)}</div></div>`).join('')}</div>`;
}
function appendBreakdownCard(){
  if(!selectedDesigner)return;const body=document.querySelector('#detailPanel .detail-body');if(!body||document.getElementById('workMixCard'))return;
  const countField=breakdownView==='product'?'products':'statuses',hoursField=breakdownView==='product'?'product_hours':'status_hours';
  const counts=breakdownTotals(selectedDesigner,countField),hours=breakdownTotals(selectedDesigner,hoursField);
  const values=breakdownMetric==='hours'?hours:counts;
  let items=Object.entries(values).filter(([,v])=>Number(v)>0).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  if(breakdownView==='product'){
    const keys=[...new Set(['I Joist','Posi Joist',...Object.keys(values)])];
    items=keys.map(k=>[k,Number(values[k]||0)]).filter(([,v])=>v>0);
  }
  const card=document.createElement('div');card.className='trend-card work-mix-card';card.id='workMixCard';
  const period=typeof rangeLabel==='function'?rangeLabel():`W/C ${fmtDate(selectedWeek)}`;
  const coverage=breakdownHourCoverage(selectedDesigner,countField);
  card.innerHTML=`<div class="trend-head work-mix-head"><div><h3>Work mix</h3><span>${breakdownMetric==='hours'?'Design hours':'Number of jobs'} • ${escapeHtml(period)}</span></div><div class="mix-controls"><div class="mix-tabs"><button type="button" data-mix="status" class="${breakdownView==='status'?'active':''}">Job type</button><button type="button" data-mix="product" class="${breakdownView==='product'?'active':''}">Product</button></div><div class="mix-tabs metric-tabs"><button type="button" data-metric="jobs" class="${breakdownMetric==='jobs'?'active':''}">Jobs</button><button type="button" data-metric="hours" class="${breakdownMetric==='hours'?'active':''}">Hours</button></div></div></div><div class="work-mix-body">${barRows(items,breakdownMetric,coverage)}</div>`;
  const weeklyTrend=body.querySelector('.trend-card:not(.holiday-card)');if(weeklyTrend)body.insertBefore(card,weeklyTrend);else body.appendChild(card);
  card.querySelectorAll('[data-mix]').forEach(btn=>btn.addEventListener('click',()=>{breakdownView=btn.dataset.mix;renderDetail()}));
  card.querySelectorAll('[data-metric]').forEach(btn=>btn.addEventListener('click',()=>{breakdownMetric=btn.dataset.metric;renderDetail()}));
}
const renderDetailBeforeBreakdowns=renderDetail;
renderDetail=function(){renderDetailBeforeBreakdowns();appendBreakdownCard()};

loadBreakdowns().then(()=>renderDetail());
