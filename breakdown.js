let breakdownState={records:[]};
let breakdownView='product';
let breakdownMetric='jobs';

function breakdownInitials(name){return String(name||'').trim().split(/\s+/).filter(Boolean).map(x=>x[0]).join('').toUpperCase()}
function canonicalBreakdownDesignerName(name){const clean=String(name||'').replace(/\s+/g,' ').trim();const aliases={'jon wilson':'Jonathan Wilson','kerry mui':'Kerry Gardiner'};return aliases[clean.toLowerCase()]||clean}
function resolveBreakdownDesigner(imported){
  const canonical=canonicalBreakdownDesignerName(imported);
  const exact=state.designers.find(d=>norm(d.name)===norm(canonical));if(exact)return exact.name;
  const ini=breakdownInitials(imported);const matches=state.designers.filter(d=>breakdownInitials(d.name)===ini);
  return matches.length===1?matches[0].name:null;
}
function normalProduct(value){
  const v=String(value||'').replace(/\s+/g,' ').trim();if(!v)return '';
  if(/\bposi\s*[- ]?\s*joist\b/i.test(v)||/^posi$/i.test(v))return 'Posi Joist';
  if(/\bi\s*[- ]?\s*joist\b/i.test(v)||/^ijoist$/i.test(v))return 'I Joist';
  return '';
}
function normalStatus(value){
  const v=String(value||'').replace(/\s+/g,' ').trim();if(!v)return '';
  if(/\bbudget(?:\s+quote)?\b/i.test(v))return 'Budget';
  if(/\botp\b/i.test(v))return 'OTP';
  if(/\bamendment\b/i.test(v))return 'Amendment';
  return '';
}
function addCount(obj,key){if(key)obj[key]=(obj[key]||0)+1}
function addHours(obj,key,value){if(key&&Number.isFinite(value))obj[key]=Math.round(((obj[key]||0)+Number(value))*100)/100}
function breakdownColumns(ws,range){
  let productCol=null,statusCol=null;
  const report=getReportColumns(ws,range);
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+60);r++)for(let c=range.s.c;c<=range.e.c;c++){
    const t=cellText(ws[XLSX.utils.encode_cell({r,c})]).replace(/\s+/g,' ').trim();
    if(productCol==null&&/^(?:product|product type|joist type|joist system|system)$/i.test(t))productCol=c;
    if(statusCol==null&&/^(?:status|job type|order type|quote type)$/i.test(t))statusCol=c;
  }
  return {productCol,estimatorCol:report.estimatorCol,statusCol,hoursCol:report.hoursCol,dateCol:dateEnquiryDoneColumn(ws,range)};
}
function rowBreakdownValue(ws,r,range,preferredCol,normaliser){
  if(preferredCol!=null){
    const preferred=normaliser(cellText(ws[XLSX.utils.encode_cell({r,c:preferredCol})]));
    if(preferred)return preferred;
  }
  for(let c=range.s.c;c<=range.e.c;c++){
    if(c===preferredCol)continue;
    const found=normaliser(cellText(ws[XLSX.utils.encode_cell({r,c})]));
    if(found)return found;
  }
  return '';
}
function extractBreakdowns(wb,defaultWeek,sourceFile){
  const combined=new Map();
  for(const sheetName of wb.SheetNames){
    const ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])continue;const range=XLSX.utils.decode_range(ws['!ref']);
    const {productCol,estimatorCol,statusCol,hoursCol,dateCol}=breakdownColumns(ws,range);const starts=[];
    for(let r=range.s.r;r<=range.e.r;r++){const name=sectionNameAtRow(ws,r,range);if(name)starts.push({r,name})}
    for(let i=0;i<starts.length;i++){
      const start=starts[i],end=i+1<starts.length?starts[i+1].r-1:range.e.r;const designer=resolveBreakdownDesigner(start.name);if(!designer)continue;
      for(let r=start.r+1;r<=end;r++){
        const rowEstimator=cellText(ws[XLSX.utils.encode_cell({r,c:estimatorCol})]);
        if(rowEstimator&&norm(rowEstimator)!==norm(start.name))continue;
        const date=dateCol!=null?cellISODate(ws[XLSX.utils.encode_cell({r,c:dateCol})]):null;
        const week=date?weekForDate(date):defaultWeek;
        if(!week)continue;
        const product=rowBreakdownValue(ws,r,range,productCol,normalProduct);
        const status=rowBreakdownValue(ws,r,range,statusCol,normalStatus);
        if(!product&&!status)continue;
        const key=`${norm(designer)}|${week}`;
        if(!combined.has(key))combined.set(key,{designer,week,products:{},statuses:{},product_hours:{},status_hours:{},source_file:sourceFile});
        const rec=combined.get(key);
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
  for(const file of excel){
    try{
      const d=parseWeekFilename(file.name);
      const defaultWeek=d&&d.getUTCDay()===1?isoDate(d):null;
      const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellFormula:true,cellNF:true,cellText:true});
      records.push(...extractBreakdowns(wb,defaultWeek,file.name));
    }catch(e){console.warn('Breakdown import skipped',file.name,e)}
  }
  if(!records.length)return;
  try{
    breakdownState=await api('/api/breakdowns',{method:'POST',body:JSON.stringify({records})});
    renderDetail();
    if(typeof renderProductTotals==='function')renderProductTotals();
    if(typeof refreshReportPeriods==='function'&&!document.getElementById('managerReportModal')?.hidden)refreshReportPeriods();
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
  card.innerHTML=`<div class="trend-head work-mix-head"><div><h3>Work mix</h3><span>${breakdownMetric==='hours'?'Design hours':'Number of jobs'} • ${escapeHtml(period)}</span></div><div class="mix-controls"><div class="mix-tabs"><button type="button" data-mix="product" class="${breakdownView==='product'?'active':''}">Job type</button><button type="button" data-mix="status" class="${breakdownView==='status'?'active':''}">Work category</button></div><div class="mix-tabs metric-tabs"><button type="button" data-metric="jobs" class="${breakdownMetric==='jobs'?'active':''}">Jobs</button><button type="button" data-metric="hours" class="${breakdownMetric==='hours'?'active':''}">Hours</button></div></div></div><div class="work-mix-body">${barRows(items,breakdownMetric,coverage)}</div>`;
  const weeklyTrend=body.querySelector('.trend-card:not(.holiday-card)');if(weeklyTrend)body.insertBefore(card,weeklyTrend);else body.appendChild(card);
  card.querySelectorAll('[data-mix]').forEach(btn=>btn.addEventListener('click',()=>{breakdownView=btn.dataset.mix;renderDetail()}));
  card.querySelectorAll('[data-metric]').forEach(btn=>btn.addEventListener('click',()=>{breakdownMetric=btn.dataset.metric;renderDetail()}));
}
const renderDetailBeforeBreakdowns=renderDetail;
renderDetail=function(){renderDetailBeforeBreakdowns();appendBreakdownCard()};

const importManyBeforeBreakdowns=importMany;
importMany=async function(files){
  await importManyBeforeBreakdowns(files);
  await saveBreakdownsFromFiles(files);
  await loadBreakdowns();
  renderDetail();
  if(typeof renderProductTotals==='function')renderProductTotals();
};

