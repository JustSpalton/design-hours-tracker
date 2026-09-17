let breakdownState={records:[]};
let breakdownView='status';

function breakdownInitials(name){return String(name||'').trim().split(/\s+/).filter(Boolean).map(x=>x[0]).join('').toUpperCase()}
function resolveBreakdownDesigner(imported){
  const exact=state.designers.find(d=>norm(d.name)===norm(imported));if(exact)return exact.name;
  const ini=breakdownInitials(imported);const matches=state.designers.filter(d=>breakdownInitials(d.name)===ini);
  return matches.length===1?matches[0].name:null;
}
function normalProduct(value){const v=String(value||'').replace(/\s+/g,' ').trim();if(!v)return '';if(/^posi\s*[- ]?\s*joist$/i.test(v))return 'Posi Joist';if(/^i\s*[- ]?\s*joist$/i.test(v))return 'I Joist';return v}
function normalStatus(value){const v=String(value||'').replace(/\s+/g,' ').trim();if(!v)return '';if(/^budget(?:\s+quote)?$/i.test(v))return 'Budget';if(/^otp$/i.test(v))return 'OTP';if(/^amendment$/i.test(v))return 'Amendment';return v}
function addCount(obj,key){if(key)obj[key]=(obj[key]||0)+1}
function breakdownColumns(ws,range){
  let productCol=null,estimatorCol=null,statusCol=null;
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+20);r++)for(let c=range.s.c;c<=range.e.c;c++){
    const t=cellText(ws[XLSX.utils.encode_cell({r,c})]);
    if(productCol==null&&/^product$/i.test(t))productCol=c;
    if(estimatorCol==null&&/^estimator\s*name$/i.test(t))estimatorCol=c;
    if(statusCol==null&&/^status$/i.test(t))statusCol=c;
  }
  return {productCol:productCol??3,estimatorCol:estimatorCol??7,statusCol:statusCol??10};
}
function extractBreakdowns(wb,week,sourceFile){
  const combined=new Map();
  for(const sheetName of wb.SheetNames){
    const ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])continue;const range=XLSX.utils.decode_range(ws['!ref']);
    const {productCol,estimatorCol,statusCol}=breakdownColumns(ws,range);const starts=[];
    for(let r=range.s.r;r<=range.e.r;r++){const name=sectionNameAtRow(ws,r,range);if(name)starts.push({r,name})}
    for(let i=0;i<starts.length;i++){
      const start=starts[i],end=i+1<starts.length?starts[i+1].r-1:range.e.r;const designer=resolveBreakdownDesigner(start.name);if(!designer)continue;
      const key=`${norm(designer)}|${week}`;if(!combined.has(key))combined.set(key,{designer,week,products:{},statuses:{},source_file:sourceFile});const rec=combined.get(key);
      for(let r=start.r+1;r<=end;r++){
        const rowEstimator=cellText(ws[XLSX.utils.encode_cell({r,c:estimatorCol})]);if(norm(rowEstimator)!==norm(start.name))continue;
        addCount(rec.products,normalProduct(cellText(ws[XLSX.utils.encode_cell({r,c:productCol})])));
        addCount(rec.statuses,normalStatus(cellText(ws[XLSX.utils.encode_cell({r,c:statusCol})])));
      }
    }
  }
  return [...combined.values()];
}
async function loadBreakdowns(){try{const res=await fetch('/api/breakdowns',{cache:'no-store'});if(res.ok)breakdownState=await res.json()}catch(_){breakdownState={records:[]}}}
async function saveBreakdownsFromFiles(files){
  const excel=[...files].filter(isExcelFile);if(!excel.length)return;const records=[];
  for(const file of excel){try{const d=parseWeekFilename(file.name);if(!d)continue;const week=isoDate(d);const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});records.push(...extractBreakdowns(wb,week,file.name))}catch(e){console.warn('Breakdown import skipped',file.name,e)}}
  if(!records.length)return;
  try{const res=await fetch('/api/breakdowns',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({records})});if(res.ok){breakdownState=await res.json();renderDetail()}}catch(e){console.warn('Breakdown save failed',e)}
}
function breakdownWeeks(){return typeof rangeWeeks==='function'?rangeWeeks():(selectedWeek?[selectedWeek]:[])}
function breakdownTotals(designer,field){
  const weeks=new Set(breakdownWeeks()),out={};
  for(const rec of breakdownState.records||[]){if(norm(rec.designer)!==norm(designer)||!weeks.has(rec.week))continue;for(const [k,v] of Object.entries(rec[field]||{}))out[k]=(out[k]||0)+Number(v||0)}
  return out;
}
function barRows(items){
  if(!items.length)return '<div class="breakdown-empty">No breakdown data for this period. Re-import the weekly estimator spreadsheet to add it.</div>';
  const max=Math.max(1,...items.map(x=>x[1]));return `<div class="breakdown-bars">${items.map(([label,count])=>`<div class="breakdown-row"><div class="breakdown-label">${escapeHtml(label)}</div><div class="breakdown-track"><div class="breakdown-fill" style="width:${Math.max(4,count/max*100)}%"></div></div><div class="breakdown-count">${count}</div></div>`).join('')}</div>`;
}
function appendBreakdownCard(){
  if(!selectedDesigner)return;const body=document.querySelector('#detailPanel .detail-body');if(!body||document.getElementById('workMixCard'))return;
  const statuses=breakdownTotals(selectedDesigner,'statuses');const products=breakdownTotals(selectedDesigner,'products');
  const statusItems=Object.entries(statuses).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  const productItems=[['I Joist',products['I Joist']||0],['Posi Joist',products['Posi Joist']||0]];
  const card=document.createElement('div');card.className='trend-card work-mix-card';card.id='workMixCard';
  const period=typeof rangeLabel==='function'?rangeLabel():`W/C ${fmtDate(selectedWeek)}`;
  card.innerHTML=`<div class="trend-head work-mix-head"><div><h3>Work mix</h3><span>Number of jobs • ${escapeHtml(period)}</span></div><div class="mix-tabs"><button type="button" data-mix="status" class="${breakdownView==='status'?'active':''}">Job type</button><button type="button" data-mix="product" class="${breakdownView==='product'?'active':''}">Product</button></div></div><div class="work-mix-body">${breakdownView==='product'?barRows(productItems):barRows(statusItems)}</div>`;
  const weeklyTrend=body.querySelector('.trend-card:not(.holiday-card)');if(weeklyTrend)body.insertBefore(card,weeklyTrend);else body.appendChild(card);
  card.querySelectorAll('[data-mix]').forEach(btn=>btn.addEventListener('click',()=>{breakdownView=btn.dataset.mix;renderDetail()}));
}
const renderDetailBeforeBreakdowns=renderDetail;
renderDetail=function(){renderDetailBeforeBreakdowns();appendBreakdownCard()};

const importInput=document.getElementById('importFiles');if(importInput)importInput.addEventListener('change',e=>saveBreakdownsFromFiles(e.target.files));
const breakdownDrop=document.getElementById('dropZone');if(breakdownDrop)breakdownDrop.addEventListener('drop',e=>saveBreakdownsFromFiles(e.dataTransfer?.files||[]));
loadBreakdowns().then(()=>renderDetail());
