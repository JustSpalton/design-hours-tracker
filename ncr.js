let ncrState={records:[],sourceFile:'',importedAt:null};
let ncrLoaded=false;
let ncrLoading=false;
let ncrFilters={period:'12',search:'',category:'',employee:'',customer:'',status:'all'};
let selectedNcr=null;

function ncrMoney(value){return new Intl.NumberFormat('en-GB',{style:'currency',currency:'GBP',maximumFractionDigits:2}).format(Number(value||0))}
function ncrNormHeader(value){return String(value??'').trim().toLowerCase().replace(/[^a-z0-9]+/g,' ').trim()}
function ncrText(cell){return String(cell?.w??cell?.v??'').replace(/\s+/g,' ').trim()}
function ncrNumber(cell){
  if(!cell)return 0;
  if(typeof cell.v==='number'&&Number.isFinite(cell.v))return Math.round(cell.v*100)/100;
  const cleaned=String(cell.v??cell.w??'').replace(/[£,$]/g,'').replace(/\s+/g,'').trim();
  const n=Number(cleaned);
  return Number.isFinite(n)?Math.round(n*100)/100:0;
}
function ncrFindSheet(wb){
  const named=wb.SheetNames.find(name=>norm(name)==='ncr');
  if(named)return wb.Sheets[named];
  for(const name of wb.SheetNames){
    const ws=wb.Sheets[name];if(!ws||!ws['!ref'])continue;
    const range=XLSX.utils.decode_range(ws['!ref']);
    for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+20);r++){
      for(let c=range.s.c;c<=range.e.c;c++){
        if(ncrNormHeader(ncrText(ws[XLSX.utils.encode_cell({r,c})]))==='ncr no')return ws;
      }
    }
  }
  return null;
}
function ncrHeaderMap(ws,range){
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+25);r++){
    const map={};
    for(let c=range.s.c;c<=range.e.c;c++){
      const h=ncrNormHeader(ncrText(ws[XLSX.utils.encode_cell({r,c})]));
      if(h)map[h]=c;
    }
    if(map['ncr no']!=null&&map['date reported']!=null&&map['customer']!=null)return {row:r,map};
  }
  return null;
}
function ncrCell(ws,r,c){return c==null?null:ws[XLSX.utils.encode_cell({r,c})]}
function ncrCol(map,...names){
  for(const name of names){const key=ncrNormHeader(name);if(map[key]!=null)return map[key]}
  return null;
}
function parseNcrWorkbook(wb){
  const ws=ncrFindSheet(wb);if(!ws||!ws['!ref'])throw new Error('No NCR worksheet was found.');
  const range=XLSX.utils.decode_range(ws['!ref']),header=ncrHeaderMap(ws,range);
  if(!header)throw new Error('The NCR column headings could not be found.');
  const m=header.map;
  const cols={
    no:ncrCol(m,'NCR No.','NCR No'),date:ncrCol(m,'Date Reported'),ifo:ncrCol(m,'IFO'),
    invoiceMonth:ncrCol(m,'Invoice Month'),customer:ncrCol(m,'Customer'),fnumber:ncrCol(m,'Fnumber','F Number'),
    site:ncrCol(m,'Site'),complaint:ncrCol(m,'NCR Complaint'),employee:ncrCol(m,'Employee'),
    approvedBy:ncrCol(m,'Approved by'),category:ncrCol(m,'NCR Category'),
    beams:ncrCol(m,'Beams'),posi:ncrCol(m,'Posi'),ancillaries:ncrCol(m,'Ancillaries'),delivery:ncrCol(m,'Delivery'),total:ncrCol(m,'Total'),
    correctiveAction:ncrCol(m,'Corrective action'),investigated:ncrCol(m,'Investigated'),finding:ncrCol(m,'NCR Finding'),
    productionComments:ncrCol(m,'Production Comments'),details:ncrCol(m,'NCR Details'),collectReplace:ncrCol(m,'Collect and Replace?'),
    notes:ncrCol(m,'Investigation / Notes','Investigation Notes')
  };
  const records=[];
  for(let r=header.row+1;r<=range.e.r;r++){
    const ncrNo=ncrText(ncrCell(ws,r,cols.no));
    if(!/^NCR\s*\d+/i.test(ncrNo))continue;
    const beams=ncrNumber(ncrCell(ws,r,cols.beams)),posi=ncrNumber(ncrCell(ws,r,cols.posi)),ancillaries=ncrNumber(ncrCell(ws,r,cols.ancillaries)),delivery=ncrNumber(ncrCell(ws,r,cols.delivery));
    const calculated=Math.round((beams+posi+ancillaries+delivery)*100)/100;
    const sheetTotal=ncrNumber(ncrCell(ws,r,cols.total));
    records.push({
      rowId:`${ncrNo}-${r+1}`,
      ncrNo,dateReported:cellISODate(ncrCell(ws,r,cols.date))||'',ifo:ncrText(ncrCell(ws,r,cols.ifo)),
      invoiceMonth:ncrText(ncrCell(ws,r,cols.invoiceMonth)),customer:ncrText(ncrCell(ws,r,cols.customer)),
      fnumber:ncrText(ncrCell(ws,r,cols.fnumber)),site:ncrText(ncrCell(ws,r,cols.site)),
      complaint:ncrText(ncrCell(ws,r,cols.complaint)),employee:ncrText(ncrCell(ws,r,cols.employee)),
      approvedBy:ncrText(ncrCell(ws,r,cols.approvedBy)),category:ncrText(ncrCell(ws,r,cols.category)),
      beams,posi,ancillaries,delivery,total:calculated||sheetTotal,
      correctiveAction:ncrText(ncrCell(ws,r,cols.correctiveAction)),investigated:ncrText(ncrCell(ws,r,cols.investigated)),
      finding:ncrText(ncrCell(ws,r,cols.finding)),productionComments:ncrText(ncrCell(ws,r,cols.productionComments)),
      details:ncrText(ncrCell(ws,r,cols.details)),collectReplace:ncrText(ncrCell(ws,r,cols.collectReplace)),
      notes:ncrText(ncrCell(ws,r,cols.notes))
    });
  }
  if(!records.length)throw new Error('No NCR records were found in the workbook.');
  return records;
}

async function importNcrMany(files){
  const excel=[...files].filter(isExcelFile);
  if(!excel.length){showToast('Drop an NCR Excel file here');return}
  const all=[];const names=[];const errors=[];
  for(const file of excel){
    try{
      const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellFormula:true,cellNF:true,cellText:true});
      all.push(...parseNcrWorkbook(wb));names.push(file.name);
    }catch(e){errors.push(`${file.name}: ${e.message}`)}
  }
  if(!all.length){if(errors.length)alert(errors.join('\n'));return}
  showToast(`Saving ${all.length} NCR records…`);
  try{
    const result=await api('/api/ncr-import',{method:'POST',body:JSON.stringify({sourceFile:names.join(', '),records:all})});
    ncrState={records:result.records||all,sourceFile:result.sourceFile||names.join(', '),importedAt:result.importedAt||new Date().toISOString()};
    ncrLoaded=true;
    renderNcr();
updateAppImportContext();
    showToast(`${ncrState.records.length} NCR records saved`);
  }catch(e){alert(e.message)}
  if(errors.length)alert(errors.join('\n'));
}

async function loadNcrData(silent=false){
  if(ncrLoading)return;
  ncrLoading=true;
  try{
    const data=await api('/api/ncr');
    ncrState={records:data.records||[],sourceFile:data.sourceFile||'',importedAt:data.importedAt||null};
    ncrLoaded=true;renderNcr();
    if(!silent&&ncrState.records.length)showToast('NCR tracker refreshed');
  }catch(e){
    const message=e?.message||'NCR data could not be loaded';
    const root=document.getElementById('ncrRoot');
    if(root)root.innerHTML=`<div class="ncr-empty-state"><h2>NCR Tracker</h2><p><strong>NCR data could not be loaded.</strong></p><p class="ncr-small">${escapeHtml(message)}</p><button type="button" class="secondary" id="ncrRetryBtn">Retry</button></div>`;
    document.getElementById('ncrRetryBtn')?.addEventListener('click',()=>loadNcrData(false));
    if(!silent)showToast(message);
    console.warn('NCR load failed',e);
  }finally{ncrLoading=false}
}

function ncrLatestDate(records){
  const dates=records.map(x=>x.dateReported).filter(Boolean).sort();return dates.at(-1)||'';
}
function ncrPeriodStart(records){
  if(ncrFilters.period==='all')return '';
  const latest=ncrLatestDate(records);if(!latest)return '';
  const d=parseISO(latest);d.setUTCMonth(d.getUTCMonth()-Number(ncrFilters.period));return isoDate(d);
}
function ncrFiltered(){
  const all=ncrState.records||[],start=ncrPeriodStart(all);
  const q=norm(ncrFilters.search);
  return all.filter(r=>{
    if(start&&(!r.dateReported||r.dateReported<start))return false;
    if(ncrFilters.category&&r.category!==ncrFilters.category)return false;
    if(ncrFilters.employee&&r.employee!==ncrFilters.employee)return false;
    if(ncrFilters.customer&&r.customer!==ncrFilters.customer)return false;
    const outstanding=!String(r.investigated||'').trim();
    if(ncrFilters.status==='open'&&!outstanding)return false;
    if(ncrFilters.status==='investigated'&&outstanding)return false;
    if(q&&!norm([r.ncrNo,r.customer,r.site,r.complaint,r.employee,r.approvedBy,r.category,r.finding,r.details,r.fnumber,r.ifo].join(' ')).includes(q))return false;
    return true;
  }).sort((a,b)=>(b.dateReported||'').localeCompare(a.dateReported||'')||String(b.ncrNo).localeCompare(String(a.ncrNo)));
}
function ncrOptions(key){
  return [...new Set((ncrState.records||[]).map(x=>String(x[key]||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
}
function ncrMonthKey(date){return date?date.slice(0,7):''}
function ncrMonthLabel(key){
  if(!/^\d{4}-\d{2}$/.test(key))return key;
  const [y,m]=key.split('-').map(Number);return new Date(Date.UTC(y,m-1,1)).toLocaleDateString('en-GB',{timeZone:'UTC',month:'short',year:'2-digit'});
}
function ncrMonthly(records){
  const map=new Map();
  for(const r of records){
    const key=ncrMonthKey(r.dateReported);if(!key)continue;
    if(!map.has(key))map.set(key,{key,count:0,cost:0});
    const row=map.get(key);row.count++;row.cost+=Number(r.total||0);
  }
  return [...map.values()].sort((a,b)=>a.key.localeCompare(b.key));
}
function ncrMonthlySvg(data){
  if(!data.length)return '<div class="ncr-empty">No dated NCRs in this view.</div>';
  const W=960,H=280,p={l:42,r:18,t:18,b:48},max=Math.max(1,...data.map(x=>x.count));
  const slot=(W-p.l-p.r)/data.length,bar=Math.max(8,Math.min(36,slot*.58));
  const y=v=>p.t+(max-v)*(H-p.t-p.b)/max;
  const grid=Array.from({length:5},(_,i)=>{const v=max*(4-i)/4,yy=p.t+i*(H-p.t-p.b)/4;return `<line x1="${p.l}" y1="${yy}" x2="${W-p.r}" y2="${yy}" stroke="#e5e7eb"/><text x="${p.l-7}" y="${yy+4}" text-anchor="end" font-size="10" fill="#6b7280">${Math.round(v)}</text>`}).join('');
  const bars=data.map((d,i)=>{const x=p.l+i*slot+(slot-bar)/2,yy=y(d.count),h=H-p.b-yy;return `<g><rect x="${x}" y="${yy}" width="${bar}" height="${h}" rx="3" fill="#d71920"><title>${ncrMonthLabel(d.key)}: ${d.count} NCRs • ${ncrMoney(d.cost)}</title></rect><text x="${x+bar/2}" y="${H-20}" text-anchor="middle" font-size="10" fill="#6b7280">${ncrMonthLabel(d.key)}</text></g>`}).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Monthly NCR count trend">${grid}${bars}</svg>`;
}
function ncrCategoryRows(records){
  const map=new Map();
  for(const r of records){
    const key=r.category||'Uncategorised';if(!map.has(key))map.set(key,{name:key,count:0,cost:0});
    const row=map.get(key);row.count++;row.cost+=Number(r.total||0);
  }
  return [...map.values()].sort((a,b)=>b.count-a.count||b.cost-a.cost);
}
function ncrCostMix(records){
  return {
    beams:records.reduce((s,r)=>s+Number(r.beams||0),0),
    posi:records.reduce((s,r)=>s+Number(r.posi||0),0),
    ancillaries:records.reduce((s,r)=>s+Number(r.ancillaries||0),0),
    delivery:records.reduce((s,r)=>s+Number(r.delivery||0),0)
  };
}
function ncrStatusBadge(r){
  return String(r.investigated||'').trim()?'<span class="ncr-status done">Investigated</span>':'<span class="ncr-status open">Outstanding</span>';
}
function renderNcr(){
  const root=document.getElementById('ncrRoot');if(!root)return;
  if(!ncrLoaded){
    root.innerHTML='<div class="ncr-loading">Loading NCR tracker…</div>';return;
  }
  if(!ncrState.records.length){
    root.innerHTML=`<div class="ncr-empty-state"><h2>NCR Tracker</h2><p>No NCR workbook has been imported yet.</p><label class="button primary" for="importFiles">Import NCR Excel</label><p class="ncr-small">You can also drag the workbook anywhere onto this page.</p></div>`;return;
  }
  const records=ncrFiltered(),totalCost=records.reduce((s,r)=>s+Number(r.total||0),0),outstanding=records.filter(r=>!String(r.investigated||'').trim()).length;
  const avg=records.length?totalCost/records.length:0,monthly=ncrMonthly(records),categories=ncrCategoryRows(records),costMix=ncrCostMix(records);
  const maxCat=Math.max(1,...categories.map(x=>x.count));
  const catRows=categories.slice(0,8).map(x=>`<button class="ncr-category-row" data-ncr-category="${escapeHtml(x.name)}"><span class="ncr-cat-name">${escapeHtml(x.name)}</span><span class="ncr-cat-bar"><i style="width:${x.count/maxCat*100}%"></i></span><strong>${x.count}</strong><em>${ncrMoney(x.cost)}</em></button>`).join('');
  const rows=records.slice(0,300).map(r=>`<tr class="ncr-row" data-ncr-row="${escapeHtml(r.rowId)}"><td><strong>${escapeHtml(r.ncrNo)}</strong></td><td>${r.dateReported?fmtDate(r.dateReported):'—'}</td><td>${escapeHtml(r.customer||'—')}</td><td>${escapeHtml(r.category||'—')}</td><td>${escapeHtml(r.employee||'—')}</td><td>${escapeHtml(r.details||r.finding||'—')}</td><td>${ncrStatusBadge(r)}</td><td class="right"><strong>${ncrMoney(r.total)}</strong></td></tr>`).join('');
  const source=ncrState.sourceFile?escapeHtml(ncrState.sourceFile):'Saved NCR data';
  const updated=ncrState.importedAt?new Date(ncrState.importedAt).toLocaleString('en-GB'):'';
  root.innerHTML=`
    <div class="ncr-head">
      <div><div class="dashboard-kicker">NCR TRACKER</div><h2>Non-Conformance Dashboard</h2><p>${source}${updated?' • updated '+escapeHtml(updated):''}</p></div>
      <label class="button primary" for="importFiles">Import / Refresh Excel</label>
    </div>
    <div class="ncr-filters">
      <select id="ncrPeriod"><option value="3">Last 3 months</option><option value="6">Last 6 months</option><option value="12">Last 12 months</option><option value="all">All data</option></select>
      <input id="ncrSearch" type="text" placeholder="Search NCR, customer, site, complaint…" value="${escapeHtml(ncrFilters.search)}" />
      <select id="ncrCategory"><option value="">All categories</option>${ncrOptions('category').map(x=>`<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`).join('')}</select>
      <select id="ncrEmployee"><option value="">All employees</option>${ncrOptions('employee').map(x=>`<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`).join('')}</select>
      <select id="ncrStatus"><option value="all">All investigations</option><option value="open">Outstanding</option><option value="investigated">Investigated</option></select>
      <button class="secondary" id="ncrClearFilters">Clear filters</button>
    </div>
    <div class="ncr-kpis">
      <div class="ncr-kpi"><span>NCRs</span><strong>${records.length}</strong><em>in current view</em></div>
      <div class="ncr-kpi"><span>Total NCR cost</span><strong>${ncrMoney(totalCost)}</strong><em>Beams + Posi + Ancillaries + Delivery</em></div>
      <div class="ncr-kpi"><span>Average cost / NCR</span><strong>${ncrMoney(avg)}</strong><em>current filtered view</em></div>
      <div class="ncr-kpi"><span>Investigation outstanding</span><strong>${outstanding}</strong><em>${records.length?Math.round(outstanding/records.length*100):0}% of current NCRs</em></div>
    </div>
    <div class="ncr-grid">
      <section class="ncr-card ncr-trend"><div class="ncr-card-head"><div><h3>Monthly NCR trend</h3><span>Number reported by month</span></div></div><div class="ncr-chart">${ncrMonthlySvg(monthly)}</div></section>
      <section class="ncr-card"><div class="ncr-card-head"><div><h3>Cost mix</h3><span>Current filtered view</span></div></div>
        <div class="ncr-cost-mix">
          <div><span>Beams</span><strong>${ncrMoney(costMix.beams)}</strong></div>
          <div><span>Posi</span><strong>${ncrMoney(costMix.posi)}</strong></div>
          <div><span>Ancillaries</span><strong>${ncrMoney(costMix.ancillaries)}</strong></div>
          <div><span>Delivery</span><strong>${ncrMoney(costMix.delivery)}</strong></div>
        </div>
      </section>
    </div>
    <section class="ncr-card"><div class="ncr-card-head"><div><h3>NCR categories</h3><span>Click a category to filter the table</span></div></div><div class="ncr-category-list">${catRows||'<div class="ncr-empty">No categories.</div>'}</div></section>
    <section class="ncr-card"><div class="ncr-card-head"><div><h3>NCR register</h3><span>${records.length>300?`Showing latest 300 of ${records.length}`:`${records.length} record${records.length===1?'':'s'}`} • click an NCR for full details</span></div></div>
      <div class="ncr-table-wrap"><table><thead><tr><th>NCR</th><th>Date</th><th>Customer</th><th>Category</th><th>Employee</th><th>Finding / detail</th><th>Investigation</th><th class="right">Cost</th></tr></thead><tbody>${rows||'<tr><td colspan="8"><div class="ncr-empty">No NCRs match the filters.</div></td></tr>'}</tbody></table></div>
    </section>`;
  const period=root.querySelector('#ncrPeriod');period.value=ncrFilters.period;
  const category=root.querySelector('#ncrCategory');category.value=ncrFilters.category;
  const employee=root.querySelector('#ncrEmployee');employee.value=ncrFilters.employee;
  const status=root.querySelector('#ncrStatus');status.value=ncrFilters.status;
  period.addEventListener('change',e=>{ncrFilters.period=e.target.value;renderNcr()});
  root.querySelector('#ncrSearch').addEventListener('input',e=>{ncrFilters.search=e.target.value;clearTimeout(renderNcr.searchTimer);renderNcr.searchTimer=setTimeout(renderNcr,180)});
  category.addEventListener('change',e=>{ncrFilters.category=e.target.value;renderNcr()});
  employee.addEventListener('change',e=>{ncrFilters.employee=e.target.value;renderNcr()});
  status.addEventListener('change',e=>{ncrFilters.status=e.target.value;renderNcr()});
  root.querySelector('#ncrClearFilters').addEventListener('click',()=>{ncrFilters={period:'12',search:'',category:'',employee:'',customer:'',status:'all'};renderNcr()});
  root.querySelectorAll('[data-ncr-category]').forEach(btn=>btn.addEventListener('click',()=>{ncrFilters.category=btn.dataset.ncrCategory;renderNcr()}));
  root.querySelectorAll('[data-ncr-row]').forEach(row=>row.addEventListener('click',()=>showNcrDetail(row.dataset.ncrRow)));
}
function ensureNcrModal(){
  let modal=document.getElementById('ncrModal');if(modal)return modal;
  modal=document.createElement('div');modal.id='ncrModal';modal.className='ncr-modal';modal.hidden=true;
  modal.innerHTML='<div class="ncr-modal-shell"><div class="ncr-modal-head"><div><div class="dashboard-kicker">NCR DETAIL</div><h2 id="ncrModalTitle">NCR</h2><p id="ncrModalSub"></p></div><button type="button" class="secondary" id="ncrModalClose">Close</button></div><div id="ncrModalBody" class="ncr-modal-body"></div></div>';
  document.body.appendChild(modal);
  modal.querySelector('#ncrModalClose').addEventListener('click',()=>modal.hidden=true);
  modal.addEventListener('click',e=>{if(e.target===modal)modal.hidden=true});
  return modal;
}
function ncrDetailBlock(title,value){
  if(!String(value||'').trim())return '';
  return `<div class="ncr-detail-block"><span>${escapeHtml(title)}</span><p>${escapeHtml(value)}</p></div>`;
}
function showNcrDetail(rowId){
  const r=(ncrState.records||[]).find(x=>x.rowId===rowId);if(!r)return;
  selectedNcr=r;const modal=ensureNcrModal();
  modal.querySelector('#ncrModalTitle').textContent=r.ncrNo||'NCR';
  modal.querySelector('#ncrModalSub').textContent=[r.dateReported?fmtDate(r.dateReported):'',r.customer,r.site].filter(Boolean).join(' • ');
  modal.querySelector('#ncrModalBody').innerHTML=`
    <div class="ncr-detail-facts">
      <div><span>IFO</span><strong>${escapeHtml(r.ifo||'—')}</strong></div><div><span>F number</span><strong>${escapeHtml(r.fnumber||'—')}</strong></div>
      <div><span>Employee</span><strong>${escapeHtml(r.employee||'—')}</strong></div><div><span>Approved by</span><strong>${escapeHtml(r.approvedBy||'—')}</strong></div>
      <div><span>Category</span><strong>${escapeHtml(r.category||'—')}</strong></div><div><span>NCR detail</span><strong>${escapeHtml(r.details||'—')}</strong></div>
    </div>
    <div class="ncr-detail-costs"><div><span>Beams</span><strong>${ncrMoney(r.beams)}</strong></div><div><span>Posi</span><strong>${ncrMoney(r.posi)}</strong></div><div><span>Ancillaries</span><strong>${ncrMoney(r.ancillaries)}</strong></div><div><span>Delivery</span><strong>${ncrMoney(r.delivery)}</strong></div><div class="total"><span>Total</span><strong>${ncrMoney(r.total)}</strong></div></div>
    ${ncrDetailBlock('Complaint',r.complaint)}
    ${ncrDetailBlock('Corrective action',r.correctiveAction)}
    ${ncrDetailBlock('Investigated',r.investigated)}
    ${ncrDetailBlock('NCR finding',r.finding)}
    ${ncrDetailBlock('Production comments',r.productionComments)}
    ${ncrDetailBlock('Collect and replace?',r.collectReplace)}
    ${ncrDetailBlock('Investigation / notes',r.notes)}`;
  modal.hidden=false;
}
function updateAppImportContext(){
  const label=document.querySelector('.import-header-btn'),overlay=document.getElementById('pageDropOverlay');
  if(label)label.textContent=currentAppTab==='ncr'?'Import NCR Excel':'Import Design Hours';
  if(overlay){
    const strong=overlay.querySelector('strong'),span=overlay.querySelector('span');
    if(currentAppTab==='ncr'){strong.textContent='Drop NCR Excel to import';span.textContent='The NCR register will replace the currently saved NCR dataset.'}
    else{strong.textContent='Drop Excel to import';span.textContent='Weekly or annual estimator files'}
  }
}
function switchAppTab(tab){
  currentAppTab=tab==='ncr'?'ncr':'design';
  const design=document.getElementById('designApp'),ncr=document.getElementById('ncrApp');
  if(design)design.hidden=currentAppTab!=='design';
  if(ncr)ncr.hidden=currentAppTab!=='ncr';
  document.querySelectorAll('[data-app-tab]').forEach(btn=>btn.classList.toggle('active',btn.dataset.appTab===currentAppTab));
  document.body.classList.toggle('ncr-active',currentAppTab==='ncr');
  updateAppImportContext();
  if(currentAppTab==='ncr'){
    document.body.classList.remove('dashboard-home');
    if(!ncrLoaded)loadNcrData(true);else renderNcr();
  }else{
    renderDetail();
  }
}
document.getElementById('designHoursTab')?.addEventListener('click',()=>switchAppTab('design'));
document.getElementById('ncrTrackerTab')?.addEventListener('click',()=>switchAppTab('ncr'));
renderNcr();
