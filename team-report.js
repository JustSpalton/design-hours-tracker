let managerReportMode='month';

function currentViewWeeks(){
  if(typeof breakdownWeeks==='function') return breakdownWeeks();
  return selectedWeek?[selectedWeek]:[];
}
function allBreakdownTotals(field,weeks=currentViewWeeks()){
  const wanted=new Set(weeks),out={};
  for(const rec of breakdownState.records||[]){
    if(!wanted.has(rec.week))continue;
    for(const [label,value] of Object.entries(rec[field]||{}))out[label]=(out[label]||0)+Number(value||0);
  }
  return out;
}
function orderedProducts(totals){
  const preferred=['I Joist','Posi Joist'];
  const rest=Object.keys(totals).filter(x=>!preferred.includes(x)).sort((a,b)=>(totals[b]||0)-(totals[a]||0)||a.localeCompare(b));
  return [...preferred,...rest].filter((x,i,a)=>a.indexOf(x)===i);
}
function breakdownHoursAvailable(field,weeks=currentViewWeeks()){
  const wanted=new Set(weeks);
  return (breakdownState.records||[]).some(rec=>wanted.has(rec.week)&&Object.prototype.hasOwnProperty.call(rec,field));
}
function installProductTotals(){
  if(document.getElementById('productTotalsPanel'))return;
  const stats=document.querySelector('.stats');if(!stats)return;
  const panel=document.createElement('div');panel.className='panel product-totals-panel';panel.id='productTotalsPanel';
  stats.insertAdjacentElement('afterend',panel);
}
function renderProductTotals(){
  installProductTotals();
  const panel=document.getElementById('productTotalsPanel');if(!panel)return;
  const weeks=currentViewWeeks(),totals=allBreakdownTotals('products',weeks),hours=allBreakdownTotals('product_hours',weeks),products=orderedProducts(totals);
  const hasData=Object.values(totals).some(v=>Number(v)>0),hasHours=breakdownHoursAvailable('product_hours',weeks);
  panel.innerHTML=`<div class="product-totals-head"><div><h3>Product totals — everyone</h3><span>${escapeHtml(typeof rangeLabel==='function'?rangeLabel():`W/C ${fmtDate(selectedWeek)}`)}</span></div></div><div class="product-total-grid">${hasData?products.map(p=>`<div class="product-total"><div class="product-total-label">${escapeHtml(p)}</div><div class="product-total-value">${Number(totals[p]||0)}</div><div class="product-total-sub">jobs • ${hasHours?`${Number(hours[p]||0).toFixed(1)} design hours`:'hours not imported'}</div></div>`).join(''):'<div class="breakdown-empty">No product breakdown data for this period yet.</div>'}</div>`;
}

function teamForDesignerId(id){return (state.teams||[]).find(t=>(t.designer_ids||[]).includes(Number(id)))||null}
function designerViewHours(name){return typeof rangeActive!=='undefined'&&rangeActive?viewActual(name):getHours(name,selectedWeek)}
function designerHolidayBadge(name){return typeof rangeActive!=='undefined'&&rangeActive?viewHolidayCount(name):holidayCount(name,selectedWeek)}
function designerButton(d){
  const h=designerViewHours(d.name),hc=designerHolidayBadge(d.name);
  const b=document.createElement('button');b.type='button';b.className='designer team-designer '+(norm(d.name)===norm(selectedDesigner)?'active':'');
  b.draggable=true;b.dataset.designerId=String(d.id);
  b.innerHTML=`<span class="drag-grip" aria-hidden="true">⋮⋮</span><span class="name">${escapeHtml(d.name)}${hc?`<span class="holiday-badge">${hc} holiday${hc===1?'':'s'}</span>`:''}</span><span class="hours ${h==null?'empty':''}">${h==null?'—':Number(h).toFixed(1)}</span>`;
  b.addEventListener('click',()=>{selectedDesigner=d.name;renderDesignerList();renderDetail()});
  b.addEventListener('dragstart',e=>{e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/designer-id',String(d.id));b.classList.add('dragging')});
  b.addEventListener('dragend',()=>b.classList.remove('dragging'));
  return b;
}
async function moveDesignerToTeam(designerId,teamId){
  try{
    await api('/api/teams',{method:'POST',body:JSON.stringify({action:'move',designerId:Number(designerId),teamId:teamId||''})});
    await loadData(true);showToast('Designer moved');
  }catch(e){alert(e.message)}
}
function wireTeamDrop(zone){
  zone.addEventListener('dragover',e=>{e.preventDefault();e.dataTransfer.dropEffect='move';zone.classList.add('drop-ready')});
  zone.addEventListener('dragleave',()=>zone.classList.remove('drop-ready'));
  zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('drop-ready');const id=e.dataTransfer.getData('text/designer-id');if(id)moveDesignerToTeam(id,zone.dataset.teamId||'')});
}
async function renameTeam(team){
  const name=prompt('Team name',team.name);if(!name||name.trim()===team.name)return;
  try{await api('/api/teams',{method:'POST',body:JSON.stringify({action:'rename',teamId:team.id,name})});await loadData(true);showToast('Team renamed')}catch(e){alert(e.message)}
}
async function addTeam(){
  const name=prompt('New team name');if(!name)return;
  try{await api('/api/teams',{method:'POST',body:JSON.stringify({action:'add',name})});await loadData(true);showToast('Team added')}catch(e){alert(e.message)}
}
async function deleteTeam(team){
  if(!confirm(`Remove "${team.name}"? Its designers will become unassigned.`))return;
  try{await api('/api/teams',{method:'POST',body:JSON.stringify({action:'delete',teamId:team.id})});await loadData(true);showToast('Team removed')}catch(e){alert(e.message)}
}
function renderDesignerListByTeams(){
  const q=norm(document.getElementById('designerSearch').value),box=document.getElementById('designerList');box.innerHTML='';
  const designers=state.designers.filter(d=>!q||norm(d.name).includes(q));
  const teams=state.teams||[];
  const assigned=new Set(teams.flatMap(t=>t.designer_ids||[]));
  for(const team of teams){
    const members=designers.filter(d=>(team.designer_ids||[]).includes(d.id));
    const group=document.createElement('div');group.className='team-group';
    group.innerHTML=`<div class="team-group-head"><div><strong>${escapeHtml(team.name)}</strong><span>${members.length} designer${members.length===1?'':'s'}</span></div><div class="team-actions"><button type="button" class="team-icon edit-team" title="Rename team">✎</button><button type="button" class="team-icon delete-team" title="Remove team">×</button></div></div><div class="team-dropzone" data-team-id="${escapeHtml(team.id)}"></div>`;
    const zone=group.querySelector('.team-dropzone');members.forEach(d=>zone.appendChild(designerButton(d)));
    if(!members.length)zone.innerHTML='<div class="team-empty">Drag designers here</div>';
    group.querySelector('.edit-team').addEventListener('click',()=>renameTeam(team));
    group.querySelector('.delete-team').addEventListener('click',()=>deleteTeam(team));
    wireTeamDrop(zone);box.appendChild(group);
  }
  const unassigned=designers.filter(d=>!assigned.has(d.id));
  const group=document.createElement('div');group.className='team-group unassigned-group';
  group.innerHTML=`<div class="team-group-head"><div><strong>Unassigned</strong><span>${unassigned.length} designer${unassigned.length===1?'':'s'}</span></div></div><div class="team-dropzone" data-team-id=""></div>`;
  const zone=group.querySelector('.team-dropzone');unassigned.forEach(d=>zone.appendChild(designerButton(d)));if(!unassigned.length)zone.innerHTML='<div class="team-empty">Everyone is assigned</div>';wireTeamDrop(zone);box.appendChild(group);
}
function installTeamControls(){
  const row=document.querySelector('.searchrow');if(!row||document.getElementById('addTeamBtn'))return;
  const b=document.createElement('button');b.type='button';b.id='addTeamBtn';b.className='secondary team-add';b.textContent='+ Team';b.title='Add team';b.addEventListener('click',addTeam);row.appendChild(b);
}

function reportPeriodKeys(mode){
  const weeks=new Set([...(state.hours||[]).map(x=>x.week),...(breakdownState.records||[]).map(x=>x.week)]);
  const keys=new Set();
  for(const week of weeks){
    const d=parseISO(week),y=d.getUTCFullYear(),m=d.getUTCMonth()+1;
    keys.add(mode==='quarter'?`${y}-Q${Math.floor((m-1)/3)+1}`:`${y}-${String(m).padStart(2,'0')}`);
  }
  return [...keys].sort().reverse();
}
function reportBounds(key,mode){
  if(mode==='quarter'){
    const [y,qraw]=key.split('-Q'),q=Number(qraw),startMonth=(q-1)*3;
    const start=new Date(Date.UTC(Number(y),startMonth,1)),end=new Date(Date.UTC(Number(y),startMonth+3,0));
    return {start:isoDate(start),end:isoDate(end),label:`Q${q} ${y}`};
  }
  const [y,m]=key.split('-').map(Number),start=new Date(Date.UTC(y,m-1,1)),end=new Date(Date.UTC(y,m,0));
  return {start:isoDate(start),end:isoDate(end),label:start.toLocaleDateString('en-GB',{timeZone:'UTC',month:'long',year:'numeric'})};
}
function weeksForBounds(bounds){
  const set=new Set([...(state.hours||[]).map(x=>x.week),...(breakdownState.records||[]).map(x=>x.week)]);
  return [...set].filter(w=>w>=bounds.start&&w<=bounds.end).sort();
}
function totalsForBreakdownWeeks(field,weeks){
  const wanted=new Set(weeks),out={};
  for(const rec of breakdownState.records||[]){
    if(!wanted.has(rec.week))continue;
    for(const [k,v] of Object.entries(rec[field]||{}))out[k]=(out[k]||0)+Number(v||0);
  }
  return out;
}
function designerReportData(designer,weeks,bounds){
  const wanted=new Set(weeks);
  const hours=(state.hours||[]).filter(x=>norm(x.designer)===norm(designer.name)&&wanted.has(x.week)).reduce((s,x)=>s+Number(x.hours||0),0);
  const holidayDays=(state.holidays||[]).filter(x=>norm(x.designer)===norm(designer.name)&&x.date>=bounds.start&&x.date<=bounds.end).length;
  const products={},statuses={};
  for(const rec of breakdownState.records||[]){
    if(norm(rec.designer)!==norm(designer.name)||!wanted.has(rec.week))continue;
    for(const [k,v] of Object.entries(rec.products||{}))products[k]=(products[k]||0)+Number(v||0);
    for(const [k,v] of Object.entries(rec.statuses||{}))statuses[k]=(statuses[k]||0)+Number(v||0);
  }
  return {hours,holidayDays,products,statuses};
}
function reportTeamTable(title,designers,weeks,bounds){
  if(!designers.length)return '';
  const rows=designers.map(d=>{const x=designerReportData(d,weeks,bounds);return `<tr><td><strong>${escapeHtml(d.name)}</strong></td><td class="right">${x.hours.toFixed(1)}</td><td class="right">${x.holidayDays}</td><td class="right">${Number(x.products['I Joist']||0)}</td><td class="right">${Number(x.products['Posi Joist']||0)}</td><td class="right">${Number(x.statuses['OTP']||0)}</td><td class="right">${Number(x.statuses['Amendment']||0)}</td><td class="right">${Number(x.statuses['Budget']||0)}</td></tr>`}).join('');
  const total=designers.reduce((s,d)=>s+designerReportData(d,weeks,bounds).hours,0);
  return `<section class="report-team"><h3>${escapeHtml(title)}</h3><div class="report-table-wrap"><table><thead><tr><th>Designer</th><th class="right">Hours</th><th class="right">Holiday days</th><th class="right">I Joist</th><th class="right">Posi Joist</th><th class="right">OTP</th><th class="right">Amendment</th><th class="right">Budget</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td><strong>Team total</strong></td><td class="right"><strong>${total.toFixed(1)}</strong></td><td colspan="6"></td></tr></tfoot></table></div></section>`;
}
function renderManagerReport(){
  const period=document.getElementById('reportPeriod')?.value;if(!period)return;
  const bounds=reportBounds(period,managerReportMode),weeks=weeksForBounds(bounds),wanted=new Set(weeks);
  const hoursRows=(state.hours||[]).filter(x=>wanted.has(x.week));
  const totalHours=hoursRows.reduce((s,x)=>s+Number(x.hours||0),0);
  const activeDesigners=new Set(hoursRows.map(x=>norm(x.designer))).size;
  const holidayDays=(state.holidays||[]).filter(x=>x.date>=bounds.start&&x.date<=bounds.end).length;
  const products=totalsForBreakdownWeeks('products',weeks),statuses=totalsForBreakdownWeeks('statuses',weeks);
  const productNames=orderedProducts(products),statusNames=Object.keys(statuses).sort((a,b)=>(statuses[b]||0)-(statuses[a]||0)||a.localeCompare(b));
  const breakdownCoverage=new Set((breakdownState.records||[]).filter(x=>wanted.has(x.week)).map(x=>x.week)).size;
  const teams=state.teams||[],assigned=new Set(teams.flatMap(t=>t.designer_ids||[]));
  const teamSections=teams.map(t=>reportTeamTable(t.name,state.designers.filter(d=>(t.designer_ids||[]).includes(d.id)),weeks,bounds)).join('');
  const unassigned=state.designers.filter(d=>!assigned.has(d.id));
  const body=document.getElementById('managerReportBody');
  body.innerHTML=`<div class="report-title"><div><div class="donaldson-report">DONALDSON</div><h2>Manager Report — ${escapeHtml(bounds.label)}</h2><p>Weeks commencing within ${fmtDate(bounds.start)} to ${fmtDate(bounds.end)}</p></div><div class="report-generated">Generated ${new Date().toLocaleDateString('en-GB')}</div></div>
  <div class="report-summary">
    <div><span>Total design hours</span><strong>${totalHours.toFixed(1)}</strong></div>
    <div><span>Designers with hours</span><strong>${activeDesigners}</strong></div>
    <div><span>Holiday days</span><strong>${holidayDays}</strong></div>
    <div><span>Weeks included</span><strong>${weeks.length}</strong></div>
  </div>
  <section class="report-block"><h3>Product totals — everyone</h3><div class="report-chips">${productNames.length?productNames.map(p=>`<div><span>${escapeHtml(p)}</span><strong>${Number(products[p]||0)}</strong></div>`).join(''):'<p>No product breakdown data imported for this period.</p>'}</div></section>
  <section class="report-block"><h3>Job type totals — everyone</h3><div class="report-chips">${statusNames.length?statusNames.map(s=>`<div><span>${escapeHtml(s)}</span><strong>${Number(statuses[s]||0)}</strong></div>`).join(''):'<p>No job type breakdown data imported for this period.</p>'}</div></section>
  ${breakdownCoverage<weeks.length?`<div class="report-note">Product and job-type figures cover ${breakdownCoverage} of ${weeks.length} week${weeks.length===1?'':'s'} in this report. Re-import older estimator spreadsheets to fill any missing breakdown history.</div>`:''}
  <div class="report-team-sections">${teamSections}${reportTeamTable('Unassigned',unassigned,weeks,bounds)}</div>`;
}
function refreshReportPeriods(){
  const sel=document.getElementById('reportPeriod');if(!sel)return;const keys=reportPeriodKeys(managerReportMode);
  sel.innerHTML=keys.map(k=>`<option value="${k}">${escapeHtml(reportBounds(k,managerReportMode).label)}</option>`).join('');
  if(keys.length)renderManagerReport();else document.getElementById('managerReportBody').innerHTML='<div class="empty-state">No report periods available yet.</div>';
}
function openManagerReport(){
  const modal=document.getElementById('managerReportModal');modal.hidden=false;document.body.classList.add('report-open');refreshReportPeriods();
}
function closeManagerReport(){document.getElementById('managerReportModal').hidden=true;document.body.classList.remove('report-open')}
function installManagerReport(){
  if(document.getElementById('managerReportBtn'))return;
  const actions=document.querySelector('header .actions');if(actions){const b=document.createElement('button');b.type='button';b.id='managerReportBtn';b.className='primary';b.textContent='Manager Report';b.addEventListener('click',openManagerReport);actions.prepend(b)}
  const modal=document.createElement('div');modal.id='managerReportModal';modal.className='manager-report-modal';modal.hidden=true;
  modal.innerHTML=`<div class="manager-report-shell"><div class="manager-report-controls"><div class="report-mode"><button type="button" data-report-mode="month" class="active">Monthly</button><button type="button" data-report-mode="quarter">Quarterly</button></div><select id="reportPeriod"></select><button type="button" class="secondary" id="printManagerReport">Print report</button><button type="button" class="secondary" id="closeManagerReport">Close</button></div><div id="managerReportBody" class="manager-report-body"></div></div>`;
  document.body.appendChild(modal);
  modal.querySelectorAll('[data-report-mode]').forEach(btn=>btn.addEventListener('click',()=>{managerReportMode=btn.dataset.reportMode;modal.querySelectorAll('[data-report-mode]').forEach(x=>x.classList.toggle('active',x===btn));refreshReportPeriods()}));
  modal.querySelector('#reportPeriod').addEventListener('change',renderManagerReport);
  modal.querySelector('#printManagerReport').addEventListener('click',()=>window.print());
  modal.querySelector('#closeManagerReport').addEventListener('click',closeManagerReport);
  modal.addEventListener('click',e=>{if(e.target===modal)closeManagerReport()});
}

const teamBaseRenderDesignerList=renderDesignerList;
renderDesignerList=function(){renderDesignerListByTeams()};
const teamBaseRenderStats=renderStats;
renderStats=function(){teamBaseRenderStats();renderProductTotals()};
const teamBaseRenderAll=renderAll;
renderAll=function(){installTeamControls();installManagerReport();teamBaseRenderAll();renderProductTotals()};

installTeamControls();installManagerReport();installProductTotals();renderDesignerList();renderProductTotals();
loadBreakdowns().then(()=>{renderProductTotals();if(!document.getElementById('managerReportModal')?.hidden)refreshReportPeriods()});
