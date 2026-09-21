let dashboardMonths=6;
let dashboardWorkMetric='jobs';

function dashboardLatestWeek(){
  const weeks=(state.hours||[]).map(x=>x.week).filter(Boolean).sort();
  return weeks.at(-1)||selectedWeek||currentMonday();
}
function dashboardPeriod(months=dashboardMonths){
  const endWeek=dashboardLatestWeek();
  const end=parseISO(endWeek);end.setUTCDate(end.getUTCDate()+4);
  const start=new Date(end);start.setUTCMonth(start.getUTCMonth()-months);
  return {start:isoDate(start),end:isoDate(end),endWeek};
}
function dashboardWeeks(months=dashboardMonths){
  const bounds=dashboardPeriod(months);
  return allWeeks().filter(w=>{
    const ws=parseISO(w),we=new Date(ws);we.setUTCDate(we.getUTCDate()+4);
    return we>=parseISO(bounds.start)&&ws<=parseISO(bounds.end);
  });
}
function dashboardBreakdownFor(designer,weeks){
  const wanted=new Set(weeks),products={},statuses={};let jobs=0;
  for(const rec of (typeof breakdownState!=='undefined'?(breakdownState.records||[]):[])){
    if(norm(rec.designer)!==norm(designer)||!wanted.has(rec.week))continue;
    for(const [key,value] of Object.entries(rec.products||{})){const n=Number(value||0);products[key]=(products[key]||0)+n;jobs+=n}
    for(const [key,value] of Object.entries(rec.statuses||{}))statuses[key]=(statuses[key]||0)+Number(value||0);
  }
  return {jobs,products,statuses};
}
function dashboardDesignerStats(designer,weeks){
  const wanted=new Set(weeks);
  const rows=(state.hours||[]).filter(x=>norm(x.designer)===norm(designer)&&wanted.has(x.week)).sort((a,b)=>a.week.localeCompare(b.week));
  const adjusted=rows.map(x=>({week:x.week,value:adjustedHours(Number(x.hours||0),holidayCount(designer,x.week))})).filter(x=>Number.isFinite(x.value));
  const total=rows.reduce((s,x)=>s+Number(x.hours||0),0);
  const avg=adjusted.length?adjusted.reduce((s,x)=>s+x.value,0)/adjusted.length:0;
  const recent=adjusted.slice(-4),previous=adjusted.slice(-8,-4);
  const recentAvg=recent.length?recent.reduce((s,x)=>s+x.value,0)/recent.length:null;
  const previousAvg=previous.length?previous.reduce((s,x)=>s+x.value,0)/previous.length:null;
  const trend=recentAvg!=null&&previousAvg!=null&&previousAvg!==0?(recentAvg-previousAvg)/previousAvg*100:null;
  const breakdown=dashboardBreakdownFor(designer,weeks);
  return {name:designer,total,avg,weeks:rows.length,trend,jobs:breakdown.jobs,iJoist:Number(breakdown.products['I Joist']||0),posiJoist:Number(breakdown.products['Posi Joist']||0)};
}
function dashboardWeeklyTotals(weeks){
  return weeks.map(week=>{
    const rows=(state.hours||[]).filter(x=>x.week===week);
    return {week,hours:rows.reduce((s,x)=>s+Number(x.hours||0),0),designers:new Set(rows.map(x=>norm(x.designer))).size};
  }).filter(x=>x.hours>0);
}
function dashboardTrendSvg(data){
  if(!data.length)return '<div class="dashboard-empty">No design-hour history is available for this period.</div>';
  const W=1000,H=300,p={l:52,r:18,t:20,b:50};const max=Math.max(10,Math.ceil(Math.max(...data.map(x=>x.hours))/20)*20);
  const x=i=>p.l+(data.length===1?(W-p.l-p.r)/2:i*(W-p.l-p.r)/(data.length-1));
  const y=v=>p.t+(max-v)*(H-p.t-p.b)/max;
  let grid='';
  for(let i=0;i<=4;i++){const v=max*(4-i)/4,yy=p.t+i*(H-p.t-p.b)/4;grid+=`<line x1="${p.l}" y1="${yy}" x2="${W-p.r}" y2="${yy}" stroke="#e5e7eb"/><text x="${p.l-9}" y="${yy+4}" text-anchor="end" font-size="10" fill="#6b7280">${v.toFixed(0)}</text>`}
  const pts=data.map((d,i)=>`${x(i)},${y(d.hours)}`).join(' ');
  const every=Math.max(1,Math.ceil(data.length/8));
  const labels=data.map((d,i)=>i%every===0||i===data.length-1?`<text x="${x(i)}" y="${H-19}" text-anchor="middle" font-size="10" fill="#6b7280">${fmtDate(d.week).slice(0,5)}</text>`:'').join('');
  const dots=data.map((d,i)=>`<circle cx="${x(i)}" cy="${y(d.hours)}" r="3.5" fill="#d71920"><title>W/C ${fmtDate(d.week)}: ${d.hours.toFixed(1)}h • ${d.designers} designer${d.designers===1?'':'s'}</title></circle>`).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Team weekly design hours trend">${grid}<polyline points="${pts}" fill="none" stroke="#d71920" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>${dots}${labels}</svg>`;
}
function dashboardMonthKey(week){
  const d=parseISO(week);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}`;
}
function dashboardMonthLabel(key){
  const [year,month]=key.split('-').map(Number);
  return new Date(Date.UTC(year,month-1,1)).toLocaleDateString('en-GB',{timeZone:'UTC',month:'short',year:'2-digit'});
}
function dashboardWorkCategoryTrend(weeks){
  const wanted=new Set(weeks),months=new Map();
  for(const week of weeks){const key=dashboardMonthKey(week);if(!months.has(key))months.set(key,{key,label:dashboardMonthLabel(key),amendmentJobs:0,otpJobs:0,amendmentHours:0,otpHours:0})}
  for(const rec of (typeof breakdownState!=='undefined'?(breakdownState.records||[]):[])){
    if(!wanted.has(rec.week))continue;
    const key=dashboardMonthKey(rec.week);
    if(!months.has(key))months.set(key,{key,label:dashboardMonthLabel(key),amendmentJobs:0,otpJobs:0,amendmentHours:0,otpHours:0});
    const row=months.get(key);
    row.amendmentJobs+=Number(rec.statuses?.Amendment||0);
    row.otpJobs+=Number(rec.statuses?.OTP||0);
    row.amendmentHours+=Number(rec.status_hours?.Amendment||0);
    row.otpHours+=Number(rec.status_hours?.OTP||0);
  }
  return [...months.values()].sort((a,b)=>a.key.localeCompare(b.key));
}
function dashboardWorkTrendSvg(data,metric){
  const amendmentKey=metric==='hours'?'amendmentHours':'amendmentJobs';
  const otpKey=metric==='hours'?'otpHours':'otpJobs';
  const usable=data.filter(x=>Number(x[amendmentKey])>0||Number(x[otpKey])>0);
  if(!usable.length)return '<div class="dashboard-empty">No Amendment / OTP breakdown data is available for this period. Re-import the weekly spreadsheets to build this trend.</div>';
  const W=1000,H=300,p={l:52,r:18,t:28,b:50};
  const rawMax=Math.max(...data.flatMap(x=>[Number(x[amendmentKey]||0),Number(x[otpKey]||0)]));
  const step=metric==='hours'?10:5;
  const max=Math.max(step,Math.ceil(rawMax/step)*step);
  const x=i=>p.l+(data.length===1?(W-p.l-p.r)/2:i*(W-p.l-p.r)/(data.length-1));
  const y=v=>p.t+(max-v)*(H-p.t-p.b)/max;
  let grid='';
  for(let i=0;i<=4;i++){
    const v=max*(4-i)/4,yy=p.t+i*(H-p.t-p.b)/4;
    grid+=`<line x1="${p.l}" y1="${yy}" x2="${W-p.r}" y2="${yy}" stroke="#e5e7eb"/><text x="${p.l-9}" y="${yy+4}" text-anchor="end" font-size="10" fill="#6b7280">${metric==='hours'?v.toFixed(1):v.toFixed(0)}</text>`;
  }
  const amendmentPts=data.map((d,i)=>`${x(i)},${y(Number(d[amendmentKey]||0))}`).join(' ');
  const otpPts=data.map((d,i)=>`${x(i)},${y(Number(d[otpKey]||0))}`).join(' ');
  const labels=data.map((d,i)=>`<text x="${x(i)}" y="${H-18}" text-anchor="middle" font-size="10" fill="#6b7280">${escapeHtml(d.label)}</text>`).join('');
  const suffix=metric==='hours'?'h':' jobs';
  const amendmentDots=data.map((d,i)=>`<circle cx="${x(i)}" cy="${y(Number(d[amendmentKey]||0))}" r="3.5" fill="#d71920"><title>${d.label} Amendments: ${Number(d[amendmentKey]||0).toFixed(metric==='hours'?1:0)}${suffix}</title></circle>`).join('');
  const otpDots=data.map((d,i)=>`<circle cx="${x(i)}" cy="${y(Number(d[otpKey]||0))}" r="3.5" fill="#222"><title>${d.label} OTPs: ${Number(d[otpKey]||0).toFixed(metric==='hours'?1:0)}${suffix}</title></circle>`).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Amendment versus OTP ${metric} trend">${grid}<polyline points="${amendmentPts}" fill="none" stroke="#d71920" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/><polyline points="${otpPts}" fill="none" stroke="#222" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>${amendmentDots}${otpDots}${labels}</svg>`;
}
function dashboardWorkCategoryTotals(data,metric){
  const amendmentKey=metric==='hours'?'amendmentHours':'amendmentJobs';
  const otpKey=metric==='hours'?'otpHours':'otpJobs';
  return {
    amendment:data.reduce((s,x)=>s+Number(x[amendmentKey]||0),0),
    otp:data.reduce((s,x)=>s+Number(x[otpKey]||0),0)
  };
}
function dashboardTrendText(value){
  if(value==null||!Number.isFinite(value))return '<span class="trend-flat">—</span>';
  const cls=value>2?'trend-up':value<-2?'trend-down':'trend-flat';
  const arrow=value>2?'↑':value<-2?'↓':'→';
  return `<span class="${cls}">${arrow} ${Math.abs(value).toFixed(1)}%</span>`;
}
function dashboardCoverage(weeks){
  const wanted=new Set(weeks);
  const covered=new Set((typeof breakdownState!=='undefined'?(breakdownState.records||[]):[]).filter(x=>wanted.has(x.week)).map(x=>x.week));
  return {covered:covered.size,total:weeks.length};
}
function openDashboardDesigner(name){
  selectedDesigner=name;
  if(typeof rangeActive!=='undefined')rangeActive=false;
  const designerWeeks=(state.hours||[]).filter(x=>norm(x.designer)===norm(name)).map(x=>x.week).sort();
  if(designerWeeks.length)selectedWeek=designerWeeks.at(-1);
  renderWeeks();if(typeof syncRangeUI==='function')syncRangeUI();renderStats();renderDesignerList();renderDetail();
}
function renderDashboard(){
  document.body.classList.add('dashboard-home');
  const box=document.getElementById('detailPanel');if(!box)return;
  const weeks=dashboardWeeks(),bounds=dashboardPeriod(),weekly=dashboardWeeklyTotals(weeks);
  const people=(state.designers||[]).map(d=>dashboardDesignerStats(d.name,weeks)).filter(x=>x.weeks>0).sort((a,b)=>a.name.localeCompare(b.name));
  const totalHours=weekly.reduce((s,x)=>s+x.hours,0);
  const avgTeamWeek=weekly.length?totalHours/weekly.length:0;
  const allBreakdowns=people.reduce((a,x)=>({jobs:a.jobs+x.jobs,iJoist:a.iJoist+x.iJoist,posiJoist:a.posiJoist+x.posiJoist}),{jobs:0,iJoist:0,posiJoist:0});
  const coverage=dashboardCoverage(weeks);
  const workTrend=dashboardWorkCategoryTrend(weeks);
  const workTotals=dashboardWorkCategoryTotals(workTrend,dashboardWorkMetric);
  const splitTotal=allBreakdowns.iJoist+allBreakdowns.posiJoist;
  const iPct=splitTotal?allBreakdowns.iJoist/splitTotal*100:0;
  const rows=people.map(p=>`<tr class="dashboard-person" data-dashboard-designer="${escapeHtml(p.name)}" tabindex="0"><td><strong>${escapeHtml(p.name)}</strong><span class="dashboard-row-hint">View details</span></td><td class="right">${p.total.toFixed(1)}</td><td class="right">${p.avg.toFixed(1)}</td><td class="right">${p.jobs||'—'}</td><td class="right">${p.iJoist||'—'}</td><td class="right">${p.posiJoist||'—'}</td><td class="right">${dashboardTrendText(p.trend)}</td></tr>`).join('');
  box.innerHTML=`<div class="dashboard-shell">
    <div class="dashboard-head"><div><div class="dashboard-kicker">TEAM OVERVIEW</div><h2>Design Hours Dashboard</h2><p>Rolling ${dashboardMonths} months • ${fmtDate(bounds.start)} to ${fmtDate(bounds.end)}</p></div><div class="dashboard-period"><button type="button" data-dashboard-months="6" class="${dashboardMonths===6?'active':''}">6 months</button><button type="button" data-dashboard-months="12" class="${dashboardMonths===12?'active':''}">12 months</button></div></div>
    <div class="dashboard-cards">
      <div class="dashboard-card"><span>Total design hours</span><strong>${totalHours.toFixed(1)}</strong><em>${weekly.length} recorded weeks</em></div>
      <div class="dashboard-card"><span>Average team hours / week</span><strong>${avgTeamWeek.toFixed(1)}</strong><em>Across recorded weeks</em></div>
      <div class="dashboard-card"><span>Jobs recorded</span><strong>${allBreakdowns.jobs}</strong><em>${coverage.covered} of ${coverage.total} weeks have job data</em></div>
      <div class="dashboard-card"><span>Active designers</span><strong>${people.length}</strong><em>With hours in this period</em></div>
    </div>
    <div class="dashboard-grid">
      <section class="dashboard-section dashboard-trend"><div class="dashboard-section-head"><div><h3>Team weekly hours trend</h3><span>Actual design hours by week</span></div></div><div class="dashboard-chart">${dashboardTrendSvg(weekly)}</div></section>
      <section class="dashboard-section dashboard-split"><div class="dashboard-section-head"><div><h3>Job type split</h3><span>I Joist vs Posi Joist</span></div></div>
        <div class="dashboard-split-body"><div class="split-numbers"><div><span>I Joist</span><strong>${allBreakdowns.iJoist}</strong></div><div><span>Posi Joist</span><strong>${allBreakdowns.posiJoist}</strong></div></div>
        ${splitTotal?`<div class="split-bar"><span style="width:${iPct}%"></span></div><div class="split-labels"><span>${iPct.toFixed(0)}% I Joist</span><span>${(100-iPct).toFixed(0)}% Posi</span></div>`:'<div class="dashboard-empty compact">Re-import historical weekly files to build the job-type trend.</div>'}
        ${coverage.covered<coverage.total? `<div class="coverage-note">Job counts currently cover ${coverage.covered} of ${coverage.total} weeks in this view.</div>`:''}</div>
      </section>
    </div>
    <section class="dashboard-section dashboard-work-trend">
      <div class="dashboard-section-head">
        <div><h3>Amendment vs OTP trend</h3><span>Monthly work-category trend across the rolling ${dashboardMonths}-month view</span></div>
        <div class="dashboard-metric-toggle"><button type="button" data-work-metric="jobs" class="${dashboardWorkMetric==='jobs'?'active':''}">Jobs</button><button type="button" data-work-metric="hours" class="${dashboardWorkMetric==='hours'?'active':''}">Hours</button></div>
      </div>
      <div class="dashboard-work-summary">
        <div><span class="legend-dot amendment"></span><span>Amendments</span><strong>${dashboardWorkMetric==='hours'?workTotals.amendment.toFixed(1)+'h':workTotals.amendment}</strong></div>
        <div><span class="legend-dot otp"></span><span>OTPs</span><strong>${dashboardWorkMetric==='hours'?workTotals.otp.toFixed(1)+'h':workTotals.otp}</strong></div>
      </div>
      <div class="dashboard-chart">${dashboardWorkTrendSvg(workTrend,dashboardWorkMetric)}</div>
    </section>
    <section class="dashboard-section dashboard-comparison"><div class="dashboard-section-head"><div><h3>Designer comparison</h3><span>Summary only — click a designer for their detailed history</span></div></div>
      <div class="dashboard-table-wrap"><table><thead><tr><th>Designer</th><th class="right">Total hrs</th><th class="right">Avg hrs / wk</th><th class="right">Jobs</th><th class="right">I Joist</th><th class="right">Posi Joist</th><th class="right">4-week trend</th></tr></thead><tbody>${rows||'<tr><td colspan="7"><div class="dashboard-empty">No designer history is available for this period.</div></td></tr>'}</tbody></table></div>
    </section>
  </div>`;
  box.querySelectorAll('[data-dashboard-months]').forEach(btn=>btn.addEventListener('click',()=>{dashboardMonths=Number(btn.dataset.dashboardMonths)||6;renderDetail()}));
  box.querySelectorAll('[data-work-metric]').forEach(btn=>btn.addEventListener('click',()=>{dashboardWorkMetric=btn.dataset.workMetric==='hours'?'hours':'jobs';renderDetail()}));
  box.querySelectorAll('[data-dashboard-designer]').forEach(row=>{
    const open=()=>openDashboardDesigner(row.dataset.dashboardDesigner);
    row.addEventListener('click',open);row.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();open()}});
  });
}
const renderDetailBeforeDashboard=renderDetail;
renderDetail=function(){
  if(!selectedDesigner){renderDashboard();return}
  document.body.classList.remove('dashboard-home');
  renderDetailBeforeDashboard();
  const head=document.querySelector('#detailPanel .detail-top');
  if(head&&!head.querySelector('.back-dashboard')){
    const back=document.createElement('button');back.type='button';back.className='secondary back-dashboard';back.textContent='← Team Dashboard';
    back.addEventListener('click',()=>{selectedDesigner=null;if(typeof rangeActive!=='undefined')rangeActive=false;renderDesignerList();renderDetail()});
    head.prepend(back);
  }
};
const renderAllBeforeDashboard=renderAll;
renderAll=function(){renderAllBeforeDashboard();renderDetail()};
renderDetail();
