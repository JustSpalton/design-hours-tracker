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
  return {name:designer,total,avg,fourWeekAvg:recentAvg,weeks:rows.length,trend,jobs:breakdown.jobs,iJoist:Number(breakdown.products['I Joist']||0),posiJoist:Number(breakdown.products['Posi Joist']||0)};
}
function dashboardWeeklyTotals(weeks){
  const data=weeks.map(week=>{
    const rows=(state.hours||[]).filter(x=>x.week===week);
    return {week,hours:rows.reduce((s,x)=>s+Number(x.hours||0),0),designers:new Set(rows.map(x=>norm(x.designer))).size};
  }).filter(x=>x.hours>0);
  return data.map((row,i)=>{
    const window=data.slice(Math.max(0,i-3),i+1);
    return {...row,rolling4:window.reduce((s,x)=>s+x.hours,0)/window.length};
  });
}
function dashboardTrendSvg(data){
  if(!data.length)return '<div class="dashboard-empty">No design-hour history is available for this period.</div>';
  const W=1000,H=300,p={l:52,r:18,t:20,b:50};const max=Math.max(10,Math.ceil(Math.max(...data.flatMap(x=>[x.hours,x.rolling4||0]))/20)*20);
  const x=i=>p.l+(data.length===1?(W-p.l-p.r)/2:i*(W-p.l-p.r)/(data.length-1));
  const y=v=>p.t+(max-v)*(H-p.t-p.b)/max;
  let grid='';
  for(let i=0;i<=4;i++){const v=max*(4-i)/4,yy=p.t+i*(H-p.t-p.b)/4;grid+=`<line x1="${p.l}" y1="${yy}" x2="${W-p.r}" y2="${yy}" stroke="#e5e7eb"/><text x="${p.l-9}" y="${yy+4}" text-anchor="end" font-size="10" fill="#6b7280">${v.toFixed(0)}</text>`}
  const pts=data.map((d,i)=>`${x(i)},${y(d.hours)}`).join(' ');
  const rollingPts=data.map((d,i)=>`${x(i)},${y(d.rolling4||0)}`).join(' ');
  const every=Math.max(1,Math.ceil(data.length/8));
  const labels=data.map((d,i)=>i%every===0||i===data.length-1?`<text x="${x(i)}" y="${H-19}" text-anchor="middle" font-size="10" fill="#6b7280">${fmtDate(d.week).slice(0,5)}</text>`:'').join('');
  const dots=data.map((d,i)=>`<circle class="dashboard-drill-point" data-drill="team-week" data-week="${d.week}" cx="${x(i)}" cy="${y(d.hours)}" r="5" fill="#d71920" tabindex="0"><title>W/C ${fmtDate(d.week)}: ${d.hours.toFixed(1)}h • click for breakdown</title></circle>`).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Team weekly design hours trend"><g class="trend-legend"><line x1="62" y1="12" x2="86" y2="12" stroke="#d71920" stroke-width="3"/><text x="91" y="15" font-size="10" fill="#6b7280">Weekly hours</text><line x1="166" y1="12" x2="190" y2="12" stroke="#222" stroke-width="2" stroke-dasharray="5 4"/><text x="195" y="15" font-size="10" fill="#6b7280">Rolling 4-week avg</text></g>${grid}<polyline points="${pts}" fill="none" stroke="#d71920" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/><polyline points="${rollingPts}" fill="none" stroke="#222" stroke-width="2" stroke-dasharray="5 4" stroke-linejoin="round" stroke-linecap="round"/>${dots}${labels}</svg>`;
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
  const amendmentDots=data.map((d,i)=>`<circle class="dashboard-drill-point" data-drill="work-month" data-month="${d.key}" data-category="Amendment" cx="${x(i)}" cy="${y(Number(d[amendmentKey]||0))}" r="5" fill="#d71920" tabindex="0"><title>${d.label} Amendments: ${Number(d[amendmentKey]||0).toFixed(metric==='hours'?1:0)}${suffix} • click for breakdown</title></circle>`).join('');
  const otpDots=data.map((d,i)=>`<circle class="dashboard-drill-point" data-drill="work-month" data-month="${d.key}" data-category="OTP" cx="${x(i)}" cy="${y(Number(d[otpKey]||0))}" r="5" fill="#222" tabindex="0"><title>${d.label} OTPs: ${Number(d[otpKey]||0).toFixed(metric==='hours'?1:0)}${suffix} • click for breakdown</title></circle>`).join('');
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
  const recordedHoursWeeks=new Set((state.hours||[]).filter(x=>wanted.has(x.week)).map(x=>x.week));
  const covered=new Set((typeof breakdownState!=='undefined'?(breakdownState.records||[]):[]).filter(x=>wanted.has(x.week)).map(x=>x.week));
  const missing=[...recordedHoursWeeks].filter(w=>!covered.has(w)).sort();
  return {covered:[...recordedHoursWeeks].filter(w=>covered.has(w)).length,total:recordedHoursWeeks.size,missing};
}
function dashboardQuantile(values,q){
  const a=values.filter(Number.isFinite).slice().sort((x,y)=>x-y);if(!a.length)return 0;
  const pos=(a.length-1)*q,base=Math.floor(pos),rest=pos-base;
  return a[base+1]!==undefined?a[base]+rest*(a[base+1]-a[base]):a[base];
}
function dashboardHourBands(weeks){
  const wanted=new Set(weeks),values=[];
  for(const row of state.hours||[]){
    if(!wanted.has(row.week))continue;
    const v=adjustedHours(Number(row.hours||0),holidayCount(row.designer,row.week));
    if(Number.isFinite(v))values.push(v);
  }
  return {low:dashboardQuantile(values,.25),high:dashboardQuantile(values,.75)};
}
function dashboardBandClass(value,bands){
  if(!Number.isFinite(value))return '';
  if(value<bands.low)return 'hours-band-low';
  if(value>bands.high)return 'hours-band-high';
  return 'hours-band-mid';
}
function dashboardSparkline(designer,weeks,bands){
  const wanted=new Set(weeks);
  const rows=(state.hours||[]).filter(x=>norm(x.designer)===norm(designer)&&wanted.has(x.week)).sort((a,b)=>a.week.localeCompare(b.week)).slice(-16);
  if(!rows.length)return '<span class="spark-empty">—</span>';
  const vals=rows.map(x=>adjustedHours(Number(x.hours||0),holidayCount(designer,x.week))).filter(Number.isFinite);
  if(!vals.length)return '<span class="spark-empty">—</span>';
  const W=118,H=34,p=3,min=Math.min(...vals,bands.low),max=Math.max(...vals,bands.high,1),span=Math.max(1,max-min);
  const x=i=>p+(vals.length===1?(W-p*2)/2:i*(W-p*2)/(vals.length-1));
  const y=v=>p+(max-v)*(H-p*2)/span;
  const yLow=y(bands.low),yHigh=y(bands.high);
  const top=Math.min(yHigh,yLow),midH=Math.abs(yLow-yHigh);
  const pts=vals.map((v,i)=>`${x(i)},${y(v)}`).join(' ');
  const last=vals.at(-1);
  return `<svg class="designer-spark" viewBox="0 0 ${W} ${H}" role="img" aria-label="Recent weekly hours trend"><rect x="0" y="0" width="${W}" height="${Math.max(0,top)}" class="spark-band-high"/><rect x="0" y="${top}" width="${W}" height="${midH}" class="spark-band-mid"/><rect x="0" y="${top+midH}" width="${W}" height="${Math.max(0,H-(top+midH))}" class="spark-band-low"/><polyline points="${pts}" fill="none" class="spark-line" stroke-width="2"/><circle cx="${x(vals.length-1)}" cy="${y(last)}" r="2.5" class="spark-dot"><title>Latest adjusted week: ${last.toFixed(1)}h</title></circle></svg>`;
}
function dashboardProductMix(p){
  const total=Number(p.iJoist||0)+Number(p.posiJoist||0);
  if(!total)return '<span class="mix-empty">—</span>';
  const i=Math.round(Number(p.iJoist||0)/total*100),posi=100-i;
  return `<div class="mix-percent"><div><strong>${i}%</strong><span>I</span></div><div><strong>${posi}%</strong><span>Posi</span></div></div>`;
}
function dashboardTeamFourWeekAverage(people){
  const vals=people.map(x=>x.fourWeekAvg).filter(Number.isFinite);
  return vals.length?vals.reduce((s,v)=>s+v,0)/vals.length:null;
}
function dashboardWorkloadIndicator(value,teamAvg){
  if(!Number.isFinite(value)||!Number.isFinite(teamAvg)||teamAvg===0)return '<span class="workload-badge workload-none">—</span>';
  const variance=(value-teamAvg)/teamAvg*100;
  const cls=variance>15?'workload-high':variance<-15?'workload-low':'workload-balanced';
  const label=variance>15?'High':variance<-15?'Low':'Balanced';
  const sign=variance>0?'+':'';
  return `<span class="workload-badge ${cls}" title="Latest 4-week adjusted average is ${sign}${variance.toFixed(1)}% vs team">${label}<small>${sign}${variance.toFixed(0)}%</small></span>`;
}
function ensureDashboardDrillModal(){
  let modal=document.getElementById('dashboardDrillModal');if(modal)return modal;
  modal=document.createElement('div');modal.id='dashboardDrillModal';modal.className='dashboard-drill-modal';modal.hidden=true;
  modal.innerHTML='<div class="dashboard-drill-shell"><div class="dashboard-drill-head"><div><div class="dashboard-kicker">DRILL DOWN</div><h2 id="dashboardDrillTitle">Details</h2><p id="dashboardDrillSub"></p></div><button type="button" class="secondary" id="closeDashboardDrill">Close</button></div><div id="dashboardDrillBody"></div></div>';
  document.body.appendChild(modal);
  modal.querySelector('#closeDashboardDrill').addEventListener('click',()=>{modal.hidden=true});
  modal.addEventListener('click',e=>{if(e.target===modal)modal.hidden=true});
  return modal;
}
function dashboardWeekDrill(week){
  const modal=ensureDashboardDrillModal(),records=(state.hours||[]).filter(x=>x.week===week).sort((a,b)=>Number(b.hours||0)-Number(a.hours||0));
  const breakdowns=(typeof breakdownState!=='undefined'?(breakdownState.records||[]):[]).filter(x=>x.week===week);
  const rows=records.map(row=>{
    const br=breakdowns.find(x=>norm(x.designer)===norm(row.designer))||{};
    const adj=adjustedHours(Number(row.hours||0),holidayCount(row.designer,week));
    const jobs=Object.values(br.products||{}).reduce((s,v)=>s+Number(v||0),0);
    return `<tr class="dashboard-drill-person" data-dashboard-designer="${escapeHtml(row.designer)}"><td><strong>${escapeHtml(row.designer)}</strong></td><td class="right">${Number(row.hours||0).toFixed(1)}</td><td class="right">${Number.isFinite(adj)?adj.toFixed(1):'—'}</td><td class="right">${jobs||'—'}</td><td class="right">${Number(br.statuses?.Amendment||0)||'—'}</td><td class="right">${Number(br.statuses?.OTP||0)||'—'}</td></tr>`;
  }).join('');
  modal.querySelector('#dashboardDrillTitle').textContent=`W/C ${fmtDate(week)}`;
  modal.querySelector('#dashboardDrillSub').textContent='Click a designer to open their full history.';
  modal.querySelector('#dashboardDrillBody').innerHTML=`<div class="dashboard-drill-table"><table><thead><tr><th>Designer</th><th class="right">Hours</th><th class="right">Adjusted</th><th class="right">Jobs</th><th class="right">Amendments</th><th class="right">OTPs</th></tr></thead><tbody>${rows||'<tr><td colspan="6">No data.</td></tr>'}</tbody></table></div>`;
  modal.hidden=false;
  modal.querySelectorAll('[data-dashboard-designer]').forEach(row=>row.addEventListener('click',()=>{modal.hidden=true;openDashboardDesigner(row.dataset.dashboardDesigner)}));
}
function dashboardWorkMonthDrill(month,category){
  const modal=ensureDashboardDrillModal(),metric=dashboardWorkMetric;
  const records=(typeof breakdownState!=='undefined'?(breakdownState.records||[]):[]).filter(x=>dashboardMonthKey(x.week)===month);
  const grouped=new Map();
  for(const rec of records){
    const key=rec.designer;if(!grouped.has(key))grouped.set(key,{designer:key,jobs:0,hours:0});
    const row=grouped.get(key);row.jobs+=Number(rec.statuses?.[category]||0);row.hours+=Number(rec.status_hours?.[category]||0);
  }
  const rows=[...grouped.values()].filter(x=>x.jobs||x.hours).sort((a,b)=>(metric==='hours'?b.hours-a.hours:b.jobs-a.jobs)).map(x=>`<tr class="dashboard-drill-person" data-dashboard-designer="${escapeHtml(x.designer)}"><td><strong>${escapeHtml(x.designer)}</strong></td><td class="right">${x.jobs}</td><td class="right">${x.hours.toFixed(1)}</td></tr>`).join('');
  modal.querySelector('#dashboardDrillTitle').textContent=`${category} — ${dashboardMonthLabel(month)}`;
  modal.querySelector('#dashboardDrillSub').textContent=`Designer contribution for this month • chart currently showing ${metric}.`;
  modal.querySelector('#dashboardDrillBody').innerHTML=`<div class="dashboard-drill-table"><table><thead><tr><th>Designer</th><th class="right">Jobs</th><th class="right">Hours</th></tr></thead><tbody>${rows||'<tr><td colspan="3">No data.</td></tr>'}</tbody></table></div>`;
  modal.hidden=false;
  modal.querySelectorAll('[data-dashboard-designer]').forEach(row=>row.addEventListener('click',()=>{modal.hidden=true;openDashboardDesigner(row.dataset.dashboardDesigner)}));
}
function bindDashboardDrilldowns(box){
  box.querySelectorAll('[data-drill]').forEach(el=>{
    const open=()=>{
      if(el.dataset.drill==='team-week')dashboardWeekDrill(el.dataset.week);
      if(el.dataset.drill==='work-month')dashboardWorkMonthDrill(el.dataset.month,el.dataset.category);
    };
    el.addEventListener('click',open);
    el.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();open()}});
  });
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
  const teamFourWeekAvg=dashboardTeamFourWeekAverage(people);
  const allBreakdowns=people.reduce((a,x)=>({jobs:a.jobs+x.jobs,iJoist:a.iJoist+x.iJoist,posiJoist:a.posiJoist+x.posiJoist}),{jobs:0,iJoist:0,posiJoist:0});
  const coverage=dashboardCoverage(weeks);
  const hourBands=dashboardHourBands(weeks);
  const workTrend=dashboardWorkCategoryTrend(weeks);
  const workTotals=dashboardWorkCategoryTotals(workTrend,dashboardWorkMetric);
  const splitTotal=allBreakdowns.iJoist+allBreakdowns.posiJoist;
  const iPct=splitTotal?allBreakdowns.iJoist/splitTotal*100:0;
  const rows=people.map(p=>`<tr class="dashboard-person" data-dashboard-designer="${escapeHtml(p.name)}" tabindex="0"><td><strong>${escapeHtml(p.name)}</strong><span class="dashboard-row-hint">View details</span></td><td class="spark-cell">${dashboardSparkline(p.name,weeks,hourBands)}</td><td class="right">${p.total.toFixed(1)}</td><td class="right"><span class="hours-band ${dashboardBandClass(p.avg,hourBands)}" title="Relative to the team distribution for this rolling period">${p.avg.toFixed(1)}</span></td><td class="right"><strong>${Number.isFinite(p.fourWeekAvg)?p.fourWeekAvg.toFixed(1):'—'}</strong></td><td>${dashboardWorkloadIndicator(p.fourWeekAvg,teamFourWeekAvg)}</td><td class="right">${p.jobs||'—'}</td><td class="right">${p.iJoist||'—'}</td><td class="right">${p.posiJoist||'—'}</td><td>${dashboardProductMix(p)}</td><td class="right">${dashboardTrendText(p.trend)}</td></tr>`).join('');
  box.innerHTML=`<div class="dashboard-shell">
    <div class="dashboard-head"><div><div class="dashboard-kicker">TEAM OVERVIEW</div><h2>Design Hours Dashboard</h2><p>Rolling ${dashboardMonths} months • ${fmtDate(bounds.start)} to ${fmtDate(bounds.end)}</p></div><div class="dashboard-period"><button type="button" data-dashboard-months="6" class="${dashboardMonths===6?'active':''}">6 months</button><button type="button" data-dashboard-months="12" class="${dashboardMonths===12?'active':''}">12 months</button></div></div>
    <div class="dashboard-cards">
      <div class="dashboard-card"><span>Total design hours</span><strong>${totalHours.toFixed(1)}</strong><em>${weekly.length} recorded weeks</em></div>
      <div class="dashboard-card"><span>Latest 4-week team average</span><strong>${Number.isFinite(teamFourWeekAvg)?teamFourWeekAvg.toFixed(1):'—'}</strong><em>Adjusted hours per designer / week</em></div>
      <div class="dashboard-card"><span>Jobs recorded</span><strong>${allBreakdowns.jobs}</strong><em>${coverage.total?`${coverage.covered} of ${coverage.total} recorded weeks have job data`:'No recorded weeks in range'}</em></div>
      <div class="dashboard-card"><span>Active designers</span><strong>${people.length}</strong><em>With hours in this period</em></div>
    </div>
    <div class="dashboard-readability-row">
      <div class="dashboard-missing ${coverage.missing.length?'has-missing':'complete'}"><strong>${coverage.missing.length?'Missing job-data weeks':'Job data complete'}</strong><span>${coverage.missing.length?coverage.missing.map(w=>`W/C ${fmtDate(w)}`).join(' • '):'Every recorded hours week in this view has job breakdown data.'}</span></div>
      <div class="hours-band-legend"><span>Weekly hour bands:</span><i class="hours-band-low"></i><b>Lower</b><i class="hours-band-mid"></i><b>Typical</b><i class="hours-band-high"></i><b>Higher</b><em>relative to this rolling period</em></div>
    </div>
    <div class="dashboard-grid">
      <section class="dashboard-section dashboard-trend"><div class="dashboard-section-head"><div><h3>Team weekly hours trend</h3><span>Actual hours + rolling 4-week average • click a point to drill down</span></div></div><div class="dashboard-chart">${dashboardTrendSvg(weekly)}</div></section>
      <section class="dashboard-section dashboard-split"><div class="dashboard-section-head"><div><h3>Job type split</h3><span>I Joist vs Posi Joist</span></div></div>
        <div class="dashboard-split-body"><div class="split-numbers"><div><span>I Joist</span><strong>${allBreakdowns.iJoist}</strong></div><div><span>Posi Joist</span><strong>${allBreakdowns.posiJoist}</strong></div></div>
        ${splitTotal?`<div class="split-bar"><span style="width:${iPct}%"></span></div><div class="split-labels"><span>${iPct.toFixed(0)}% I Joist</span><span>${(100-iPct).toFixed(0)}% Posi</span></div>`:'<div class="dashboard-empty compact">Re-import historical weekly files to build the job-type trend.</div>'}
        ${coverage.missing.length? `<div class="coverage-note">${coverage.missing.length} recorded week${coverage.missing.length===1?' is':'s are'} missing job breakdown data. See the list above.</div>`:''}</div>
      </section>
    </div>
    <section class="dashboard-section dashboard-work-trend">
      <div class="dashboard-section-head">
        <div><h3>Amendment vs OTP trend</h3><span>Monthly trend across the rolling ${dashboardMonths}-month view • click a point to drill down</span></div>
        <div class="dashboard-metric-toggle"><button type="button" data-work-metric="jobs" class="${dashboardWorkMetric==='jobs'?'active':''}">Jobs</button><button type="button" data-work-metric="hours" class="${dashboardWorkMetric==='hours'?'active':''}">Hours</button></div>
      </div>
      <div class="dashboard-work-summary">
        <div><span class="legend-dot amendment"></span><span>Amendments</span><strong>${dashboardWorkMetric==='hours'?workTotals.amendment.toFixed(1)+'h':workTotals.amendment}</strong></div>
        <div><span class="legend-dot otp"></span><span>OTPs</span><strong>${dashboardWorkMetric==='hours'?workTotals.otp.toFixed(1)+'h':workTotals.otp}</strong></div>
      </div>
      <div class="dashboard-chart">${dashboardWorkTrendSvg(workTrend,dashboardWorkMetric)}</div>
    </section>
    <section class="dashboard-section dashboard-comparison"><div class="dashboard-section-head"><div><h3>Designer comparison</h3><span>Summary only — click a designer for their detailed history</span></div></div>
      <div class="dashboard-table-wrap"><table><thead><tr><th>Designer</th><th>Recent trend</th><th class="right">Total hrs</th><th class="right">Period avg</th><th class="right">4-wk avg</th><th>Workload</th><th class="right">Jobs</th><th class="right">I Joist</th><th class="right">Posi Joist</th><th>Product mix</th><th class="right">4-week change</th></tr></thead><tbody>${rows||'<tr><td colspan="11"><div class="dashboard-empty">No designer history is available for this period.</div></td></tr>'}</tbody></table></div>
    </section>
  </div>`;
  box.querySelectorAll('[data-dashboard-months]').forEach(btn=>btn.addEventListener('click',()=>{dashboardMonths=Number(btn.dataset.dashboardMonths)||6;renderDetail()}));
  box.querySelectorAll('[data-work-metric]').forEach(btn=>btn.addEventListener('click',()=>{dashboardWorkMetric=btn.dataset.workMetric==='hours'?'hours':'jobs';renderDetail()}));
  bindDashboardDrilldowns(box);
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
