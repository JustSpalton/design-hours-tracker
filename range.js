let rangeActive=false;
let rangeStart='';
let rangeEnd='';

function addDaysISO(iso,days){const d=parseISO(iso);d.setUTCDate(d.getUTCDate()+days);return isoDate(d)}
function rangeLabel(){return rangeActive?`${fmtDate(rangeStart)} – ${fmtDate(rangeEnd)}`:`W/C ${fmtDate(selectedWeek)}`}
function rangeWeeks(){
  if(!rangeActive)return selectedWeek?[selectedWeek]:[];
  const start=parseISO(rangeStart),end=parseISO(rangeEnd);
  return allWeeks().filter(w=>{
    const ws=parseISO(w),we=parseISO(addDaysISO(w,4));
    return we>=start&&ws<=end;
  });
}
function recordsForDesignerInView(name){
  const weeks=new Set(rangeWeeks());
  return state.hours.filter(x=>norm(x.designer)===norm(name)&&weeks.has(x.week));
}
function viewHolidayCount(name){return rangeWeeks().reduce((sum,w)=>sum+holidayCount(name,w),0)}
function viewActual(name){const rows=recordsForDesignerInView(name);return rows.length?rows.reduce((s,x)=>s+Number(x.hours||0),0):null}
function viewAdjusted(name){
  const rows=recordsForDesignerInView(name);
  if(!rows.length)return null;
  return rows.reduce((s,x)=>{const a=adjustedHours(Number(x.hours),holidayCount(name,x.week));return s+(Number.isFinite(a)?a:0)},0);
}

function installRangeControls(){
  const weekbox=document.querySelector('.weekbox');
  const select=document.getElementById('weekSelect');
  if(!weekbox||!select||document.getElementById('dateRangeControls'))return;
  const wrap=document.createElement('div');
  wrap.id='dateRangeControls';
  wrap.className='date-range-controls';
  wrap.innerHTML=`<div class="range-divider"><span>or choose a date range</span></div><div class="range-fields"><label>From<input id="rangeStart" type="date"></label><label>To<input id="rangeEnd" type="date"></label></div><div class="range-actions"><button type="button" class="primary" id="applyRange">Apply range</button><button type="button" class="secondary" id="clearRange" hidden>Back to single week</button></div><div class="range-summary" id="rangeSummary"></div>`;
  weekbox.insertBefore(wrap,weekbox.querySelector('.live'));
  const weeks=allWeeks();
  const first=weeks[0]||selectedWeek||currentMonday();
  const last=weeks.at(-1)||selectedWeek||currentMonday();
  document.getElementById('rangeStart').value=first;
  document.getElementById('rangeEnd').value=addDaysISO(last,4);
  document.getElementById('applyRange').addEventListener('click',()=>{
    const from=document.getElementById('rangeStart').value;
    const to=document.getElementById('rangeEnd').value;
    if(!from||!to){showToast('Choose both a From and To date');return}
    if(from>to){showToast('The From date must be before the To date');return}
    rangeStart=from;rangeEnd=to;rangeActive=true;
    syncRangeUI();renderStats();renderDesignerList();renderDetail();
  });
  document.getElementById('clearRange').addEventListener('click',()=>{
    rangeActive=false;syncRangeUI();renderStats();renderDesignerList();renderDetail();
  });
  select.addEventListener('change',()=>{
    if(!rangeActive)return;
    rangeActive=false;syncRangeUI();renderStats();renderDesignerList();renderDetail();
  });
  syncRangeUI();
}
function syncRangeUI(){
  const clear=document.getElementById('clearRange'),summary=document.getElementById('rangeSummary');
  if(!clear||!summary)return;
  clear.hidden=!rangeActive;
  const weeks=rangeWeeks();
  summary.textContent=rangeActive?`${weeks.length} weekly record period${weeks.length===1?'':'s'} overlapping ${fmtDate(rangeStart)} to ${fmtDate(rangeEnd)}`:'';
  document.querySelector('.weekbox')?.classList.toggle('range-active',rangeActive);
}

const baseRenderWeeks=renderWeeks;
renderWeeks=function(){baseRenderWeeks();if(document.getElementById('dateRangeControls'))syncRangeUI()};

renderStats=function(){
  const weeks=rangeWeeks();
  const entries=[];
  for(const d of state.designers){
    for(const w of weeks){
      const h=getHours(d.name,w);
      if(typeof h!=='number')continue;
      const hols=holidayCount(d.name,w),adj=adjustedHours(h,hols);
      entries.push({designer:d.name,h,hols,adj});
    }
  }
  const total=entries.reduce((a,b)=>a+b.h,0);
  const adjusted=entries.map(x=>x.adj).filter(v=>typeof v==='number'&&Number.isFinite(v));
  const holidayDays=entries.reduce((a,b)=>a+b.hols,0);
  const designers=new Set(entries.map(x=>norm(x.designer))).size;
  document.getElementById('teamTotal').textContent=total.toFixed(1);
  document.getElementById('teamTotalSub').textContent=rangeLabel();
  document.getElementById('withHours').textContent=designers;
  document.getElementById('withHoursSub').textContent=rangeActive?`of ${state.designers.length} designers in range`:`of ${state.designers.length} designers`;
  document.getElementById('averageHours').textContent=(adjusted.length?adjusted.reduce((a,b)=>a+b,0)/adjusted.length:0).toFixed(1);
  document.getElementById('averageHoursSub').textContent=rangeActive?`${weeks.length} week${weeks.length===1?'':'s'} • ${holidayDays} holiday day${holidayDays===1?'':'s'} • per designer/week`:(holidayDays?`${holidayDays} holiday day${holidayDays===1?'':'s'} adjusted to a 5-day week`:'holiday-adjusted to a 5-day week');
};

renderDesignerList=function(){
  const q=norm(document.getElementById('designerSearch').value),box=document.getElementById('designerList');box.innerHTML='';
  state.designers.filter(d=>!q||norm(d.name).includes(q)).forEach(d=>{
    const h=rangeActive?viewActual(d.name):getHours(d.name,selectedWeek),hc=rangeActive?viewHolidayCount(d.name):holidayCount(d.name,selectedWeek);
    const b=document.createElement('button');b.className='designer '+(norm(d.name)===norm(selectedDesigner)?'active':'');
    b.innerHTML=`<span class="name">${escapeHtml(d.name)}${hc?`<span class="holiday-badge">${hc} holiday${hc===1?'':'s'}</span>`:''}</span><span class="hours ${h==null?'empty':''}">${h==null?'—':Number(h).toFixed(1)}</span>`;
    b.addEventListener('click',()=>{selectedDesigner=d.name;renderDesignerList();renderDetail()});box.appendChild(b)
  });
  if(!box.children.length)box.innerHTML='<div class="empty-state">No designers found.</div>';
};

renderDetail=function(){
  const box=document.getElementById('detailPanel');if(!selectedDesigner){box.innerHTML='<div class="empty-state">Select a designer.</div>';return}
  const weeks=rangeWeeks(),weekSet=new Set(weeks);
  const history=state.hours.filter(x=>norm(x.designer)===norm(selectedDesigner)&&weekSet.has(x.week)).sort((a,b)=>a.week.localeCompare(b.week)).map(x=>{const holidayDays=holidayCount(selectedDesigner,x.week);return {...x,holidayDays,adjusted:adjustedHours(Number(x.hours),holidayDays)}});
  const actual=history.length?history.reduce((s,x)=>s+Number(x.hours),0):null;
  const adjustedVals=history.map(x=>x.adjusted).filter(v=>typeof v==='number'&&Number.isFinite(v));
  const adjustedTotal=adjustedVals.length?adjustedVals.reduce((s,v)=>s+v,0):null;
  const avg=adjustedVals.length?adjustedVals.reduce((s,v)=>s+v,0)/adjustedVals.length:0;
  const hols=viewHolidayCount(selectedDesigner);
  const rows=weeks.slice().reverse().map(w=>{const x=hoursRecord(selectedDesigner,w);const hc=holidayCount(selectedDesigner,w);const adj=x?adjustedHours(Number(x.hours),hc):null;return `<tr><td>${fmtDate(w)}</td><td class="right"><strong>${x?Number(x.hours).toFixed(1):'—'}</strong></td><td>${hc?`${hc} day${hc===1?'':'s'}`:'—'}</td><td class="right"><strong>${adj!=null?Number(adj).toFixed(1):'—'}</strong></td><td>${x?`<span class="badge imported">Saved</span> ${escapeHtml(x.source_file||'')}`:'<span class="badge missing">No entry</span>'}</td></tr>`}).join('');
  let holidayCard='';
  if(rangeActive){
    holidayCard=`<div class="trend-card holiday-card"><div class="holiday-head"><div><h3>Holiday days in range</h3><span>${hols?`${hols} day${hols===1?'':'s'} recorded`:'No holiday recorded'}</span></div><div class="adjusted">${actual!=null&&hols?`${actual.toFixed(1)}h actual → ${adjustedTotal!=null?adjustedTotal.toFixed(1):'—'}h adjusted`:''}</div></div><div class="holiday-note">Holiday days are shown across the selected range. Switch back to a single week to add or remove individual holiday days.</div></div>`;
  }else{
    const currentHolidayDates=holidayDates(selectedDesigner,selectedWeek),currentHours=actual,currentAdjusted=adjustedTotal;const dayNames=['Mon','Tue','Wed','Thu','Fri'];const days=weekDays(selectedWeek);
    const holidayButtons=days.map((date,i)=>{const active=currentHolidayDates.includes(date);return `<button type="button" class="holiday-day ${active?'active':''}" data-holiday-date="${date}" ${(!editorKey||!state.holidayReady)?'disabled':''}><span class="dow">${dayNames[i]}</span><span class="date">${fmtDate(date).slice(0,5)}</span></button>`}).join('');
    holidayCard=`<div class="trend-card holiday-card"><div class="holiday-head"><div><h3>Holiday days this week</h3><span>${currentHolidayDates.length?`${currentHolidayDates.length} day${currentHolidayDates.length===1?'':'s'} recorded`:'No holiday recorded'}</span></div><div class="adjusted">${currentHours!=null&&currentHolidayDates.length?`${currentHours.toFixed(1)}h → ${currentAdjusted!=null?currentAdjusted.toFixed(1):'—'}h adjusted`:''}</div></div><div class="holiday-grid">${holidayButtons}</div><div class="holiday-note">${state.holidayReady?'Click a day to mark or unmark holiday.':'Holiday tracking database update is pending.'}</div></div>`;
  }
  box.innerHTML=`<div class="detail-top"><div><h2>${escapeHtml(selectedDesigner)}</h2><div class="range-detail-label">${rangeLabel()}</div></div><div class="metric-pair"><div class="metric-small"><div class="label">Adjusted${rangeActive?' total':''}</div><div class="value">${adjustedTotal!=null?adjustedTotal.toFixed(1):'—'}</div></div><div class="hero-hours"><div class="big">${actual!=null?actual.toFixed(1):'—'}</div><div class="small">actual hours • ${rangeActive?'selected range':`W/C ${fmtDate(selectedWeek)}`}</div></div></div></div><div class="detail-body">${holidayCard}<div class="trend-card"><div class="trend-head"><div><h3>Weekly trend</h3><span>${history.length} recorded week${history.length===1?'':'s'} • adjusted weekly average ${avg.toFixed(1)}h</span></div></div><div class="chart-wrap">${svgChart(history)}</div></div><div class="trend-card"><div class="trend-head"><div><h3>Weekly hours log</h3></div></div><div class="section-body"><table><thead><tr><th>Week commencing</th><th class="right">Actual hours</th><th>Holiday</th><th class="right">Adjusted</th><th>Source</th></tr></thead><tbody>${rows}</tbody></table></div></div></div>`;
  box.querySelectorAll('[data-holiday-date]').forEach(btn=>btn.addEventListener('click',()=>setHoliday(btn.dataset.holidayDate,!btn.classList.contains('active'))));
};

const baseRenderAll=renderAll;
renderAll=function(){baseRenderAll();installRangeControls();syncRangeUI()};

installRangeControls();
renderAll();
