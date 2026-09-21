let state={designers:[],hours:[],holidays:[],holidayReady:false,importLog:[],teams:[]};
let selectedWeek=null, selectedDesigner=null;
let editorKey=sessionStorage.getItem('designHoursEditorKey')||'';

function norm(s){return String(s??'').trim().toLowerCase().replace(/[^a-z0-9]+/g,' ')}
function canonicalDesignerName(name){const clean=String(name??'').replace(/\s+/g,' ').trim();const aliases={'jon wilson':'Jonathan Wilson','kerry mui':'Kerry Gardiner'};return aliases[clean.toLowerCase()]||clean}
function escapeHtml(s){return String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function parseISO(s){const [y,m,d]=s.split('-').map(Number);return new Date(Date.UTC(y,m-1,d))}
function fmtDate(k){if(!k)return '—';return parseISO(k).toLocaleDateString('en-GB',{timeZone:'UTC',day:'2-digit',month:'2-digit',year:'numeric'})}
function isoDate(d){return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`}
function currentMonday(){const d=new Date();const x=new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate()));const day=x.getUTCDay();x.setUTCDate(x.getUTCDate()-((day+6)%7));return isoDate(x)}
function parseWeekFilename(name){const m=String(name).match(/(?:^|[^0-9])(\d{1,2})[._-](\d{1,2})[._-](\d{2,4})(?:[^0-9]|$)/);if(!m)return null;let y=+m[3];if(y<100)y+=2000;const d=new Date(Date.UTC(y,+m[2]-1,+m[1]));return isNaN(d)?null:d}
function showToast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');clearTimeout(showToast.t);showToast.t=setTimeout(()=>t.classList.remove('show'),2500)}
function showError(msg=''){const e=document.getElementById('status');e.textContent=msg;e.classList.toggle('show',!!msg)}
function isExcelFile(file){return /\.(xls|xlsx|xlsm|xlsb)$/i.test(String(file?.name||''))}
function cellText(cell){return String(cell?.w??cell?.v??'').trim()}
function numericHours(cell){if(!cell)return null;const formatted=cellText(cell);if(/^sum\s*=/i.test(formatted))return null;const raw=cell.v;if(typeof raw==='number'&&Number.isFinite(raw))return raw;const n=Number(String(raw??formatted).replace(/,/g,'').trim());return Number.isFinite(n)?n:null}
function cellISODate(cell){
  if(!cell)return null;
  const raw=cell.v;
  if(raw instanceof Date&&!isNaN(raw)){return isoDate(new Date(Date.UTC(raw.getFullYear(),raw.getMonth(),raw.getDate())))}
  if(typeof raw==='number'&&Number.isFinite(raw)&&raw>20000&&raw<80000){
    try{const p=XLSX.SSF.parse_date_code(raw);if(p&&p.y&&p.m&&p.d)return isoDate(new Date(Date.UTC(p.y,p.m-1,p.d)))}catch(_){}
  }
  const text=cellText(cell).replace(/\s+/g,' ').trim();
  let m=text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(?:\s|$)/);
  if(m){let y=Number(m[3]);if(y<100)y+=2000;const d=new Date(Date.UTC(y,Number(m[2])-1,Number(m[1])));if(!isNaN(d)&&d.getUTCDate()===Number(m[1])&&d.getUTCMonth()===Number(m[2])-1)return isoDate(d)}
  m=text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s|$)/);
  if(m){const d=new Date(Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3])));if(!isNaN(d))return isoDate(d)}
  return null;
}
function findHeaderColumn(ws,range,patterns,maxRows=40){
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+maxRows);r++)for(let col=range.s.c;col<=range.e.c;col++){
    const t=cellText(ws[XLSX.utils.encode_cell({r,c:col})]).replace(/\s+/g,' ').trim();
    if(patterns.some(re=>re.test(t)))return col;
  }
  return null;
}
function dateEnquiryDoneColumn(ws,range){return findHeaderColumn(ws,range,[/^date\s+enquiry\s+done$/i,/^enquiry\s+done\s+date$/i,/^date\s+done$/i],60)}
function extractDatedHours(wb){
  const grouped=new Map(),ignored=new Set();let datedRows=0;
  for(const sheetName of wb.SheetNames){
    const ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])continue;
    const range=XLSX.utils.decode_range(ws['!ref']),dateCol=dateEnquiryDoneColumn(ws,range);
    if(dateCol==null)continue;
    const {estimatorCol,hoursCol}=getReportColumns(ws,range),starts=[];
    for(let r=range.s.r;r<=range.e.r;r++){const name=sectionNameAtRow(ws,r,range);if(name)starts.push({r,name})}
    for(let i=0;i<starts.length;i++){
      const start=starts[i],end=i+1<starts.length?starts[i+1].r-1:range.e.r;
      const designer=knownDesigner(start.name);if(!designer){ignored.add(start.name);continue}
      for(let r=start.r+1;r<=end;r++){
        const rowEstimator=cellText(ws[XLSX.utils.encode_cell({r,c:estimatorCol})]);
        if(rowEstimator&&norm(rowEstimator)!==norm(start.name))continue;
        const date=cellISODate(ws[XLSX.utils.encode_cell({r,c:dateCol})]);if(!date)continue;
        const hours=numericHours(ws[XLSX.utils.encode_cell({r,c:hoursCol})]);if(hours==null||hours<0||hours>1000)continue;
        const week=weekForDate(date),key=`${week}|${norm(designer)}`,previous=grouped.get(key);
        grouped.set(key,{week,name:designer,hours:Math.round(((previous?.hours||0)+hours)*100)/100});
        datedRows++;
      }
    }
  }
  return {records:[...grouped.values()],ignored:[...ignored],datedRows};
}
function getReportColumns(ws,range){let estimatorCol=null,hoursCol=null;for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+15);r++){for(let c=range.s.c;c<=range.e.c;c++){const txt=cellText(ws[XLSX.utils.encode_cell({r,c})]);if(estimatorCol==null&&/^estimator\s*name\s*$/i.test(txt))estimatorCol=c;if(hoursCol==null&&/^est\.?\s*time\s*est\s*$/i.test(txt))hoursCol=c}}return {estimatorCol:estimatorCol??8,hoursCol:hoursCol??9}}
function sectionNameAtRow(ws,r,range){for(let c=range.s.c;c<=range.e.c;c++){const txt=cellText(ws[XLSX.utils.encode_cell({r,c})]);if(!/^estimator\s*name\s*:/i.test(txt))continue;const inline=txt.replace(/^estimator\s*name\s*:\s*/i,'').trim();if(inline)return inline.replace(/\s+/g,' ');for(let n=c+1;n<=Math.min(range.e.c,c+3);n++){const next=cellText(ws[XLSX.utils.encode_cell({r,c:n})]).replace(/\s+/g,' ').trim();if(next)return next}}return null}
function findEstimatorSections(wb){const sections=[];for(const sheetName of wb.SheetNames){const ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])continue;const range=XLSX.utils.decode_range(ws['!ref']);const {estimatorCol,hoursCol}=getReportColumns(ws,range);const starts=[];for(let r=range.s.r;r<=range.e.r;r++){const name=sectionNameAtRow(ws,r,range);if(name)starts.push({r,name})}for(let i=0;i<starts.length;i++){const start=starts[i],end=i+1<starts.length?starts[i+1].r-1:range.e.r;let sum=0,count=0;for(let r=start.r+1;r<=end;r++){const rowName=cellText(ws[XLSX.utils.encode_cell({r,c:estimatorCol})]);if(norm(rowName)!==norm(start.name))continue;const v=numericHours(ws[XLSX.utils.encode_cell({r,c:hoursCol})]);if(v==null||v<0||v>1000)continue;sum+=v;count++}if(!count){sum=0;for(let r=start.r+1;r<=end;r++){const v=numericHours(ws[XLSX.utils.encode_cell({r,c:hoursCol})]);if(v==null||v<0||v>1000)continue;sum+=v;count++}}if(count)sections.push({name:start.name,hours:Math.round(sum*100)/100,count})}}return sections}
function knownDesigner(name){const canonical=canonicalDesignerName(name);return state.designers.find(d=>norm(d.name)===norm(canonical))?.name||null}
function weekForDate(date){const d=parseISO(date);const day=d.getUTCDay();d.setUTCDate(d.getUTCDate()-((day+6)%7));return isoDate(d)}
function weekDays(week){const start=parseISO(week);return Array.from({length:5},(_,i)=>{const d=new Date(start);d.setUTCDate(start.getUTCDate()+i);return isoDate(d)})}
function holidayDates(designer,week){const days=new Set(weekDays(week));return (state.holidays||[]).filter(x=>norm(x.designer)===norm(designer)&&days.has(x.date)).map(x=>x.date).sort()}
function holidayCount(designer,week){return holidayDates(designer,week).length}
function adjustedHours(hours,holidays){const working=5-Number(holidays||0);return typeof hours==='number'&&Number.isFinite(hours)&&working>0?hours*5/working:null}
function allWeeks(){const set=new Set(state.hours.map(x=>x.week));(state.holidays||[]).forEach(x=>set.add(weekForDate(x.date)));set.add(currentMonday());return [...set].sort()}
function hoursRecord(designer,week){return state.hours.find(x=>norm(x.designer)===norm(designer)&&x.week===week)||null}
function getHours(designer,week){const x=hoursRecord(designer,week);return x?Number(x.hours):null}
function setEditingUI(){const edit=!!editorKey;document.getElementById('editState').textContent=edit?'Editing unlocked':'View only';document.getElementById('editState').className='lock'+(edit?' edit':'');document.getElementById('unlockBtn').textContent=edit?'Lock editing':'Unlock editing';document.querySelectorAll('.edit-control').forEach(el=>{el.style.display=edit?'':'none'});document.querySelectorAll('.holiday-day').forEach(el=>{el.disabled=!edit||!state.holidayReady})}
async function api(path,options={}){while(typeof window.firebaseApi!=='function')await new Promise(resolve=>setTimeout(resolve,20));return window.firebaseApi(path,options)}
async function loadData(silent=false){try{document.getElementById('syncText').textContent='Syncing…';const data=await api('/api/data');const unchanged=silent&&JSON.stringify(data)===JSON.stringify(state);state=data;const weeks=allWeeks();selectedWeek=weeks.includes(selectedWeek)?selectedWeek:(weeks.at(-1)||currentMonday());if(selectedDesigner&&!state.designers.some(d=>norm(d.name)===norm(selectedDesigner)))selectedDesigner=null;showError('');if(!unchanged)renderAll();document.getElementById('syncText').textContent='Shared tracker';if(!silent)showToast('Tracker refreshed')}catch(e){showError(e.message);document.getElementById('syncText').textContent='Sync failed'}}
function updateWeekNav(){
  const weeks=allWeeks(),i=weeks.indexOf(selectedWeek),prev=document.getElementById('previousWeekBtn'),next=document.getElementById('nextWeekBtn');
  if(prev)prev.disabled=i<=0;
  if(next)next.disabled=i<0||i>=weeks.length-1;
}
function renderWeeks(){const sel=document.getElementById('weekSelect');sel.innerHTML='';allWeeks().slice().reverse().forEach(k=>{const o=document.createElement('option');o.value=k;o.textContent=`W/C ${fmtDate(k)}`;o.selected=k===selectedWeek;sel.appendChild(o)});updateWeekNav()}
function selectTrackerWeek(week){
  if(!week||week===selectedWeek)return;
  selectedWeek=week;
  renderWeeks();
  renderStats();
  renderDesignerList();
  renderDetail();
  if(typeof syncRangeUI==='function')syncRangeUI();
}
function moveTrackerWeek(direction){
  const weeks=allWeeks(),i=weeks.indexOf(selectedWeek),nextIndex=i+direction;
  if(i<0||nextIndex<0||nextIndex>=weeks.length)return;
  selectTrackerWeek(weeks[nextIndex]);
}
function renderStats(){const entries=state.designers.map(d=>{const h=getHours(d.name,selectedWeek),hols=holidayCount(d.name,selectedWeek);return {h,hols,adj:adjustedHours(h,hols)}}).filter(x=>typeof x.h==='number');const total=entries.reduce((a,b)=>a+b.h,0);const adjusted=entries.map(x=>x.adj).filter(v=>typeof v==='number'&&Number.isFinite(v));const holidayDays=entries.reduce((a,b)=>a+b.hols,0);document.getElementById('teamTotal').textContent=total.toFixed(1);document.getElementById('teamTotalSub').textContent=`W/C ${fmtDate(selectedWeek)}`;document.getElementById('withHours').textContent=entries.length;document.getElementById('withHoursSub').textContent=`of ${state.designers.length} designers`;document.getElementById('averageHours').textContent=(adjusted.length?adjusted.reduce((a,b)=>a+b,0)/adjusted.length:0).toFixed(1);document.getElementById('averageHoursSub').textContent=holidayDays?`${holidayDays} holiday day${holidayDays===1?'':'s'} adjusted to a 5-day week`:'holiday-adjusted to a 5-day week'}
function renderDesignerList(){const q=norm(document.getElementById('designerSearch').value),box=document.getElementById('designerList');box.innerHTML='';state.designers.filter(d=>!q||norm(d.name).includes(q)).forEach(d=>{const h=getHours(d.name,selectedWeek),hc=holidayCount(d.name,selectedWeek);const b=document.createElement('button');b.className='designer '+(norm(d.name)===norm(selectedDesigner)?'active':'');b.innerHTML=`<span class="name">${escapeHtml(d.name)}${hc?`<span class="holiday-badge">${hc} holiday</span>`:''}</span><span class="hours ${h==null?'empty':''}">${h==null?'—':Number(h).toFixed(1)}</span>`;b.addEventListener('click',()=>{selectedDesigner=d.name;renderDesignerList();renderDetail()});box.appendChild(b)});if(!box.children.length)box.innerHTML='<div class="empty-state">No designers found.</div>'}
function svgChart(data){if(!data.length)return '<div class="empty-state">No recorded hours yet.</div>';const W=760,H=250,p={l:46,r:16,t:20,b:46};const vals=data.map(d=>d.hours);const max=Math.max(10,Math.ceil(Math.max(...vals)/10)*10);const x=i=>p.l+(data.length===1?(W-p.l-p.r)/2:i*(W-p.l-p.r)/(data.length-1));const y=v=>p.t+(max-v)*(H-p.t-p.b)/max;let grid='';for(let i=0;i<=4;i++){const v=max*(4-i)/4,yy=p.t+i*(H-p.t-p.b)/4;grid+=`<line x1="${p.l}" y1="${yy}" x2="${W-p.r}" y2="${yy}" stroke="#e5e7eb"/><text x="${p.l-8}" y="${yy+4}" text-anchor="end" font-size="10" fill="#6b7280">${v.toFixed(0)}</text>`}const pts=data.map((d,i)=>`${x(i)},${y(d.hours)}`).join(' ');const dots=data.map((d,i)=>`<circle cx="${x(i)}" cy="${y(d.hours)}" r="4" fill="#d71920"><title>${fmtDate(d.week)}: ${Number(d.hours).toFixed(1)} hours${d.holidayDays?` • ${d.holidayDays} holiday day${d.holidayDays===1?'':'s'} • adjusted ${Number(d.adjusted).toFixed(1)}h`:''}</title></circle>`).join('');const labels=data.map((d,i)=>{if(data.length>12&&i%2===1)return '';const dt=parseISO(d.week);return `<text x="${x(i)}" y="${H-18}" text-anchor="middle" font-size="10" fill="#6b7280">${String(dt.getUTCDate()).padStart(2,'0')}/${String(dt.getUTCMonth()+1).padStart(2,'0')}</text>`}).join('');return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Weekly design hours trend">${grid}<polyline points="${pts}" fill="none" stroke="#d71920" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>${dots}${labels}</svg>`}
function renderDetail(){const box=document.getElementById('detailPanel');if(!selectedDesigner){box.innerHTML='<div class="empty-state">Select a designer.</div>';return}const current=hoursRecord(selectedDesigner,selectedWeek);const currentHours=current?Number(current.hours):null;const currentHolidayDates=holidayDates(selectedDesigner,selectedWeek);const currentAdjusted=adjustedHours(currentHours,currentHolidayDates.length);const history=state.hours.filter(x=>norm(x.designer)===norm(selectedDesigner)).sort((a,b)=>a.week.localeCompare(b.week)).map(x=>{const holidayDays=holidayCount(selectedDesigner,x.week);return {...x,holidayDays,adjusted:adjustedHours(Number(x.hours),holidayDays)}});const trendHistory=history.filter(x=>x.week<=selectedWeek).slice(-12);const recentAdjusted=trendHistory.map(x=>x.adjusted).filter(v=>typeof v==='number'&&Number.isFinite(v));const avg=recentAdjusted.length?recentAdjusted.reduce((s,v)=>s+v,0)/recentAdjusted.length:0;const dayNames=['Mon','Tue','Wed','Thu','Fri'];const days=weekDays(selectedWeek);const holidayButtons=days.map((date,i)=>{const active=currentHolidayDates.includes(date);return `<button type="button" class="holiday-day ${active?'active':''}" data-holiday-date="${date}" ${(!editorKey||!state.holidayReady)?'disabled':''}><span class="dow">${dayNames[i]}</span><span class="date">${fmtDate(date).slice(0,5)}</span></button>`}).join('');const rows=allWeeks().slice().reverse().map(w=>{const x=hoursRecord(selectedDesigner,w);const hc=holidayCount(selectedDesigner,w);const adj=x?adjustedHours(Number(x.hours),hc):null;return `<tr><td>${fmtDate(w)}</td><td class="right"><strong>${x?Number(x.hours).toFixed(1):'—'}</strong></td><td>${hc?`${hc} day${hc===1?'':'s'}`:'—'}</td><td class="right"><strong>${adj!=null?Number(adj).toFixed(1):'—'}</strong></td><td>${x?`<span class="badge imported">Saved</span> ${escapeHtml(x.source_file||'')}`:'<span class="badge missing">No entry</span>'}</td></tr>`}).join('');box.innerHTML=`<div class="detail-top"><div><h2>${escapeHtml(selectedDesigner)}</h2></div><div class="metric-pair"><div class="metric-small"><div class="label">Adjusted</div><div class="value">${currentAdjusted!=null?currentAdjusted.toFixed(1):'—'}</div></div><div class="hero-hours"><div class="big">${currentHours!=null?currentHours.toFixed(1):'—'}</div><div class="small">actual hours • W/C ${fmtDate(selectedWeek)}</div></div></div></div><div class="detail-body"><div class="trend-card holiday-card"><div class="holiday-head"><div><h3>Holiday days this week</h3><span>${currentHolidayDates.length?`${currentHolidayDates.length} day${currentHolidayDates.length===1?'':'s'} recorded`:'No holiday recorded'}</span></div><div class="adjusted">${currentHours!=null&&currentHolidayDates.length?`${currentHours.toFixed(1)}h → ${currentAdjusted!=null?currentAdjusted.toFixed(1):'—'}h adjusted`:''}</div></div><div class="holiday-grid">${holidayButtons}</div><div class="holiday-note">${state.holidayReady?(editorKey?'Click a day to mark or unmark holiday.':'Unlock editing to change holiday days.'):'Holiday tracking database update is pending.'}</div></div><div class="trend-card"><div class="trend-head"><div><h3>12-week trend</h3><span>${trendHistory.length} recorded week${trendHistory.length===1?'':'s'} ending W/C ${fmtDate(selectedWeek)} • adjusted average ${avg.toFixed(1)}h</span></div></div><div class="chart-wrap">${svgChart(trendHistory)}</div></div><div class="trend-card"><div class="trend-head"><div><h3>Weekly hours log</h3></div></div><div class="section-body"><table><thead><tr><th>Week commencing</th><th class="right">Actual hours</th><th>Holiday</th><th class="right">Adjusted</th><th>Source</th></tr></thead><tbody>${rows}</tbody></table></div></div></div>`;box.querySelectorAll('[data-holiday-date]').forEach(btn=>btn.addEventListener('click',()=>setHoliday(btn.dataset.holidayDate,!btn.classList.contains('active'))))}
function renderLog(){const box=document.getElementById('importLog');if(!state.importLog?.length){box.innerHTML='<div class="empty-state">No imports recorded yet.</div>';return}box.innerHTML=`<table><thead><tr><th>Week</th><th>File</th><th class="right">Designers saved</th><th>Imported</th></tr></thead><tbody>${state.importLog.map(x=>`<tr><td>${fmtDate(x.week)}</td><td>${escapeHtml(x.source_file)}</td><td class="right">${Number(x.rows_imported||0)}</td><td>${new Date(x.imported_at).toLocaleString('en-GB')}</td></tr>`).join('')}</tbody></table>`}
function renderAll(){renderWeeks();renderStats();renderDesignerList();renderDetail();renderLog();setEditingUI()}
async function setHoliday(date,holiday){if(!editorKey){showToast('Unlock editing first');return}if(!state.holidayReady){showToast('Holiday tracking is not ready yet');return}try{await api('/api/holidays',{method:'POST',body:JSON.stringify({designer:selectedDesigner,date,holiday})});await loadData(true);showToast(holiday?`Holiday added for ${fmtDate(date)}`:`Holiday removed for ${fmtDate(date)}`)}catch(e){alert(e.message)}}
async function importEstimatorFile(file){
  if(!editorKey)throw new Error('Unlock editing before importing.');
  if(!isExcelFile(file))throw new Error(`${file.name}: unsupported Excel file.`);
  const buf=await file.arrayBuffer();let wb;
  try{wb=XLSX.read(buf,{type:'array',cellDates:true,cellFormula:true,cellStyles:true,cellNF:true,cellText:true})}
  catch(e){throw new Error(`${file.name}: Excel could not be read (${e.message}).`)}

  const dated=extractDatedHours(wb);
  if(dated.records.length){
    const byWeek=new Map();
    for(const row of dated.records){if(!byWeek.has(row.week))byWeek.set(row.week,[]);byWeek.get(row.week).push({name:row.name,hours:row.hours})}
    const weeks=[...byWeek.keys()].sort();
    const result=await api('/api/import-batch',{method:'POST',body:JSON.stringify({sourceFile:file.name,weeks:weeks.map(week=>({week,records:byWeek.get(week)}))})});
    if(weeks.length)selectedWeek=weeks.at(-1);
    if(dated.records.length)selectedDesigner=dated.records.at(-1).name;
    return {imported:result.imported?.length||0,ignored:(result.skipped?.length||0)+dated.ignored.length,weeks:result.weeks||weeks.length,mode:'dated'};
  }

  const wd=parseWeekFilename(file.name);
  if(!wd)throw new Error(`${file.name}: no Date Enquiry Done values were found and no week date was found in the filename.`);
  if(wd.getUTCDay()!==1)throw new Error(`${file.name}: the date in the filename is not a Monday.`);
  const sections=findEstimatorSections(wb);
  if(!sections.length)throw new Error(`${file.name}: no Estimator Name sections with design hours were found.`);
  const grouped=new Map(),ignored=[];
  for(const sec of sections){
    const designer=knownDesigner(sec.name);if(!designer){ignored.push(sec.name);continue}
    const k=norm(designer);grouped.set(k,{name:designer,hours:Math.round(((grouped.get(k)?.hours||0)+sec.hours)*100)/100});
  }
  const records=[...grouped.values()];if(!records.length)throw new Error(`${file.name}: none of the estimator names are in this tracker.`);
  const week=isoDate(wd);
  const result=await api('/api/import',{method:'POST',body:JSON.stringify({week,sourceFile:file.name,records})});
  selectedWeek=week;if(result.imported?.length)selectedDesigner=result.imported[0].name;
  return {imported:result.imported?.length||0,ignored:(result.skipped?.length||0)+ignored.length,weeks:1,mode:'filename'};
}
async function importMany(files){const excel=[...files].filter(isExcelFile);if(!excel.length){showToast('Drop an Excel file here');return}let imported=0,ignored=0,weeks=0,errors=[];for(const f of excel){try{const r=await importEstimatorFile(f);imported+=r.imported;ignored+=r.ignored;weeks+=r.weeks||0}catch(e){errors.push(e.message)}}await loadData(true);const bits=[];if(imported)bits.push(`${imported} designer-week total${imported===1?'':'s'} saved`);if(weeks)bits.push(`${weeks} week${weeks===1?'':'s'} detected`);if(ignored)bits.push(`${ignored} other designer${ignored===1?'':'s'} skipped`);if(errors.length)bits.push(`${errors.length} file${errors.length===1?'':'s'} failed`);showToast(bits.join(' • ')||'No data imported');if(errors.length)alert(errors.join('\n'))}

document.getElementById('weekSelect').addEventListener('change',e=>selectTrackerWeek(e.target.value));
document.getElementById('previousWeekBtn').addEventListener('click',()=>moveTrackerWeek(-1));
document.getElementById('nextWeekBtn').addEventListener('click',()=>moveTrackerWeek(1));
document.getElementById('designerSearch').addEventListener('input',renderDesignerList);
document.getElementById('refreshBtn').addEventListener('click',()=>loadData());
document.getElementById('unlockBtn').addEventListener('click',async()=>{if(editorKey){editorKey='';sessionStorage.removeItem('designHoursEditorKey');setEditingUI();showToast('Editing locked');return}const pin=prompt('Enter the editor PIN');if(!pin)return;const old=editorKey;editorKey=pin;try{await api('/api/auth',{method:'POST'});sessionStorage.setItem('designHoursEditorKey',editorKey);setEditingUI();showToast('Editing unlocked')}catch(e){editorKey=old;alert(e.message)}});
document.getElementById('addDesignerBtn').addEventListener('click',async()=>{const name=prompt('New designer name');if(!name)return;try{await api('/api/designers',{method:'POST',body:JSON.stringify({name})});await loadData(true);selectedDesigner=name.replace(/\s+/g,' ').trim();renderAll();showToast(`${selectedDesigner} added`)}catch(e){alert(e.message)}});
document.getElementById('importFiles').addEventListener('change',e=>{importMany([...e.target.files]);e.target.value='' });

let excelDragDepth=0;
function hasDraggedFiles(e){return Array.from(e.dataTransfer?.types||[]).includes('Files')}
function clearExcelDrag(){excelDragDepth=0;document.body.classList.remove('excel-drag-active')}
document.addEventListener('dragenter',e=>{
  if(!hasDraggedFiles(e))return;
  e.preventDefault();
  excelDragDepth++;
  document.body.classList.add('excel-drag-active');
});
document.addEventListener('dragover',e=>{
  if(!hasDraggedFiles(e))return;
  e.preventDefault();
  if(e.dataTransfer)e.dataTransfer.dropEffect='copy';
  document.body.classList.add('excel-drag-active');
});
document.addEventListener('dragleave',e=>{
  if(!hasDraggedFiles(e))return;
  excelDragDepth=Math.max(0,excelDragDepth-1);
  if(!excelDragDepth)document.body.classList.remove('excel-drag-active');
});
document.addEventListener('drop',e=>{
  if(!hasDraggedFiles(e))return;
  e.preventDefault();
  clearExcelDrag();
  const files=[...(e.dataTransfer?.files||[])];
  if(!editorKey){showToast('Editing is locked');return}
  importMany(files);
});
window.addEventListener('blur',clearExcelDrag);
setEditingUI();loadData(true);setInterval(()=>loadData(true),30000);document.addEventListener('visibilitychange',()=>{if(!document.hidden)loadData(true)});
