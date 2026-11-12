/* Year-on-year comparison. Data stays in the browser. */
(() => {
  'use strict';
  const key = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const iso = d => {
    const x = new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
    const day = x.getUTCDay() || 7;
    x.setUTCDate(x.getUTCDate()+4-day);
    const year = x.getUTCFullYear();
    return {year,week:Math.ceil((((x-new Date(Date.UTC(year,0,1)))/86400000)+1)/7),day};
  };
  const canonical = s => String(s).trim().toLowerCase().replace(/\s+/g,' ');
  const stamp = (d,mode) => mode === 'week' ? `${iso(d).week}-${iso(d).day}` : `${d.getMonth()+1}-${d.getDate()}`;
  const yearOf = (d,mode) => mode === 'week' ? iso(d).year : d.getFullYear();
  const change = (a,b) => a === 0 ? (b === 0 ? 'No change' : 'New / no baseline') : `${b-a>=0?'+':''}${((b-a)/a*100).toFixed(1)}%`;
  function compare(a,b,mode,today) {
    const [old,newer] = a.dates[0] < b.dates[0] ? [a,b] : [b,a];
    const offset=yearOf(newer.dates[0],mode)-yearOf(old.dates[0],mode);
    if (offset<1) throw new Error('Choose files from different years with comparable starting periods.');
    for(const model of [old,newer]) {
      if(new Set(model.dates.map(d=>stamp(d,mode))).size!==model.dates.length) throw new Error('A file contains more than one occurrence of the same comparison day. Export one calendar or school year per file.');
    }
    if (canonical(old.school)!==canonical(newer.school)) throw new Error('Choose the same school in both files before comparing.');
    const previous = new Map(old.dates.map(d=>[stamp(d,mode),d]));
    const pairs = newer.dates.filter(d=>!today || key(d)<=today).flatMap(d=>{
      const prior=previous.get(stamp(d,mode));
      return prior && yearOf(d,mode)-yearOf(prior,mode)===offset ? [{old:prior,newer:d}] : [];
    });
    if (!pairs.length) throw new Error('No matching dates found. Try the other comparison view or check the date coverage of both files.');
    const sum = (model,dates,family) => dates.reduce((n,d)=>n+(family ? (Array.isArray(family)?family:[family]).reduce((v,f)=>v+(model.familyDaily[f]?.[key(d)] || 0),0) : model.dailyTotal[key(d)]),0);
    const metric = (label,ps,family) => {
      const before=sum(old,ps.map(p=>p.old),family), after=sum(newer,ps.map(p=>p.newer),family);
      return {label,days:ps.length,before,after,oldAverage:before/ps.length,newAverage:after/ps.length,delta:after-before};
    };
    const groups=new Map();
    for(const p of pairs) {
      const g=mode==='week' ? `Week ${String(iso(p.newer).week).padStart(2,'0')}` : p.newer.toLocaleDateString('en-GB',{month:'long'});
      if(!groups.has(g))groups.set(g,[]);groups.get(g).push(p);
    }
    const families=[...new Set([...old.families,...newer.families])];
    const scopes=[{label:'Rise then Shine (breakfast)',match:s=>/^(risethenshine|breakfast|breakfastclub)$/.test(s)},{label:'Stay and Play',match:s=>s==='stayandplay'},{label:'Whole day',match:null}].map(scope=>{
      const fs=scope.match?families.filter(f=>scope.match(canonical(f).replace(/[^a-z]/g,''))):null;
      return {label:scope.label,missing:fs ? !old.families.some(f=>fs.includes(f)) || !newer.families.some(f=>fs.includes(f)) : false,total:metric(scope.label,pairs,fs),periods:[...groups].map(([g,ps])=>metric(g,ps,fs))};
    });
    return {old,newer,pairs,scopes,total:metric('Whole day',pairs),periods:[...groups].map(([g,ps])=>metric(g,ps)),programmes:families.map(f=>metric(f,pairs,f)),excludedOld:old.dates.length-pairs.length,excludedNew:newer.dates.length-pairs.length};
  }
  window.AttendanceTrends = {compare,iso,change};
  const $=id=>document.getElementById(id);
  if (!$('trendTool')) return;
  const escape=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const files=[null,null]; const tickets=[0,0]; let current=null;
  const exportButtons=()=>['trendDownload','trendExcel','trendPdf'].map($).filter(Boolean);
  const invalidate=()=>{current=null;$('trendResults').innerHTML='';exportButtons().forEach(b=>b.disabled=true);};
  for(const tab of ['report','trend']) $(tab+'Tab').addEventListener('click',()=>{
    for(const name of ['report','trend']) {$(name+'Tool').classList.toggle('hidden',name!==tab);$(name+'Tab').setAttribute('aria-pressed',String(name===tab));}
  });
  [0,1].forEach(i=>{
    $('trendFile'+i).addEventListener('change',async e=>{
      const ticket=++tickets[i];files[i]=null;invalidate();$('trendError').textContent='';$('trendSchool'+i).replaceChildren();
      const file=e.target.files[0];if(!file)return;
      $('trendStatus'+i).textContent='Reading file…';
      try {
        if(!/\.xlsx?$/i.test(file.name))throw new Error('Choose an Excel .xlsx or .xls file.');
        const workbook=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false,raw:true});
        const parsed=window.AttendanceTools.parseAttendanceWorkbook(workbook);
        if(ticket!==tickets[i])return;
        files[i]=parsed.records;
        const schools=[...new Set(parsed.records.map(r=>r.centre))].sort();
        $('trendSchool'+i).replaceChildren(...schools.map(s=>new Option(s,s)));
        $('trendStatus'+i).textContent=`${file.name} · ${schools.length} school${schools.length===1?'':'s'} detected`;
      } catch(err) {if(ticket!==tickets[i])return;$('trendStatus'+i).textContent='Could not read file';$('trendError').textContent=err.message;}
    });
    $('trendSchool'+i).addEventListener('change',invalidate);
  });
  ['trendMode','trendPast'].forEach(id=>$(id).addEventListener('change',invalidate));
  $('trendCompare').addEventListener('click',()=>{
    invalidate();$('trendError').textContent='';
    try {
      if(files.some(f=>!f))throw new Error('Upload both files first.');
      const models=files.map((rows,i)=>window.AttendanceTools.buildReportModel(rows.filter(r=>r.centre===$('trendSchool'+i).value)));
      const mode=$('trendMode').value;
      const c=compare(...models,mode,$('trendPast').checked?key(new Date()):null);
      const year=m=>{const first=yearOf(m.dates[0],mode),last=yearOf(m.dates[m.dates.length-1],mode);return first===last?first:`${first}–${last}`;};
      const range=ds=>`${ds[0].toLocaleDateString('en-GB')} – ${ds[ds.length-1].toLocaleDateString('en-GB')}`;
      const t=c.total;
      const insight=[];
      c.scopes.forEach(s=>{const r=s.total;insight.push(`${s.label}: ${r.before.toLocaleString()} to ${r.after.toLocaleString()} bookings (${change(r.before,r.after)}); average per matched day ${r.oldAverage.toFixed(1)} to ${r.newAverage.toFixed(1)}.${s.missing?' This programme is absent in at least one file; check coverage or naming before interpreting zero counts.':''}`);});
      insight.push(`Bookings ${t.delta===0?'were unchanged':t.delta>0?'increased':'decreased'} from ${t.before.toLocaleString()} to ${t.after.toLocaleString()} (${change(t.before,t.after)}) across ${t.days} matched dates in each year.`);
      insight.push(`Average bookings per matched day changed from ${t.oldAverage.toFixed(1)} to ${t.newAverage.toFixed(1)}.`);
      const ranked=[...c.programmes].sort((a,b)=>b.delta-a.delta);
      if(ranked[0]?.delta>0)insight.push(`${ranked[0].label} contributed the largest increase: ${ranked[0].delta.toLocaleString()} additional bookings.`);
      const decline=ranked[ranked.length-1];if(decline?.delta<0)insight.push(`${decline.label} contributed the largest decrease: ${Math.abs(decline.delta).toLocaleString()} fewer bookings.`);
      const peak=[...c.periods].sort((a,b)=>b.delta-a.delta)[0];if(peak && c.periods.length>1)insight.push(`${peak.label} had the strongest change in booking numbers (${peak.delta>=0?'+':''}${peak.delta}; ${change(peak.before,peak.after)}).`);
      if(c.programmes.some(p=>!c.old.families.includes(p.label)||!c.newer.families.includes(p.label)))insight.push('Some programmes appear in only one file. Their missing-year counts are shown as zero; check whether the programme changed name or was introduced or removed.');
      const table=(title,rows)=>`<h3>${title}</h3><div class="preview-table-wrap"><table class="preview-table"><thead><tr><th>Period / programme</th><th>Matched days / year</th><th>${year(c.old)} bookings</th><th>${year(c.newer)} bookings</th><th>Change</th><th>Change %</th><th>${year(c.old)} avg/day</th><th>${year(c.newer)} avg/day</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${escape(r.label)}</td><td>${r.days}</td><td>${r.before}</td><td>${r.after}</td><td>${r.delta>0?'+':''}${r.delta}</td><td>${change(r.before,r.after)}</td><td>${r.oldAverage.toFixed(1)}</td><td>${r.newAverage.toFixed(1)}</td></tr>`).join('')}</tbody></table></div>`;
      const max=Math.max(...c.periods.flatMap(p=>[p.before,p.after]),1);
      $('trendResults').innerHTML=`<h2>${escape(c.old.school)} · ${year(c.old)} to ${year(c.newer)}</h2><p><strong>Earlier year:</strong> ${range(c.pairs.map(p=>p.old))}<br><strong>Later year:</strong> ${range(c.pairs.map(p=>p.newer))}</p><p class="data-notice">${t.days} matched dates per year. Excluded: ${c.excludedOld} earlier-year dates and ${c.excludedNew} later-year dates. ${mode==='week'?'Matched by ISO week number and weekday (Monday–Sunday).':'Matched by calendar month and day. Weekday mix can differ between years.'} Gaps and unmatched dates are excluded. Blank cells in summary exports count as zero. Partial weeks/months contain only matching dates.</p><div class="detected-grid"><div class="data-card"><span class="data-label">Booking change</span><strong>${change(t.before,t.after)}</strong></div><div class="data-card"><span class="data-label">Additional bookings</span><strong>${t.delta>0?'+':''}${t.delta}</strong></div><div class="data-card"><span class="data-label">Earlier avg/day</span><strong>${t.oldAverage.toFixed(1)}</strong></div><div class="data-card"><span class="data-label">Later avg/day</span><strong>${t.newAverage.toFixed(1)}</strong></div></div><h3 style="margin-top:24px">What changed</h3><ul>${insight.map(s=>`<li>${escape(s)}</li>`).join('')}</ul><p>These are booking trends; the export cannot establish the reasons for a change or confirm actual attendance.</p><h3>Booking trend</h3><p>🟣 ${year(c.old)} &nbsp; 🟠 ${year(c.newer)} · total bookings on matched dates</p>${c.periods.map(p=>`<div style="margin:14px 0"><strong>${escape(p.label)}</strong><div style="width:${p.before/max*100}%;min-width:2px;background:#54208a;height:14px" title="${year(c.old)}: ${p.before}"></div><div style="width:${p.after/max*100}%;min-width:2px;background:#ff8a00;height:14px;margin-top:3px" title="${year(c.newer)}: ${p.after}"></div><small>${p.before} → ${p.after} bookings</small></div>`).join('')}${table(mode==='week'?'Week-by-week comparison':'Month-by-month comparison',c.periods)}${table('Programme comparison',c.programmes)}<details class="preview-details"><summary>Check every matched date</summary><div class="preview-table-wrap"><table class="preview-table"><thead><tr><th>Earlier date</th><th>Later date</th><th>Earlier bookings</th><th>Later bookings</th></tr></thead><tbody>${c.pairs.map(p=>`<tr><td>${p.old.toLocaleDateString('en-GB',{weekday:'short',year:'numeric',month:'short',day:'numeric'})}</td><td>${p.newer.toLocaleDateString('en-GB',{weekday:'short',year:'numeric',month:'short',day:'numeric'})}</td><td>${c.old.dailyTotal[key(p.old)]}</td><td>${c.newer.dailyTotal[key(p.newer)]}</td></tr>`).join('')}</tbody></table></div></details>`;
      $('trendResults').innerHTML+=c.scopes.map(s=>table(`${escape(s.label)} — ${mode==='week'?'weekly':'monthly'} comparison`,s.periods)).join('');
      current={c,insight,mode,futureExcluded:$('trendPast').checked,cutoff:key(new Date())};exportButtons().forEach(b=>b.disabled=false);
    } catch(err) {$('trendError').textContent=err.message;}
  });
  for(const [id,type] of [['trendExcel','excel'],['trendPdf','pdf']]) {
    $(id)?.addEventListener('click',async()=>{
      if(!current)return;
      const snapshot=current;exportButtons().forEach(b=>b.disabled=true);$('trendError').textContent='';
      try {await window.TrendExports[type](snapshot);} catch(err) {$('trendError').textContent=`Export failed: ${err.message}`;}
      finally {exportButtons().forEach(b=>b.disabled=!current);}
    });
  }
  $('trendDownload').addEventListener('click',()=>{
    if(!current)return;
    const {c,insight,mode}=current;
    const rows=[['School',c.old.school],['Comparison',mode],['Earlier year',yearOf(c.old.dates[0],mode)],['Later year',yearOf(c.newer.dates[0],mode)],['Matched dates per year',c.pairs.length],['Excluded earlier dates',c.excludedOld],['Excluded later dates',c.excludedNew],['Method',mode==='week'?'ISO week and weekday':'Calendar month and day'],['Notes','Only matched dates; summary blanks = zero; bookings do not confirm attendance.'],[],['Group','Matched days','Earlier bookings','Later bookings','Change','Change %','Earlier avg/day','Later avg/day'],...[c.total,...c.periods,...c.programmes].map(r=>[r.label,r.days,r.before,r.after,r.delta,change(r.before,r.after),r.oldAverage.toFixed(2),r.newAverage.toFixed(2)]),[],['Insights'],...insight.map(s=>[s]),[],['Earlier date','Later date','Earlier bookings','Later bookings'],...c.pairs.map(p=>[key(p.old),key(p.newer),c.old.dailyTotal[key(p.old)],c.newer.dailyTotal[key(p.newer)]])];
    const csv=rows.map(row=>row.map(v=>'"'+String(v).replace(/^[=+@-]/,"'"+'$&').replace(/"/g,'""')+'"').join(',')).join('\r\n');
    const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='Year_on_year_booking_trends.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  });
})();
