/* Branded, browser-local year-on-year exports. */
(() => {
  'use strict';
  const purple='54208A', pale='F1EAF7', yellow='FFD21A';
  const key=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const range=ds=>`${ds[0].toLocaleDateString('en-GB')} to ${ds[ds.length-1].toLocaleDateString('en-GB')}`;
  const headers=['Period','Matched days / year','Earlier bookings','Later bookings','Change','Change %','Earlier avg/day','Later avg/day'];
  const row=r=>[r.label,r.days,r.before,r.after,r.delta,r.before?(r.after-r.before)/r.before:r.after?'No baseline':0,r.oldAverage,r.newAverage];
  function meta(s) {
    const {c,mode}=s;
    return [['School',c.old.school],['Earlier file coverage',range(c.old.dates)],['Later file coverage',range(c.newer.dates)],['Earlier dates compared',range(c.pairs.map(p=>p.old))],['Later dates compared',range(c.pairs.map(p=>p.newer))],['Matched dates per year',c.pairs.length],['Excluded earlier / later dates',`${c.excludedOld} / ${c.excludedNew}`],['Date matching',mode==='week'?'Same ISO week number and weekday (Monday–Sunday)':'Same calendar month and day; weekday mix can differ'],['Future dates',s.futureExcluded?`Later-year dates after ${s.cutoff} excluded`:'Included if represented in both files']];
  }
  const notes='Only matched dates contribute. Gaps and unmatched dates are excluded; blank summary cells count as zero. Partial periods use only matching dates. Whole day is the sum of all booking options, not a distinct-child count. The same child may have breakfast and after-school bookings. Bookings do not confirm attendance or explain the causes of change.';
  function filename(c,ext){return `${c.old.school.replace(/[^a-z0-9]+/gi,'_')}_Year_on_Year_Trends.${ext}`;}
  function download(blob,name){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);}
  function workbook(s) {
    const {c,insight}=s;const wb=XLSX.utils.book_new();
    function sheet(name,rows,widths,headerRow,kind='metrics') {
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws['!cols']=widths.map(w=>({wch:w}));ws['!rows']=rows.map((r,i)=>({hpt:i===0?30:Math.max(24,Math.ceil(String(r[1]??r[0]??'').length/90)*16)}));
      for(const addr of Object.keys(ws)) {
        if(addr.startsWith('!'))continue;
        const cell=ws[addr],pos=XLSX.utils.decode_cell(addr);
        cell.s={font:{name:'Montserrat',sz:11,color:{rgb:'2C2631'}},alignment:{vertical:'center',wrapText:true},fill:{fgColor:{rgb:pos.r%2?pale:'FFFFFF'}}};
        if(pos.r===0||pos.r===headerRow)cell.s={...cell.s,font:{name:'Montserrat',sz:pos.r===0?16:10,bold:true,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:purple}}};
        if(headerRow!=null&&pos.r>headerRow&&typeof cell.v==='number')cell.z=kind==='dates'?'#,##0':pos.c===5?'0.0%;-0.0%;0.0%':pos.c>=6?'0.0':'#,##0;[Red]-#,##0';
      }
      if(headerRow!=null)ws['!autofilter']={ref:XLSX.utils.encode_range({r:headerRow,c:0},{r:rows.length-1,c:rows[headerRow].length-1})};
      ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:widths.length-1}}];
      XLSX.utils.book_append_sheet(wb,ws,name);
    }
    sheet('Comparison summary',[['JAG | Year-on-year booking trends'],...meta(s),[],['How to read this report',notes],[],['Key insights'],...insight.map((t,i)=>[`Insight ${i+1}`,t])],[30,110]);
    sheet('Programme overview',[['Like-for-like programme comparison'],headers,...c.scopes.map(g=>row(g.total))],[32,18,20,20,16,22,20,20],1);
    c.scopes.forEach((g,i)=>sheet(['Breakfast trend','Stay and Play trend','Whole day trend'][i],[[g.label+' | '+(s.mode==='week'?'weekly':'monthly')],headers,...g.periods.map(row)],[30,18,20,20,16,22,20,20],1));
    const dailyHeaders=['Earlier date','Later date',...c.scopes.flatMap(g=>[g.label+' earlier',g.label+' later'])];
    const count=(m,d,i)=>i===2?m.dailyTotal[key(d)]:m.families.filter(f=>i===0?/^(risethenshine|breakfast|breakfastclub)$/.test(f.toLowerCase().replace(/[^a-z]/g,'')):f.toLowerCase().replace(/[^a-z]/g,'')==='stayandplay').reduce((n,f)=>n+(m.familyDaily[f]?.[key(d)]||0),0);
    sheet('Matched dates',[['Audit of dates and booking counts'],dailyHeaders,...c.pairs.map(p=>[key(p.old),key(p.newer),...c.scopes.flatMap((g,i)=>[count(c.old,p.old,i),count(c.newer,p.newer,i)])])],[18,18,24,24,24,24,24,24],1,'dates');
    return wb;
  }
  async function pdf(s) {
    const {c,insight}=s;const {jsPDF}=window.jspdf;
    const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4',compress:true});
    const logo=await new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>reject(new Error('Could not load the JAG logo. Please retry.'));img.src='assets/jag-logo.png';});
    let y=0,page=0;
    function brand(title,landscape=false,first=false){
      if(!first)doc.addPage('a4',landscape?'landscape':'portrait');page++;
      const width=landscape?297:210,height=landscape?210:297;
      doc.setFillColor('#'+purple);doc.rect(0,0,width,9,'F');doc.setFillColor('#'+yellow);doc.rect(0,9,width,2,'F');
      doc.addImage(logo,'PNG',15,17,19,19);doc.setTextColor('#'+purple);doc.setFont('helvetica','bold');doc.setFontSize(16);doc.text('Junior Adventures Group',39,24);
      doc.setFontSize(9);doc.setFont('helvetica','normal');doc.text(c.old.school,39,31,{maxWidth:width-55});
      doc.setFontSize(21);doc.setFont('helvetica','bold');doc.text(title,15,48);y=58;
      doc.setTextColor('#6D6672');doc.setFontSize(8);doc.text('Year-on-year booking trends',15,height-9);doc.text(`Page ${page}`,width-28,height-9);
    }
    function paragraph(text,size=10,bold=false,width=180){
      doc.setFont('helvetica',bold?'bold':'normal');doc.setFontSize(size);doc.setTextColor(bold?'#'+purple:'#2C2631');
      const lines=doc.splitTextToSize(String(text),width);const height=lines.length*size*.40+3;
      if(y+height>278)brand('Comparison notes and insights');
      doc.text(lines,15,y);y+=height;
    }
    brand('Year-on-year comparison',false,true);
    paragraph('Periods, matching method and key findings',12,true);
    for(const [label,value] of meta(s))paragraph(`${label}: ${value}`,9);
    paragraph('Like-for-like findings',12,true);
    // Three programme findings lead the front page; supporting insights follow.
    insight.slice(0,3).forEach(t=>paragraph(t,10));
    paragraph('Whole day combines booking entries across all programmes; it is not a count of distinct children.',9);
    brand('Insights and interpretation');
    insight.slice(3).forEach(t=>paragraph(t,11));
    paragraph('Comparison rules',12,true);paragraph(notes,10);
    if(c.scopes.some(g=>g.missing))paragraph('At least one programme is absent in a file. Zero bookings can reflect missing coverage or a programme name change; check the source before drawing conclusions.',10,true);
    function table(title,metrics) {
      let first=true;
      const widths=[49,23,30,30,25,34,34,42];
      function start(){brand(title,true);const hs=['Period','Matched\ndays/year','Earlier\nbookings','Later\nbookings','Change','Change %','Earlier\navg/day','Later\navg/day'];let x=15;doc.setFont('helvetica','bold');doc.setFontSize(9);hs.forEach((h,i)=>{doc.setFillColor('#'+purple);doc.rect(x,y,widths[i],14,'F');doc.setTextColor('#FFFFFF');doc.text(h,x+2,y+5);x+=widths[i];});y+=14;}
      start();
      metrics.forEach((r,i)=>{if(y+14>194)start();let x=15;const values=[r.label,r.days,r.before.toLocaleString(),r.after.toLocaleString(),`${r.delta>0?'+':''}${r.delta}`,window.AttendanceTrends.change(r.before,r.after),r.oldAverage.toFixed(1),r.newAverage.toFixed(1)];
        values.forEach((v,j)=>{doc.setFillColor(i%2?'#'+pale:'#FFFFFF');doc.rect(x,y,widths[j],12,'F');doc.setFont('helvetica',j===0?'bold':'normal');doc.setFontSize(9);doc.setTextColor('#2C2631');doc.text(doc.splitTextToSize(String(v),widths[j]-4),x+2,y+5);x+=widths[j];});y+=12;});
    }
    // Landscape content width is 267mm, leaving 15mm margins.
    table('Programme comparison',c.scopes.map(g=>g.total));
    for(const g of c.scopes)table(`${g.label} - ${s.mode==='week'?'weekly':'monthly'}`,g.periods);
    return doc.output('blob');
  }
  const excelData=s=>XLSX.write(workbook(s),{bookType:'xlsx',type:'array',cellStyles:true});
  async function pack(s) {
    const zip=new JSZip();
    zip.file(filename(s.c,'xlsx'),excelData(s));
    zip.file(filename(s.c,'pdf'),await (await pdf(s)).arrayBuffer());
    const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
    download(blob,filename(s.c,'zip'));
  }
  window.TrendExports={workbook,buildPdf:pdf,pack,excel:s=>download(new Blob([excelData(s)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),filename(s.c,'xlsx')),pdf:async s=>download(await pdf(s),filename(s.c,'pdf'))};
})();
