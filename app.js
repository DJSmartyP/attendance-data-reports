/*
 * Junior Adventures Group Attendance Report Generator
 * Static, client-side only. Designed for GitHub Pages.
 */

(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const els = {
    fileInput: $('fileInput'), dropZone: $('dropZone'), errorBox: $('errorBox'), successBox: $('successBox'),
    detectedPanel: $('detectedPanel'), generatePanel: $('generatePanel'), resetBtn: $('resetBtn'),
    schoolName: $('schoolName'), periodValue: $('periodValue'), daysValue: $('daysValue'), entriesValue: $('entriesValue'),
    familyChips: $('familyChips'), sessionChips: $('sessionChips'), previewTableWrap: $('previewTableWrap'),
    excelBtn: $('excelBtn'), pdfBtn: $('pdfBtn'), packBtn: $('packBtn'),
    progressWrap: $('progressWrap'), progressBar: $('progressBar'), progressText: $('progressText'),
    pdfStage: $('pdfStage')
  };

  let report = null;
  let sourceFileName = '';

  // ---------- Upload handling ----------
  els.fileInput.addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (file) loadFile(file);
  });

  ['dragenter', 'dragover'].forEach(type => els.dropZone.addEventListener(type, (e) => {
    e.preventDefault();
    els.dropZone.classList.add('dragover');
  }));
  ['dragleave', 'drop'].forEach(type => els.dropZone.addEventListener(type, (e) => {
    e.preventDefault();
    els.dropZone.classList.remove('dragover');
  }));
  els.dropZone.addEventListener('drop', (e) => {
    const file = e.dataTransfer?.files?.[0];
    if (file) loadFile(file);
  });

  els.resetBtn.addEventListener('click', resetApp);
  els.excelBtn.addEventListener('click', async () => {
    try {
      setBusy(true, 'Creating Excel workbook…', 20);
      const blob = buildExcelBlob();
      downloadBlob(blob, `${baseFileName()}.xlsx`);
      setBusy(false);
    } catch (err) { handleGenerationError(err); }
  });
  els.pdfBtn.addEventListener('click', async () => {
    try {
      setBusy(true, 'Building branded PDF pages…', 5);
      const blob = await buildPdfBlob();
      downloadBlob(blob, `${baseFileName()}.pdf`);
      setBusy(false);
    } catch (err) { handleGenerationError(err); }
  });
  els.packBtn.addEventListener('click', async () => {
    try {
      setBusy(true, 'Creating Excel workbook…', 5);
      const xlsxBlob = buildExcelBlob();
      setProgress(30, 'Building branded PDF pages…');
      const pdfBlob = await buildPdfBlob(30, 88);
      setProgress(92, 'Packaging files…');
      const zip = new JSZip();
      const base = baseFileName();
      zip.file(`${base}.xlsx`, xlsxBlob);
      zip.file(`${base}.pdf`, pdfBlob);
      zip.file('README.txt', `Junior Adventures Group attendance report pack\nSchool: ${currentSchool()}\nPeriod: ${formatLongDate(report.minDate)} to ${formatLongDate(report.maxDate)}\nSource: ${sourceFileName}\n\nGenerated locally in the browser.`);
      const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } }, (meta) => {
        setProgress(92 + Math.round(meta.percent * .08), 'Packaging files…');
      });
      downloadBlob(zipBlob, `${base}_pack.zip`);
      setBusy(false);
    } catch (err) { handleGenerationError(err); }
  });

  async function loadFile(file) {
    clearMessages();
    try {
      if (!/\.xlsx?$/i.test(file.name)) throw new Error('Please choose an Excel .xlsx or .xls file.');
      showMessage('success', `Reading ${file.name}…`);
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
      const parsed = parseAttendanceWorkbook(workbook);
      report = buildReportModel(parsed.records);
      sourceFileName = file.name;
      renderDetected();
      showMessage('success', `Ready: ${report.records.length.toLocaleString()} attendance rows detected from ${file.name}.`);
    } catch (err) {
      report = null;
      els.detectedPanel.classList.add('hidden');
      els.generatePanel.classList.add('hidden');
      showMessage('error', err.message || 'Could not read this attendance export.');
      console.error(err);
    }
  }

  function parseAttendanceWorkbook(workbook) {
    let best = null;
    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      for (let i = 0; i < Math.min(rows.length, 60); i++) {
        const row = rows[i] || [];
        const normalized = row.map(normalizeHeader);
        const sessionIdx = findHeaderIndex(normalized, ['sessions', 'session']);
        const dateIdx = findHeaderIndex(normalized, ['bookeddate', 'bookingdate', 'datebooked', 'attendancedate', 'date']);
        const centreIdx = findHeaderIndex(normalized, ['centre', 'center', 'school', 'site']);
        if (sessionIdx >= 0 && dateIdx >= 0 && centreIdx >= 0) {
          best = { sheetName, rows, headerRowIndex: i, header: row };
          break;
        }
      }
      if (best) break;
    }
    if (!best) {
      throw new Error('I could not find the expected attendance columns. The export needs a session column, a booked/attendance date column, and a Centre/School/Site column.');
    }

    const headerKeys = best.header.map(normalizeHeader);
    const idx = {
      session: findHeaderIndex(headerKeys, ['sessions', 'session']),
      label: findHeaderIndex(headerKeys, ['label', 'sessionlabel']),
      bookingId: findHeaderIndex(headerKeys, ['bookingid', 'booking_id', 'id']),
      date: findHeaderIndex(headerKeys, ['bookeddate', 'bookingdate', 'datebooked', 'attendancedate', 'date']),
      centre: findHeaderIndex(headerKeys, ['centre', 'center', 'school', 'site']),
      dayName: findHeaderIndex(headerKeys, ['dayname', 'day'])
    };

    const records = [];
    for (let r = best.headerRowIndex + 1; r < best.rows.length; r++) {
      const row = best.rows[r] || [];
      const session = cleanText(row[idx.session]);
      const date = parseExcelDate(row[idx.date]);
      if (!session || !date) continue;
      records.push({
        session,
        label: idx.label >= 0 ? cleanText(row[idx.label]) : '',
        bookingId: idx.bookingId >= 0 ? cleanText(row[idx.bookingId]) : `row-${r + 1}`,
        date,
        centre: cleanText(row[idx.centre]) || 'School',
        dayName: idx.dayName >= 0 ? cleanText(row[idx.dayName]) : ''
      });
    }
    if (!records.length) throw new Error('The workbook contains the expected headers, but I could not find any attendance rows beneath them.');
    return { ...best, records };
  }

  function buildReportModel(records) {
    const centreCounts = countBy(records.map(r => r.centre).filter(Boolean));
    const school = mostCommon(centreCounts) || 'School';
    const minDate = new Date(Math.min(...records.map(r => r.date.getTime())));
    const maxDate = new Date(Math.max(...records.map(r => r.date.getTime())));

    const sessionSourceOrder = [];
    const labelsBySession = new Map();
    for (const r of records) {
      if (!sessionSourceOrder.includes(r.session)) sessionSourceOrder.push(r.session);
      if (!labelsBySession.has(r.session)) labelsBySession.set(r.session, []);
      if (r.label) labelsBySession.get(r.session).push(r.label);
    }

    const sessionMeta = {};
    for (const session of sessionSourceOrder) {
      const label = mostCommon(countBy(labelsBySession.get(session) || [])) || '';
      sessionMeta[session] = {
        family: familyName(session),
        label,
        display: displaySessionName(session, label)
      };
    }

    const families = [];
    for (const s of sessionSourceOrder) {
      const fam = sessionMeta[s].family;
      if (!families.includes(fam)) families.push(fam);
    }

    const sessions = [];
    for (const family of families) {
      const famSessions = sessionSourceOrder.filter(s => sessionMeta[s].family === family);
      famSessions.sort((a, b) => {
        if (family.toLowerCase() === 'stay and play') {
          const aFull = /3\s*[-–]\s*6\b/.test(sessionMeta[a].display) ? -100 : 0;
          const bFull = /3\s*[-–]\s*6\b/.test(sessionMeta[b].display) ? -100 : 0;
          if (aFull !== bFull) return aFull - bFull;
        }
        return sessionSortValue(sessionMeta[a]) - sessionSortValue(sessionMeta[b]);
      });
      sessions.push(...famSessions);
    }

    const familySessions = Object.fromEntries(families.map(f => [f, sessions.filter(s => sessionMeta[s].family === f)]));
    const subtotalFamilies = families.filter(f => familySessions[f].length > 1);

    const dateSessionIds = new Map();
    for (let i = 0; i < records.length; i++) {
      const r = records[i];
      const dk = dateKey(r.date);
      if (!dateSessionIds.has(dk)) dateSessionIds.set(dk, new Map());
      const sm = dateSessionIds.get(dk);
      if (!sm.has(r.session)) sm.set(r.session, new Set());
      sm.get(r.session).add(r.bookingId || `row-${i}`);
    }

    const dates = Array.from(dateSessionIds.keys()).sort().map(parseDateKey);
    const counts = {};
    const dailyTotal = {};
    const familyDaily = Object.fromEntries(families.map(f => [f, {}]));

    for (const d of dates) {
      const dk = dateKey(d);
      counts[dk] = {};
      for (const s of sessions) counts[dk][s] = dateSessionIds.get(dk).get(s)?.size || 0;
      dailyTotal[dk] = sessions.reduce((sum, s) => sum + counts[dk][s], 0);
      for (const fam of families) {
        familyDaily[fam][dk] = familySessions[fam].reduce((sum, s) => sum + counts[dk][s], 0);
      }
    }

    const weeksMap = new Map();
    for (const d of dates) {
      const monday = startOfWeek(d);
      const wk = dateKey(monday);
      if (!weeksMap.has(wk)) weeksMap.set(wk, []);
      weeksMap.get(wk).push(d);
    }
    const weeks = Array.from(weeksMap.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([mondayKey, weekDates]) => ({ monday: parseDateKey(mondayKey), dates: weekDates.sort((a,b) => a-b) }));

    const familyTotals = Object.fromEntries(families.map(f => [f, dates.reduce((sum, d) => sum + familyDaily[f][dateKey(d)], 0)]));
    const sessionTotals = Object.fromEntries(sessions.map(s => [s, dates.reduce((sum, d) => sum + counts[dateKey(d)][s], 0)]));
    const totalAttendance = dates.reduce((sum, d) => sum + dailyTotal[dateKey(d)], 0);
    const busiestDay = dates.reduce((best, d) => dailyTotal[dateKey(d)] > dailyTotal[dateKey(best)] ? d : best, dates[0]);

    const monthlyMap = new Map();
    for (const d of dates) {
      const k = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      if (!monthlyMap.has(k)) monthlyMap.set(k, { date: new Date(d.getFullYear(), d.getMonth(), 1), days: 0, total: 0, families: Object.fromEntries(families.map(f => [f, 0])) });
      const m = monthlyMap.get(k);
      m.days += 1;
      m.total += dailyTotal[dateKey(d)];
      for (const f of families) m.families[f] += familyDaily[f][dateKey(d)];
    }
    const monthly = Array.from(monthlyMap.values()).sort((a,b) => a.date-b.date);

    const weekday = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'].map(name => {
      const ds = dates.filter(d => d.toLocaleDateString('en-GB', { weekday: 'long' }) === name);
      if (!ds.length) return null;
      const total = ds.reduce((sum,d) => sum + dailyTotal[dateKey(d)], 0);
      return { name, days: ds.length, total, average: total / ds.length };
    }).filter(Boolean);

    return {
      records, school, minDate, maxDate, sessions, sessionMeta, families, familySessions, subtotalFamilies,
      dates, counts, dailyTotal, familyDaily, weeks, familyTotals, sessionTotals, totalAttendance,
      operatingDays: dates.length, averageDaily: totalAttendance / dates.length, busiestDay, monthly, weekday
    };
  }

  // ---------- Detected data UI ----------
  function renderDetected() {
    els.schoolName.value = report.school;
    els.periodValue.textContent = `${formatLongDate(report.minDate)} – ${formatLongDate(report.maxDate)}`;
    els.daysValue.textContent = report.operatingDays.toLocaleString();
    els.entriesValue.textContent = report.totalAttendance.toLocaleString();
    els.familyChips.innerHTML = report.families.map(f => `<span class="chip">${escapeHtml(f)}</span>`).join('');
    els.sessionChips.innerHTML = report.sessions.map(s => `<span class="chip">${escapeHtml(report.sessionMeta[s].display)}</span>`).join('');
    renderPreview();
    els.detectedPanel.classList.remove('hidden');
    els.generatePanel.classList.remove('hidden');
    els.detectedPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderPreview() {
    const firstWeeks = report.weeks.slice(0, 2);
    const headers = ['Date', 'Day', ...report.sessions.map(s => report.sessionMeta[s].display), 'Day total', ...report.subtotalFamilies.map(f => `${f} Only`)];
    let html = `<table class="preview-table"><thead><tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>`;
    for (const week of firstWeeks) {
      for (const d of week.dates) {
        const dk = dateKey(d);
        html += '<tr>';
        html += `<td>${escapeHtml(formatShortDate(d))}</td><td>${escapeHtml(shortWeekday(d))}</td>`;
        for (const s of report.sessions) {
          const fam = report.sessionMeta[s].family;
          html += `<td>${report.familyDaily[fam][dk] === 0 ? 'NA' : report.counts[dk][s]}</td>`;
        }
        html += `<td><strong>${report.dailyTotal[dk]}</strong></td>`;
        for (const f of report.subtotalFamilies) html += `<td>${report.familyDaily[f][dk]}</td>`;
        html += '</tr>';
      }
    }
    html += '</tbody></table>';
    if (report.weeks.length > 2) html += `<p style="font-size:.75rem;color:#6d6672">Showing the first two weeks. The generated reports include all ${report.weeks.length} weeks.</p>`;
    els.previewTableWrap.innerHTML = html;
  }

  // ---------- Excel generator ----------
  function buildExcelBlob() {
    ensureReport();
    const rows = [];
    const merges = [];
    const rowMeta = [];
    const headers = ['W/C', 'Day of Week', ...report.sessions.map(s => report.sessionMeta[s].display), 'Day total:', ...report.subtotalFamilies.map(f => `${f} Only`)];
    const nCols = headers.length;

    rows.push([`${currentSchool()} Attendance - ${formatOrdinalDate(report.minDate, true)} - ${formatOrdinalDate(report.maxDate, true)}`]);
    rowMeta.push({ type: 'title' });
    rows.push([]); rowMeta.push({ type: 'blank' });
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: nCols - 1 } });

    for (const week of report.weeks) {
      rows.push(headers); rowMeta.push({ type: 'header' });
      for (const d of week.dates) {
        const dk = dateKey(d);
        const row = [formatOrdinalDate(d, false), shortWeekday(d)];
        for (const s of report.sessions) {
          const fam = report.sessionMeta[s].family;
          row.push(report.familyDaily[fam][dk] === 0 ? 'NA' : report.counts[dk][s]);
        }
        row.push(report.dailyTotal[dk]);
        for (const f of report.subtotalFamilies) row.push(report.familyDaily[f][dk]);
        rows.push(row); rowMeta.push({ type: 'data' });
      }
      rows.push([]); rowMeta.push({ type: 'blank' });
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!merges'] = merges;
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!cols'] = headers.map((_, i) => ({ wch: i === 0 ? 14 : i === 1 ? 13 : i >= 2 + report.sessions.length ? 17 : 20 }));
    ws['!rows'] = rowMeta.map(m => ({ hpt: m.type === 'title' ? 26 : m.type === 'header' ? 30 : m.type === 'blank' ? 8 : 20 }));

    const purple = '54208A', purpleDark = '42146F', pale = 'F1EAF7', line = 'D8CFE0', yellowPale = 'FFF8D9';
    const defaultFont = { name: 'Montserrat', sz: 10, color: { rgb: '2C2631' } };
    const range = XLSX.utils.decode_range(ws['!ref']);

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: 's', v: '' };
        const meta = rowMeta[r];
        ws[addr].s = {
          font: { ...defaultFont },
          alignment: { vertical: 'center', horizontal: c < 2 ? 'left' : 'center', wrapText: true },
          border: meta.type === 'data' ? { bottom: { style: 'hair', color: { rgb: 'E7E0EB' } } } : undefined
        };
        if (meta.type === 'title') {
          ws[addr].s = { font: { name: 'Montserrat', sz: 17, bold: true, color: { rgb: purple } }, alignment: { vertical: 'center', horizontal: 'left' } };
        } else if (meta.type === 'header') {
          ws[addr].s = {
            font: { name: 'Montserrat', sz: 10, bold: true, color: { rgb: purpleDark } },
            fill: { fgColor: { rgb: pale } },
            alignment: { vertical: 'center', horizontal: 'center', wrapText: true },
            border: { bottom: { style: 'thin', color: { rgb: line } } }
          };
        } else if (meta.type === 'data') {
          const dayTotalCol = 2 + report.sessions.length;
          if (c === dayTotalCol) ws[addr].s.fill = { fgColor: { rgb: yellowPale } };
          if (c > dayTotalCol) ws[addr].s.fill = { fgColor: { rgb: 'F7F1FB' } };
        }
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Attendance');
    wb.Props = {
      Title: `${currentSchool()} Attendance Report`,
      Subject: `${formatLongDate(report.minDate)} to ${formatLongDate(report.maxDate)}`,
      Author: 'Junior Adventures Group',
      Company: 'Junior Adventures Group'
    };
    const array = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
    setProgress(100, 'Excel ready.');
    return new Blob([array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ---------- PDF generator ----------
  async function buildPdfBlob(progressStart = 0, progressEnd = 100) {
    ensureReport();
    await ensureFontsAndImages();
    els.pdfStage.innerHTML = '';
    const pages = buildPdfPages();
    const { jsPDF } = window.jspdf;
    let doc = null;

    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      els.pdfStage.appendChild(page.el);
      await nextFrame();
      const canvas = await html2canvas(page.el, {
        scale: 1.55,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false,
        windowWidth: page.orientation === 'landscape' ? 1123 : 794,
        windowHeight: page.orientation === 'landscape' ? 794 : 1123
      });
      const imgData = canvas.toDataURL('image/jpeg', 0.92);
      const orientation = page.orientation === 'landscape' ? 'landscape' : 'portrait';
      if (!doc) {
        doc = new jsPDF({ orientation, unit: 'mm', format: 'a4', compress: true });
      } else {
        doc.addPage('a4', orientation);
      }
      const pw = orientation === 'landscape' ? 297 : 210;
      const ph = orientation === 'landscape' ? 210 : 297;
      doc.addImage(imgData, 'JPEG', 0, 0, pw, ph, undefined, 'FAST');
      page.el.remove();
      const ratio = (i + 1) / pages.length;
      const pct = progressStart + ratio * (progressEnd - progressStart);
      setProgress(Math.round(pct), `Rendering PDF page ${i + 1} of ${pages.length}…`);
      await sleep(8);
    }
    els.pdfStage.innerHTML = '';
    return doc.output('blob');
  }

  function buildPdfPages() {
    const pages = [];
    const totalPages = 3 + report.weeks.length;
    const logo = 'assets/jag-logo.png';
    const period = `${formatLongDate(report.minDate)} to ${formatLongDate(report.maxDate)}`;

    // Page 1 - summary
    const p1 = createPdfPage('portrait', 1, totalPages, logo, period);
    const peakMonth = report.monthly.reduce((a,b) => (a.total/a.days) > (b.total/b.days) ? a : b);
    const peakWeekday = report.weekday.reduce((a,b) => a.average > b.average ? a : b);
    p1.content.innerHTML = `
      <h1 class="pdf-title">School Attendance Report</h1>
      <p class="pdf-subtitle"><strong>${escapeHtml(currentSchool())}</strong><br>${escapeHtml(period)}</p>
      <div class="pdf-kpis">
        ${pdfKpi(report.totalAttendance.toLocaleString(), 'Attendance entries')}
        ${pdfKpi(report.operatingDays.toLocaleString(), 'Operating days in export')}
        ${pdfKpi(report.averageDaily.toFixed(1), 'Average per operating day')}
        ${pdfKpi(report.dailyTotal[dateKey(report.busiestDay)].toLocaleString(), `Highest day – ${formatShortDate(report.busiestDay)}`)}
      </div>
      <h2 class="pdf-section-title">Programme overview</h2>
      ${programmeOverviewTable()}
      <h2 class="pdf-section-title">At a glance</h2>
      <p class="pdf-body">The strongest monthly average in the uploaded period was <strong>${monthYear(peakMonth.date)}</strong> at <strong>${(peakMonth.total/peakMonth.days).toFixed(1)}</strong> attendance entries per operating day. <strong>${escapeHtml(peakWeekday.name)}</strong> was the strongest weekday on average at <strong>${peakWeekday.average.toFixed(1)}</strong>. These figures count attendance entries across bookable sessions and are not a unique-child count.</p>
    `;
    pages.push({ el: p1.el, orientation: 'portrait' });

    // Page 2 - monthly and weekday
    const p2 = createPdfPage('portrait', 2, totalPages, logo, period);
    p2.content.innerHTML = `
      <h1 class="pdf-title">Attendance patterns</h1>
      <p class="pdf-subtitle">Monthly and weekday attendance across the detected reporting period.</p>
      <h2 class="pdf-section-title">Monthly attendance</h2>
      ${monthlyTable()}
      <h2 class="pdf-section-title">Weekday pattern</h2>
      ${weekdayTable()}
      <p class="pdf-note">Monthly averages are based only on operating days represented in the uploaded data. Partial months therefore reflect the portion of the month present in the export.</p>
    `;
    pages.push({ el: p2.el, orientation: 'portrait' });

    // Page 3 - session summary
    const p3 = createPdfPage('portrait', 3, totalPages, logo, period);
    p3.content.innerHTML = `
      <h1 class="pdf-title">Bookable session breakdown</h1>
      <p class="pdf-subtitle">Attendance entries recorded against each session category detected in the uploaded file.</p>
      ${sessionBreakdownTable()}
      <h2 class="pdf-section-title">Detailed weekly tables</h2>
      <p class="pdf-body">The following appendix reproduces the detailed day-by-day structure of the Excel workbook. Each week is clearly separated and shows attendance against every detected session, plus daily and programme totals.</p>
    `;
    pages.push({ el: p3.el, orientation: 'portrait' });

    // Weekly landscape pages
    report.weeks.forEach((week, idx) => {
      const pageNo = 4 + idx;
      const p = createPdfPage('landscape', pageNo, totalPages, logo, period);
      const monday = week.monday;
      const sunday = addDays(monday, 6);
      const total = week.dates.reduce((sum,d) => sum + report.dailyTotal[dateKey(d)], 0);
      p.content.innerHTML = `
        <h1 class="pdf-week-title">Week commencing ${escapeHtml(formatLongDate(monday))}</h1>
        <div class="pdf-week-meta">${escapeHtml(currentSchool())} &nbsp; | &nbsp; ${escapeHtml(formatShortDate(monday))} – ${escapeHtml(formatShortDate(sunday))} &nbsp; | &nbsp; Week total: <strong>${total.toLocaleString()}</strong> attendance entries</div>
        ${weeklyTable(week)}
        <p class="pdf-note">Attendance entries are shown by bookable session. “NA” indicates that no attendance was recorded for that programme family on the day.</p>
      `;
      pages.push({ el: p.el, orientation: 'landscape' });
    });
    return pages;
  }

  function createPdfPage(orientation, pageNo, totalPages, logo, period) {
    const el = document.createElement('section');
    el.className = `pdf-page ${orientation}`;
    el.innerHTML = `
      <div class="pdf-brandbar"></div>
      <img class="pdf-header-logo" src="${logo}" alt="" />
      <div class="pdf-header-text"><strong>Junior Adventures Group</strong><span>${escapeHtml(currentSchool())} – Attendance report</span></div>
      <div class="pdf-content"></div>
      <div class="pdf-footer"><span>${escapeHtml(period)}</span><span>Page ${pageNo} of ${totalPages}</span></div>
    `;
    return { el, content: el.querySelector('.pdf-content') };
  }

  function pdfKpi(value, label) {
    return `<div class="pdf-kpi"><b>${escapeHtml(value)}</b><span>${escapeHtml(label)}</span></div>`;
  }

  function programmeOverviewTable() {
    const head = '<tr><th>Programme</th><th>Attendance entries</th><th>Share of total</th><th>Average / day</th></tr>';
    const body = report.families.map(f => {
      const total = report.familyTotals[f];
      return `<tr><td>${escapeHtml(f)}</td><td>${total.toLocaleString()}</td><td>${(100*total/report.totalAttendance).toFixed(1)}%</td><td>${(total/report.operatingDays).toFixed(1)}</td></tr>`;
    }).join('');
    return `<table class="pdf-table"><thead>${head}</thead><tbody>${body}</tbody></table>`;
  }

  function monthlyTable() {
    const famHeads = report.families.map(f => `<th>${escapeHtml(f)}</th>`).join('');
    const rows = report.monthly.map(m => `<tr><td>${escapeHtml(monthYear(m.date))}</td><td>${m.days}</td><td>${m.total.toLocaleString()}</td><td>${(m.total/m.days).toFixed(1)}</td>${report.families.map(f => `<td>${m.families[f].toLocaleString()}</td>`).join('')}</tr>`).join('');
    return `<table class="pdf-table compact"><thead><tr><th>Month</th><th>Operating days</th><th>Total</th><th>Avg/day</th>${famHeads}</tr></thead><tbody>${rows}</tbody></table>`;
  }

  function weekdayTable() {
    const rows = report.weekday.map(w => `<tr><td>${escapeHtml(w.name)}</td><td>${w.days}</td><td>${w.total.toLocaleString()}</td><td>${w.average.toFixed(1)}</td></tr>`).join('');
    return `<table class="pdf-table"><thead><tr><th>Weekday</th><th>Operating days</th><th>Total entries</th><th>Average/day</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function sessionBreakdownTable() {
    const rows = report.sessions.map(s => `<tr><td>${escapeHtml(report.sessionMeta[s].family)}</td><td>${escapeHtml(report.sessionMeta[s].display)}</td><td>${report.sessionTotals[s].toLocaleString()}</td><td>${(report.sessionTotals[s]/report.operatingDays).toFixed(1)}</td></tr>`).join('');
    return `<table class="pdf-table"><thead><tr><th>Programme</th><th>Session</th><th>Attendance entries</th><th>Average/day</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function weeklyTable(week) {
    const totalCols = 2 + report.sessions.length + 1 + report.subtotalFamilies.length;
    const cls = totalCols > 11 ? 'very-compact' : totalCols > 9 ? 'compact' : '';
    const headers = ['Date','Day',...report.sessions.map(s => report.sessionMeta[s].display),'Day total',...report.subtotalFamilies.map(f => `${f} Only`)];
    const dayTotalIndex = 2 + report.sessions.length;
    const th = headers.map((h, i) => `<th class="${i === dayTotalIndex ? 'total-col' : i > dayTotalIndex ? 'subtotal-col' : ''}">${escapeHtml(h)}</th>`).join('');
    const rows = week.dates.map(d => {
      const dk = dateKey(d);
      const sessionCells = report.sessions.map(s => {
        const fam = report.sessionMeta[s].family;
        return `<td>${report.familyDaily[fam][dk] === 0 ? 'NA' : report.counts[dk][s]}</td>`;
      }).join('');
      const subtotalCells = report.subtotalFamilies.map(f => `<td class="subtotal-col">${report.familyDaily[f][dk]}</td>`).join('');
      return `<tr><td>${escapeHtml(formatShortDate(d))}</td><td>${escapeHtml(shortWeekday(d))}</td>${sessionCells}<td class="total-col">${report.dailyTotal[dk]}</td>${subtotalCells}</tr>`;
    }).join('');
    return `<table class="pdf-table ${cls}"><thead><tr>${th}</tr></thead><tbody>${rows}</tbody></table>`;
  }

  // ---------- Utilities ----------
  function normalizeHeader(v) { return cleanText(v).toLowerCase().replace(/[^a-z0-9]/g, ''); }
  function findHeaderIndex(normalizedRow, choices) {
    const normalizedChoices = choices.map(c => c.toLowerCase().replace(/[^a-z0-9]/g, ''));
    for (const c of normalizedChoices) {
      const i = normalizedRow.indexOf(c);
      if (i >= 0) return i;
    }
    return -1;
  }
  function cleanText(v) { return v == null ? '' : String(v).trim(); }

  function parseExcelDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !Number.isNaN(v.valueOf())) return stripTime(v);
    if (typeof v === 'number' && Number.isFinite(v)) {
      const parsed = XLSX.SSF.parse_date_code(v);
      if (!parsed) return null;
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    const text = String(v).trim();
    if (/^\d+(\.\d+)?$/.test(text)) return parseExcelDate(Number(text));
    const d = new Date(text);
    return Number.isNaN(d.valueOf()) ? null : stripTime(d);
  }

  function familyName(session) {
    let s = cleanText(session).replace(/\s*\([^)]*\)\s*$/, '').trim();
    s = s.replace(/\s+\d+\s*$/, '').trim();
    return s || cleanText(session);
  }

  function displaySessionName(session, label) {
    const family = familyName(session);
    if (/school\s*starts/i.test(label || '')) return titleCaseSpecial(family);
    const lm = String(label || '').match(/(\d{1,2})\s*(?:pm)?\s*(?:to|-|–)\s*(\d{1,2})\s*pm/i);
    if (lm) return `${titleCaseSpecial(family)} ${lm[1]}-${lm[2]}`;
    const sm = String(session || '').match(/\(\s*(\d{1,2}):?(\d{2})?\s*-\s*(\d{1,2}):?(\d{2})?\s*\)/);
    if (sm && Number(sm[1]) >= 12) {
      const a = Number(sm[1]) > 12 ? Number(sm[1]) - 12 : Number(sm[1]);
      const b = Number(sm[3]) > 12 ? Number(sm[3]) - 12 : Number(sm[3]);
      return `${titleCaseSpecial(family)} ${a}-${b}`;
    }
    return cleanText(session).replace(/\s+/g, ' ');
  }

  function sessionSortValue(meta) {
    const text = `${meta.label} ${meta.display}`;
    const m = text.match(/(\d{1,2})(?::\d{2})?\s*(?:pm)?\s*(?:to|-|–)\s*(\d{1,2})/i);
    if (!m) return -1;
    return Number(m[1]) * 100 + Number(m[2]);
  }

  function titleCaseSpecial(text) {
    return String(text || '').toLowerCase().split(/\s+/).map(w => ['and','then','to'].includes(w) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ')
      .replace(/\bthen\b/g, 'Then');
  }

  function countBy(values) {
    const m = new Map();
    for (const v of values) m.set(v, (m.get(v) || 0) + 1);
    return m;
  }
  function mostCommon(map) {
    let best = null, count = -1;
    for (const [k,v] of map.entries()) if (v > count) { best = k; count = v; }
    return best;
  }
  function stripTime(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function dateKey(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
  function parseDateKey(s) { const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); }
  function startOfWeek(d) { const x = stripTime(d); const day = x.getDay(); const diff = day === 0 ? -6 : 1-day; return addDays(x, diff); }
  function addDays(d, n) { const x = stripTime(d); x.setDate(x.getDate()+n); return x; }
  function shortWeekday(d) { return ['Sun','Mon','Tues','Wed','Thurs','Fri','Sat'][d.getDay()]; }
  function ordinal(n) { const v=n%100; return `${n}${v>=11&&v<=13?'th':({1:'st',2:'nd',3:'rd'}[n%10]||'th')}`; }
  function formatOrdinalDate(d, includeYear=false) { return `${ordinal(d.getDate())} ${d.toLocaleDateString('en-GB',{month:'short'})}${includeYear ? ` ${d.getFullYear()}` : ''}`; }
  function formatShortDate(d) { return `${d.getDate()} ${d.toLocaleDateString('en-GB',{month:'short'})}`; }
  function formatLongDate(d) { return d.toLocaleDateString('en-GB', { day:'numeric', month:'long', year:'numeric' }); }
  function monthYear(d) { return d.toLocaleDateString('en-GB',{month:'long',year:'numeric'}); }
  function currentSchool() { return cleanText(els.schoolName.value) || report?.school || 'School'; }
  function slug(text) { return text.normalize('NFKD').replace(/[’']/g,'').replace(/[^a-zA-Z0-9]+/g,'_').replace(/^_+|_+$/g,''); }
  function baseFileName() { return `${slug(currentSchool())}_Attendance_Report_${dateKey(report.minDate)}_to_${dateKey(report.maxDate)}`; }
  function escapeHtml(value) { return String(value ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#039;','"':'&quot;'}[ch])); }
  function ensureReport() { if (!report) throw new Error('Upload an attendance export first.'); }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function setBusy(busy, text='', percent=0) {
    [els.excelBtn, els.pdfBtn, els.packBtn, els.resetBtn].forEach(b => b.disabled = busy);
    if (busy) {
      els.progressWrap.classList.remove('hidden');
      setProgress(percent, text);
    } else {
      setProgress(100, 'Report ready.');
      setTimeout(() => els.progressWrap.classList.add('hidden'), 900);
      [els.excelBtn, els.pdfBtn, els.packBtn, els.resetBtn].forEach(b => b.disabled = false);
    }
  }
  function setProgress(percent, text) {
    els.progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    els.progressText.textContent = text;
  }
  function handleGenerationError(err) {
    console.error(err);
    setBusy(false);
    showMessage('error', `Report generation failed: ${err.message || err}`);
  }
  function showMessage(type, text) {
    if (type === 'error') {
      els.errorBox.textContent = text; els.errorBox.classList.remove('hidden'); els.successBox.classList.add('hidden');
    } else {
      els.successBox.textContent = text; els.successBox.classList.remove('hidden'); els.errorBox.classList.add('hidden');
    }
  }
  function clearMessages() { els.errorBox.classList.add('hidden'); els.successBox.classList.add('hidden'); }
  function resetApp() {
    report = null; sourceFileName = ''; els.fileInput.value = '';
    els.detectedPanel.classList.add('hidden'); els.generatePanel.classList.add('hidden');
    els.previewTableWrap.innerHTML = ''; clearMessages();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
  function nextFrame() { return new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))); }
  async function ensureFontsAndImages() {
    try { await document.fonts.ready; } catch (_) {}
    const logo = new window.Image();
    logo.src = 'assets/jag-logo.png';
    if (!logo.complete) await new Promise((resolve,reject) => { logo.onload=resolve; logo.onerror=reject; });
  }
})();
