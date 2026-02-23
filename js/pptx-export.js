/* ======================================================
   pptx-export.js – Think-cell style PowerPoint Gantt export
   Updated for scopes & activity milestones
   ====================================================== */

const PptxExport = (() => {

  let pptx = null;
  let S = {};

  function shape(name) { return (S && S[name]) || name; }

  const SLIDE_W = 13.33, SLIDE_H = 7.5;
  const MARGIN = { L: 0.3, R: 0.2, T: 0.15, B: 0.15 };
  const SCOPE_COL_W = 0.12;
  const LABEL_COL_W = 2.3;
  const STATUS_COL_W = 1.0;
  const COMMENT_COL_W = 1.3;

  const CHART_L = MARGIN.L + SCOPE_COL_W + LABEL_COL_W;
  const CHART_R = SLIDE_W - MARGIN.R - STATUS_COL_W - COMMENT_COL_W;
  const CHART_W = CHART_R - CHART_L;

  const PHASE_BAND_H = 0.22, MONTH_ROW_H = 0.22, DAY_ROW_H = 0.18;
  const HEADER_TOTAL_H = PHASE_BAND_H + MONTH_ROW_H + DAY_ROW_H;
  const ROW_H = 0.28, BAR_H = 0.17, PHASE_GAP = 0.06, MILESTONE_AREA_H = 0.7;
  const FONT = 'Segoe UI';

  const C = {
    GRID: 'D0D4DB', GRID_LIGHT: 'ECEDF0', HEADER_BG: 'F0F2F5',
    HEADER_TEXT: '4A5568', LABEL_TEXT: '1A1F36', MUTED_TEXT: '718096',
    WHITE: 'FFFFFF', TODAY: 'C62828',
    STATUS_COMPLETE: '2E7D32', STATUS_PROGRESS: '1565C0', STATUS_NOTSTARTED: '9E9E9E',
    MILESTONE_MARKER: '37474F',
  };

  function getMsColor() {
    const settings = App.getSettings();
    return cleanHex(settings.msColor || '#37474F');
  }

  function exportToPptx(tasks) {
    if (!tasks || !tasks.length) { alert('No tasks to export.'); return; }
    pptx = new PptxGenJS();
    S = pptx.shapes || {};
    pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
    pptx.layout = 'WIDE';
    pptx.author = 'Gantt Chart Generator';
    const slide = pptx.addSlide();

    const regularTasks = tasks.filter(t => !t.isMilestone);
    const globalMs = tasks.filter(t => t.isMilestone && !t.linkedTaskId);
    const linkedMs = tasks.filter(t => t.isMilestone && t.linkedTaskId);
    const { minDate, maxDate, totalDays } = computeTimeline(tasks);
    const weeks = computeWeekTicks(minDate, maxDate, totalDays);
    const months = computeMonths(minDate, maxDate, totalDays);
    const scopeSpans = computeScopeSpans(regularTasks, minDate, totalDays);
    const scopeGaps = computeScopeGapPositions(regularTasks);

    const chartT = MARGIN.T + HEADER_TOTAL_H;
    const bodyH = computeBodyHeight(regularTasks, scopeGaps);

    drawScopeBands(slide, scopeSpans, MARGIN.T);
    drawMonthHeader(slide, months, MARGIN.T + PHASE_BAND_H);
    drawDayHeader(slide, weeks, MARGIN.T + PHASE_BAND_H + MONTH_ROW_H);
    drawColumnHeaders(slide, MARGIN.T + PHASE_BAND_H + MONTH_ROW_H);
    drawGridLines(slide, weeks, chartT, bodyH);
    drawTaskRows(slide, regularTasks, linkedMs, scopeGaps, minDate, totalDays, chartT);
    drawTodayLine(slide, minDate, totalDays, chartT, bodyH);
    drawMilestones(slide, globalMs, minDate, totalDays, chartT + bodyH);
    drawBorders(slide, chartT, bodyH);

    pptx.writeFile({ fileName: 'Gantt-Chart.pptx' })
      .then(() => console.log('PPTX exported'))
      .catch(err => { console.error(err); alert('Export failed: ' + err.message); });
  }

  /* ===== Timeline ===== */
  function computeTimeline(tasks) {
    let min = new Date(tasks[0].start), max = new Date(tasks[0].end || tasks[0].start);
    tasks.forEach(t => {
      const s = new Date(t.start), e = new Date(t.end || t.start);
      if (s < min) min = s; if (e > max) max = e;
      (t.milestones || []).forEach(m => { const md = new Date(m.date); if (md < min) min = md; if (md > max) max = md; });
    });
    min.setDate(min.getDate() - 5); max.setDate(max.getDate() + 10);
    min = startOfWeek(min); max = endOfWeek(max);
    return { minDate: min, maxDate: max, totalDays: daysBetween(min, max) + 1 };
  }

  function computeWeekTicks(minDate, maxDate, totalDays) {
    const t = [], cur = new Date(minDate);
    while (cur <= maxDate) { t.push({ date: new Date(cur), pct: daysBetween(minDate, cur) / totalDays, label: cur.getDate() + '.' }); cur.setDate(cur.getDate() + 7); }
    return t;
  }

  function computeMonths(minDate, maxDate, totalDays) {
    const months = [], cur = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
    while (cur <= maxDate) {
      const mS = new Date(Math.max(cur.getTime(), minDate.getTime())), mE = new Date(cur.getFullYear(), cur.getMonth()+1, 0), cE = new Date(Math.min(mE.getTime(), maxDate.getTime()));
      const lo = daysBetween(minDate, mS), ro = daysBetween(minDate, cE);
      months.push({ leftPct: lo / totalDays, widthPct: (ro - lo + 1) / totalDays, label: mS.toLocaleString('en', { month: 'short' }) });
      cur.setMonth(cur.getMonth() + 1);
    }
    return months;
  }

  function computeScopeSpans(tasks, minDate, totalDays) {
    const map = new Map();
    tasks.forEach(t => {
      if (!t.scope) return;
      const s = new Date(t.start), e = new Date(t.end || t.start), sc = App.getScopeById(t.scope);
      if (!sc) return;
      if (!map.has(t.scope)) map.set(t.scope, { min: s, max: e, color: cleanHex(sc.color), name: sc.name });
      else { const p = map.get(t.scope); if (s < p.min) p.min = s; if (e > p.max) p.max = e; }
    });
    return [...map.values()].map(p => {
      const lo = Math.max(0, daysBetween(minDate, p.min)), ro = daysBetween(minDate, p.max);
      return { leftPct: lo / totalDays, widthPct: (ro - lo + 1) / totalDays, color: p.color, name: p.name };
    });
  }

  function computeScopeGapPositions(tasks) {
    const gaps = new Set(); let prev = null;
    tasks.forEach((t, i) => { if (prev && t.scope && prev !== t.scope) gaps.add(i); if (t.scope) prev = t.scope; });
    return gaps;
  }

  function computeBodyHeight(tasks, gaps) { let h = tasks.length * ROW_H; gaps.forEach(() => { h += PHASE_GAP; }); return h; }

  function getRowY(chartT, idx, gaps) {
    let y = chartT, gc = 0;
    for (const g of gaps) { if (g <= idx) gc++; }
    return y + idx * ROW_H + gc * PHASE_GAP;
  }

  /* ===== Drawing ===== */
  function drawScopeBands(slide, spans, topY) {
    spans.forEach(p => {
      const x = CHART_L + p.leftPct * CHART_W, w = p.widthPct * CHART_W;
      if (w < 0.01) return;
      slide.addShape(shape('RECTANGLE'), { x, y: topY, w, h: PHASE_BAND_H, fill: { color: p.color }, line: { type: 'none' } });
      slide.addText(p.name.toUpperCase(), { x, y: topY, w, h: PHASE_BAND_H, fontSize: 7, fontFace: FONT, bold: true, color: C.WHITE, align: 'center', valign: 'middle' });
    });
  }

  function drawMonthHeader(slide, months, topY) {
    slide.addShape(shape('RECTANGLE'), { x: CHART_L, y: topY, w: CHART_W, h: MONTH_ROW_H, fill: { color: C.WHITE }, line: { color: C.GRID, width: 0.3 } });
    months.forEach(m => {
      const x = CHART_L + m.leftPct * CHART_W, w = m.widthPct * CHART_W;
      if (w < 0.01) return;
      slide.addShape(shape('LINE'), { x, y: topY, w: 0, h: MONTH_ROW_H, line: { color: C.GRID, width: 0.4 } });
      slide.addText(m.label, { x, y: topY, w, h: MONTH_ROW_H, fontSize: 8, fontFace: FONT, bold: true, color: C.HEADER_TEXT, align: 'center', valign: 'middle' });
    });
  }

  function drawDayHeader(slide, weeks, topY) {
    weeks.forEach(w => { slide.addText(w.label, { x: CHART_L + w.pct * CHART_W - 0.12, y: topY, w: 0.24, h: DAY_ROW_H, fontSize: 7, fontFace: FONT, color: C.MUTED_TEXT, align: 'center', valign: 'middle' }); });
  }

  function drawColumnHeaders(slide, topY) {
    slide.addText('Activity', { x: MARGIN.L + SCOPE_COL_W + 0.08, y: topY, w: LABEL_COL_W - 0.16, h: DAY_ROW_H, fontSize: 7, fontFace: FONT, bold: true, color: C.HEADER_TEXT, valign: 'middle' });
    slide.addText('Status', { x: CHART_R + 0.05, y: topY, w: STATUS_COL_W - 0.1, h: DAY_ROW_H, fontSize: 7, fontFace: FONT, bold: true, color: C.HEADER_TEXT, valign: 'middle' });
    slide.addText('Comment', { x: CHART_R + STATUS_COL_W + 0.05, y: topY, w: COMMENT_COL_W - 0.1, h: DAY_ROW_H, fontSize: 7, fontFace: FONT, bold: true, color: C.HEADER_TEXT, valign: 'middle' });
  }

  function drawGridLines(slide, weeks, chartT, bodyH) {
    weeks.forEach(w => { slide.addShape(shape('LINE'), { x: CHART_L + w.pct * CHART_W, y: chartT, w: 0, h: bodyH, line: { color: C.GRID, width: 0.3, dashType: 'dash' } }); });
  }

  function drawTaskRows(slide, tasks, linkedMs, scopeGaps, minDate, totalDays, chartT) {
    tasks.forEach((t, i) => {
      const y = getRowY(chartT, i, scopeGaps);
      slide.addShape(shape('LINE'), { x: MARGIN.L, y: y + ROW_H, w: SLIDE_W - MARGIN.L - MARGIN.R, h: 0, line: { color: C.GRID_LIGHT, width: 0.3 } });

      // Scope color bar
      const sc = App.getScopeById(t.scope);
      if (sc) {
        slide.addShape(shape('RECTANGLE'), { x: MARGIN.L, y, w: SCOPE_COL_W - 0.02, h: ROW_H, fill: { color: cleanHex(sc.color) }, line: { type: 'none' } });
      }

      slide.addText(t.name, { x: MARGIN.L + SCOPE_COL_W + 0.08, y, w: LABEL_COL_W - 0.16, h: ROW_H, fontSize: 8, fontFace: FONT, color: C.LABEL_TEXT, valign: 'middle', wrap: false });

      const sOff = daysBetween(minDate, new Date(t.start)), eOff = daysBetween(minDate, new Date(t.end));
      const x1 = CHART_L + (sOff / totalDays) * CHART_W, x2 = CHART_L + ((eOff + 1) / totalDays) * CHART_W;
      const bW = Math.max(x2 - x1, 0.04), bY = y + (ROW_H - BAR_H) / 2;
      const bColor = cleanHex(App.getTaskColor(t));

      if (t.barStyle === 'dashed') {
        slide.addShape(shape('RECTANGLE'), { x: x1, y: bY, w: bW, h: BAR_H, fill: { type: 'none' }, line: { color: bColor, width: 1.2, dashType: 'dash' } });
      } else {
        slide.addShape(shape('RECTANGLE'), { x: x1, y: bY, w: bW, h: BAR_H, fill: { color: bColor }, line: { type: 'none' } });
        if (t.barStyle === 'hatched') {
          for (let hx = 0; hx < bW; hx += 0.06) {
            slide.addShape(shape('LINE'), { x: x1 + hx, y: bY, w: 0.04, h: BAR_H, line: { color: C.WHITE, width: 0.4, transparency: 50 } });
          }
        }
      }

      // Activity milestones
      (t.milestones || []).forEach(am => {
        const amOff = daysBetween(minDate, new Date(am.date));
        const ax = CHART_L + ((amOff + 0.5) / totalDays) * CHART_W;
        const mt = (App.getMsTypes() || []).find(m => m.type === am.type);
        const amCol = cleanHex(am.color || (mt ? mt.color : null) || App.getSettings().msColor || '#607D8B');
        slide.addText('▲', { x: ax - 0.06, y: y, w: 0.12, h: ROW_H * 0.6, fontSize: 8, fontFace: FONT, color: amCol, align: 'center', valign: 'middle' });
        slide.addText(am.type, { x: ax - 0.15, y: y + ROW_H * 0.5, w: 0.3, h: ROW_H * 0.5, fontSize: 5, fontFace: FONT, bold: true, color: C.HEADER_TEXT, align: 'center', valign: 'top' });
      });

      // Linked milestones
      const lms = linkedMs.filter(m => m.linkedTaskId === t.id);
      lms.forEach(ms => {
        const mOff = daysBetween(minDate, new Date(ms.start));
        const mx = CHART_L + ((mOff + 0.5) / totalDays) * CHART_W;
        slide.addText('▲', { x: mx - 0.08, y, w: 0.16, h: ROW_H * 0.6, fontSize: 8, fontFace: FONT, color: getMsColor(), align: 'center', valign: 'middle' });
        slide.addText(ms.name, { x: mx - 0.4, y: y + ROW_H * 0.55, w: 0.8, h: ROW_H * 0.45, fontSize: 5, fontFace: FONT, color: C.HEADER_TEXT, align: 'center', valign: 'top', wrap: true });
      });

      // Status
      const sColor = t.status === 'complete' ? C.STATUS_COMPLETE : t.status === 'in-progress' ? C.STATUS_PROGRESS : C.STATUS_NOTSTARTED;
      const sDot = t.status === 'complete' ? '●' : t.status === 'in-progress' ? '◐' : '○';
      const sLbl = t.status === 'complete' ? 'Complete' : t.status === 'in-progress' ? 'In progress' : 'Not started';
      slide.addText([{ text: sDot + ' ', options: { fontSize: 8, color: sColor } }, { text: sLbl, options: { fontSize: 7, color: C.MUTED_TEXT } }], { x: CHART_R + 0.05, y, w: STATUS_COL_W - 0.1, h: ROW_H, fontFace: FONT, valign: 'middle' });

      if (t.comment) slide.addText(t.comment, { x: CHART_R + STATUS_COL_W + 0.05, y, w: COMMENT_COL_W - 0.1, h: ROW_H, fontSize: 7, fontFace: FONT, color: C.MUTED_TEXT, valign: 'middle', wrap: false });
    });
  }

  function drawTodayLine(slide, minDate, totalDays, chartT, bodyH) {
    const settings = App.getSettings();
    if (!settings.showToday) return;
    const today = new Date(); today.setHours(0,0,0,0);
    const off = daysBetween(minDate, today);
    if (off < 0 || off > totalDays) return;
    const x = CHART_L + (off / totalDays) * CHART_W;
    slide.addShape(shape('LINE'), { x, y: chartT, w: 0, h: bodyH, line: { color: C.TODAY, width: 1.2, dashType: 'dash' } });
    slide.addText('TODAY', { x: x - 0.22, y: chartT + bodyH + 0.02, w: 0.44, h: 0.14, fontSize: 6, fontFace: FONT, bold: true, color: C.TODAY, align: 'center', valign: 'middle' });
  }

  function drawMilestones(slide, milestones, minDate, totalDays, topY) {
    if (!milestones.length) return;
    slide.addShape(shape('LINE'), { x: CHART_L, y: topY, w: CHART_W, h: 0, line: { color: C.GRID, width: 0.8 } });
    const chartT = MARGIN.T + HEADER_TOTAL_H;
    milestones.forEach(ms => {
      const off = daysBetween(minDate, new Date(ms.start));
      const x = CHART_L + ((off + 0.5) / totalDays) * CHART_W;
      slide.addShape(shape('LINE'), { x, y: chartT, w: 0, h: topY - chartT, line: { color: C.GRID, width: 0.5, dashType: 'dash' } });
      slide.addText('▲', { x: x - 0.08, y: topY + 0.02, w: 0.16, h: 0.14, fontSize: 8, fontFace: FONT, color: getMsColor(), align: 'center', valign: 'middle' });
      slide.addText(ms.name, { x: x - 0.5, y: topY + 0.16, w: 1.0, h: 0.35, fontSize: 6, fontFace: FONT, color: C.HEADER_TEXT, align: 'center', valign: 'top', wrap: true });
    });
  }

  function drawBorders(slide, chartT, bodyH) {
    slide.addShape(shape('LINE'), { x: MARGIN.L, y: chartT, w: SLIDE_W - MARGIN.L - MARGIN.R, h: 0, line: { color: C.GRID, width: 0.6 } });
    slide.addShape(shape('LINE'), { x: MARGIN.L, y: chartT + bodyH, w: SLIDE_W - MARGIN.L - MARGIN.R, h: 0, line: { color: C.GRID, width: 0.6 } });
    slide.addShape(shape('LINE'), { x: CHART_L, y: chartT, w: 0, h: bodyH, line: { color: C.GRID, width: 0.4 } });
    slide.addShape(shape('LINE'), { x: CHART_R, y: MARGIN.T + PHASE_BAND_H + MONTH_ROW_H, w: 0, h: chartT + bodyH - (MARGIN.T + PHASE_BAND_H + MONTH_ROW_H), line: { color: C.GRID, width: 0.3 } });
    slide.addShape(shape('LINE'), { x: CHART_R + STATUS_COL_W, y: MARGIN.T + PHASE_BAND_H + MONTH_ROW_H, w: 0, h: chartT + bodyH - (MARGIN.T + PHASE_BAND_H + MONTH_ROW_H), line: { color: C.GRID, width: 0.3 } });
  }

  /* ===== Helpers ===== */
  function daysBetween(d1, d2) { const a = new Date(d1), b = new Date(d2); a.setHours(0,0,0,0); b.setHours(0,0,0,0); return Math.round((b - a) / 86400000); }
  function startOfWeek(d) { const day = d.getDay(); return new Date(d.getFullYear(), d.getMonth(), d.getDate() - day + (day === 0 ? -6 : 1)); }
  function endOfWeek(d) { const s = startOfWeek(d); s.setDate(s.getDate() + 6); return s; }
  function cleanHex(hex) { return hex ? hex.replace('#', '').toUpperCase() : '1565C0'; }

  return { exportToPptx };
})();
