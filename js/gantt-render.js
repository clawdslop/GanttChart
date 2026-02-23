/* ======================================================
   gantt-render.js – Think-cell style Gantt renderer
   Milestones on task lanes, scope legend, zoom, reorder
   ====================================================== */

const GanttRender = (() => {

  let _timeline = null;
  let _zoomLevel = 100;

  function setZoom(val) { _zoomLevel = val; render(App.getTasks()); }
  function getZoom() { return _zoomLevel; }

  /* ===== Drag Tooltip ===== */
  function showDragTooltip(ev, startDate, endDate) {
    let tip = document.getElementById('tc-drag-tooltip');
    if (!tip) {
      tip = document.createElement('div');
      tip.id = 'tc-drag-tooltip';
      document.body.appendChild(tip);
    }
    tip.textContent = endDate
      ? App.formatDate(startDate) + '  →  ' + App.formatDate(endDate)
      : App.formatDate(startDate);
    tip.style.left = (ev.clientX + 14) + 'px';
    tip.style.top = (ev.clientY - 30) + 'px';
    tip.style.display = 'block';
  }

  function hideDragTooltip() {
    const tip = document.getElementById('tc-drag-tooltip');
    if (tip) tip.style.display = 'none';
  }

  /* ===== Compute date from cursor position ===== */
  function dateFromCursorX(cell, clientX) {
    if (!_timeline) return null;
    const rect = cell.getBoundingClientRect();
    const pct = (clientX - rect.left) / rect.width;
    const dayOff = Math.round(pct * _timeline.totalDays);
    return ds(addD(_timeline.minDate, dayOff));
  }

  function render(tasks) {
    const container = document.getElementById('gantt-container');
    if (!container) return;
    if (!tasks || !tasks.length) {
      container.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#6b7280;font-size:14px;">Add tasks to see the Gantt chart</div>';
      _timeline = null; return;
    }

    const s = App.getSettings();
    const ts = s.timeScale || {};
    const regularTasks = tasks.filter(t => !t.isMilestone);
    const milestones = tasks.filter(t => t.isMilestone);

    if (!regularTasks.length) {
      container.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#6b7280;font-size:14px;">Add tasks to see the Gantt chart</div>';
      _timeline = null; return;
    }

    const { minDate, maxDate, totalDays } = computeTimeline(tasks);
    const weeks = computeWeekTicks(minDate, maxDate);
    const months = computeMonths(minDate, maxDate, ts.showYears !== false);
    const years = computeYears(minDate, maxDate);
    const todayPct = s.showToday ? computeTodayPct(minDate, totalDays) : null;
    _timeline = { minDate, maxDate, totalDays };

    const laneMap = new Map();
    milestones.forEach(ms => {
      let tid = ms.linkedTaskId;
      if (!tid || !regularTasks.find(r => r.id === tid)) {
        const msT = new Date(ms.start).getTime();
        let best = regularTasks[0], bestD = Infinity;
        regularTasks.forEach(r => {
          const rs = new Date(r.start).getTime(), re = new Date(r.end).getTime();
          const d = msT < rs ? rs - msT : (msT > re ? msT - re : 0);
          if (d < bestD) { bestD = d; best = r; }
        });
        tid = best.id;
      }
      if (!laneMap.has(tid)) laneMap.set(tid, []);
      laneMap.get(tid).push(ms);
    });

    const chartMinW = Math.round(900 * (_zoomLevel / 100));
    let html = '<div class="tc-chart">';

    // Year header
    if (ts.showYears !== false) {
      html += '<div class="tc-row tc-header-years">';
      html += '<div class="tc-cell tc-col-activity"></div>';
      html += `<div class="tc-cell tc-col-chart" style="min-width:${chartMinW}px"><div class="tc-year-labels">`;
      years.forEach(y => {
        html += `<div class="tc-year-label" style="left:${y.leftPct}%;width:${y.widthPct}%">${y.label}</div>`;
      });
      html += '</div></div>';
      html += '<div class="tc-cell tc-col-status"></div>';
      html += '<div class="tc-cell tc-col-comment"></div>';
      html += '</div>';
    }

    // Month header
    if (ts.showMonths !== false) {
      html += '<div class="tc-row tc-header-months">';
      html += '<div class="tc-cell tc-col-activity"></div>';
      html += `<div class="tc-cell tc-col-chart" style="min-width:${chartMinW}px"><div class="tc-month-labels">`;
      months.forEach(m => {
        html += `<div class="tc-month-label" style="left:${m.leftPct}%;width:${m.widthPct}%">${m.label}</div>`;
      });
      html += '</div></div>';
      html += '<div class="tc-cell tc-col-status"></div>';
      html += '<div class="tc-cell tc-col-comment"></div>';
      html += '</div>';
    }

    // Week/day header
    if (ts.showWeeks !== false) {
      html += '<div class="tc-row tc-header-days">';
      html += '<div class="tc-cell tc-col-activity tc-label-header">Activity</div>';
      html += `<div class="tc-cell tc-col-chart" style="min-width:${chartMinW}px"><div class="tc-day-labels">`;
      weeks.forEach(w => {
        html += `<div class="tc-day-label" style="left:${w.pct}%">${w.label}</div>`;
      });
      html += '</div></div>';
      html += '<div class="tc-cell tc-col-status tc-label-header">Status</div>';
      html += '<div class="tc-cell tc-col-comment tc-label-header">Comment</div>';
      html += '</div>';
    }

    if (ts.showWeeks === false) {
      html += '<div class="tc-row tc-header-days">';
      html += '<div class="tc-cell tc-col-activity tc-label-header">Activity</div>';
      html += `<div class="tc-cell tc-col-chart" style="min-width:${chartMinW}px"></div>`;
      html += '<div class="tc-cell tc-col-status tc-label-header">Status</div>';
      html += '<div class="tc-cell tc-col-comment tc-label-header">Comment</div>';
      html += '</div>';
    }

    // Task rows
    let lastRegularScope = null;
    regularTasks.forEach((t, rIdx) => {
      const allIdx = tasks.indexOf(t);

      if (t.scope && t.scope !== lastRegularScope) {
        if (lastRegularScope !== null) {
          html += '<div class="tc-row tc-phase-separator"><div class="tc-cell tc-col-activity"></div><div class="tc-cell tc-col-chart"></div><div class="tc-cell tc-col-status"></div><div class="tc-cell tc-col-comment"></div></div>';
        }
        lastRegularScope = t.scope;
      }

      const startOff = daysBetween(minDate, new Date(t.start));
      const endOff = daysBetween(minDate, new Date(t.end));
      const leftPct = (startOff / totalDays) * 100;
      const widthPct = ((endOff - startOff + 1) / totalDays) * 100;
      const barColor = App.getTaskColor(t);
      const sel = t.id === App.getSelected() ? ' tc-selected' : '';

      html += `<div class="tc-row tc-task-row${sel}" data-id="${t.id}" data-all-idx="${allIdx}">`;
      html += `<div class="tc-cell tc-col-activity tc-activity-name"><span class="tc-drag-handle" title="Drag to reorder">⠿</span>${esc(t.name)}</div>`;
      html += `<div class="tc-cell tc-col-chart tc-bar-cell" style="min-width:${chartMinW}px">`;
      weeks.forEach(w => { html += `<div class="tc-grid-line" style="left:${w.pct}%"></div>`; });
      if (todayPct !== null) html += `<div class="tc-today-line" style="left:${todayPct}%"></div>`;

      // Primary bar (skip if barStyle is 'none')
      if ((t.barStyle || 'solid') !== 'none') {
        const barExtra = t.barStyle === 'dashed' ? `color:${barColor};` : '';
        html += `<div class="tc-bar tc-bar-${t.barStyle || 'solid'}" data-task-id="${t.id}" style="left:${leftPct}%;width:${widthPct}%;background-color:${barColor};${barExtra}">`;
        html += '<div class="tc-bar-handle tc-bar-handle-l"></div>';
        html += '<div class="tc-bar-handle tc-bar-handle-r"></div>';
        html += '</div>';
      }

      // Additional segment bars
      (t.segments || []).forEach((seg, si) => {
        const segStartOff = daysBetween(minDate, new Date(seg.start));
        const segEndOff = daysBetween(minDate, new Date(seg.end));
        const segLeftPct = (segStartOff / totalDays) * 100;
        const segWidthPct = ((segEndOff - segStartOff + 1) / totalDays) * 100;
        const segColor = seg.color || (seg.scope ? App.getScopeColor(seg.scope) : null) || barColor;
        const segStyle = seg.barStyle || 'solid';
        const segBarExtra = segStyle === 'dashed' ? `color:${segColor};` : '';
        html += `<div class="tc-bar tc-bar-${segStyle}" data-task-id="${t.id}" data-seg-idx="${si}" style="left:${segLeftPct}%;width:${segWidthPct}%;background-color:${segColor};${segBarExtra}">`;
        html += '<div class="tc-bar-handle tc-bar-handle-l"></div>';
        html += '<div class="tc-bar-handle tc-bar-handle-r"></div>';
        html += '</div>';
      });

      // Activity milestones
      (t.milestones || []).forEach((am, ai) => {
        const amOff = daysBetween(minDate, new Date(am.date));
        const amPct = ((amOff + 0.5) / totalDays) * 100;
        const mt = (App.getMsTypes() || []).find(x => x.type === am.type);
        const amColor = am.done ? '#2E7D32' : (am.color || (mt ? mt.color : null) || s.msColor || '#607D8B');
        const displayLabel = am.label || am.type;
        const dateLabel = s.showMsDates ? `<div class="tc-ams-date">${App.formatDate(am.date)}</div>` : '';
        const pinCls = am.pin ? ' tc-ams-pinned' : '';
        html += `<div class="tc-ams${pinCls}" style="left:${amPct}%" title="${esc(am.type + ': ' + displayLabel)}" data-task-id="${t.id}" data-ams-idx="${ai}">`;
        html += `<div class="tc-ams-marker" style="color:${amColor}">${am.done ? '✓' : '▲'}</div>`;
        html += `<div class="tc-ams-label">${esc(displayLabel)}</div>`;
        html += dateLabel;
        html += '</div>';
      });

      // Standalone milestones on this task's lane
      const laneMs = laneMap.get(t.id) || [];
      laneMs.forEach(ms => {
        const off = daysBetween(minDate, new Date(ms.start));
        const pct = ((off + 0.5) / totalDays) * 100;
        const isDone = ms.status === 'complete';
        const msColor = isDone ? '#2E7D32' : (ms.color || s.msColor || '#37474F');
        const dateLabel = s.showMsDates ? `<div class="tc-ms-date">${App.formatDate(ms.start)}</div>` : '';
        html += `<div class="tc-milestone" data-ms-id="${ms.id}" style="left:${pct}%">`;
        html += `<div class="tc-ms-marker" style="color:${msColor}">${isDone ? '✓' : '▲'}</div>`;
        html += `<div class="tc-ms-label">${esc(ms.name)}</div>`;
        html += dateLabel;
        html += '</div>';
      });

      html += '</div>';
      html += `<div class="tc-cell tc-col-status"><span class="tc-status-dot ${t.status}">${statusIcon(t.status)}</span> ${statusLabel(t.status)}</div>`;
      html += `<div class="tc-cell tc-col-comment">${esc(t.comment || '')}</div>`;
      html += '</div>';
    });

    // TODAY label row
    if (todayPct !== null) {
      html += '<div class="tc-row" style="min-height:22px">';
      html += '<div class="tc-cell tc-col-activity"></div>';
      html += `<div class="tc-cell tc-col-chart tc-bar-cell" style="position:relative;min-width:${chartMinW}px">`;
      html += `<div class="tc-today-line" style="left:${todayPct}%"></div>`;
      html += `<div class="tc-today-label" style="left:${todayPct}%">TODAY</div>`;
      html += '</div>';
      html += '<div class="tc-cell tc-col-status"></div>';
      html += '<div class="tc-cell tc-col-comment"></div>';
      html += '</div>';
    }

    // Scope legend (if enabled in settings)
    const allScopes = App.getScopes();
    if (allScopes.length && s.showLegend !== false) {
      html += '<div class="tc-scope-legend">';
      allScopes.forEach(sc => {
        html += `<div class="tc-scope-legend-item"><span class="tc-scope-legend-swatch" style="background:${sc.color}"></span>${esc(sc.name)}</div>`;
      });
      html += '</div>';
    }

    html += '</div>';
    container.innerHTML = html;
    bindBarDrag(container);
    bindRowClick(container);
    bindRowReorder(container);
  }

  function refresh(tasks) { render(tasks); }

  /* ===== Row Reorder (drag up/down) ===== */
  function bindRowReorder(container) {
    container.querySelectorAll('.tc-drag-handle').forEach(handle => {
      handle.addEventListener('mousedown', onRowDragStart);
    });
  }

  function onRowDragStart(e) {
    e.preventDefault(); e.stopPropagation();
    const row = e.target.closest('.tc-task-row');
    if (!row) return;
    const taskId = row.dataset.id;
    const allIdx = parseInt(row.dataset.allIdx);
    if (isNaN(allIdx)) return;

    App.selectTask(taskId);
    const container = document.getElementById('gantt-container');
    const rowH = row.getBoundingClientRect().height;
    const startY = e.clientY;
    let dragDelta = 0;

    row.classList.add('tc-row-dragging');
    document.body.style.cursor = 'grabbing';
    document.body.style.userSelect = 'none';

    function onMove(ev) {
      dragDelta = ev.clientY - startY;
      row.style.transform = `translateY(${dragDelta}px)`;
      row.style.zIndex = '50';
    }

    function onUp() {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      row.classList.remove('tc-row-dragging');
      row.style.transform = '';
      row.style.zIndex = '';
      document.body.style.cursor = '';
      document.body.style.userSelect = '';

      const rowsMoved = Math.round(dragDelta / rowH);
      if (rowsMoved !== 0) {
        const targetIdx = Math.max(0, Math.min(App.getTasks().length - 1, allIdx + rowsMoved));
        App.reorderTask(allIdx, targetIdx);
      }
    }

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  /* ===== Bar Drag ===== */
  function bindBarDrag(container) {
    container.querySelectorAll('.tc-bar[data-task-id]').forEach(b => b.addEventListener('mousedown', onBarMouseDown));
    container.querySelectorAll('.tc-milestone[data-ms-id]').forEach(m => m.addEventListener('mousedown', onStandaloneMsDown));
    container.querySelectorAll('.tc-ams[data-task-id]').forEach(a => a.addEventListener('mousedown', onAmsMouseDown));
  }

  function onBarMouseDown(e) {
    e.stopPropagation();
    const bar = e.currentTarget;
    const taskId = bar.dataset.taskId;
    const segIdx = bar.dataset.segIdx !== undefined ? parseInt(bar.dataset.segIdx) : -1;
    const task = App.getTaskById(taskId);
    if (!task || !_timeline) return;
    App.selectTask(taskId);
    const cell = bar.closest('.tc-bar-cell'); if (!cell) return;
    const cellW = cell.getBoundingClientRect().width;
    const handle = e.target.closest('.tc-bar-handle');
    let mode = handle ? (handle.classList.contains('tc-bar-handle-l') ? 'resize-start' : 'resize-end') : 'move';

    let origS, origE;
    if (segIdx >= 0 && task.segments && task.segments[segIdx]) {
      origS = new Date(task.segments[segIdx].start);
      origE = new Date(task.segments[segIdx].end);
    } else {
      origS = new Date(task.start);
      origE = new Date(task.end);
    }

    const startX = e.clientX;
    const startY = e.clientY;
    bar.classList.add('tc-dragging');
    document.body.style.cursor = mode === 'move' ? 'grabbing' : 'ew-resize';
    document.body.style.userSelect = 'none';
    const linkedStandaloneMsIds = segIdx < 0
      ? App.getTasks().filter(t => t.isMilestone && t.linkedTaskId === taskId).map(t => t.id)
      : [];

    let highlightedRow = null;
    function clearHighlight() {
      if (highlightedRow) { highlightedRow.classList.remove('tc-lane-drop-target'); highlightedRow = null; }
    }

    function currentDraggedBarEl() {
      const safeTaskId = (window.CSS && CSS.escape) ? CSS.escape(taskId) : taskId;
      if (segIdx >= 0) {
        return document.querySelector(`.tc-bar[data-task-id="${safeTaskId}"][data-seg-idx="${segIdx}"]`);
      }
      return document.querySelector(`.tc-bar[data-task-id="${safeTaskId}"]:not([data-seg-idx])`);
    }

    function applyDragPreview(dy) {
      if (mode !== 'move') return;
      const draggedBar = currentDraggedBarEl();
      if (draggedBar) {
        draggedBar.classList.add('tc-dragging');
        draggedBar.style.transform = `translateY(${dy}px)`;
        draggedBar.style.zIndex = '100';
      }

      if (segIdx >= 0) return;

      const safeTaskId = (window.CSS && CSS.escape) ? CSS.escape(taskId) : taskId;
      document.querySelectorAll(`.tc-ams[data-task-id="${safeTaskId}"]`).forEach(el => {
        el.style.transform = `translateY(${dy}px)`;
        el.style.zIndex = '101';
      });

      linkedStandaloneMsIds.forEach(msId => {
        const safeMsId = (window.CSS && CSS.escape) ? CSS.escape(msId) : msId;
        document.querySelectorAll(`.tc-milestone[data-ms-id="${safeMsId}"]`).forEach(el => {
          el.style.transform = `translate(-50%, calc(-50% + ${dy}px))`;
          el.style.zIndex = '101';
        });
      });
    }

    function clearDragPreview() {
      const draggedBar = currentDraggedBarEl();
      if (draggedBar) {
        draggedBar.style.transform = '';
        draggedBar.style.zIndex = '';
        draggedBar.classList.remove('tc-dragging');
      }

      if (segIdx >= 0) return;

      const safeTaskId = (window.CSS && CSS.escape) ? CSS.escape(taskId) : taskId;
      document.querySelectorAll(`.tc-ams[data-task-id="${safeTaskId}"]`).forEach(el => {
        el.style.transform = '';
        el.style.zIndex = '';
      });
      linkedStandaloneMsIds.forEach(msId => {
        const safeMsId = (window.CSS && CSS.escape) ? CSS.escape(msId) : msId;
        document.querySelectorAll(`.tc-milestone[data-ms-id="${safeMsId}"]`).forEach(el => {
          el.style.transform = '';
          el.style.zIndex = '';
        });
      });
    }

    function onMove(ev) {
      const dd = Math.round(((ev.clientX - startX) / cellW) * _timeline.totalDays);
      const dy = ev.clientY - startY;
      let ns, ne;
      if (mode === 'move') { ns = addD(origS, dd); ne = addD(origE, dd); }
      else if (mode === 'resize-start') { ns = addD(origS, dd); ne = new Date(origE); if (ns > ne) ns = new Date(ne); }
      else { ns = new Date(origS); ne = addD(origE, dd); if (ne < ns) ne = new Date(ns); }

      showDragTooltip(ev, ds(ns), ds(ne));

      if (mode === 'move') {
        const container = document.getElementById('gantt-container');
        const rows = container.querySelectorAll('.tc-task-row[data-id]');
        clearHighlight();
        rows.forEach(row => {
          if (row.dataset.id === taskId) return;
          const rect = row.getBoundingClientRect();
          if (ev.clientY >= rect.top && ev.clientY <= rect.bottom) {
            row.classList.add('tc-lane-drop-target');
            highlightedRow = row;
          }
        });
      }

      if (dd) {
        if (segIdx >= 0) {
          App.updateSegment(taskId, segIdx, { start: ds(ns), end: ds(ne) }, { chartOnly: true });
        } else {
          App.updateTask(taskId, { start: ds(ns), end: ds(ne) }, { chartOnly: true });
        }
      }
      applyDragPreview(dy);
    }

    function onUp(ev) {
      clearDragPreview();
      bar.classList.remove('tc-dragging');
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      hideDragTooltip();
      clearHighlight();

      if (mode === 'move') {
        const container = document.getElementById('gantt-container');
        const rows = [...container.querySelectorAll('.tc-task-row[data-id]')];
        let targetRow = null;
        rows.forEach(row => {
          const rect = row.getBoundingClientRect();
          if (ev.clientY >= rect.top && ev.clientY <= rect.bottom && row.dataset.id !== taskId) {
            targetRow = row;
          }
        });

        if (targetRow) {
          const targetId = targetRow.dataset.id;
          if (segIdx >= 0) {
            App.moveSegmentToTask(taskId, segIdx, targetId);
            return;
          } else {
            App.moveBarToTask(taskId, targetId);
            linkedStandaloneMsIds.forEach(msId => {
              App.updateTask(msId, { linkedTaskId: targetId }, { chartOnly: true });
            });
            return;
          }
        }
      }

      App.renderTable();
    }

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  function onStandaloneMsDown(e) {
    e.stopPropagation();
    const el = e.currentTarget;
    const msId = el.dataset.msId;
    const task = App.getTaskById(msId);
    if (!task || !_timeline) return;

    App.selectTask(msId);
    const container = document.getElementById('gantt-container');
    const cell = el.closest('.tc-bar-cell');
    if (!cell) return;

    const cellW = cell.getBoundingClientRect().width;
    const startX = e.clientX;
    const startY = e.clientY;
    const origDate = new Date(task.start);

    el.style.zIndex = '100';
    document.body.style.cursor = 'grabbing';
    document.body.style.userSelect = 'none';

    function onMove(ev) {
      const dx = ev.clientX - startX;
      const dy = ev.clientY - startY;
      el.style.transform = `translate(calc(-50% + ${dx}px), calc(-50% + ${dy}px))`;

      const dd = Math.round((dx / cellW) * _timeline.totalDays);
      const newDate = ds(addD(origDate, dd));
      showDragTooltip(ev, newDate);
    }

    function onUp(ev) {
      el.style.transform = '';
      el.style.zIndex = '';
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      hideDragTooltip();

      const dx = ev.clientX - startX;
      const dd = Math.round((dx / cellW) * _timeline.totalDays);
      const newDate = dd ? ds(addD(origDate, dd)) : task.start;

      const rows = [...container.querySelectorAll('.tc-task-row[data-id]')];
      const mouseY = ev.clientY;
      let closestRow = null, closestDist = Infinity;
      rows.forEach(row => {
        const rect = row.getBoundingClientRect();
        const mid = rect.top + rect.height / 2;
        const dist = Math.abs(mouseY - mid);
        if (dist < closestDist) { closestDist = dist; closestRow = row; }
      });

      const updates = { start: newDate, end: newDate };
      if (closestRow) updates.linkedTaskId = closestRow.dataset.id;

      App.updateTask(msId, updates);
      App.renderTable();
    }

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  function onAmsMouseDown(e) {
    const el = e.currentTarget, taskId = el.dataset.taskId, idx = parseInt(el.dataset.amsIdx);
    const task = App.getTaskById(taskId);
    if (!task || !task.milestones || !task.milestones[idx] || !_timeline) return;
    if (task.milestones[idx].pin) return;
    e.stopPropagation();
    const cell = el.closest('.tc-bar-cell'); if (!cell) return;
    const cellW = cell.getBoundingClientRect().width, startX = e.clientX, origDate = new Date(task.milestones[idx].date);
    document.body.style.cursor = 'grabbing'; document.body.style.userSelect = 'none';

    function onMove(ev) {
      const dd = Math.round(((ev.clientX - startX) / cellW) * _timeline.totalDays);
      const newDate = ds(addD(origDate, dd));
      showDragTooltip(ev, newDate);
      if (!dd) return;
      App.updateActivityMilestone(taskId, idx, { date: newDate });
    }

    function onUp() {
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      hideDragTooltip();
    }

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  /* ===== Row Click ===== */
  function bindRowClick(container) {
    container.querySelectorAll('.tc-task-row[data-id]').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('.tc-bar, .tc-milestone, .tc-ams, .tc-drag-handle')) return;
        App.selectTask(row.dataset.id);
        App.highlightSelectedRow();
        const tr = document.querySelector(`#task-tbody tr[data-id="${row.dataset.id}"]`);
        if (tr) tr.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      });
    });
  }

  /* ===== Timeline Computations ===== */
  function computeTimeline(tasks) {
    let min = null, max = null;
    tasks.forEach(t => {
      if (t.start && (t.barStyle || 'solid') !== 'none') {
        const s = new Date(t.start), e = new Date(t.end || t.start);
        if (!min || s < min) min = s;
        if (!max || e > max) max = e;
      }
      (t.milestones || []).forEach(m => { if (!m.date) return; const md = new Date(m.date); if (!min || md < min) min = md; if (!max || md > max) max = md; });
      (t.segments || []).forEach(seg => {
        if (!seg.start) return;
        const ss = new Date(seg.start), se = new Date(seg.end || seg.start);
        if (!min || ss < min) min = ss; if (!max || se > max) max = se;
      });
    });
    if (!min) { const d = new Date(); min = new Date(d); max = new Date(d); max.setDate(max.getDate() + 30); }
    min.setDate(min.getDate() - 5); max.setDate(max.getDate() + 10);
    min = startOfWeek(min); max = endOfWeek(max);
    return { minDate: min, maxDate: max, totalDays: daysBetween(min, max) + 1 };
  }

  function computeWeekTicks(minDate, maxDate) {
    const ticks = [], total = daysBetween(minDate, maxDate) + 1, cur = new Date(minDate);
    while (cur <= maxDate) { ticks.push({ date: new Date(cur), pct: (daysBetween(minDate, cur) / total) * 100, label: cur.getDate() + '.' }); cur.setDate(cur.getDate() + 7); }
    return ticks;
  }

  function computeMonths(minDate, maxDate, withYear) {
    const months = [], total = daysBetween(minDate, maxDate) + 1, cur = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
    while (cur <= maxDate) {
      const mS = new Date(Math.max(cur.getTime(), minDate.getTime())), mE = new Date(cur.getFullYear(), cur.getMonth() + 1, 0), cE = new Date(Math.min(mE.getTime(), maxDate.getTime()));
      const lo = daysBetween(minDate, mS), ro = daysBetween(minDate, cE);
      let label = mS.toLocaleString('en', { month: 'short' });
      if (withYear) label += " '" + String(mS.getFullYear()).slice(2);
      months.push({ leftPct: (lo / total) * 100, widthPct: ((ro - lo + 1) / total) * 100, label });
      cur.setMonth(cur.getMonth() + 1);
    }
    return months;
  }

  function computeYears(minDate, maxDate) {
    const years = [], total = daysBetween(minDate, maxDate) + 1;
    let cur = new Date(minDate.getFullYear(), 0, 1);
    while (cur <= maxDate) {
      const yS = new Date(Math.max(cur.getTime(), minDate.getTime()));
      const yE = new Date(cur.getFullYear(), 11, 31);
      const cE = new Date(Math.min(yE.getTime(), maxDate.getTime()));
      const lo = daysBetween(minDate, yS), ro = daysBetween(minDate, cE);
      years.push({ leftPct: (lo / total) * 100, widthPct: ((ro - lo + 1) / total) * 100, label: String(cur.getFullYear()) });
      cur.setFullYear(cur.getFullYear() + 1);
    }
    return years;
  }

  function computeTodayPct(minDate, totalDays) {
    const today = new Date(); today.setHours(0,0,0,0);
    const off = daysBetween(minDate, today);
    return (off < 0 || off > totalDays) ? null : (off / totalDays) * 100;
  }

  /* ===== Helpers ===== */
  function daysBetween(d1, d2) { const a = new Date(d1), b = new Date(d2); a.setHours(0,0,0,0); b.setHours(0,0,0,0); return Math.round((b - a) / 86400000); }
  function addD(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
  function ds(d) { return d.toISOString().slice(0, 10); }
  function startOfWeek(d) { const day = d.getDay(); return new Date(d.getFullYear(), d.getMonth(), d.getDate() - day + (day === 0 ? -6 : 1)); }
  function endOfWeek(d) { const s = startOfWeek(d); s.setDate(s.getDate() + 6); return s; }
  function statusIcon(s) { return s === 'complete' ? '●' : s === 'in-progress' ? '◐' : '○'; }
  function statusLabel(s) { return s === 'complete' ? 'Complete' : s === 'in-progress' ? 'In progress' : s === 'not-started' ? 'Not started' : (s || ''); }
  function esc(s) { return !s ? '' : s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

  /* ===== Init ===== */
  function init() {
    App.onChange(() => { if (!App.shouldSuppressTableRender()) App.renderTable(); refresh(App.getTasks()); });
    render(App.getTasks());

    const slider = document.getElementById('zoom-slider');
    const valLabel = document.getElementById('zoom-value');
    if (slider) {
      slider.addEventListener('input', e => {
        _zoomLevel = parseInt(e.target.value);
        if (valLabel) valLabel.textContent = _zoomLevel + '%';
        render(App.getTasks());
      });
    }
  }
  document.addEventListener('DOMContentLoaded', () => setTimeout(init, 50));

  return { render, refresh, setZoom, getZoom };
})();
