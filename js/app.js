/* ======================================================
   app.js – State management, table UI, persistence
   Gantt Chart Generator
   ====================================================== */

const DEFAULT_SCOPE_PALETTE = [
  '#388E3C', '#1565C0', '#880E4F', '#E65100',
  '#4527A0', '#00695C', '#37474F', '#AD1457',
  '#00838F', '#BF360C', '#283593', '#558B2F',
];

const DEFAULT_MS_TYPES = [
  { type: 'IA', label: 'Initial Approval', color: '#E65100' },
  { type: 'FA', label: 'Final Approval',   color: '#2E7D32' },
  { type: 'TL', label: 'Technical Launch',  color: '#1565C0' },
];

const DEFAULT_SETTINGS = {
  dateFormat: 'DD.MM.YYYY',
  showToday: true,
  showMsDates: true,
  msColor: '#37474F',
  timeScale: { showYears: true, showMonths: true, showWeeks: true },
  presetMsColors: ['#2E7D32', '#F9A825', '#C62828'],
};

const STORAGE_KEY = 'gantt-v3';

const App = (() => {
  let tasks = [];
  let scopes = [];
  let msTypes = [...DEFAULT_MS_TYPES];
  let settings = { ...DEFAULT_SETTINGS };
  let selectedId = null;
  let _onChange = () => {};
  let _suppressTableRender = false;
  let panelCollapsed = false;

  /* ===== Helpers ===== */
  function genId(prefix) { return (prefix || 'id') + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
  function todayStr() { return new Date().toISOString().slice(0, 10); }
  function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x.toISOString().slice(0, 10); }
  function daysDiff(d1, d2) { const a = new Date(d1), b = new Date(d2); a.setHours(0,0,0,0); b.setHours(0,0,0,0); return Math.round((b - a) / 86400000); }

  function formatDate(isoStr) {
    if (!isoStr) return '';
    const d = new Date(isoStr);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    const mmm = d.toLocaleString('en', { month: 'short' });
    switch (settings.dateFormat) {
      case 'MM/DD/YYYY': return `${mm}/${dd}/${yyyy}`;
      case 'YYYY-MM-DD': return isoStr;
      case 'DD MMM YYYY': return `${dd} ${mmm} ${yyyy}`;
      case 'MMM DD': return `${mmm} ${d.getDate()}`;
      default: return `${dd}.${mm}.${yyyy}`;
    }
  }

  /* ===== Settings ===== */
  function getSettings() { return settings; }
  function updateSettings(changes) {
    Object.assign(settings, changes);
    save(); _onChange();
  }

  /* ===== Scope CRUD ===== */
  function addScope(name, color, description) {
    const s = { id: genId('sc'), name: name || 'New Scope', color: color || nextScopeColor(), description: description || '' };
    scopes.push(s); save(); _onChange(); return s;
  }
  function updateScope(id, changes) { const s = scopes.find(x => x.id === id); if (!s) return; Object.assign(s, changes); save(); _onChange(); }
  function deleteScope(id) { scopes = scopes.filter(x => x.id !== id); tasks.forEach(t => { if (t.scope === id) t.scope = ''; }); save(); _onChange(); }
  function getScopes() { return scopes; }
  function getScopeById(id) { return scopes.find(x => x.id === id); }
  function getScopeByName(name) { return scopes.find(x => x.name === name); }
  function nextScopeColor() { return DEFAULT_SCOPE_PALETTE[scopes.length % DEFAULT_SCOPE_PALETTE.length]; }
  function getScopeColor(scopeId) { const s = getScopeById(scopeId); return s ? s.color : '#607D8B'; }
  function getTaskColor(task) { if (task.color) return task.color; if (task.scope) return getScopeColor(task.scope); return '#1565C0'; }

  /* ===== Milestone Types ===== */
  function getMsTypes() { return msTypes; }
  function addMsType(type, label, color) { msTypes.push({ type, label, color: color || '#607D8B' }); save(); }

  /* ===== Task CRUD ===== */
  function createTask(overrides = {}) {
    const today = todayStr();
    const lastTask = tasks.filter(t => !t.isMilestone).pop();
    return {
      id: genId('t'), name: 'New Task', start: today, end: addDays(today, 7),
      scope: lastTask ? lastTask.scope : '', status: 'not-started', comment: '',
      barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '',
      milestones: [], segments: [], dependencies: '', progress: 0, ...overrides,
    };
  }

  function addTask(overrides = {}) { const t = createTask(overrides); tasks.push(t); selectedId = t.id; save(); _onChange(); return t; }
  function addMilestone() { return addTask({ name: 'Milestone', start: todayStr(), end: todayStr(), isMilestone: true, barStyle: 'solid' }); }

  function updateTask(id, changes, opts = {}) {
    const t = tasks.find(x => x.id === id); if (!t) return;
    const oldStart = t.start;
    Object.assign(t, changes);
    if (t.isMilestone && changes.start && !changes.end) t.end = t.start;
    if (t.milestones && (changes.start !== undefined || changes.end !== undefined)) {
      t.milestones.forEach(m => {
        if (m.pin === 'start') m.date = t.start;
        else if (m.pin === 'end') m.date = t.end;
      });
    }
    save();
    if (opts.chartOnly) { _suppressTableRender = true; _onChange(); _suppressTableRender = false; }
    else _onChange();
  }

  function deleteTask(id) {
    tasks = tasks.filter(x => x.id !== id);
    tasks.forEach(t => {
      if (t.dependencies) t.dependencies = t.dependencies.split(',').map(s => s.trim()).filter(d => d !== id).join(', ');
      if (t.linkedTaskId === id) t.linkedTaskId = '';
    });
    if (selectedId === id) selectedId = tasks.length ? tasks[0].id : null;
    save(); _onChange();
  }

  function moveTask(id, direction) {
    const idx = tasks.findIndex(t => t.id === id);
    if (idx < 0) return;
    const target = direction === 'up' ? idx - 1 : idx + 1;
    if (target < 0 || target >= tasks.length) return;
    [tasks[idx], tasks[target]] = [tasks[target], tasks[idx]];
    save(); _onChange();
  }

  function reorderTask(fromIdx, toIdx) {
    if (fromIdx === toIdx || fromIdx < 0 || toIdx < 0 || fromIdx >= tasks.length || toIdx >= tasks.length) return;
    const [item] = tasks.splice(fromIdx, 1);
    tasks.splice(toIdx, 0, item);
    save(); _onChange();
  }

  /* ===== Activity Milestones ===== */
  function addActivityMilestone(taskId, type, date) {
    const t = tasks.find(x => x.id === taskId); if (!t) return;
    if (!t.milestones) t.milestones = [];
    t.milestones.push({ type, date: date || t.end, label: type, done: false, pin: null });
    save(); _onChange();
  }
  function updateActivityMilestone(taskId, idx, changes) {
    const t = tasks.find(x => x.id === taskId); if (!t || !t.milestones || !t.milestones[idx]) return;
    Object.assign(t.milestones[idx], changes);
    if (changes.pin === 'start') t.milestones[idx].date = t.start;
    else if (changes.pin === 'end') t.milestones[idx].date = t.end;
    save(); _onChange();
  }
  function removeActivityMilestone(taskId, idx) {
    const t = tasks.find(x => x.id === taskId); if (!t || !t.milestones) return;
    t.milestones.splice(idx, 1); save(); _onChange();
  }

  /* ===== Segments ===== */
  function addSegment(taskId) {
    const t = tasks.find(x => x.id === taskId); if (!t || t.isMilestone) return;
    if (!t.segments) t.segments = [];
    const lastEnd = t.segments.length ? t.segments[t.segments.length - 1].end : t.end;
    t.segments.push({ id: genId('seg'), name: '', start: addDays(lastEnd, 1), end: addDays(lastEnd, 7), barStyle: t.barStyle || 'solid', color: '', scope: '', status: 'not-started', comment: '' });
    save(); _onChange();
  }
  function updateSegment(taskId, segIdx, changes, opts = {}) {
    const t = tasks.find(x => x.id === taskId); if (!t || !t.segments || !t.segments[segIdx]) return;
    Object.assign(t.segments[segIdx], changes); save();
    if (opts.chartOnly) { _suppressTableRender = true; _onChange(); _suppressTableRender = false; }
    else _onChange();
  }
  function removeSegment(taskId, segIdx) {
    const t = tasks.find(x => x.id === taskId); if (!t || !t.segments) return;
    t.segments.splice(segIdx, 1); save(); _onChange();
  }
  function moveSegmentToTask(fromTaskId, segIdx, toTaskId) {
    const from = tasks.find(x => x.id === fromTaskId);
    const to = tasks.find(x => x.id === toTaskId);
    if (!from || !to || to.isMilestone || !from.segments || !from.segments[segIdx]) return;
    if (!to.segments) to.segments = [];
    const [seg] = from.segments.splice(segIdx, 1);
    to.segments.push(seg);
    save(); _onChange();
  }
  function moveBarToTask(fromTaskId, toTaskId) {
    const from = tasks.find(x => x.id === fromTaskId);
    const to = tasks.find(x => x.id === toTaskId);
    if (!from || !to || to.isMilestone || from.isMilestone) return;
    if (!to.segments) to.segments = [];
    to.segments.push({
      id: genId('seg'), name: from.name || '', start: from.start, end: from.end,
      barStyle: (from.barStyle && from.barStyle !== 'none') ? from.barStyle : 'solid',
      color: from.color || '', scope: from.scope || '', status: from.status || 'not-started',
      comment: from.comment || '',
    });
    from.barStyle = 'none';
    save(); _onChange();
  }

  /* ===== Selection ===== */
  function selectTask(id) {
    if (selectedId === id) return;
    selectedId = id; highlightSelectedRow(); updateDeleteBtn(); renderActivityMilestones();
  }
  function highlightSelectedRow() {
    const tbody = document.getElementById('task-tbody');
    if (tbody) tbody.querySelectorAll('tr[data-id]').forEach(tr => tr.classList.toggle('selected', tr.dataset.id === selectedId));
    document.querySelectorAll('.tc-task-row[data-id]').forEach(r => r.classList.toggle('tc-selected', r.dataset.id === selectedId));
  }
  function getSelected() { return selectedId; }
  function getTasks() { return tasks; }
  function getTaskById(id) { return tasks.find(x => x.id === id); }
  function shouldSuppressTableRender() { return _suppressTableRender; }

  /* ===== Persistence ===== */
  function save() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ tasks, scopes, msTypes, settings })); } catch (_) {}
  }

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const data = JSON.parse(raw);
        if (Array.isArray(data)) { migrateV2(data); }
        else {
          tasks = data.tasks || [];
          scopes = data.scopes || [];
          msTypes = data.msTypes || [...DEFAULT_MS_TYPES];
          settings = { ...DEFAULT_SETTINGS, ...(data.settings || {}) };
          if (!settings.timeScale) settings.timeScale = { ...DEFAULT_SETTINGS.timeScale };
          if (!settings.presetMsColors) settings.presetMsColors = [...DEFAULT_SETTINGS.presetMsColors];
        }
        if (tasks.length) selectedId = tasks[0].id;
      }
    } catch (_) {
      const old = localStorage.getItem('gantt-chart-tasks-v2');
      if (old) { try { migrateV2(JSON.parse(old)); } catch (_) {} }
    }
  }

  function migrateV2(oldTasks) {
    const phaseMap = new Map(); let ci = 0;
    oldTasks.forEach(t => {
      const pName = t.phase || t.scope;
      if (pName && !phaseMap.has(pName)) { phaseMap.set(pName, { id: genId('sc'), name: pName, color: DEFAULT_SCOPE_PALETTE[ci % DEFAULT_SCOPE_PALETTE.length], description: '' }); ci++; }
    });
    scopes = [...phaseMap.values()];
    tasks = oldTasks.map(t => {
      const pName = t.phase || t.scope || ''; const sc = phaseMap.get(pName);
      return { ...t, scope: sc ? sc.id : '', milestones: t.milestones || [], segments: t.segments || [], linkedTaskId: t.linkedTaskId || '' };
    });
    msTypes = [...DEFAULT_MS_TYPES]; settings = { ...DEFAULT_SETTINGS }; save();
  }

  function setData(newTasks, newScopes, newMsTypes, newSettings) {
    tasks = newTasks || []; scopes = newScopes || [];
    msTypes = newMsTypes || [...DEFAULT_MS_TYPES];
    settings = { ...DEFAULT_SETTINGS, ...(newSettings || {}) };
    selectedId = tasks.length ? tasks[0].id : null;
    save(); _onChange();
  }

  function exportJSON() { return JSON.stringify({ tasks, scopes, msTypes, settings }, null, 2); }
  function onChange(fn) { _onChange = fn; }

  /* ===== Preset swatches HTML helper ===== */
  function presetSwatchHtml(idx) {
    const presets = settings.presetMsColors || [];
    if (!presets.length) return '';
    return presets.map(c =>
      `<span class="color-swatch" style="background:${c}" data-color="${c}" data-idx="${idx}"></span>`
    ).join('');
  }

  /* ===== Table Rendering ===== */
  function renderTable() {
    const tbody = document.getElementById('task-tbody');
    const empty = document.getElementById('empty-state');
    if (!tasks.length) { tbody.innerHTML = ''; empty.classList.remove('hidden'); renderActivityMilestones(); return; }
    empty.classList.add('hidden');
    let rows = '';
    tasks.forEach(t => {
      const sc = getScopeById(t.scope);
      const scopeColor = sc ? sc.color : '';
      rows += `
      <tr data-id="${t.id}" class="${t.id === selectedId ? 'selected' : ''}">
        <td class="col-handle"><span class="row-handle" title="Drag to reorder">⠿</span></td>
        <td class="col-name">
          ${t.isMilestone ? '<span class="milestone-badge">◆</span> ' : ''}
          <input type="text" value="${esc(t.name)}" data-field="name" />
        </td>
        <td class="col-start"><input type="date" value="${t.start}" data-field="start" /></td>
        <td class="col-end"><input type="date" value="${t.end}" data-field="end" ${t.isMilestone ? 'disabled' : ''} /></td>
        <td class="col-scope">
          <div class="scope-cell">
            ${scopeColor ? `<span class="scope-dot" style="background:${scopeColor}"></span>` : ''}
            <select data-field="scope">
              <option value="">—</option>
              ${scopes.map(s => `<option value="${s.id}" ${t.scope === s.id ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}
            </select>
          </div>
        </td>
        <td class="col-status">
          <select data-field="status">
            <option value="complete" ${t.status === 'complete' ? 'selected' : ''}>Complete</option>
            <option value="in-progress" ${t.status === 'in-progress' ? 'selected' : ''}>In progress</option>
            <option value="not-started" ${t.status === 'not-started' ? 'selected' : ''}>Not started</option>
          </select>
        </td>
        <td class="col-comment"><input type="text" value="${esc(t.comment || '')}" data-field="comment" placeholder="Comment…" /></td>
        <td class="col-style">
          ${t.isMilestone
            ? `<div class="color-pick-row">${presetSwatchHtml('')}<input type="color" value="${t.color || settings.msColor || '#37474F'}" data-field="color" class="ms-color-picker" title="Milestone color" /></div>`
            : `<div class="style-seg-row"><select data-field="barStyle">
                <option value="solid" ${t.barStyle === 'solid' ? 'selected' : ''}>Solid</option>
                <option value="hatched" ${t.barStyle === 'hatched' ? 'selected' : ''}>Hatched</option>
                <option value="dashed" ${t.barStyle === 'dashed' ? 'selected' : ''}>Dashed</option>
                <option value="none" ${t.barStyle === 'none' ? 'selected' : ''}>None</option>
              </select><button class="btn-add-seg" data-task-id="${t.id}" title="Add segment">+</button></div>`
          }
        </td>
      </tr>`;
      if (!t.isMilestone && t.segments && t.segments.length) {
        t.segments.forEach((seg, si) => {
          const segSc = getScopeById(seg.scope);
          const segScopeColor = segSc ? segSc.color : '';
          rows += `<tr data-id="${t.id}" data-seg-idx="${si}" class="segment-row">
            <td class="col-handle"></td>
            <td class="col-name"><span class="seg-badge">↳</span><input type="text" value="${esc(seg.name || '')}" data-field="name" placeholder="Segment ${si + 1}" /></td>
            <td class="col-start"><input type="date" value="${seg.start}" data-field="start" /></td>
            <td class="col-end"><input type="date" value="${seg.end}" data-field="end" /></td>
            <td class="col-scope">
              <div class="scope-cell">
                ${segScopeColor ? `<span class="scope-dot" style="background:${segScopeColor}"></span>` : ''}
                <select data-field="scope">
                  <option value="">—</option>
                  ${scopes.map(s => `<option value="${s.id}" ${seg.scope === s.id ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}
                </select>
              </div>
            </td>
            <td class="col-status">
              <select data-field="status">
                <option value="complete" ${seg.status === 'complete' ? 'selected' : ''}>Complete</option>
                <option value="in-progress" ${seg.status === 'in-progress' ? 'selected' : ''}>In progress</option>
                <option value="not-started" ${(seg.status || 'not-started') === 'not-started' ? 'selected' : ''}>Not started</option>
              </select>
            </td>
            <td class="col-comment"><input type="text" value="${esc(seg.comment || '')}" data-field="comment" placeholder="Comment…" /></td>
            <td class="col-style"><div class="style-seg-row"><select data-field="barStyle">
                <option value="solid" ${seg.barStyle === 'solid' ? 'selected' : ''}>Solid</option>
                <option value="hatched" ${seg.barStyle === 'hatched' ? 'selected' : ''}>Hatched</option>
                <option value="dashed" ${seg.barStyle === 'dashed' ? 'selected' : ''}>Dashed</option>
              </select><button class="seg-del" data-task-id="${t.id}" data-seg-idx="${si}" title="Remove segment">&times;</button></div></td>
          </tr>`;
        });
      }
    });
    tbody.innerHTML = rows;
    updateDeleteBtn(); renderActivityMilestones();
  }

  /* ===== Activity Milestones Panel ===== */
  function renderActivityMilestones() {
    const panel = document.getElementById('activity-ms-panel');
    const list = document.getElementById('ams-list');
    if (!panel || !list) return;
    const task = tasks.find(t => t.id === selectedId && !t.isMilestone);
    if (!task) { panel.classList.add('hidden'); return; }
    panel.classList.remove('hidden');
    const ms = task.milestones || [];
    if (!ms.length) { list.innerHTML = '<div class="ams-empty">No activity milestones. Click + Add.</div>'; return; }
    const presets = settings.presetMsColors || [];
    list.innerHTML = ms.map((m, i) => {
      const mt = msTypes.find(x => x.type === m.type);
      const amColor = m.color || (mt ? mt.color : null) || settings.msColor || '#607D8B';
      return `<div class="ams-row" data-idx="${i}">
        <input type="checkbox" class="ams-done" data-idx="${i}" ${m.done ? 'checked' : ''} title="Mark done" />
        <div class="color-pick-row">
          ${presets.map(c => `<span class="color-swatch" style="background:${c}" data-color="${c}" data-idx="${i}"></span>`).join('')}
          <input type="color" class="ams-color" value="${amColor}" data-idx="${i}" title="Milestone color" />
        </div>
        <select class="ams-type" data-idx="${i}">
          ${msTypes.map(x => `<option value="${x.type}" ${m.type === x.type ? 'selected' : ''}>${x.type}</option>`).join('')}
          <option value="_custom" ${msTypes.every(x => x.type !== m.type) ? 'selected' : ''}>Custom…</option>
        </select>
        <input type="text" class="ams-label-input" value="${esc(m.label || m.type)}" data-idx="${i}" placeholder="Label" />
        <input type="date" class="ams-date" value="${m.date || ''}" data-idx="${i}" />
        <select class="ams-pin" data-idx="${i}">
          <option value="">Manual</option>
          <option value="fixed" ${m.pin === 'fixed' ? 'selected' : ''}>Fixed</option>
          <option value="start" ${m.pin === 'start' ? 'selected' : ''}>Pin Start</option>
          <option value="end" ${m.pin === 'end' ? 'selected' : ''}>Pin End</option>
        </select>
        <button class="ams-del" data-idx="${i}" title="Remove">&times;</button>
      </div>`;
    }).join('');
  }

  function esc(s) { return !s ? '' : s.replace(/"/g, '&quot;').replace(/</g, '&lt;'); }
  function updateDeleteBtn() { const btn = document.getElementById('btn-delete-task'); if (btn) btn.disabled = !selectedId; }

  /* ===== Scope Modal ===== */
  function renderScopeModal() {
    const list = document.getElementById('scope-list'); if (!list) return;
    list.innerHTML = scopes.map(s => `
      <div class="scope-row" data-id="${s.id}">
        <input type="color" class="scope-color-picker" value="${s.color}" data-field="color" />
        <input type="text" class="scope-name-input" value="${esc(s.name)}" data-field="name" placeholder="Scope name" />
        <input type="text" class="scope-desc-input" value="${esc(s.description)}" data-field="description" placeholder="Description…" />
        <button class="scope-del" data-id="${s.id}" title="Delete">&times;</button>
      </div>
    `).join('');
  }

  /* ===== Settings Modal ===== */
  function renderSettingsModal() {
    const fmt = document.getElementById('setting-date-format');
    const today = document.getElementById('setting-show-today');
    const msDates = document.getElementById('setting-show-ms-dates');
    const msCol = document.getElementById('setting-ms-color');
    if (fmt) fmt.value = settings.dateFormat;
    if (today) today.checked = settings.showToday;
    if (msDates) msDates.checked = settings.showMsDates;
    if (msCol) msCol.value = settings.msColor;
    const ts = settings.timeScale || {};
    const tsY = document.getElementById('setting-ts-years');
    const tsM = document.getElementById('setting-ts-months');
    const tsW = document.getElementById('setting-ts-weeks');
    if (tsY) tsY.checked = ts.showYears !== false;
    if (tsM) tsM.checked = ts.showMonths !== false;
    if (tsW) tsW.checked = ts.showWeeks !== false;
    renderPresetColors();
  }

  function renderPresetColors() {
    const list = document.getElementById('preset-ms-colors-list');
    if (!list) return;
    const presets = settings.presetMsColors || [];
    list.innerHTML = presets.length
      ? presets.map((c, i) => `<span class="color-swatch color-swatch-lg" style="background:${c}" data-idx="${i}" title="Click to remove"></span>`).join('')
      : '<span style="font-size:11px;color:var(--text-muted)">No presets</span>';
  }

  /* ===== Event Binding ===== */
  function bindEvents() {
    const tbody = document.getElementById('task-tbody');

    tbody.addEventListener('click', e => {
      if (e.target.classList.contains('seg-del')) {
        const tid = e.target.dataset.taskId, si = parseInt(e.target.dataset.segIdx);
        if (tid && !isNaN(si)) removeSegment(tid, si);
        return;
      }
      if (e.target.classList.contains('btn-add-seg')) {
        const tid = e.target.dataset.taskId;
        if (tid) addSegment(tid);
        return;
      }
      if (e.target.classList.contains('color-swatch')) {
        const color = e.target.dataset.color;
        const tr = e.target.closest('tr[data-id]');
        if (tr && color) {
          updateTask(tr.dataset.id, { color });
          const picker = tr.querySelector('.ms-color-picker');
          if (picker) picker.value = color;
        }
        return;
      }
      if (e.target.closest('input, select, button')) return;
      const tr = e.target.closest('tr[data-id]');
      if (tr) selectTask(tr.dataset.id);
    });

    tbody.addEventListener('mousedown', e => {
      const input = e.target.closest('input, select');
      if (input) {
        const tr = input.closest('tr[data-id]');
        if (tr && tr.dataset.id !== selectedId) { selectedId = tr.dataset.id; highlightSelectedRow(); updateDeleteBtn(); renderActivityMilestones(); }
      }
    });

    tbody.addEventListener('change', e => {
      const input = e.target; const field = input.dataset.field; if (!field) return;
      const tr = input.closest('tr[data-id]'); if (!tr) return;
      const segIdx = tr.dataset.segIdx;
      if (segIdx !== undefined) {
        updateSegment(tr.dataset.id, parseInt(segIdx), { [field]: input.value });
      } else {
        updateTask(tr.dataset.id, { [field]: input.value });
      }
    });

    let inputTimer = null;
    tbody.addEventListener('input', e => {
      const input = e.target; if (input.tagName !== 'INPUT') return;
      const field = input.dataset.field; if (!field) return;
      const tr = input.closest('tr[data-id]'); if (!tr) return;
      clearTimeout(inputTimer);
      inputTimer = setTimeout(() => {
        const segIdx = tr.dataset.segIdx;
        if (segIdx !== undefined) {
          const t = tasks.find(x => x.id === tr.dataset.id);
          if (t && t.segments && t.segments[parseInt(segIdx)]) {
            t.segments[parseInt(segIdx)][field] = input.value; save();
            if (typeof GanttRender !== 'undefined') GanttRender.refresh(tasks);
          }
        } else {
          const t = tasks.find(x => x.id === tr.dataset.id); if (!t) return;
          t[field] = input.value; save();
          if (typeof GanttRender !== 'undefined') GanttRender.refresh(tasks);
        }
      }, 300);
    });

    // Activity milestones
    const amsPanel = document.getElementById('activity-ms-panel');
    document.getElementById('btn-add-ams').addEventListener('click', () => {
      if (!selectedId) return; const t = getTaskById(selectedId);
      if (!t || t.isMilestone) return;
      addActivityMilestone(selectedId, msTypes[0]?.type || 'IA');
    });
    amsPanel.addEventListener('change', e => {
      if (!selectedId) return; const idx = parseInt(e.target.dataset.idx); if (isNaN(idx)) return;
      if (e.target.classList.contains('ams-type')) {
        const newType = e.target.value;
        const oldMs = getTaskById(selectedId).milestones[idx];
        const changes = { type: newType };
        if (!oldMs.label || oldMs.label === oldMs.type || msTypes.some(mt => mt.label === oldMs.label || mt.type === oldMs.label)) {
          changes.label = newType;
        }
        updateActivityMilestone(selectedId, idx, changes);
      }
      if (e.target.classList.contains('ams-date')) updateActivityMilestone(selectedId, idx, { date: e.target.value });
      if (e.target.classList.contains('ams-color')) updateActivityMilestone(selectedId, idx, { color: e.target.value });
      if (e.target.classList.contains('ams-label-input')) updateActivityMilestone(selectedId, idx, { label: e.target.value });
      if (e.target.classList.contains('ams-pin')) updateActivityMilestone(selectedId, idx, { pin: e.target.value || null });
      if (e.target.classList.contains('ams-done')) updateActivityMilestone(selectedId, idx, { done: e.target.checked });
    });
    amsPanel.addEventListener('input', e => {
      if (!selectedId) return;
      if (e.target.classList.contains('ams-color')) {
        const idx = parseInt(e.target.dataset.idx); if (isNaN(idx)) return;
        updateActivityMilestone(selectedId, idx, { color: e.target.value });
      }
    });
    amsPanel.addEventListener('click', e => {
      if (e.target.classList.contains('ams-del')) { const idx = parseInt(e.target.dataset.idx); if (!isNaN(idx) && selectedId) removeActivityMilestone(selectedId, idx); return; }
      if (e.target.classList.contains('color-swatch')) {
        const color = e.target.dataset.color;
        const idx = parseInt(e.target.dataset.idx);
        if (!isNaN(idx) && selectedId && color) {
          updateActivityMilestone(selectedId, idx, { color });
        }
        return;
      }
    });

    // Toolbar
    document.getElementById('btn-toggle-panel').addEventListener('click', togglePanel);
    document.getElementById('btn-add-task').addEventListener('click', () => { addTask(); focusLastTaskName(); });
    document.getElementById('btn-add-milestone').addEventListener('click', () => { addMilestone(); focusLastTaskName(); });
    document.getElementById('btn-delete-task').addEventListener('click', () => { if (selectedId && confirm('Delete this task?')) deleteTask(selectedId); });
    document.getElementById('btn-sample').addEventListener('click', loadSample);

    // Scope modal
    document.getElementById('btn-scopes').addEventListener('click', () => { renderScopeModal(); document.getElementById('scope-modal').classList.remove('hidden'); });
    document.getElementById('btn-close-scopes').addEventListener('click', () => document.getElementById('scope-modal').classList.add('hidden'));
    document.getElementById('btn-add-scope').addEventListener('click', () => { addScope(); renderScopeModal(); });
    const scopeList = document.getElementById('scope-list');
    scopeList.addEventListener('input', e => { const row = e.target.closest('.scope-row'); if (!row) return; const field = e.target.dataset.field; if (field) updateScope(row.dataset.id, { [field]: e.target.value }); });
    scopeList.addEventListener('click', e => { if (e.target.classList.contains('scope-del')) deleteScope(e.target.dataset.id); renderScopeModal(); });

    // Settings modal
    document.getElementById('btn-settings').addEventListener('click', () => { renderSettingsModal(); document.getElementById('settings-modal').classList.remove('hidden'); });
    document.getElementById('btn-close-settings').addEventListener('click', () => document.getElementById('settings-modal').classList.add('hidden'));
    document.getElementById('setting-date-format').addEventListener('change', e => updateSettings({ dateFormat: e.target.value }));
    document.getElementById('setting-show-today').addEventListener('change', e => updateSettings({ showToday: e.target.checked }));
    document.getElementById('setting-show-ms-dates').addEventListener('change', e => updateSettings({ showMsDates: e.target.checked }));
    document.getElementById('setting-ms-color').addEventListener('input', e => updateSettings({ msColor: e.target.value }));

    // Time scale settings
    function tsChange() {
      updateSettings({ timeScale: {
        showYears: document.getElementById('setting-ts-years').checked,
        showMonths: document.getElementById('setting-ts-months').checked,
        showWeeks: document.getElementById('setting-ts-weeks').checked,
      }});
    }
    document.getElementById('setting-ts-years').addEventListener('change', tsChange);
    document.getElementById('setting-ts-months').addEventListener('change', tsChange);
    document.getElementById('setting-ts-weeks').addEventListener('change', tsChange);

    // Preset milestone colors
    document.getElementById('btn-add-preset-color').addEventListener('click', () => {
      const picker = document.getElementById('new-preset-color');
      const presets = settings.presetMsColors || [];
      if (!presets.includes(picker.value)) {
        presets.push(picker.value);
        updateSettings({ presetMsColors: [...presets] });
        renderPresetColors();
      }
    });
    document.getElementById('preset-ms-colors-list').addEventListener('click', e => {
      if (e.target.classList.contains('color-swatch')) {
        const idx = parseInt(e.target.dataset.idx);
        const presets = [...(settings.presetMsColors || [])];
        if (!isNaN(idx)) {
          presets.splice(idx, 1);
          updateSettings({ presetMsColors: presets });
          renderPresetColors();
        }
      }
    });

    // Import / Export
    document.getElementById('btn-import').addEventListener('click', () => document.getElementById('file-import').click());
    document.getElementById('file-import').addEventListener('change', e => {
      const file = e.target.files[0]; if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        try {
          const data = JSON.parse(ev.target.result);
          if (Array.isArray(data)) { migrateV2(data); _onChange(); }
          else if (data.tasks) setData(data.tasks, data.scopes, data.msTypes, data.settings);
          else alert('Invalid JSON.');
        } catch (_) { alert('Could not parse JSON file.'); }
      };
      reader.readAsText(file); e.target.value = '';
    });

    document.getElementById('btn-export-json').addEventListener('click', () => {
      const blob = new Blob([exportJSON()], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = 'gantt-project.json'; a.click();
      URL.revokeObjectURL(url);
    });

    document.getElementById('btn-export-pptx').addEventListener('click', () => {
      if (typeof PptxExport !== 'undefined') PptxExport.exportToPptx(tasks);
    });

    document.addEventListener('keydown', e => {
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
      if ((e.ctrlKey || e.metaKey) && e.key === 'n') { e.preventDefault(); addTask(); focusLastTaskName(); }
      if ((e.ctrlKey || e.metaKey) && e.key === 'e') { e.preventDefault(); document.getElementById('btn-export-pptx').click(); }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        if (selectedId && !e.target.closest('input,select,textarea')) { e.preventDefault(); if (confirm('Delete this task?')) deleteTask(selectedId); }
      }
    });

    initResizer();
  }

  function focusLastTaskName() {
    requestAnimationFrame(() => {
      const inputs = document.querySelectorAll('#task-tbody tr:last-child input[data-field="name"]');
      const last = inputs[inputs.length - 1]; if (last) { last.focus(); last.select(); }
    });
  }

  function togglePanel() {
    panelCollapsed = !panelCollapsed;
    document.getElementById('task-panel').classList.toggle('collapsed', panelCollapsed);
    document.getElementById('panel-resizer').classList.toggle('hidden', panelCollapsed);
    _onChange();
  }

  function initResizer() {
    const resizer = document.getElementById('panel-resizer'), panel = document.getElementById('task-panel');
    let startX, startW;
    resizer.addEventListener('mousedown', e => {
      e.preventDefault(); startX = e.clientX; startW = panel.offsetWidth;
      resizer.classList.add('dragging');
      document.addEventListener('mousemove', onMove); document.addEventListener('mouseup', onUp);
    });
    function onMove(e) { panel.style.width = Math.max(280, Math.min(window.innerWidth * 0.6, startW + (e.clientX - startX))) + 'px'; }
    function onUp() { resizer.classList.remove('dragging'); document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); _onChange(); }
  }

  /* ===== Sample Project ===== */
  function loadSample() {
    const base = addDays(todayStr(), -35);
    const sampleScopes = [
      { id: 'sc_mkt', name: 'Marketing', color: '#388E3C', description: 'Market positioning and outreach' },
      { id: 'sc_vdd', name: 'Valuation & DD', color: '#1565C0', description: 'Valuation, due diligence, analysis' },
      { id: 'sc_neg', name: 'Negotiation & Closing', color: '#880E4F', description: 'Deal negotiation and closing' },
    ];
    const sampleTasks = [
      { id: 't1', name: 'Teaser sent', start: base, end: addDays(base,6), scope: 'sc_mkt', status: 'complete', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 100 },
      { id: 't2', name: 'NDA signed', start: addDays(base,14), end: addDays(base,20), scope: 'sc_mkt', status: 'complete', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [{type:'FA',date:addDays(base,20),label:'FA',done:true,pin:'end'}], segments: [], dependencies: '', progress: 100 },
      { id: 't3', name: 'CIM sent', start: addDays(base,14), end: addDays(base,27), scope: 'sc_mkt', status: 'complete', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [{type:'IA',date:addDays(base,14),label:'IA',done:true,pin:'start'}], segments: [], dependencies: '', progress: 100 },
      { id: 't4', name: 'Calls with Management', start: addDays(base,28), end: addDays(base,42), scope: 'sc_vdd', status: 'complete', comment: '', barStyle: 'hatched', color: '', isMilestone: false, linkedTaskId: '', milestones: [], segments: [{ id: 'seg1', start: addDays(base,45), end: addDays(base,48), barStyle: 'hatched', color: '' }], dependencies: '', progress: 100 },
      { id: 't5', name: 'Financial Model & Valuation', start: addDays(base,35), end: addDays(base,55), scope: 'sc_vdd', status: 'complete', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [{type:'IA',date:addDays(base,35),label:'IA',done:true,pin:'start'},{type:'FA',date:addDays(base,55),label:'FA',done:false,pin:'end'}], segments: [], dependencies: '', progress: 100 },
      { id: 't6', name: 'Expression of Interest', start: addDays(base,42), end: addDays(base,52), scope: 'sc_vdd', status: 'complete', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 100 },
      { id: 't7', name: 'Data Room Access', start: addDays(base,42), end: addDays(base,62), scope: 'sc_vdd', status: 'in-progress', comment: '1 week delay', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [{type:'TL',date:addDays(base,62),label:'TL',done:false,pin:'end'}], segments: [], dependencies: '', progress: 60 },
      { id: 't8', name: 'Mgmt Meetings', start: addDays(base,49), end: addDays(base,69), scope: 'sc_vdd', status: 'in-progress', comment: '3 of 6 done', barStyle: 'hatched', color: '', isMilestone: false, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 50 },
      { id: 't9', name: 'Final Due Diligence', start: addDays(base,56), end: addDays(base,76), scope: 'sc_vdd', status: 'in-progress', comment: 'Delay', barStyle: 'dashed', color: '', isMilestone: false, linkedTaskId: '', milestones: [{type:'FA',date:addDays(base,76),label:'FA',done:false,pin:'end'}], segments: [], dependencies: '', progress: 20 },
      { id: 't10', name: 'Quality of Earnings', start: addDays(base,77), end: addDays(base,90), scope: 'sc_neg', status: 'not-started', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 0 },
      { id: 't11', name: 'Definitive Agreements', start: addDays(base,84), end: addDays(base,104), scope: 'sc_neg', status: 'not-started', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [{type:'TL',date:addDays(base,104),label:'TL',done:false,pin:'end'}], segments: [], dependencies: '', progress: 0 },
      { id: 't12', name: "Shareholders' Agreement", start: addDays(base,91), end: addDays(base,97), scope: 'sc_neg', status: 'not-started', comment: '', barStyle: 'solid', color: '', isMilestone: false, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 0 },
      { id: 'ms1', name: 'CIM reviewed', start: addDays(base,30), end: addDays(base,30), scope: '', status: 'complete', comment: '', barStyle: 'solid', color: '', isMilestone: true, linkedTaskId: 't3', milestones: [], segments: [], dependencies: '', progress: 0 },
      { id: 'ms2', name: 'Non-Binding Offer', start: addDays(base,50), end: addDays(base,50), scope: '', status: 'not-started', comment: '', barStyle: 'solid', color: '', isMilestone: true, linkedTaskId: 't6', milestones: [], segments: [], dependencies: '', progress: 0 },
      { id: 'ms3', name: 'Letter of Intent', start: addDays(base,70), end: addDays(base,70), scope: '', status: 'not-started', comment: '', barStyle: 'solid', color: '', isMilestone: true, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 0 },
      { id: 'ms4', name: 'Agreements signed', start: addDays(base,98), end: addDays(base,98), scope: '', status: 'not-started', comment: '', barStyle: 'solid', color: '', isMilestone: true, linkedTaskId: '', milestones: [], segments: [], dependencies: '', progress: 0 },
    ];
    setData(sampleTasks, sampleScopes, [...DEFAULT_MS_TYPES], { ...DEFAULT_SETTINGS });
  }

  function init() { load(); renderTable(); bindEvents(); }

  return {
    init, getTasks, getTaskById, getSelected, addTask, addMilestone,
    updateTask, deleteTask, selectTask, setData, onChange, renderTable,
    exportJSON, loadSample, getScopes, getScopeById, getScopeByName,
    getScopeColor, getTaskColor, addScope, updateScope, deleteScope,
    getMsTypes, addMsType, addActivityMilestone, updateActivityMilestone, removeActivityMilestone,
    highlightSelectedRow, shouldSuppressTableRender, renderActivityMilestones,
    addDays, todayStr, formatDate, getSettings, updateSettings,
    moveTask, reorderTask, DEFAULT_SCOPE_PALETTE,
    addSegment, updateSegment, removeSegment, moveSegmentToTask, moveBarToTask,
  };
})();

document.addEventListener('DOMContentLoaded', () => App.init());
