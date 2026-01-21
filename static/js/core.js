
function parsePlotlyDate(x) {
  if (typeof x === "number") return x;
  if (!x) return NaN;
  if (x instanceof Date) return x.getTime();
  let t = Date.parse(x);
  if (!Number.isNaN(t)) return t;
  if (typeof x === "string") {
    // Plotly иногда отдаёт 'YYYY-MM-DD HH:MM:SS' без 'T' — поправим
    if (x.includes(" ") && !x.includes("T")) {
      t = Date.parse(x.replace(" ", "T"));
      if (!Number.isNaN(t)) return t;
    }
    // Последняя попытка: убрать миллисекунды
    t = Date.parse(x.replace(/\.\d+/, ""));
    if (!Number.isNaN(t)) return t;
  }
  return NaN;
}

function getVisibleRangeFromPlot() {
  const plot = document.getElementById("plot");
  try {
    const xa =
      (plot && plot.layout && plot.layout.xaxis) ||
      (plot && plot._fullLayout && plot._fullLayout.xaxis);
    const rng = xa && xa.range;
    if (rng && rng.length === 2) {
      const a = parsePlotlyDate(rng[0]);
      const b = parsePlotlyDate(rng[1]);
      if (!Number.isNaN(a) && !Number.isNaN(b)) return [plotXToMs(a), plotXToMs(b)];
    }
  } catch (e) {}
  return null;
}

function parseRelayoutRange(ev) {
  try {
    if (!ev) return null;
    if (ev["xaxis.range"] && Array.isArray(ev["xaxis.range"])) {
      const a = parsePlotlyDate(ev["xaxis.range"][0]);
      const b = parsePlotlyDate(ev["xaxis.range"][1]);
      if (!Number.isNaN(a) && !Number.isNaN(b)) return [plotXToMs(a), plotXToMs(b)];
    }
    if (ev["xaxis.range[0]"] != null && ev["xaxis.range[1]"] != null) {
      const a = parsePlotlyDate(ev["xaxis.range[0]"]);
      const b = parsePlotlyDate(ev["xaxis.range[1]"]);
      if (!Number.isNaN(a) && !Number.isNaN(b)) return [plotXToMs(a), plotXToMs(b)];
    }
  } catch (e) {}
  return null;
}


// Plotly рисует ось времени в UTC. Чтобы на графике было локальное время (как в данных/сводке),
// сдвигаем X на локальный offset. При чтении диапазона с графика делаем обратное преобразование.
//function msToPlotX(ms) {
//  const offMin = new Date(ms).getTimezoneOffset(); // minutes (UTC - local)
//  return ms - offMin * 60000;
//}
//function plotXToMs(plotMs) {
//  const offMin = new Date(plotMs).getTimezoneOffset();
//  return plotMs + offMin * 60000;
//}

function msToPlotX(ms) {
  return ms;
}

function plotXToMs(plotMs) {
  return plotMs;
}

// LeMuRe Viewer UI (selection like <select multiple> + drag reorder + export by visible X range)
let LOADED = false;

let CHANNELS_FILE = [];       // channels in the order they appear in file
let CHANNELS_VIEW = [];       // currently shown order in list
let CHANNELS_VIEW_ALL = [];   // full order (sorted) ignoring filters; used for plotting/export order
let SAVED_ORDER = [];         // saved custom order (from channel_order.json) - not auto-applied
let SUMMARY = null;
let currentRange = null;      // [start_ms, end_ms]

// При загрузке НОВОГО теста нужно сбросить диапазон графика к полному диапазону теста.
// При смене каналов внутри одного теста диапазон должен сохраняться.
let _forceResetRangeOnNextPlot = false;

let VIEWER_SETTINGS = null;  // server-side settings for template styling

let redrawTimer = null;
let saveOrderTimer = null;

let selected = new Set();     // selected channel codes (like <select multiple>)
let anchorCode = null;        // last clicked code (for Shift selection)
let dragData = null;          // data for dragging multiple selected items

const el = (id) => document.getElementById(id);
const log = (msg) => {
  const L = el("log");
  if(!L) return;
  const ts = new Date().toISOString().replace("T"," ").slice(0,19);
  L.textContent = `[${ts}] ${msg}\n` + L.textContent;
};

// Show JS errors in the log so you see them immediately
window.addEventListener("error", (e) => {
  log("JS error: " + (e && e.message ? e.message : e));
});


// ===== UX/UI helpers (busy overlay, toasts, filters, auto-step) =====
let BUSY_GUARD = 0;
let _busyTimer = null;
let _busyStopwatch = null;

function _setOverlayVisible(vis, text) {
  const ov = el('busyOverlay');
  const tx = el('busyText');
  if(!ov) return;
  if(tx && text) tx.textContent = text;
  ov.classList.toggle('hidden', !vis);
  ov.setAttribute('aria-hidden', vis ? 'false' : 'true');
}

function beginBusy(text, opts) {
  BUSY_GUARD++;
  const my = BUSY_GUARD;
  let shown = false;

  // Optional elapsed timer on overlay (useful for long exports)
  const withTimer = !!(opts && opts.timer);
  const tEl = el('busyTimer');
  const t0 = (withTimer && window.performance && performance.now) ? performance.now() : 0;

  // Stop any previous overlay timers
  if(_busyTimer) clearTimeout(_busyTimer);
  if(_busyStopwatch) { try { clearInterval(_busyStopwatch); } catch(_) {} _busyStopwatch = null; }

  function fmtElapsed(ms) {
    try {
      const total = Math.max(0, Math.floor(ms));
      const sec = Math.floor(total / 1000);
      const min = Math.floor(sec / 60);
      const s2 = sec % 60;
      const t = Math.floor((total % 1000) / 100);
      return String(min).padStart(2,'0') + ':' + String(s2).padStart(2,'0') + '.' + String(t);
    } catch(e) {
      return '';
    }
  }

  if(withTimer && tEl) {
    tEl.textContent = '00:00.0';
    _busyStopwatch = setInterval(() => {
      if(my !== BUSY_GUARD) return;
      const now = (window.performance && performance.now) ? performance.now() : Date.now();
      tEl.textContent = fmtElapsed(now - t0);
    }, 100);
  } else {
    if(tEl) tEl.textContent = '';
  }

  _busyTimer = setTimeout(() => {
    if(my !== BUSY_GUARD) return;
    _setOverlayVisible(true, text || 'Работаю…');
    shown = true;
  }, 250);

  return () => {
    if(my !== BUSY_GUARD) return;
    if(_busyTimer) clearTimeout(_busyTimer);
    if(_busyStopwatch) { try { clearInterval(_busyStopwatch); } catch(_) {} _busyStopwatch = null; }
    if(tEl) tEl.textContent = '';
    if(shown) _setOverlayVisible(false, '');
  };
}

function toast(title, body, kind='info', ms=3500) {
  const c = el('toastContainer');
  if(!c) { log((title||'') + ' ' + (body||'')); return; }
  const t = document.createElement('div');
  t.className = 'toast ' + kind;
  const tt = document.createElement('div');
  tt.className = 'tTitle';
  tt.textContent = title || '';
  const tb = document.createElement('div');
  tb.className = 'tBody';
  tb.textContent = body || '';
  if(title) t.appendChild(tt);
  if(body) t.appendChild(tb);
  c.appendChild(t);
  setTimeout(() => {
    try { t.remove(); } catch(_) {}
  }, ms);
}

// Filters state
let FILTER_TEXT = '';
let FILTER_ONLY_SELECTED = false;
let FILTER_CHIP = null; // {type:'prefix'|'unit', value:string}

function hasActiveFilter() {
  return !!(FILTER_TEXT || FILTER_ONLY_SELECTED || FILTER_CHIP);
}

function applyChannelFilters(arr) {
  let out = arr.slice();
  if(FILTER_CHIP) {
    const t = FILTER_CHIP.type;
    const v = (FILTER_CHIP.value || '').toLowerCase();
    if(t === 'prefix') out = out.filter(ch => (ch.code || '').toLowerCase().startsWith(v));
    if(t === 'unit') out = out.filter(ch => ((ch.unit || '').toLowerCase() == v));
  }
  if(FILTER_TEXT) {
    const q = FILTER_TEXT.toLowerCase();
    out = out.filter(ch => {
      const hay = ((ch.code||'') + ' ' + (ch.label||'') + ' ' + (ch.unit||'')).toLowerCase();
      return hay.includes(q);
    });
  }
  if(FILTER_ONLY_SELECTED) {
    out = out.filter(ch => selected.has(ch.code));
  }
  return out;
}

function updateSelectedCountUI() {
  const sp = el('selCount');
  if(!sp) return;
  const total = selected ? selected.size : 0;
  // visible list = CHANNELS_VIEW
  const vis = (CHANNELS_VIEW || []).filter(ch => selected.has(ch.code)).length;
  sp.textContent = `Выбрано: ${total}` + (hasActiveFilter() ? ` (видно: ${vis})` : '');
}


// ===== Last selection (localStorage) =====
// Сохраняем последний выбор каналов + основные настройки UI локально в браузере,
// чтобы после перезапуска/обновления можно было быстро продолжить.
const LS_LAST_STATE_KEY = 'lemure_last_state_v1';
const LS_VIEW_OPTS_KEY = 'lemure_view_opts_v1';
const LS_RECENT_FOLDERS_KEY = 'lemure_recent_folders_v1';
const LS_REFRIGERANT_KEY = 'lemure_refrigerant_v1';

let VIEW_COMPACT = false;
let VIEW_GROUPS = false;
let COLLAPSED_GROUPS = new Set();

let RECENT_FOLDERS = [];
const RECENT_MAX = 10;

// Range stats cache (точное число точек в текущем диапазоне)
let RANGE_STATS = null;
let _rangeStatsTimer = null;
let _rangeStatsToken = 0;

let _lastStateTimer = null;

function _collectLastState() {
  try {
    const stepAuto = el('stepAuto') ? !!el('stepAuto').checked : true;
    const stepTarget = el('stepTarget') ? parseInt(el('stepTarget').value || '5000', 10) : 5000;
    const stepManual = el('step') ? parseInt(el('step').value || '1', 10) : 1;
    const showLegend = el('showLegend') ? !!el('showLegend').checked : true;

    return {
      v: 1,
      saved_at: new Date().toISOString(),
      selected: Array.from(selected || []),
      step_auto: stepAuto,
      step_target: (isFinite(stepTarget) && stepTarget > 0) ? stepTarget : 5000,
      step: (isFinite(stepManual) && stepManual > 0) ? stepManual : 1,
      show_legend: showLegend,
    };
  } catch(e) {
    return null;
  }
}

function saveLastStateNow() {
  try {
    if(!LOADED) return;
    const st = _collectLastState();
    if(!st) return;
    localStorage.setItem(LS_LAST_STATE_KEY, JSON.stringify(st));
  } catch(e) {}
}

function scheduleSaveLastState() {
  try {
    if(_lastStateTimer) clearTimeout(_lastStateTimer);
    _lastStateTimer = setTimeout(() => saveLastStateNow(), 600);
  } catch(e) {}
}

function applyLastStateIfAny() {
  try {
    const raw = localStorage.getItem(LS_LAST_STATE_KEY);
    if(!raw) return false;
    const st = JSON.parse(raw);
    if(!st || !Array.isArray(st.selected)) return false;

    const available = new Set((CHANNELS_FILE||[]).map(c => c.code));
    const sel = (st.selected||[]).map(c => String(c)).filter(c => available.has(c));

    if(el('stepAuto') && typeof st.step_auto === 'boolean') el('stepAuto').checked = st.step_auto;
    if(el('stepTarget') && st.step_target) el('stepTarget').value = String(parseInt(st.step_target, 10));
    if(el('step') && st.step) el('step').value = String(parseInt(st.step, 10));
    if(el('showLegend') && typeof st.show_legend === 'boolean') el('showLegend').checked = st.show_legend;

    updateStepUI();

    if(sel && sel.length) {
      setSelection(sel);
      anchorCode = sel[0];
      scheduleRedraw();
    }
    return true;
  } catch(e) {
    return false;
  }
}

// ===== Stage 3: view modes, recent folders, range stats, export info =====

function _safeJsonParse(raw, defv) {
  try { return JSON.parse(raw); } catch(e) { return defv; }
}

function loadViewOpts() {
  // Сейчас в UI оставлен только «компактно». Режим «группы» убран, поэтому принудительно выключаем.
  try {
    const raw = localStorage.getItem(LS_VIEW_OPTS_KEY);
    const st = raw ? _safeJsonParse(raw, null) : null;
    const compact = !!(st && st.compact);

    VIEW_COMPACT = compact;
    VIEW_GROUPS  = false;
    COLLAPSED_GROUPS = new Set();

    if(el('viewCompact')) el('viewCompact').checked = VIEW_COMPACT;
    // viewGroups больше нет в интерфейсе
  } catch(e) {}
}

function saveViewOpts() {
  try {
    const collapsed = {};
    (COLLAPSED_GROUPS || new Set()).forEach(k => { collapsed[String(k)] = true; });
    const st = { compact: !!VIEW_COMPACT, groups: !!VIEW_GROUPS, collapsed_groups: collapsed };
    localStorage.setItem(LS_VIEW_OPTS_KEY, JSON.stringify(st));
  } catch(e) {}
}

function syncViewOptsFromUI(save=false) {
  try {
    const vc = el('viewCompact');
    const vg = el('viewGroups');
    if(vc) VIEW_COMPACT = !!vc.checked;
    if(vg) VIEW_GROUPS  = !!vg.checked;
    if(save) saveViewOpts();
  } catch(e) {}
}


function initChannelControlsHints() {
  // Динамические подсказки в блоке «Управление каналами»
  const root = el('channelControls');
  const body = el('controlsHintsBody');
  if(!root || !body) return;

  const defText = String(body.getAttribute('data-default') || body.textContent || '').trim();
  const setText = (t) => { body.textContent = (t && String(t).trim()) ? String(t) : defText; };

  // Навешиваем на все элементы с data-hint внутри блока
  const nodes = root.querySelectorAll('[data-hint]');
  nodes.forEach(n => {
    n.addEventListener('mouseenter', () => setText(n.getAttribute('data-hint')));
    n.addEventListener('mouseleave', () => setText(defText));
  });

  // На всякий случай: уход мыши за блок
  root.addEventListener('mouseleave', () => setText(defText));
  setText(defText);
}

function applyViewClasses() {
  const list = el('channelList');
  if(!list) return;
  list.classList.toggle('compact', !!VIEW_COMPACT);
}

function groupOfCode(code) {
  const s = String(code || '');
  const i = s.indexOf('-');
  if(i > 0) return s.slice(0, i);
  return 'Другие';
}

function toggleGroupCollapsed(groupKey) {
  const k = String(groupKey);
  if(COLLAPSED_GROUPS.has(k)) COLLAPSED_GROUPS.delete(k);
  else COLLAPSED_GROUPS.add(k);
  saveViewOpts();

  const list = el('channelList');
  if(!list) return;
  // update items visibility
  list.querySelectorAll(`.chanItem[data-group="${CSS.escape(k)}"]`).forEach(li => {
    li.classList.toggle('groupHidden', COLLAPSED_GROUPS.has(k));
  });
  // update caret
  const hdr = list.querySelector(`.chanGroupHeader[data-group="${CSS.escape(k)}"]`);
  if(hdr) {
    const c = hdr.querySelector('.chanGroupCaret');
    if(c) c.textContent = COLLAPSED_GROUPS.has(k) ? '▶' : '▼';
  }
}

// ---- Recent folders ----
function loadRecentFolders() {
  try {
    const raw = localStorage.getItem(LS_RECENT_FOLDERS_KEY);
    const arr = raw ? _safeJsonParse(raw, []) : [];
    if(Array.isArray(arr)) RECENT_FOLDERS = arr.map(x=>String(x)).filter(Boolean);
    else RECENT_FOLDERS = [];
  } catch(e) { RECENT_FOLDERS = []; }
}

function saveRecentFolders() {
  try { localStorage.setItem(LS_RECENT_FOLDERS_KEY, JSON.stringify(RECENT_FOLDERS || [])); } catch(e) {}
}

function refreshRecentFoldersUI() {
  const sel = el('recentFolders');
  if(!sel) return;
  sel.innerHTML = '';

  const opt0 = document.createElement('option');
  opt0.value = '';
  opt0.textContent = (RECENT_FOLDERS && RECENT_FOLDERS.length) ? '— недавние —' : '— нет недавних —';
  sel.appendChild(opt0);

  (RECENT_FOLDERS || []).forEach(path => {
    const o = document.createElement('option');
    o.value = path;
    o.textContent = path;
    sel.appendChild(o);
  });
}

function addRecentFolder(path) {
  const pth = String(path || '').trim();
  if(!pth) return;
  const arr = (RECENT_FOLDERS || []).filter(x => x && x !== pth);
  arr.unshift(pth);
  RECENT_FOLDERS = arr.slice(0, RECENT_MAX);
  saveRecentFolders();
  refreshRecentFoldersUI();
}

function clearRecentFolders() {
  RECENT_FOLDERS = [];
  saveRecentFolders();
  refreshRecentFoldersUI();
  toast('Недавние очищены', 'Список недавних папок очищен', 'ok');
}

// ---- Refrigerant (template cell B1) ----
const ALLOWED_REFRIGERANTS = ['R290', 'R600a'];

function getRefrigerant() {
  try {
    const sel = el('refrigerant');
    let v = sel ? String(sel.value || '').trim() : '';
    if(!ALLOWED_REFRIGERANTS.includes(v)) {
      const saved = localStorage.getItem(LS_REFRIGERANT_KEY);
      if(ALLOWED_REFRIGERANTS.includes(saved)) v = saved;
    }
    if(!ALLOWED_REFRIGERANTS.includes(v)) v = 'R290';
    return v;
  } catch(e) {
    return 'R290';
  }
}

function initRefrigerantUI() {
  const sel = el('refrigerant');
  if(!sel) return;
  try {
    const saved = localStorage.getItem(LS_REFRIGERANT_KEY);
    if(ALLOWED_REFRIGERANTS.includes(saved)) sel.value = saved;
  } catch(e) {}

  // гарантируем значение по умолчанию
  if(!ALLOWED_REFRIGERANTS.includes(String(sel.value || '').trim())) sel.value = 'R290';

  sel.addEventListener('change', () => {
    try { localStorage.setItem(LS_REFRIGERANT_KEY, getRefrigerant()); } catch(e) {}
  });
}

async function copyTextToClipboard(txt) {
  try {
    if(navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(txt);
      return true;
    }
  } catch(e) {}
  // fallback
  try {
    const ta = document.createElement('textarea');
    ta.value = txt;
    ta.style.position = 'fixed';
    ta.style.left = '-2000px';
    ta.style.top = '0';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    const ok = document.execCommand('copy');
    ta.remove();
    return !!ok;
  } catch(e) {}
  return false;
}

async function copyFolderPath() {
  const fld = el('folder');
  const v = fld ? String(fld.value || '').trim() : '';
  if(!v) { toast('Путь пуст', 'Сначала укажи папку', 'warn'); return; }
  const ok = await copyTextToClipboard(v);
  if(ok) toast('Скопировано', 'Путь скопирован в буфер обмена', 'ok');
  else toast('Не удалось', 'Не получилось скопировать путь', 'warn');
}

// ---- Range stats + export info ----
function getEffectiveRange() {
  if(!SUMMARY) return null;
  const vr = getVisibleRangeFromPlot();
  if(vr && vr.length === 2) return [Math.min(vr[0], vr[1]), Math.max(vr[0], vr[1])];
  if(currentRange && currentRange.length === 2) return [Math.min(currentRange[0], currentRange[1]), Math.max(currentRange[0], currentRange[1])];
  return [SUMMARY.start_ms, SUMMARY.end_ms];
}

function estimatePointsForRange(r) {
  if(!SUMMARY || !SUMMARY.points) return 1;
  const totalDur = Math.max(1, (SUMMARY.end_ms - SUMMARY.start_ms));
  const rangeDur = Math.max(1, (r[1] - r[0]));
  return Math.max(1, Math.round(SUMMARY.points * (rangeDur / totalDur)));
}

function requestRangeStatsDebounced() {
  if(!_rangeStatsTimer) {
    // no-op
  } else {
    try { clearTimeout(_rangeStatsTimer); } catch(e) {}
  }
  const r = getEffectiveRange();
  if(!r) return;
  _rangeStatsTimer = setTimeout(() => _fetchRangeStatsNow(r), 250);
}

function _fetchRangeStatsNow(r) {
  if(!LOADED || !SUMMARY) return;
  const token = ++_rangeStatsToken;
  const qs = new URLSearchParams();
  qs.set('start_ms', String(r[0]));
  qs.set('end_ms', String(r[1]));
  fetch('/api/range_stats?' + qs.toString())
    .then(resp => resp.json())
    .then(j => {
      if(token != _rangeStatsToken) return;
      if(j && j.ok) {
        RANGE_STATS = { start_ms: j.start_ms, end_ms: j.end_ms, points: j.points, total: j.total };
      }
    })
    .catch(_=>{})
    .finally(() => {
      updateStepUI();
      updateExportInfo();
    });
}

function updateExportInfo() {
  const box = el('exportInfo');
  if(!box || !SUMMARY) return;

  const codes = getSelectedCodes();
  const range = getEffectiveRange() || [SUMMARY.start_ms, SUMMARY.end_ms];

  let pts = null;
  if(RANGE_STATS && RANGE_STATS.start_ms === range[0] && RANGE_STATS.end_ms === range[1]) pts = RANGE_STATS.points;
  if(pts == null) pts = estimatePointsForRange(range);

  const auto = !!(el('stepAuto') && el('stepAuto').checked);
  const step = auto ? computeAutoStep() : (parseInt(el('step')?.value || '1', 10) || 1);
  const after = Math.max(1, Math.ceil(pts / Math.max(1, step)));

  const extra = !!(el('tplExtra') && el('tplExtra').checked);

  // Сделаем компактно и читабельно
  box.innerHTML =
    `Каналы: <b>${codes.length}</b> &nbsp;•&nbsp; ` +
    `Точек в диапазоне: <b>${pts}</b> &nbsp;•&nbsp; ` +
    `После шага: <b>${after}</b> &nbsp;•&nbsp; ` +
    `Шаг: <b>${Math.max(1, step)}</b> <span class="mini">(${auto ? 'авто' : 'вручную'})</span> &nbsp;•&nbsp; ` +
    `Z: <b>${extra ? 'да' : 'нет'}</b>`;
}

function onRangeChanged() {
  requestRangeStatsDebounced();
  updateExportInfo();
}

function clearFilters() {
  FILTER_TEXT = '';
  FILTER_ONLY_SELECTED = false;
  FILTER_CHIP = null;
  const si = el('chanSearch');
  if(si) si.value = '';
  const cb = el('onlySelected');
  if(cb) cb.checked = false;
  document.querySelectorAll('.chip.active').forEach(b => b.classList.remove('active'));
  applySortAndRender(true, false);
}

// Auto-step helpers
function getStepTarget() {
  const sel = el('stepTarget');
  const v = sel ? parseInt(sel.value || '5000', 10) : 5000;
  return (isFinite(v) && v > 0) ? v : 5000;
}

function computeAutoStep() {
  if(!SUMMARY || !SUMMARY.points) return 1;
  const target = getStepTarget();

  const r = getEffectiveRange() || [SUMMARY.start_ms, SUMMARY.end_ms];

  let pts = null;
  if(RANGE_STATS && RANGE_STATS.start_ms === r[0] && RANGE_STATS.end_ms === r[1]) pts = RANGE_STATS.points;
  if(pts == null) pts = estimatePointsForRange(r);

  const step = Math.ceil(Math.max(1, pts) / target);
  return Math.max(1, step);
}

function updateStepUI() {
  const cb = el('stepAuto');
  const inp = el('step');
  const info = el('stepAutoInfo');
  if(!cb || !inp) return;
  const auto = cb.checked;
  inp.disabled = auto;

  if(info) {
    if(!auto) {
      info.textContent = '';
    } else {
      const st = computeAutoStep();
      const r = getEffectiveRange() || (SUMMARY ? [SUMMARY.start_ms, SUMMARY.end_ms] : null);
      let pts = null;
      if(r && RANGE_STATS && RANGE_STATS.start_ms === r[0] && RANGE_STATS.end_ms === r[1]) pts = RANGE_STATS.points;
      if(r && pts == null) pts = estimatePointsForRange(r);

      let estAfter = '';
      if(pts != null) {
        const left = Math.max(1, Math.ceil(Math.max(1, pts) / Math.max(1, st)));
        const exact = (RANGE_STATS && r && RANGE_STATS.start_ms === r[0] && RANGE_STATS.end_ms === r[1]);
        estAfter = `, ${left} точек${exact ? '' : ' ~'}`;
      }
      info.textContent = `(шаг: ${st}${estAfter})`;
    }
  }

  updateExportInfo();
}

function getShowLegend() {
  const cb = el('showLegend');
  return cb ? !!cb.checked : true;
}


// Priority order (used only when sort mode = "priority")
const PRIORITY = [
  "A-Pc","A-Pe","UR-sie","T-sie","A-Tc","A-Te",
  "A-T1","A-T2","A-T3","A-T4","A-T5","A-T6","A-T7",
  "A-I","A-F","A-V","A-W",
];

function fmtRange(ms) {
  if(!ms || ms.length !== 2) return "—";
  const a = new Date(ms[0]);
  const b = new Date(ms[1]);
  const pad = (n) => (n < 10 ? "0"+n : ""+n);
  const fmt = (d) => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  const durS = Math.max(0, Math.round((ms[1]-ms[0])/1000));
  const h = Math.floor(durS/3600), m = Math.floor((durS%3600)/60), s = durS%60;
  return `${fmt(a)}  →  ${fmt(b)}  (${h}h ${m}m ${s}s)`;
}

function updateRangeText() {
  const r = (currentRange && currentRange.length===2)
    ? [Math.min(currentRange[0], currentRange[1]), Math.max(currentRange[0], currentRange[1])]
    : null;
  el("rangeText").textContent = r ? fmtRange(r) : "—";
  onRangeChanged();
}


function pickFolder() {
  fetch("/api/pick_folder", {method:"POST"})
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) { log("Pick folder error: " + j.error); return; }
      if(j.folder) el("folder").value = j.folder;
    })
    .catch(e=>log("pickFolder error: " + e));
}

// Natural sort helper
function tokenizeNatural(s) {
  const re = /(\d+)|(\D+)/g;
  const out = [];
  let m;
  while((m = re.exec(s)) !== null) {
    if(m[1]) out.push({t:"n", v: parseInt(m[1],10)});
    else out.push({t:"s", v: m[2].toLowerCase()});
  }
  return out;
}
function naturalCompare(a, b) {
  const ta = tokenizeNatural(a);
  const tb = tokenizeNatural(b);
  const n = Math.max(ta.length, tb.length);
  for(let i=0;i<n;i++) {
    const xa = ta[i], xb = tb[i];
    if(!xa) return -1;
    if(!xb) return 1;
    if(xa.t === xb.t) {
      if(xa.v < xb.v) return -1;
      if(xa.v > xb.v) return 1;
    } else {
      // letters before numbers
      if(xa.t === "s") return -1;
      return 1;
    }
  }
  return 0;
}

function sortByNaturalCode(arr) {
  return arr.slice().sort((x,y)=>naturalCompare(x.code, y.code));
}
function sortByLabel(arr) {
  return arr.slice().sort((x,y) => (x.label||"").localeCompare((y.label||""), undefined, {sensitivity:"base"}));
}
function sortByUnitThenCode(arr) {
  return arr.slice().sort((x,y) => {
    const ux = x.unit || "";
    const uy = y.unit || "";
    if(ux !== uy) return ux.localeCompare(uy);
    return naturalCompare(x.code, y.code);
  });
}

function applyPriority(arr) {
  const byCode = new Map(arr.map(c => [c.code, c]));
  const used = new Set();
  const out = [];
  PRIORITY.forEach(code => {
    if(byCode.has(code) && !used.has(code)) {
      out.push(byCode.get(code));
      used.add(code);
    }
  });
  const rest = arr.filter(c => !used.has(c.code));
  sortByNaturalCode(rest).forEach(c => out.push(c));
  return out;
}

// "custom saved": first saved order, then remaining channels in FILE order
function applyCustomSaved(arr, savedOrder) {
  const byCode = new Map(arr.map(c => [c.code, c]));
  const used = new Set();
  const out = [];
  (savedOrder||[]).forEach(code => {
    if(byCode.has(code) && !used.has(code)) {
      out.push(byCode.get(code));
      used.add(code);
    }
  });
  arr.forEach(c => {
    if(!used.has(c.code)) out.push(c);
  });
  return out;
}

function getSortMode() {
  const s = el("sortMode");
  return s ? (s.value || "file") : "file";
}
function setSortMode(mode) {
  const s = el("sortMode");
  if(s) s.value = mode;
}

function buildViewOrder() {
  const mode = getSortMode();
  if(mode === "file") return CHANNELS_FILE.slice();
  if(mode === "custom") return applyCustomSaved(CHANNELS_FILE, SAVED_ORDER);
  if(mode === "priority") return applyPriority(CHANNELS_FILE);
  if(mode === "natural") return sortByNaturalCode(CHANNELS_FILE);
  if(mode === "label") return sortByLabel(CHANNELS_FILE);
  if(mode === "unit") return sortByUnitThenCode(CHANNELS_FILE);
  return CHANNELS_FILE.slice();
}

function currentOrderCodes() {
  const list = el("channelList");
  const items = list.querySelectorAll(".chanItem[data-code]");
  return Array.from(items).map(li => li.getAttribute("data-code"));
}

function getSelectedCodes() {
  // Preserve order for plotting/export.
  // Important: фильтры списка каналов не должны "выкидывать" выбранные каналы из графика.
  // Поэтому порядок берём из полного (отсортированного) списка, а не из видимого DOM.
  const order = (CHANNELS_VIEW_ALL && CHANNELS_VIEW_ALL.length)
    ? CHANNELS_VIEW_ALL.map(c => c.code)
    : currentOrderCodes();
  return order.filter(code => selected.has(code));
}

function setSelection(codes) {
  selected = new Set(codes);
  refreshSelectionUI();
}

function refreshSelectionUI() {
  const list = el("channelList");
  const items = list.querySelectorAll(".chanItem[data-code]");
  items.forEach(li => {
    const code = li.getAttribute("data-code");
    if(selected.has(code)) li.classList.add("selected");
    else li.classList.remove("selected");
    li.setAttribute("aria-selected", selected.has(code) ? "true" : "false");
  });
  updateSelectedCountUI();
  updateExportInfo();
}

function selectAll() {
  setSelection(CHANNELS_VIEW.map(c => c.code));
  scheduleRedraw();
}
function clearAll() {
  // По требованию UI: «Снять все» оставляет выбранным один канал — самый первый в списке.
  let first = null;
  try {
    const list = el('channelList');
    if(list) {
      const li = list.querySelector('.chanItem[data-code]');
      if(li) first = li.getAttribute('data-code');
    }
  } catch(e) {}

  if(!first && CHANNELS_VIEW_ALL && CHANNELS_VIEW_ALL.length) first = CHANNELS_VIEW_ALL[0].code;
  if(!first && CHANNELS_VIEW && CHANNELS_VIEW.length) first = CHANNELS_VIEW[0].code;

  if(first) {
    setSelection([first]);
    anchorCode = first;
  } else {
    setSelection([]);
    anchorCode = null;
  }
  scheduleRedraw();
}

function indexOfCode(code) {
  const order = currentOrderCodes();
  return order.indexOf(code);
}

function onItemClick(ev, code) {
  const isCtrl = ev.ctrlKey || ev.metaKey;
  const isShift = ev.shiftKey;

  const list = el("channelList");
  const items = Array.from(list.querySelectorAll(".chanItem[data-code]"));
  const curIndex = items.findIndex(li => li.getAttribute("data-code") === code);

  if(isShift && anchorCode !== null) {
    const aIdx = indexOfCode(anchorCode);
    const bIdx = curIndex;
    if(aIdx !== -1 && bIdx !== -1) {
      const a = Math.min(aIdx, bIdx);
      const b = Math.max(aIdx, bIdx);
      const rangeCodes = items.slice(a, b+1).map(li => li.getAttribute("data-code"));
      if(!isCtrl) selected.clear();
      rangeCodes.forEach(c => selected.add(c));
    } else {
      if(!isCtrl) selected.clear();
      selected.add(code);
    }
  } else if(isCtrl) {
    if(selected.has(code)) selected.delete(code);
    else selected.add(code);
    anchorCode = code;
  } else {
    selected.clear();
    selected.add(code);
    anchorCode = code;
  }

  refreshSelectionUI();
  scheduleRedraw();
}

function onItemDblClick(ev, code) {
  // "solo": keep only this channel
  selected.clear();
  selected.add(code);
  anchorCode = code;
  refreshSelectionUI();
  drawPlot(); // immediate
}

