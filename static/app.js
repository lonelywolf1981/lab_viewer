
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
function msToPlotX(ms) {
  const offMin = new Date(ms).getTimezoneOffset(); // minutes (UTC - local)
  return ms - offMin * 60000;
}
function plotXToMs(plotMs) {
  const offMin = new Date(plotMs).getTimezoneOffset();
  return plotMs + offMin * 60000;
}
// LeMuRe Viewer UI (selection like <select multiple> + drag reorder + export by visible X range)
let LOADED = false;

let CHANNELS_FILE = [];       // channels in the order they appear in file
let CHANNELS_VIEW = [];       // currently shown order in list
let CHANNELS_VIEW_ALL = [];   // full order (sorted) ignoring filters; used for plotting/export order
let SAVED_ORDER = [];         // saved custom order (from channel_order.json) - not auto-applied
let SUMMARY = null;
let currentRange = null;      // [start_ms, end_ms]

let VIEWER_SETTINGS = null;  // server-side settings for template styling

let redrawTimer = null;
let saveOrderTimer = null;

let selected = new Set();     // selected channel codes (like <select multiple>)
let anchorCode = null;        // last clicked code (for Shift selection)

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

function _setOverlayVisible(vis, text) {
  const ov = el('busyOverlay');
  const tx = el('busyText');
  if(!ov) return;
  if(tx && text) tx.textContent = text;
  ov.classList.toggle('hidden', !vis);
  ov.setAttribute('aria-hidden', vis ? 'false' : 'true');
}

function beginBusy(text) {
  BUSY_GUARD++;
  const my = BUSY_GUARD;
  let shown = false;
  if(_busyTimer) clearTimeout(_busyTimer);
  _busyTimer = setTimeout(() => {
    if(my !== BUSY_GUARD) return;
    _setOverlayVisible(true, text || 'Работаю…');
    shown = true;
  }, 250);
  return () => {
    if(my !== BUSY_GUARD) return;
    if(_busyTimer) clearTimeout(_busyTimer);
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
  try {
    const raw = localStorage.getItem(LS_VIEW_OPTS_KEY);
    const st = raw ? _safeJsonParse(raw, null) : null;
    const compact = !!(st && st.compact);
    const groups  = !!(st && st.groups);
    const collapsed = (st && st.collapsed_groups && typeof st.collapsed_groups === 'object') ? st.collapsed_groups : {};

    VIEW_COMPACT = compact;
    VIEW_GROUPS  = groups;
    COLLAPSED_GROUPS = new Set(Object.keys(collapsed).filter(k => collapsed[k]));

    if(el('viewCompact')) el('viewCompact').checked = VIEW_COMPACT;
    if(el('viewGroups')) el('viewGroups').checked = VIEW_GROUPS;
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
  setSelection([]);
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

function renderChannelList(arr, keepSelection=true, pruneSelection=true) {
  const list = el("channelList");
  if(!list) return;
  list.innerHTML = "";
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-multiselectable", "true");

  applyViewClasses();

  const canDrag = !hasActiveFilter() && !VIEW_GROUPS;

  // Keep selection intersection
  const prev = new Set(selected);

  CHANNELS_VIEW = arr.slice();

  // group counts (for headers)
  const groupCounts = {};
  if(VIEW_GROUPS) {
    arr.forEach(ch => {
      const g = groupOfCode(ch.code);
      groupCounts[g] = (groupCounts[g] || 0) + 1;
    });
  }

  let lastGroup = null;

  arr.forEach((ch) => {
    const grp = groupOfCode(ch.code);

    if(VIEW_GROUPS && grp !== lastGroup) {
      lastGroup = grp;
      const hdr = document.createElement('li');
      hdr.className = 'chanGroupHeader';
      hdr.setAttribute('data-group', grp);

      const caret = document.createElement('div');
      caret.className = 'chanGroupCaret';
      caret.textContent = COLLAPSED_GROUPS.has(grp) ? '▶' : '▼';
      hdr.appendChild(caret);

      const title = document.createElement('div');
      title.className = 'chanGroupTitle';
      title.textContent = grp;
      hdr.appendChild(title);

      const cnt = document.createElement('div');
      cnt.className = 'chanGroupCount';
      cnt.textContent = String(groupCounts[grp] || 0);
      hdr.appendChild(cnt);

      hdr.addEventListener('click', () => toggleGroupCollapsed(grp));

      list.appendChild(hdr);
    }

    const li = document.createElement("li");
    li.className = "chanItem";
    li.setAttribute("draggable", "false");
    li.setAttribute("data-code", ch.code);
    li.setAttribute("data-group", grp);
    li.setAttribute("role", "option");

    if(VIEW_GROUPS && COLLAPSED_GROUPS.has(grp)) li.classList.add('groupHidden');

    // Drag handle: draggable only here
    const handle = document.createElement("div");
    handle.className = "dragHandle";
    handle.textContent = "≡";
    if(canDrag) {
      handle.title = "Перетащить";
      handle.setAttribute("draggable", "true");
    } else {
      const why = VIEW_GROUPS ? 'Перетаскивание недоступно в режиме «группы»' : 'Перетаскивание недоступно при включённых фильтрах';
      handle.title = why;
      handle.setAttribute("draggable", "false");
      handle.classList.add("dragDisabled");
    }
    // prevent selection clicks on handle
    handle.addEventListener("click", (ev) => ev.stopPropagation());
    handle.addEventListener("dblclick", (ev) => ev.stopPropagation());
    li.appendChild(handle);

    const code = document.createElement("div");
    code.className = "chanCode";
    code.textContent = ch.code;
    li.appendChild(code);

    const lab = document.createElement("div");
    lab.className = "chanLabel";
    lab.textContent = ch.label;
    li.appendChild(lab);

    if(ch.unit) {
      const ub = document.createElement('div');
      ub.className = 'chanUnitBadge';
      ub.textContent = ch.unit;
      li.appendChild(ub);
    }

    // Selection clicks on row
    li.addEventListener("click", (ev) => onItemClick(ev, ch.code));
    li.addEventListener("dblclick", (ev) => onItemDblClick(ev, ch.code));

    // Drag start/end on handle
    handle.addEventListener("dragstart", (ev) => {
      if(!canDrag) {
        ev.preventDefault();
        if(VIEW_GROUPS) toast('Перетаскивание отключено', 'Отключите режим «группы» и попробуйте снова', 'warn', 5000);
        else toast('Перетаскивание отключено', 'Снимите фильтры (поиск/чипы/только выбранные) и попробуйте снова', 'warn', 5000);
        return;
      }
      li.classList.add("dragging");
      ev.dataTransfer.setData("text/plain", ch.code);
      ev.dataTransfer.effectAllowed = "move";
      if(ev.dataTransfer.setDragImage) {
        try { ev.dataTransfer.setDragImage(li, 20, 12); } catch(_) {}
      }
    });
    handle.addEventListener("dragend", () => {
      li.classList.remove("dragging");
      list.querySelectorAll(".chanItem").forEach(x => x.classList.remove("dragOver"));
    });

    // Drag over/drop on row
    li.addEventListener("dragover", (ev) => {
      ev.preventDefault();
      li.classList.add("dragOver");
      ev.dataTransfer.dropEffect = "move";
    });
    li.addEventListener("dragleave", () => li.classList.remove("dragOver"));
    li.addEventListener("drop", (ev) => {
      ev.preventDefault();
      li.classList.remove("dragOver");
      const draggedCode = ev.dataTransfer.getData("text/plain");
      if(!draggedCode || draggedCode === ch.code) return;
      moveCodeBefore(draggedCode, ch.code);
    });

    list.appendChild(li);
  });

  // Restore selection (optionally prune to visible list)
  if(keepSelection && prev.size) {
    if(pruneSelection) {
      const available = new Set(arr.map(c => c.code));
      const inter = Array.from(prev).filter(c => available.has(c));
      if(inter.length) selected = new Set(inter);
    } else {
      selected = new Set(prev);
    }
  }

  // If nothing selected, select first few temps
  if(selected.size === 0 && arr.length) {
    const temps = arr.filter(c => c.unit === "°C").slice(0, 6).map(c => c.code);
    if(temps.length) selected = new Set(temps);
    else selected = new Set(arr.slice(0,4).map(c => c.code));
    anchorCode = arr[0].code;
  }

  refreshSelectionUI();
}

function moveCodeBefore(draggedCode, targetCode) {
  const list = el("channelList");
  const items = Array.from(list.querySelectorAll(".chanItem"));
  const dragged = items.find(li => li.getAttribute("data-code") === draggedCode);
  const target  = items.find(li => li.getAttribute("data-code") === targetCode);
  if(!dragged || !target) return;

  list.insertBefore(dragged, target);
  refreshSelectionUI();

  // Dragging means "custom" order now
  setSortMode("custom");
  scheduleSaveOrder();
  scheduleSaveLastState();
}

function scheduleSaveOrder() {
  if(saveOrderTimer) clearTimeout(saveOrderTimer);
  saveOrderTimer = setTimeout(() => {
    const order = currentOrderCodes();
    fetch("/api/save_order", {
      method:"POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({order})
    }).then(r=>r.json())
      .then(j=>{
        if(j.ok) {
          SAVED_ORDER = order.slice();
          log("Order saved (" + j.saved + " items)");
        } else log("Order save error: " + (j.error||""));
      })
      .catch(e=>log("Order save error: " + e));
  }, 350);
}



// ===== Named saved orders =====
let NAMED_ORDERS = []; // [{key,name,count,saved_at}]

function refreshNamedOrders(selectKey=null) {
  fetch('/api/orders_list')
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) { log('orders_list error: ' + (j.error||'')); return; }
      NAMED_ORDERS = j.orders || [];
      const sel = el('orderSelect');
      if(!sel) return;
      sel.innerHTML = '';
      const opt0 = document.createElement('option');
      opt0.value = '';
      opt0.textContent = NAMED_ORDERS.length ? '— выбери сохранённый порядок —' : '— нет сохранённых —';
      sel.appendChild(opt0);

      NAMED_ORDERS.forEach(o=>{
        const opt = document.createElement('option');
        opt.value = o.key;
        const meta = o.count ? ` (${o.count})` : '';
        opt.textContent = (o.name || o.key) + meta;
        if(o.saved_at) opt.title = 'Сохранено: ' + o.saved_at;
        sel.appendChild(opt);
      });

      if(selectKey) sel.value = selectKey;
    })
    .catch(e=>log('orders_list error: ' + e));
}

function saveNamedOrder() {
  if(!CHANNELS_FILE || CHANNELS_FILE.length === 0) {
    toast('Нет данных', 'Сначала загрузите тест, чтобы были каналы', 'warn');
    return;
  }
  const name = (el('orderName') ? el('orderName').value : '').trim();
  if(!name) {
    toast('Имя не указано', 'Введите имя порядка', 'warn');
    return;
  }
  const order = currentOrderCodes();
  fetch('/api/orders_save', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({name, order})
  })
  .then(r=>r.json())
  .then(j=>{
    if(!j.ok) { toast('Ошибка сохранения', j.error || 'Не удалось сохранить', 'err', 6000); log('orders_save error: ' + (j.error||'')); return; }
    toast('Порядок сохранён', `"${name}"`, 'ok');
    log(`Named order saved: "${name}" (key=${j.key}, items=${j.saved})`);
    refreshNamedOrders(j.key);
    // also update "custom" saved order to the current list order
    SAVED_ORDER = order.slice();
    setSortMode('custom');
    // persist as last custom
    fetch('/api/save_order', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({order})
    }).catch(()=>{});
  })
  .catch(e=>{ toast('Ошибка сохранения', (e && e.message) ? e.message : String(e), 'err', 6000); log('orders_save error: ' + e); });
}

function loadNamedOrder() {
  if(!CHANNELS_FILE || CHANNELS_FILE.length === 0) {
    toast('Нет данных', 'Сначала загрузите тест, чтобы были каналы', 'warn');
    return;
  }
  const key = el('orderSelect') ? el('orderSelect').value : '';
  if(!key) {
    toast('Не выбран порядок', 'Выберите сохранённый порядок из списка', 'warn');
    return;
  }
  fetch('/api/orders_load?key=' + encodeURIComponent(key))
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) { toast('Ошибка загрузки', j.error || 'Не удалось загрузить', 'err', 6000); log('orders_load error: ' + (j.error||'')); return; }
      const order = j.order || [];
      SAVED_ORDER = order.slice();
      setSortMode('custom');
      applySortAndRender(true, true);
      toast('Порядок загружен', `"${j.name || key}"`, 'ok');
      // persist as last custom so it remembers on next start
      fetch('/api/save_order', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({order})
      }).then(()=>{
        log(`Named order loaded: "${j.name || key}" (items=${order.length})`);
      }).catch(()=>{
        log(`Named order loaded: "${j.name || key}" (items=${order.length})`);
      });
    })
    .catch(e=>{ toast('Ошибка загрузки', (e && e.message) ? e.message : String(e), 'err', 6000); log('orders_load error: ' + e); });
}

// ===== Presets (наборы каналов + настройки) =====
let PRESETS = []; // [{key,name,count,saved_at}]

function refreshPresets(selectKey=null) {
  fetch('/api/presets_list')
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) { log('presets_list error: ' + (j.error||'')); return; }
      PRESETS = j.presets || [];
      const sel = el('presetSelect');
      if(!sel) return;
      sel.innerHTML = '';
      const opt0 = document.createElement('option');
      opt0.value = '';
      opt0.textContent = PRESETS.length ? '— выбери набор —' : '— нет сохранённых наборов —';
      sel.appendChild(opt0);

      PRESETS.forEach(p=>{
        const opt = document.createElement('option');
        opt.value = p.key;
        const meta = p.count ? ` (${p.count})` : '';
        opt.textContent = (p.name || p.key) + meta;
        if(p.saved_at) opt.title = 'Сохранено: ' + p.saved_at;
        sel.appendChild(opt);
      });

      if(selectKey) sel.value = selectKey;
    })
    .catch(e=>log('presets_list error: ' + e));
}

function _collectPresetPayload() {
  // Важно: сохраняем полный порядок (без учёта фильтров списка), иначе можно «урезать» каналы.
  const order = (CHANNELS_VIEW_ALL && CHANNELS_VIEW_ALL.length)
    ? CHANNELS_VIEW_ALL.map(c => c.code)
    : currentOrderCodes();

  const stepAuto = el('stepAuto') ? !!el('stepAuto').checked : true;
  const stepTarget = el('stepTarget') ? parseInt(el('stepTarget').value || '5000', 10) : 5000;
  const stepManual = el('step') ? parseInt(el('step').value || '1', 10) : 1;
  const showLegend = el('showLegend') ? !!el('showLegend').checked : true;

  return {
    channels: getSelectedCodes(),
    sort_mode: getSortMode(),
    order: order,
    step_auto: stepAuto,
    step_target: (isFinite(stepTarget) && stepTarget > 0) ? stepTarget : 5000,
    step: (isFinite(stepManual) && stepManual > 0) ? stepManual : 1,
    show_legend: showLegend,
  };
}

function savePreset() {
  if(!CHANNELS_FILE || CHANNELS_FILE.length === 0) {
    toast('Нет данных', 'Сначала загрузите тест, чтобы были каналы', 'warn');
    return;
  }

  let name = (el('presetName') ? el('presetName').value : '').trim();
  const selKey = el('presetSelect') ? el('presetSelect').value : '';
  if(!name && selKey) {
    // если имя не введено — обновим выбранный набор
    const opt = el('presetSelect').selectedOptions && el('presetSelect').selectedOptions[0];
    if(opt) name = (opt.textContent || '').replace(/\s*\(\d+\)\s*$/,'').trim();
  }
  if(!name) {
    toast('Имя не указано', 'Введите имя набора или выберите существующий', 'warn');
    return;
  }

  const payload = _collectPresetPayload();

  fetch('/api/presets_save', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({name, preset: payload})
  })
  .then(r=>r.json())
  .then(j=>{
    if(!j.ok) {
      toast('Ошибка сохранения', j.error || 'Не удалось сохранить набор', 'err', 6000);
      log('presets_save error: ' + (j.error||''));
      return;
    }
    toast('Набор сохранён', `"${name}"`, 'ok');
    log(`Preset saved: "${name}" (key=${j.key}, channels=${j.count})`);
    refreshPresets(j.key);
    if(el('presetName')) el('presetName').value = name;
  })
  .catch(e=>{
    toast('Ошибка сохранения', (e && e.message) ? e.message : String(e), 'err', 6000);
    log('presets_save error: ' + e);
  });
}

function _applyPresetObject(preset) {
  if(!preset) return;

  // 1) настройки UI
  if(el('showLegend') && typeof preset.show_legend === 'boolean') el('showLegend').checked = preset.show_legend;
  if(el('stepAuto') && typeof preset.step_auto === 'boolean') el('stepAuto').checked = preset.step_auto;
  if(el('stepTarget') && preset.step_target) el('stepTarget').value = String(parseInt(preset.step_target, 10));
  if(el('step') && preset.step) el('step').value = String(parseInt(preset.step, 10));
  updateStepUI();

  // 2) порядок/сортировка
  if(Array.isArray(preset.order) && preset.order.length) {
    SAVED_ORDER = preset.order.slice();
    // Также сохраним как "Мой порядок", чтобы было единообразно (drag + custom).
    fetch('/api/save_order', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({order: SAVED_ORDER})
    }).catch(()=>{});
  }

  if(preset.sort_mode) setSortMode(preset.sort_mode);

  // 3) перерисуем список каналов, затем применим выбор
  applySortAndRender(false, false);

  const available = new Set((CHANNELS_FILE||[]).map(c => c.code));
  const sel = (preset.channels || []).map(c => String(c)).filter(c => available.has(c));
  if(sel.length) {
    setSelection(sel);
    anchorCode = sel[0];
  }

  updateSelectedCountUI();
  scheduleSaveLastState();
  scheduleRedraw();
}

function loadPreset() {
  if(!CHANNELS_FILE || CHANNELS_FILE.length === 0) {
    toast('Нет данных', 'Сначала загрузите тест, чтобы были каналы', 'warn');
    return;
  }
  const key = el('presetSelect') ? el('presetSelect').value : '';
  if(!key) {
    toast('Не выбран набор', 'Выберите набор из списка', 'warn');
    return;
  }

  fetch('/api/presets_load?key=' + encodeURIComponent(key))
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) {
        toast('Ошибка загрузки', j.error || 'Не удалось загрузить набор', 'err', 6000);
        log('presets_load error: ' + (j.error||''));
        return;
      }
      if(el('presetName')) el('presetName').value = j.name || key;
      _applyPresetObject(j.preset || j);
      toast('Набор применён', `"${j.name || key}"`, 'ok');
      log(`Preset loaded: "${j.name || key}"`);
    })
    .catch(e=>{
      toast('Ошибка загрузки', (e && e.message) ? e.message : String(e), 'err', 6000);
      log('presets_load error: ' + e);
    });
}

function deletePreset() {
  const key = el('presetSelect') ? el('presetSelect').value : '';
  if(!key) {
    toast('Не выбран набор', 'Выберите набор из списка', 'warn');
    return;
  }
  const name = (el('presetSelect').selectedOptions && el('presetSelect').selectedOptions[0])
    ? el('presetSelect').selectedOptions[0].textContent
    : key;
  if(!confirm(`Удалить набор "${name}"?`)) return;

  fetch('/api/presets_delete', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({key})
  })
  .then(r=>r.json())
  .then(j=>{
    if(!j.ok) {
      toast('Ошибка удаления', j.error || 'Не удалось удалить набор', 'err', 6000);
      log('presets_delete error: ' + (j.error||''));
      return;
    }
    toast('Набор удалён', name, 'ok');
    if(el('presetName')) el('presetName').value = '';
    refreshPresets();
  })
  .catch(e=>{
    toast('Ошибка удаления', (e && e.message) ? e.message : String(e), 'err', 6000);
    log('presets_delete error: ' + e);
  });
}

function applySortAndRender(keepSelection=true, doRedraw=true) {
  syncViewOptsFromUI(false);
  applyViewClasses();

  // Сначала строим полный порядок (сортировка), потом применяем фильтры только для отображения.
  const base = buildViewOrder();
  CHANNELS_VIEW_ALL = base.slice();

  const filtered = applyChannelFilters(base);
  // Если включён фильтр, НЕ обрезаем selection до видимых элементов
  const prune = !hasActiveFilter();
  renderChannelList(filtered, keepSelection, prune);

  // Фильтрация списка не должна пересчитывать график. Перерисовка только при изменении выбора/порядка.
  if(doRedraw) scheduleRedraw();
  updateSelectedCountUI();
}

function loadTest() {
  const folder = el("folder").value.trim();
  if(!folder) { toast('Папка не указана', 'Укажи папку с тестом', 'warn'); return; }

  const endBusy = beginBusy('Загружаю тест…');
  const btn = el('btnLoad');
  const btnPick = el('btnPick');
  if(btn) btn.disabled = true;
  if(btnPick) btnPick.disabled = true;

  fetch("/api/load", {
    method:"POST",
    headers: {"Content-Type":"application/json"},
    body: JSON.stringify({folder})
  })
  .then(r=>r.json())
  .then(j=>{
    if(!j.ok) {
      toast('Ошибка загрузки', j.error || 'Не удалось загрузить тест', 'err', 6000);
      log("load error: " + j.error);
      return;
    }
    LOADED = true;

    RANGE_STATS = null;
    _rangeStatsToken++;

    CHANNELS_FILE = j.channels || [];
    SAVED_ORDER = j.saved_order || [];

    SUMMARY = j.summary || null;
    if(SUMMARY) {
      el("summary").innerHTML =
        `Точек: <b>${SUMMARY.points}</b> | ` +
        `Начало: <b>${SUMMARY.start}</b> | ` +
        `Конец: <b>${SUMMARY.end}</b>`;
      currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
      updateRangeText();
    } else {
      el("summary").textContent = "Нет данных";
    }

    // Сбрасываем фильтры на новом тесте, чтобы список не оказался пустым
    clearFilters();

    // Requirement: no sorting on load -> "file" mode
    setSortMode("file");
    applySortAndRender(false, false);
    updateStepUI();

    // Попробуем восстановить последний выбор (localStorage)
    applyLastStateIfAny();

    // Stage 3: недавние папки
    addRecentFolder(j.folder || folder);
    updateExportInfo();

    log("Loaded: " + (j.folder || folder));
    toast('Тест загружен', `Каналов: ${CHANNELS_FILE.length}, точек: ${SUMMARY ? SUMMARY.points : '—'}`, 'ok');
    drawPlot();
  })
  .catch(e=>{
    toast('Ошибка загрузки', (e && e.message) ? e.message : String(e), 'err', 6000);
    log("loadTest error: " + e);
  })
  .finally(()=>{
    if(btn) btn.disabled = false;
    if(btnPick) btnPick.disabled = false;
    endBusy();
  });
}

function getStep() {
  const cb = el('stepAuto');
  if(cb && cb.checked) return computeAutoStep();
  const v = parseInt(el("step").value || "1", 10);
  return isFinite(v) && v > 0 ? v : 1;
}

function labelFor(code) {
  const c = CHANNELS_FILE.find(x => x.code === code);
  return c ? c.label : code;
}


function drawPlot() {
  if(!LOADED || !SUMMARY) return;
  const codes = getSelectedCodes();
  if(!codes.length) { toast('Каналы не выбраны', 'Выбери хотя бы один канал', 'warn'); return; }

  // обновим подсказку авто-шага под текущий диапазон
  updateStepUI();
  const step = getStep();

  const endBusy = beginBusy('Строю график…');
  const btn = el('btnDraw');
  if(btn) btn.disabled = true;

  // Запоминаем текущий видимый диапазон ДО перерисовки, чтобы он не сбрасывался при смене датчиков.
  // Важно: НЕ переводим в ISO-строки (toISOString), иначе на некоторых ПК появляется смещение по времени/оси.
  // Держим диапазон как [ms, ms] и (если нужно) отдаём Plotly как Date-объекты.
  let desiredRange = null;
  const vrBefore = getVisibleRangeFromPlot();
  if(vrBefore && vrBefore.length === 2) {
    desiredRange = [Math.min(vrBefore[0], vrBefore[1]), Math.max(vrBefore[0], vrBefore[1])];
  } else if(currentRange && currentRange.length === 2) {
    desiredRange = [Math.min(currentRange[0], currentRange[1]), Math.max(currentRange[0], currentRange[1])];
  }

  const start_ms = SUMMARY.start_ms;
  const end_ms = SUMMARY.end_ms;

  const qs = new URLSearchParams();
  qs.set('channels', codes.join(','));
  qs.set('start_ms', String(start_ms));
  qs.set('end_ms', String(end_ms));

  // Если включён "Авто шаг" — просим сервер ограничить количество точек.
  // Это защищает Plotly от зависаний на больших тестах.
  const isAuto = !!(el('stepAuto') && el('stepAuto').checked);
  if(isAuto) {
    qs.set('max_points', String(getStepTarget()));
  } else {
    qs.set('step', String(step));
  }

  fetch(`/api/series?${qs.toString()}`)
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) {
        toast('Ошибка данных', j.error || 'Не удалось получить серию', 'err', 6000);
        log("series error: " + j.error);
        return;
      }
      // IMPORTANT: use local strings instead of Date objects to avoid a fixed timezone offset
      // on some machines/browsers when Plotly formats dates.
      const t = (j.t_ms || []).map((ms) => new Date(msToPlotX(ms)));
      const series = j.series || {};
      const traces = [];

      // Производительность: при большом числе точек лучше использовать WebGL (scattergl)
      const useGL = (t.length > 8000) || (codes.length * t.length > 300000);

      codes.forEach(code => {
        const y = series[code] || [];
        traces.push({
          type: useGL ? "scattergl" : "scatter",
          mode: "lines",
          name: labelFor(code),
          x: t,
          y: y,
          hovertemplate: "%{x|%Y-%m-%d %H:%M:%S}<br>%{y:.2f}<extra></extra>",
        });
      });

      const layout = {
        // Важно: uirevision + Plotly.react сохраняют zoom/диапазон при обновлении данных
        // Привяжем uirevision к конкретному тесту (чтобы при загрузке нового теста zoom сбрасывался,
        // а при смене датчиков внутри одного теста — сохранялся).
        uirevision: `lemure-${SUMMARY.start_ms}-${SUMMARY.end_ms}`,
        showlegend: getShowLegend(),
        margin: {l: 70, r: 20, t: getShowLegend() ? 140 : 40, b: 90},
        hovermode: "x unified",
        xaxis: {
          type: "date",
          tickformat: "%H:%M<br>%d.%m",
          showspikes: true,
          spikemode: "across",
          spikesnap: "cursor",
          // Всегда фиксируем общий диапазон range-slider на весь тест,
          // чтобы при сохранённом zoom он не "прыгал" и визуально не уезжал.
          rangeslider: {
            visible: true,
            autorange: false,
            range: [new Date(msToPlotX(start_ms)), new Date(msToPlotX(end_ms))],
          },
        },
        yaxis: {automargin: true},
        legend: getShowLegend() ? {
          orientation: "h",
          yanchor: "bottom",
          y: 1.22,
          xanchor: "left",
          x: 0,
          font: {size: 10},
          entrywidthmode: "pixels",
          entrywidth: 240,
          itemclick: "toggle",
          itemdoubleclick: "toggleothers",
        } : undefined,
      };

      // Если до перерисовки пользователь уже выбрал диапазон (zoom/range-slider),
      // задаём его прямо в layout (без Plotly.relayout), так Plotly не делает
      // дополнительных пересчётов, из-за которых у некоторых ПК/локалей график
      // визуально "уезжал" вправо.
      if(desiredRange && desiredRange.length === 2) {
        layout.xaxis.autorange = false;
        layout.xaxis.range = [new Date(msToPlotX(desiredRange[0])) , new Date(msToPlotX(desiredRange[1]))];
      }

      const config = {
        responsive: true,
        displaylogo: false,
        scrollZoom: true,
      };

      const plotDiv = el("plot");
      const p = Plotly.react("plot", traces, layout, config);

      Promise.resolve(p).then(() => {
        // Не накапливаем слушатели при каждом redraw
        try {
          if(plotDiv && plotDiv.removeAllListeners) plotDiv.removeAllListeners("plotly_relayout");
        } catch(_) {}

        plotDiv.on("plotly_relayout", (ev) => {
          if(ev && ev["xaxis.autorange"] === true) {
            currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
            updateRangeText();
            return;
          }
          const rr = parseRelayoutRange(ev) || getVisibleRangeFromPlot();
          if(rr && rr.length === 2) {
            currentRange = [Math.min(rr[0], rr[1]), Math.max(rr[0], rr[1])];
            updateRangeText();
          }
        });

        // Если нужно удержать диапазон, делаем это без relayout() (он иногда вызывает визуальный "уезд" вправо).
        // Достаточно задать range прямо в layout до Plotly.react. Если uirevision работает — Plotly сам сохранит.
        const vrAfter = getVisibleRangeFromPlot();
        if(vrAfter && vrAfter.length === 2) {
          currentRange = [Math.min(vrAfter[0], vrAfter[1]), Math.max(vrAfter[0], vrAfter[1])];
        } else if(desiredRange && desiredRange.length === 2) {
          currentRange = desiredRange.slice();
        } else if(!currentRange) {
          currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
        }
        updateRangeText();
      });

      // Текст диапазона обновим сразу (для экспорта)
      if(desiredRange && desiredRange.length === 2) currentRange = desiredRange.slice();
      if(!currentRange) currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
      updateRangeText();

      log(`Plot updated: channels=${codes.length}, step=${(j && typeof j.step === 'number') ? j.step : step}${(el('stepAuto')?.checked ? ' (auto)' : '')}`);
    })
    .catch(e=>{
      toast('Ошибка графика', (e && e.message) ? e.message : String(e), 'err', 6000);
      log("drawPlot error: " + e);
    })
    .finally(()=>{
      if(btn) btn.disabled = false;
      endBusy();
    });
}




function exportData(fmt) {
  if(!LOADED || !SUMMARY) { toast('Нет данных', 'Сначала загрузите тест', 'warn'); return; }

  fmt = (fmt || "csv").toLowerCase();
  if(fmt !== "csv" && fmt !== "xlsx") fmt = "csv";

  const codes = getSelectedCodes();
  if(!codes || !codes.length) {
    toast('Каналы не выбраны', 'Выбери хотя бы один канал', 'warn');
    return;
  }

  // Интервал: сначала из видимой области графика (zoom/range-slider), иначе из currentRange
  let start_ms = SUMMARY.start_ms;
  let end_ms = SUMMARY.end_ms;
  const vr = getVisibleRangeFromPlot();
  if(vr && vr.length === 2) {
    start_ms = Math.min(vr[0], vr[1]);
    end_ms   = Math.max(vr[0], vr[1]);
  } else if(currentRange && currentRange.length === 2) {
    start_ms = Math.min(currentRange[0], currentRange[1]);
    end_ms   = Math.max(currentRange[0], currentRange[1]);
  }

  updateStepUI();
  const step = getStep();

  const endBusy = beginBusy(`Экспортирую ${fmt.toUpperCase()}…`);

  const st  = el("tplStatus");
  const bc = el("btnCsv");
  const bx = el("btnXlsx");
  if(bc) bc.disabled = true;
  if(bx) bx.disabled = true;
  if(st) st.innerHTML = `<span class="spinner"></span>Экспортирую ${fmt.toUpperCase()}… <span style="opacity:.85">(каналы: ${codes.length}, диапазон: ${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()})</span>`;

  const qs = new URLSearchParams();
  qs.set("format", fmt);
  qs.set("start_ms", String(start_ms));
  qs.set("end_ms", String(end_ms));
  qs.set("channels", codes.join(","));
  qs.set("step", String(step));

  fetch("/api/export?" + qs.toString())
    .then(async (r) => {
      if(!r.ok) {
        let msg = "HTTP " + r.status;
        try {
          const ct = r.headers.get("content-type") || "";
          if(ct.includes("application/json")) {
            const j = await r.json();
            if(j && j.error) msg = j.error;
          }
        } catch(e) {}
        throw new Error(msg);
      }
      return r.blob();
    })
    .then((blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = (fmt === "xlsx") ? "export.xlsx" : "export.csv";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(()=>URL.revokeObjectURL(url), 2000);

      if(st) st.textContent = "Готово. Файл скачивается…";
      toast('Экспорт готов', `${fmt.toUpperCase()}: ${codes.length} каналов`, 'ok');
      log(`Export ${fmt.toUpperCase()} OK: channels=${codes.length}, range=${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()}, step=${step}`);
    })
    .catch((e) => {
      const msg = (e && e.message) ? e.message : String(e);
      if(st) st.textContent = "Ошибка: " + msg;
      toast('Ошибка экспорта', msg, 'err', 6000);
      log("Export " + fmt + " error: " + e);
    })
    .finally(() => {
      if(bc) bc.disabled = false;
      if(bx) bx.disabled = false;
      endBusy();
    });
}
function exportTemplate() {
  if(!LOADED || !SUMMARY) { toast('Нет данных', 'Сначала загрузите тест', 'warn'); return; }

  const codes = getSelectedCodes();
  if(!codes || !codes.length) {
    toast('Каналы не выбраны', 'Выбери хотя бы один канал', 'warn');
    return;
  }

  // Интервал: сначала из видимой области графика (zoom/range-slider), иначе из currentRange
  let start_ms = SUMMARY.start_ms;
  let end_ms = SUMMARY.end_ms;
  const vr = getVisibleRangeFromPlot();
  if(vr && vr.length === 2) {
    start_ms = Math.min(vr[0], vr[1]);
    end_ms   = Math.max(vr[0], vr[1]);
  } else if(currentRange && currentRange.length === 2) {
    start_ms = Math.min(currentRange[0], currentRange[1]);
    end_ms   = Math.max(currentRange[0], currentRange[1]);
  }

  updateStepUI();
  const step = getStep();

  const endBusy = beginBusy('Формирую XLSX по шаблону…');

  const btn = el("btnTpl");
  const st  = el("tplStatus");

  const includeExtra = (el('tplExtra') && el('tplExtra').checked) ? 1 : 0;
  if(btn) { btn.disabled = true; btn.textContent = "Готовлю шаблон…"; }
  if(st)  { st.innerHTML = `<span class="spinner"></span>Формирую Excel… <span style="opacity:.85">(каналы: ${codes.length}, диапазон: ${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()}, Z: ${includeExtra ? 'да' : 'нет'})</span>`; }

  const qs = new URLSearchParams();
  qs.set("start_ms", String(start_ms));
  qs.set("end_ms", String(end_ms));
  qs.set("channels", codes.join(","));
  qs.set("step", String(step));
  qs.set('include_extra', String(includeExtra));

  fetch("/api/export_template?" + qs.toString())
    .then(async (r) => {
      if(!r.ok) {
        let msg = "HTTP " + r.status;
        try {
          const ct = r.headers.get("content-type") || "";
          if(ct.includes("application/json")) {
            const j = await r.json();
            if(j && j.error) msg = j.error;
          }
        } catch(e) {}
        throw new Error(msg);
      }
      return r.blob();
    })
    .then((blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "template_filled.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(()=>URL.revokeObjectURL(url), 2000);

      if(st) st.textContent = "Готово. Файл скачивается…";
      toast('Шаблон готов', `${codes.length} каналов`, 'ok');
      log(`Export TEMPLATE OK: channels=${codes.length}, range=${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()}, step=${step}`);
    })
    .catch((e) => {
      const msg = (e && e.message) ? e.message : String(e);
      if(st) st.textContent = "Ошибка: " + msg;
      toast('Ошибка шаблона', msg, 'err', 6000);
      log("Export TEMPLATE error: " + e);
    })
    .finally(() => {
      if(btn) { btn.disabled = false; btn.textContent = "В шаблон XLSX"; }
      endBusy();
    });
}




// ---------------- Export template styling settings (server-side) ----------------
function _num(v, defv) {
  const x = parseFloat(v);
  return Number.isFinite(x) ? x : defv;
}

function _set(id, val) {
  const e = el(id);
  if(!e) return;
  e.value = (val === undefined || val === null) ? '' : String(val);
}

function _setText(id, txt) {
  const e = el(id);
  if(!e) return;
  e.textContent = txt;
}

function _styleStatus(msg, isErr) {
  const e = el('styleStatus');
  if(!e) return;
  e.textContent = msg || '';
  e.style.color = isErr ? '#b00020' : '#333';
}

// ---------------- Color picker label (HEX/RGB) ----------------
// Важно: браузерный pop-up у input[type=color] нельзя заставить всегда показывать HEX.
// Поэтому показываем подпись под цветом и даём переключатель формата.
const COLOR_FMT_KEY = 'lemure_color_format';

function _clamp(v, a, b){ return Math.max(a, Math.min(b, v)); }

function _normHex(s){
  if(!s) return null;
  let t = String(s).trim();
  if(!t) return null;
  if(t[0] !== '#') t = '#' + t;
  const m3 = /^#([0-9a-fA-F]{3})$/.exec(t);
  if(m3){
    const a = m3[1].split('').map(ch => ch + ch).join('');
    return ('#' + a).toUpperCase();
  }
  const m6 = /^#([0-9a-fA-F]{6})$/.exec(t);
  if(m6) return ('#' + m6[1]).toUpperCase();
  return null;
}

function _hexToRgb(hex){
  const h = _normHex(hex);
  if(!h) return null;
  const n = parseInt(h.slice(1), 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function _rgbToHex(r,g,b){
  const rr = _clamp(parseInt(r,10), 0, 255);
  const gg = _clamp(parseInt(g,10), 0, 255);
  const bb = _clamp(parseInt(b,10), 0, 255);
  const to2 = (x) => x.toString(16).padStart(2,'0').toUpperCase();
  return '#' + to2(rr) + to2(gg) + to2(bb);
}

function _getColorFmt(){
  const sel = el('colorFormat');
  return (sel && sel.value) ? sel.value : 'hex';
}

function _formatColorLabel(hex){
  const fmt = _getColorFmt();
  const h = _normHex(hex) || String(hex || '').trim();
  if(fmt === 'rgb'){
    const rgb = _hexToRgb(h);
    if(!rgb) return h;
    return `rgb(${rgb.r}, ${rgb.g}, ${rgb.b})`;
  }
  return h;
}

function _parseColorLabel(txt){
  const fmt = _getColorFmt();
  const t = String(txt || '').trim();
  if(!t) return null;
  const parseRgb = (str) => {
    let s = String(str || '').trim().toLowerCase();
    s = s.replace(/^rgb\(/,'').replace(/\)$/,'');
    const parts = s.split(/[\s,]+/).filter(Boolean);
    if(parts.length !== 3) return null;
    const r = parseInt(parts[0],10);
    const g = parseInt(parts[1],10);
    const b = parseInt(parts[2],10);
    if(!Number.isFinite(r) || !Number.isFinite(g) || !Number.isFinite(b)) return null;
    return _rgbToHex(r,g,b);
  };

  if(fmt === 'rgb'){
    const asRgb = parseRgb(t);
    if(asRgb) return asRgb;
    return _normHex(t);
  }

  // HEX mode
  const asHex = _normHex(t);
  if(asHex) return asHex;
  if(/^rgb\(/i.test(t)) return parseRgb(t);
  return null;
}

function refreshColorCodes(){
  document.querySelectorAll('input[type=color][data-hascode="1"]').forEach(inp => {
    const sid = inp.getAttribute('data-code-input');
    const ci = sid ? document.getElementById(sid) : null;
    if(ci) ci.value = _formatColorLabel(inp.value);
  });
}

function initColorCodes(){
  const sel = el('colorFormat');
  if(sel){
    const saved = localStorage.getItem(COLOR_FMT_KEY);
    if(saved === 'hex' || saved === 'rgb') sel.value = saved;
    sel.addEventListener('change', () => {
      localStorage.setItem(COLOR_FMT_KEY, sel.value);
      refreshColorCodes();
    });
  }

  const bind = (colorInp, codeInp) => {
    const upd = () => { codeInp.value = _formatColorLabel(colorInp.value); };
    const commit = () => {
      const parsed = _parseColorLabel(codeInp.value);
      if(parsed){
        colorInp.value = parsed;
        upd();
      } else {
        upd();
        toast('Неверный цвет', 'Используй #RRGGBB или rgb(r,g,b)', 'warn', 3500);
      }
    };

    colorInp.addEventListener('input', upd);
    colorInp.addEventListener('change', upd);

    codeInp.addEventListener('focus', () => { try{ codeInp.select(); } catch(e){} });
    codeInp.addEventListener('keydown', (e) => {
      if(e.key === 'Enter'){
        commit();
        codeInp.blur();
      }
    });
    codeInp.addEventListener('blur', commit);
    upd();
  };

  // 1) Row mark color (rmColor) — поле рядом
  const rm = el('rmColor');
  if(rm && rm.getAttribute('data-hascode') !== '1'){
    const ci = document.createElement('input');
    ci.type = 'text';
    ci.className = 'clrCodeInput';
    ci.id = 'rmColorCode';
    ci.autocomplete = 'off';
    ci.spellcheck = false;
    rm.insertAdjacentElement('afterend', ci);
    rm.setAttribute('data-hascode', '1');
    rm.setAttribute('data-code-input', ci.id);
    bind(rm, ci);
  }

  // 2) Colors in W/X/Y scales — поле под цветом
  document.querySelectorAll('.scaleGroup input[type=color]').forEach(inp => {
    if(inp.getAttribute('data-hascode') === '1') return;
    const ci = document.createElement('input');
    ci.type = 'text';
    ci.className = 'clrCodeInput';
    ci.id = (inp.id ? `${inp.id}_code` : `clr_${Math.random().toString(16).slice(2)}_code`);
    ci.autocomplete = 'off';
    ci.spellcheck = false;
    inp.insertAdjacentElement('afterend', ci);
    inp.setAttribute('data-hascode', '1');
    inp.setAttribute('data-code-input', ci.id);
    bind(inp, ci);
  });

  refreshColorCodes();
}

function _collectStyleSettingsFromUI() {
  return {
    row_mark: {
      threshold_T: _num(el('rmThreshold')?.value, 150),
      color: (el('rmColor')?.value || '#EAD706'),
      intensity: Math.max(0, Math.min(100, parseInt(el('rmIntensity')?.value || '100', 10)))
    },
    scales: {
      W: {
        min: _num(el('wMin')?.value, 0),
        opt: _num(el('wOpt')?.value, 1),
        max: _num(el('wMax')?.value, 2),
        colors: {
          min: (el('wCMin')?.value || '#007BFF'),
          opt: (el('wCOpt')?.value || '#00FF00'),
          max: (el('wCMax')?.value || '#FE3448'),
        }
      },
      X: {
        min: _num(el('xMin')?.value, 0),
        opt: _num(el('xOpt')?.value, 9),
        max: _num(el('xMax')?.value, 10),
        colors: {
          min: (el('xCMin')?.value || '#007BFF'),
          opt: (el('xCOpt')?.value || '#00FF00'),
          max: (el('xCMax')?.value || '#FE3448'),
        }
      },
      Y: {
        min: _num(el('yMin')?.value, 0),
        opt: _num(el('yOpt')?.value, 5),
        max: _num(el('yMax')?.value, 6),
        colors: {
          min: (el('yCMin')?.value || '#007BFF'),
          opt: (el('yCOpt')?.value || '#00FF00'),
          max: (el('yCMax')?.value || '#FE3448'),
        }
      }
    }
  };
}

function _applyStyleSettingsToUI(s) {
  if(!s) return;
  const rm = s.row_mark || {};
  _set('rmThreshold', rm.threshold_T ?? 150);
  _set('rmColor', rm.color || '#EAD706');
  _set('rmIntensity', rm.intensity ?? 100);
  _setText('rmIntensityVal', String(rm.intensity ?? 100));

  const sc = s.scales || {};
  const w = sc.W || {}; const x = sc.X || {}; const y = sc.Y || {};
  _set('wMin', w.min ?? 0); _set('wOpt', w.opt ?? 1); _set('wMax', w.max ?? 2);
  _set('xMin', x.min ?? 0); _set('xOpt', x.opt ?? 9); _set('xMax', x.max ?? 10);
  _set('yMin', y.min ?? 0); _set('yOpt', y.opt ?? 5); _set('yMax', y.max ?? 6);

  const wc = (w.colors || {});
  const xc = (x.colors || {});
  const yc = (y.colors || {});
  _set('wCMin', wc.min || '#007BFF'); _set('wCOpt', wc.opt || '#00FF00'); _set('wCMax', wc.max || '#FE3448');
  _set('xCMin', xc.min || '#007BFF'); _set('xCOpt', xc.opt || '#00FF00'); _set('xCMax', xc.max || '#FE3448');
  _set('yCMin', yc.min || '#007BFF'); _set('yCOpt', yc.opt || '#00FF00'); _set('yCMax', yc.max || '#FE3448');

  // если значения выставлены программно — обновим подписи
  try{ refreshColorCodes(); } catch(e) {}
}

function loadStyleSettings() {
  return fetch('/api/settings')
    .then(r=>r.json())
    .then(j=>{
      if(!j || !j.ok) throw new Error((j && j.error) ? j.error : 'Не удалось загрузить настройки');
      VIEWER_SETTINGS = j.settings;
      _applyStyleSettingsToUI(VIEWER_SETTINGS);
      _styleStatus('Настройки загружены', false);
    })
    .catch(e=>{
      _styleStatus('Ошибка загрузки настроек', true);
      log('Settings load error: ' + e);
    });
}

function saveStyleSettings() {
  const btn = el('btnSaveStyle');
  const s = _collectStyleSettingsFromUI();
  if(btn) btn.disabled = true;
  _styleStatus('Сохраняю...', false);

  return fetch('/api/settings', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(s)
  })
    .then(r=>r.json())
    .then(j=>{
      if(!j || !j.ok) throw new Error((j && j.error) ? j.error : 'Не удалось сохранить настройки');
      VIEWER_SETTINGS = j.settings;
      _applyStyleSettingsToUI(VIEWER_SETTINGS);
      _styleStatus('Сохранено', false);
      log('Settings saved');
    })
    .catch(e=>{
      _styleStatus('Ошибка сохранения', true);
      log('Settings save error: ' + e);
      toast('Ошибка сохранения', (e && e.message) ? e.message : String(e), 'err', 6000);
    })
    .finally(()=>{ if(btn) btn.disabled = false; });
}

function scheduleRedraw() {
  if(redrawTimer) clearTimeout(redrawTimer);
  // сохраняем последний выбор/настройки, чтобы после перезапуска быстро продолжить
  scheduleSaveLastState();
  redrawTimer = setTimeout(() => {
    if(LOADED) drawPlot();
  }, 150);
}

// wiring
function wire() {
  const bp = el("btnPick");
  if(bp) bp.addEventListener("click", pickFolder);

  const bl = el("btnLoad");
  if(bl) bl.addEventListener("click", loadTest);

  const bd = el("btnDraw");
  if(bd) bd.addEventListener("click", drawPlot);

  // Enter в поле папки = загрузить
  const fld = el('folder');
  if(fld) fld.addEventListener('keydown', (e)=>{
    if(e.key === 'Enter') loadTest();
  });

  // ===== Stage 3: view modes + recent folders + template options =====
  loadViewOpts();
  loadRecentFolders();
  refreshRecentFoldersUI();
  applyViewClasses();

  const rf = el('recentFolders');
  if(rf) rf.addEventListener('change', () => {
    const v = String(rf.value || '').trim();
    if(v && el('folder')) el('folder').value = v;
  });

  const bcopy = el('btnCopyFolder');
  if(bcopy) bcopy.addEventListener('click', copyFolderPath);

  const bclr = el('btnClearRecent');
  if(bclr) bclr.addEventListener('click', clearRecentFolders);

  const vc = el('viewCompact');
  if(vc) vc.addEventListener('change', () => {
    syncViewOptsFromUI(true);
    applyViewClasses();
    applySortAndRender(true, false);
  });
  const vg = el('viewGroups');
  if(vg) vg.addEventListener('change', () => {
    syncViewOptsFromUI(true);
    applyViewClasses();
    applySortAndRender(true, false);
  });

  const te = el('tplExtra');
  if(te) te.addEventListener('change', () => updateExportInfo());

  const bc = el("btnCsv");
  if(bc) bc.addEventListener("click", () => exportData("csv"));

  const bx = el("btnXlsx");
  if(bx) bx.addEventListener("click", () => exportData("xlsx"));

  const bt = el("btnTpl");
  if(bt) bt.addEventListener("click", exportTemplate);

  const ba = el("btnAll");
  if(ba) ba.addEventListener("click", selectAll);

  const bn = el("btnNone");
  if(bn) bn.addEventListener("click", clearAll);

  const sm = el("sortMode");
  if(sm) sm.addEventListener("change", () => applySortAndRender(true, true));

  // ===== Filters (search / only selected / chips) =====
  let _ft = null;
  const applyFiltersDebounced = () => {
    if(_ft) clearTimeout(_ft);
    _ft = setTimeout(() => applySortAndRender(true, false), 150);
  };

  const si = el('chanSearch');
  if(si) si.addEventListener('input', () => {
    FILTER_TEXT = (si.value || '').trim();
    applyFiltersDebounced();
  });
  if(si) si.addEventListener('keydown', (e)=>{
    if(e.key === 'Escape') { si.value=''; FILTER_TEXT=''; applySortAndRender(true, false); }
  });

  const onlySel = el('onlySelected');
  if(onlySel) onlySel.addEventListener('change', () => {
    FILTER_ONLY_SELECTED = !!onlySel.checked;
    applySortAndRender(true, false);
  });

  const clr = el('btnClearFilter');
  if(clr) clr.addEventListener('click', clearFilters);

  const chipBox = el('chipRow');
  if(chipBox) chipBox.addEventListener('click', (e)=>{
    const btn = e.target && e.target.closest ? e.target.closest('button.chip') : null;
    if(!btn) return;
    // clear button
    if(btn.id === 'btnClearFilter') return;
    const type = btn.getAttribute('data-ftype');
    const val = btn.getAttribute('data-fval');
    if(!type || !val) return;

    const isActive = btn.classList.contains('active');
    document.querySelectorAll('.chip.active').forEach(b=>b.classList.remove('active'));
    if(isActive) {
      FILTER_CHIP = null;
    } else {
      btn.classList.add('active');
      FILTER_CHIP = {type, value: val};
    }
    applySortAndRender(true, false);
  });

  // ===== Step controls =====
  const stAuto = el('stepAuto');
  if(stAuto) stAuto.addEventListener('change', () => { updateStepUI(); scheduleRedraw(); });
  const stTarget = el('stepTarget');
  if(stTarget) stTarget.addEventListener('change', () => { updateStepUI(); scheduleRedraw(); });
  const stInp = el('step');
  if(stInp) stInp.addEventListener('change', () => { updateStepUI(); scheduleRedraw(); });

  // Legend toggle
  const lg = el('showLegend');
  if(lg) lg.addEventListener('change', () => scheduleRedraw());

  const bs = el("btnSaveNamed");
  if(bs) bs.addEventListener("click", saveNamedOrder);

  const bln = el("btnLoadNamed");
  if(bln) bln.addEventListener("click", loadNamedOrder);

  const br = el("btnRefreshOrders");
  if(br) br.addEventListener("click", () => refreshNamedOrders());

  const bstyle = el("btnSaveStyle");
  if(bstyle) bstyle.addEventListener("click", saveStyleSettings);

  const rint = el("rmIntensity");
  if(rint) rint.addEventListener("input", () => {
    const sp = el("rmIntensityVal");
    if(sp) sp.textContent = rint.value;
  });



  // ===== Presets =====
  const bpS = el('btnSavePreset');
  if(bpS) bpS.addEventListener('click', savePreset);
  const bpL = el('btnLoadPreset');
  if(bpL) bpL.addEventListener('click', loadPreset);
  const bpD = el('btnDeletePreset');
  if(bpD) bpD.addEventListener('click', deletePreset);
  const bpR = el('btnRefreshPresets');
  if(bpR) bpR.addEventListener('click', () => refreshPresets());

  const ps = el('presetSelect');
  if(ps) ps.addEventListener('change', () => {
    // подставим имя в поле, чтобы удобно было перезаписать
    const opt = ps.selectedOptions && ps.selectedOptions[0];
    if(opt && ps.value && el('presetName')) {
      el('presetName').value = (opt.textContent || '').replace(/\s*\(\d+\)\s*$/,'').trim();
    }
  });

  // load export styling settings
  initColorCodes();
  loadStyleSettings();


  // populate lists on startup
  refreshNamedOrders();
  refreshPresets();

  // init step UI
  updateStepUI();

  log("Ready. По умолчанию порядок каналов = как в файле. Сортировку можно менять, а перетаскивание сохраняет 'Мой порядок'.");
}

wire();
