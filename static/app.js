
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
      if (!Number.isNaN(a) && !Number.isNaN(b)) return [a, b];
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
      if (!Number.isNaN(a) && !Number.isNaN(b)) return [a, b];
    }
    if (ev["xaxis.range[0]"] != null && ev["xaxis.range[1]"] != null) {
      const a = parsePlotlyDate(ev["xaxis.range[0]"]);
      const b = parsePlotlyDate(ev["xaxis.range[1]"]);
      if (!Number.isNaN(a) && !Number.isNaN(b)) return [a, b];
    }
  } catch (e) {}
  return null;
}
// LeMuRe Viewer UI (selection like <select multiple> + drag reorder + export by visible X range)
let LOADED = false;

let CHANNELS_FILE = [];       // channels in the order they appear in file
let CHANNELS_VIEW = [];       // currently shown order in list
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
  // preserve visual order (top-to-bottom)
  const order = currentOrderCodes();
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

function renderChannelList(arr, keepSelection=true) {
  const list = el("channelList");
  list.innerHTML = "";
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-multiselectable", "true");

  // Keep selection intersection
  const prev = new Set(selected);

  CHANNELS_VIEW = arr.slice();

  arr.forEach((ch) => {
    const li = document.createElement("li");
    li.className = "chanItem";
    li.setAttribute("draggable", "false");
    li.setAttribute("data-code", ch.code);
    li.setAttribute("role", "option");

    // Drag handle: draggable only here
    const handle = document.createElement("div");
    handle.className = "dragHandle";
    handle.textContent = "≡";
    handle.title = "Перетащить";
    handle.setAttribute("draggable", "true");
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

    // Selection clicks on row
    li.addEventListener("click", (ev) => onItemClick(ev, ch.code));
    li.addEventListener("dblclick", (ev) => onItemDblClick(ev, ch.code));

    // Drag start/end on handle
    handle.addEventListener("dragstart", (ev) => {
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

  // Restore selection
  if(keepSelection && prev.size) {
    const available = new Set(arr.map(c => c.code));
    const inter = Array.from(prev).filter(c => available.has(c));
    if(inter.length) selected = new Set(inter);
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
}

function scheduleSaveOrder() {
  if(saveOrderTimer) clearTimeout(saveOrderTimer);
  saveOrderTimer = setTimeout(() => {
    const order = currentOrderCodes();
    const prevOrder = SAVED_ORDER.join(',');
    const newOrder = order.join(',');
    if(prevOrder === newOrder) return;
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
    alert('Сначала загрузите тест, чтобы были каналы.');
    return;
  }
  const name = (el('orderName') ? el('orderName').value : '').trim();
  if(!name) {
    alert('Введите имя порядка.');
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
    if(!j.ok) { alert('Ошибка сохранения: ' + (j.error||'')); log('orders_save error: ' + (j.error||'')); return; }
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
  .catch(e=>{ alert('Ошибка сохранения: ' + e); log('orders_save error: ' + e); });
}

function loadNamedOrder() {
  if(!CHANNELS_FILE || CHANNELS_FILE.length === 0) {
    alert('Сначала загрузите тест, чтобы были каналы.');
    return;
  }
  const key = el('orderSelect') ? el('orderSelect').value : '';
  if(!key) {
    alert('Выберите сохранённый порядок из списка.');
    return;
  }
  fetch('/api/orders_load?key=' + encodeURIComponent(key))
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) { alert('Ошибка загрузки: ' + (j.error||'')); log('orders_load error: ' + (j.error||'')); return; }
      const order = j.order || [];
      SAVED_ORDER = order.slice();
      setSortMode('custom');
      applySortAndRender(true);
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
    .catch(e=>{ alert('Ошибка загрузки: ' + e); log('orders_load error: ' + e); });
}
function applySortAndRender(keepSelection=true) {
  const arr = buildViewOrder();
  renderChannelList(arr, keepSelection);
  scheduleRedraw();
}

function loadTest() {
  const folder = el("folder").value.trim();
  if(!folder) { alert("Укажи папку с тестом"); return; }

  fetch("/api/load", {
    method:"POST",
    headers: {"Content-Type":"application/json"},
    body: JSON.stringify({folder})
  })
  .then(r=>r.json())
  .then(j=>{
    if(!j.ok) { alert("Ошибка: " + j.error); log("load error: " + j.error); return; }
    LOADED = true;

    CHANNELS_FILE = j.channels || [];
    SAVED_ORDER = j.saved_order || [];

    SUMMARY = j.summary || null;
    if(SUMMARY) {
      el("summary").innerHTML =
        `Точек: <b>${SUMMARY.points}</b><br>` +
        `Начало: <b>${SUMMARY.start}</b><br>` +
        `Конец: <b>${SUMMARY.end}</b>`;
      currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
      updateRangeText();
    } else {
      el("summary").textContent = "Нет данных";
    }

    // Requirement: no sorting on load -> "file" mode
    setSortMode("file");
    applySortAndRender(false);

    log("Loaded: " + (j.folder || folder));
    drawPlot();
  })
  .catch(e=>log("loadTest error: " + e));
}

function getStep() {
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
  if(!codes.length) { alert("Выбери хотя бы один канал"); return; }

  const step = getStep();

  // Запоминаем текущий видимый диапазон ДО перерисовки, чтобы он не сбрасывался при смене датчиков
  let desiredRange = null;
  const vrBefore = getVisibleRangeFromPlot();
  if(vrBefore && vrBefore.length === 2) {
    desiredRange = [Math.min(vrBefore[0], vrBefore[1]), Math.max(vrBefore[0], vrBefore[1])];
  } else if(currentRange && currentRange.length === 2) {
    desiredRange = [Math.min(currentRange[0], currentRange[1]), Math.max(currentRange[0], currentRange[1])];
  }

  const start_ms = SUMMARY.start_ms;
  const end_ms = SUMMARY.end_ms;

  fetch(`/api/series?channels=${encodeURIComponent(codes.join(","))}&start_ms=${start_ms}&end_ms=${end_ms}&step=${step}`)
    .then(r=>r.json())
    .then(j=>{
      if(!j.ok) { alert("Ошибка: " + j.error); log("series error: " + j.error); return; }
      const t = (j.t_ms || []).map(x => new Date(x));
      const series = j.series || {};
      const traces = [];

      codes.forEach(code => {
        const y = series[code] || [];
        traces.push({
          type: "scatter",
          mode: "lines",
          name: labelFor(code),
          x: t,
          y: y,
          hovertemplate: "%{x|%Y-%m-%d %H:%M:%S}<br>%{y:.2f}<extra></extra>",
        });
      });

      const layout = {
        // Важно: uirevision + Plotly.react сохраняют zoom/диапазон при обновлении данных
        uirevision: "lemure-v14",
        margin: {l: 70, r: 20, t: 140, b: 90},
        hovermode: "x unified",
        xaxis: {
          type: "date",
          tickformat: "%H:%M<br>%d.%m",
          showspikes: true,
          spikemode: "across",
          spikesnap: "cursor",
          rangeslider: {visible: true},
        },
        yaxis: {automargin: true},
        legend: {
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
        },
      };

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

        // Восстановить диапазон (если был выбран) после обновления датчиков
        if(desiredRange && desiredRange.length === 2) {
          const r0 = new Date(desiredRange[0]).toISOString();
          const r1 = new Date(desiredRange[1]).toISOString();
          try {
            Plotly.relayout(plotDiv, {"xaxis.range": [r0, r1]});
          } catch(_) {}
          currentRange = desiredRange.slice();
        } else if(!currentRange) {
          currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
        }
        updateRangeText();
      });

      // Текст диапазона обновим сразу (для экспорта), даже если relayout применится чуть позже
      if(desiredRange && desiredRange.length === 2) currentRange = desiredRange.slice();
      if(!currentRange) currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
      updateRangeText();

      log(`Plot updated: channels=${codes.length}, step=${step}`);
    })
    .catch(e=>log("drawPlot error: " + e));
}




function exportData(fmt) {
  if(!LOADED || !SUMMARY) { alert("Сначала загрузите тест"); return; }

  fmt = (fmt || "csv").toLowerCase();
  if(fmt !== "csv" && fmt !== "xlsx") fmt = "csv";

  const codes = getSelectedCodes();
  if(!codes || !codes.length) {
    alert("Выбери хотя бы один канал");
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

  const step = getStep();

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
      log(`Export ${fmt.toUpperCase()} OK: channels=${codes.length}, range=${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()}, step=${step}`);
    })
    .catch((e) => {
      if(st) st.textContent = "Ошибка: " + (e && e.message ? e.message : e);
      log("Export " + fmt + " error: " + e);
      alert("Ошибка экспорта. Смотри блок 'Лог'.\n\n" + (e && e.message ? e.message : e));
    })
    .finally(() => {
      if(bc) bc.disabled = false;
      if(bx) bx.disabled = false;
    });
}
function exportTemplate() {
  if(!LOADED || !SUMMARY) { alert("Сначала загрузите тест"); return; }

  const codes = getSelectedCodes();
  if(!codes || !codes.length) {
    alert("Выбери хотя бы один канал");
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

  const step = getStep();

  const btn = el("btnTpl");
  const st  = el("tplStatus");
  if(btn) { btn.disabled = true; btn.textContent = "Готовлю шаблон…"; }
  if(st)  { st.innerHTML = `<span class="spinner"></span>Формирую Excel… <span style="opacity:.85">(каналы: ${codes.length}, диапазон: ${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()})</span>`; }

  const qs = new URLSearchParams();
  qs.set("start_ms", String(start_ms));
  qs.set("end_ms", String(end_ms));
  qs.set("channels", codes.join(","));
  qs.set("step", String(step));

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
      log(`Export TEMPLATE OK: channels=${codes.length}, range=${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()}, step=${step}`);
    })
    .catch((e) => {
      if(st) st.textContent = "Ошибка: " + (e && e.message ? e.message : e);
      log("Export TEMPLATE error: " + e);
      alert("Ошибка при формировании шаблона. Смотри блок 'Лог'.\n\n" + (e && e.message ? e.message : e));
    })
    .finally(() => {
      if(btn) { btn.disabled = false; btn.textContent = "В шаблон XLSX"; }
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

function _collectStyleSettingsFromUI() {
  return {
    row_mark: {
      threshold_T: _num(el('rmThreshold')?.value, 150),
      color: (el('rmColor')?.value || '#FFF2CC'),
      intensity: Math.max(0, Math.min(100, parseInt(el('rmIntensity')?.value || '100', 10)))
    },
    scales: {
      W: {
        min: _num(el('wMin')?.value, 0),
        opt: _num(el('wOpt')?.value, 1),
        max: _num(el('wMax')?.value, 2),
        colors: {
          min: (el('wCMin')?.value || '#0000FF'),
          opt: (el('wCOpt')?.value || '#00FF00'),
          max: (el('wCMax')?.value || '#FF0000'),
        }
      },
      X: {
        min: _num(el('xMin')?.value, 0),
        opt: _num(el('xOpt')?.value, 9),
        max: _num(el('xMax')?.value, 18),
        colors: {
          min: (el('xCMin')?.value || '#0000FF'),
          opt: (el('xCOpt')?.value || '#00FF00'),
          max: (el('xCMax')?.value || '#FF0000'),
        }
      },
      Y: {
        min: _num(el('yMin')?.value, 0),
        opt: _num(el('yOpt')?.value, 5),
        max: _num(el('yMax')?.value, 10),
        colors: {
          min: (el('yCMin')?.value || '#0000FF'),
          opt: (el('yCOpt')?.value || '#00FF00'),
          max: (el('yCMax')?.value || '#FF0000'),
        }
      }
    }
  };
}

function _applyStyleSettingsToUI(s) {
  if(!s) return;
  const rm = s.row_mark || {};
  _set('rmThreshold', rm.threshold_T ?? 150);
  _set('rmColor', rm.color || '#FFF2CC');
  _set('rmIntensity', rm.intensity ?? 100);
  _setText('rmIntensityVal', String(rm.intensity ?? 100));

  const sc = s.scales || {};
  const w = sc.W || {}; const x = sc.X || {}; const y = sc.Y || {};
  _set('wMin', w.min ?? 0); _set('wOpt', w.opt ?? 1); _set('wMax', w.max ?? 2);
  _set('xMin', x.min ?? 0); _set('xOpt', x.opt ?? 9); _set('xMax', x.max ?? 18);
  _set('yMin', y.min ?? 0); _set('yOpt', y.opt ?? 5); _set('yMax', y.max ?? 10);

  const wc = (w.colors || {});
  const xc = (x.colors || {});
  const yc = (y.colors || {});
  _set('wCMin', wc.min || '#0000FF'); _set('wCOpt', wc.opt || '#00FF00'); _set('wCMax', wc.max || '#FF0000');
  _set('xCMin', xc.min || '#0000FF'); _set('xCOpt', xc.opt || '#00FF00'); _set('xCMax', xc.max || '#FF0000');
  _set('yCMin', yc.min || '#0000FF'); _set('yCOpt', yc.opt || '#00FF00'); _set('yCMax', yc.max || '#FF0000');
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
      alert('Не удалось сохранить настройки оформления.\n\n' + (e && e.message ? e.message : e));
    })
    .finally(()=>{ if(btn) btn.disabled = false; });
}

function scheduleRedraw() {
  if(redrawTimer) clearTimeout(redrawTimer);
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
  if(sm) sm.addEventListener("change", () => applySortAndRender(true));

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

  // load export styling settings
  loadStyleSettings();


  // populate list on startup
  refreshNamedOrders();

  log("Ready. По умолчанию порядок каналов = как в файле. Сортировку можно менять, а перетаскивание сохраняет 'Мой порядок'.");
}

wire();
