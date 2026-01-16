// LeMuRe Viewer UI (selection like <select multiple> + drag reorder + export by visible X range)
let LOADED = false;

let CHANNELS_FILE = [];       // channels in the order they appear in file
let CHANNELS_VIEW = [];       // currently shown order in list
let SAVED_ORDER = [];         // saved custom order (from channel_order.json) - not auto-applied
let SUMMARY = null;
let currentRange = null;      // [start_ms, end_ms]

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

function setRangeText() {
  const r = currentRange && currentRange.length===2 ? [Math.min(currentRange[0], currentRange[1]), Math.max(currentRange[0], currentRange[1])] : null;
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
      setRangeText();
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

      Plotly.newPlot("plot", traces, layout, config);

      const plotDiv = el("plot");
      plotDiv.on("plotly_relayout", (ev) => {
        if(ev["xaxis.autorange"] === true) {
          currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
          setRangeText();
          return;
        }
        const r0 = ev["xaxis.range[0]"];
        const r1 = ev["xaxis.range[1]"];
        if(r0 && r1) {
          const s = Date.parse(r0);
          const e = Date.parse(r1);
          if(!isNaN(s) && !isNaN(e)) {
            currentRange = [Math.min(s,e), Math.max(s,e)];
            setRangeText();
          }
        }
      });

      currentRange = [SUMMARY.start_ms, SUMMARY.end_ms];
      setRangeText();
      log(`Plot updated: channels=${codes.length}, step=${step}`);
    })
    .catch(e=>log("drawPlot error: " + e));
}



function exportTemplate() {
  if(!LOADED || !SUMMARY) { alert("Сначала загрузите тест"); return; }

  // Respect selected channels (same as plot/CSV/XLSX exports)
  const codes = getSelectedCodes();
  if(!codes || !codes.length) {
    alert("Выбери хотя бы один канал");
    return;
  }

  let start_ms = SUMMARY.start_ms;
  let end_ms = SUMMARY.end_ms;
  // Prefer the remembered range; if it's missing, try to read the current Plotly x-range
  if(currentRange && currentRange.length === 2) {
    start_ms = Math.min(currentRange[0], currentRange[1]);
    end_ms   = Math.max(currentRange[0], currentRange[1]);
  } else {
    try {
      const plotDiv = document.getElementById("plot");
      const rng = plotDiv && plotDiv._fullLayout && plotDiv._fullLayout.xaxis && plotDiv._fullLayout.xaxis.range;
      if(rng && rng.length === 2) {
        const s = Date.parse(rng[0]);
        const e = Date.parse(rng[1]);
        if(!isNaN(s) && !isNaN(e)) {
          start_ms = Math.min(s, e);
          end_ms = Math.max(s, e);
        }
      }
    } catch(_) {}
  }

  const btn = el("btnTpl");
  const st = el("tplStatus");

  const url = `/api/export_template?start_ms=${start_ms}&end_ms=${end_ms}&channels=${encodeURIComponent(codes.join(","))}`;
  log(`Export TEMPLATE: ${new Date(start_ms).toISOString()} -> ${new Date(end_ms).toISOString()} (20s grid)`);

  if(btn) { btn.classList.add("btnBusy"); btn.textContent = "Готовлю шаблон…"; }
  if(st)  { st.innerHTML = `<span class="spinner"></span>Формирую Excel… <span style="opacity:.85">(каналы: ${codes.length}, диапазон: ${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()})</span>`; }

  fetch(url, {method:"GET"})
    .then(async (resp) => {
      if(!resp.ok) {
        const txt = await resp.text().catch(()=> "");
        throw new Error(`HTTP ${resp.status} ${resp.statusText} ${txt}`.trim());
      }

      // get filename from Content-Disposition if present
      let filename = "template_filled.xlsx";
      const cd = resp.headers.get("Content-Disposition") || resp.headers.get("content-disposition");
      if(cd) {
        const m = /filename\*?=(?:UTF-8''|")?([^\";]+)/i.exec(cd);
        if(m && m[1]) {
          filename = decodeURIComponent(m[1].replace(/"/g,"").trim());
        }
      }

      const blob = await resp.blob();
      const blobUrl = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();

      setTimeout(()=>URL.revokeObjectURL(blobUrl), 15000);

      log(`Template ready: ${filename} (${Math.round(blob.size/1024)} KB)`);
      if(st) st.innerHTML = `Готово: <b>${filename}</b> (скачивание началось)`;
      setTimeout(()=>{ if(st) st.textContent = ""; }, 7000);
    })
    .catch((e) => {
      console.error(e);
      log("Template export error: " + e.message);
      alert("Ошибка при формировании шаблона. Смотри блок 'Лог'.\n\n" + e.message);
      if(st) st.textContent = "Ошибка (см. лог)";
      setTimeout(()=>{ if(st) st.textContent = ""; }, 7000);
    })
    .finally(() => {
      if(btn) { btn.classList.remove("btnBusy"); btn.textContent = "В шаблон XLSX"; }
    });
}

function exportData(fmt) {
  if(!LOADED || !SUMMARY) { alert("Сначала загрузите тест"); return; }
  const codes = getSelectedCodes();
  if(!codes.length) { alert("Выбери хотя бы один канал"); return; }

  const step = getStep();
  let start_ms = SUMMARY.start_ms;
  let end_ms = SUMMARY.end_ms;

  if(currentRange && currentRange.length === 2) {
    start_ms = Math.min(currentRange[0], currentRange[1]);
    end_ms   = Math.max(currentRange[0], currentRange[1]);
  }

  const url = `/api/export?format=${encodeURIComponent(fmt)}&channels=${encodeURIComponent(codes.join(","))}&start_ms=${start_ms}&end_ms=${end_ms}&step=${step}`;
  log(`Export ${fmt}: ${new Date(start_ms).toISOString()} -> ${new Date(end_ms).toISOString()} | step=${step}`);
  window.location.href = url;
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

  // populate list on startup
  refreshNamedOrders();

  log("Ready. По умолчанию порядок каналов = как в файле. Сортировку можно менять, а перетаскивание сохраняет 'Мой порядок'.");
}

wire();
