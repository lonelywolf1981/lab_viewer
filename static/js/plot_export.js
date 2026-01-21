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
  if(_forceResetRangeOnNextPlot) {
    // Загружен новый тест: НЕ переносим zoom/диапазон со старого графика.
    desiredRange = [SUMMARY.start_ms, SUMMARY.end_ms];
  } else {
    const vrBefore = getVisibleRangeFromPlot();
    if(vrBefore && vrBefore.length === 2) {
      desiredRange = [Math.min(vrBefore[0], vrBefore[1]), Math.max(vrBefore[0], vrBefore[1])];
    } else if(currentRange && currentRange.length === 2) {
      desiredRange = [Math.min(currentRange[0], currentRange[1]), Math.max(currentRange[0], currentRange[1])];
    }
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
        // Сброс диапазона применён (если был нужен) — дальше сохраняем диапазон при смене каналов.
        _forceResetRangeOnNextPlot = false;
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
      try {
        _tplServerTotalS = r.headers.get('X-Export-Total-S');
        _tplServerTiming = r.headers.get('X-Export-Timing');
      } catch(e) {}
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

      if(st) {
        const t = (_tplServerTiming ? (' (сервер: ' + _tplServerTiming + 's)') : '');
        st.textContent = 'Готово. Файл скачивается…' + t;
      }
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

  const endBusy = beginBusy('Формирую XLSX по шаблону…', {timer:true});

  const btn = el("btnTpl");
  const st  = el("tplStatus");

  const includeExtra = (el('tplExtra') && el('tplExtra').checked) ? 1 : 0;
  const refrigerant = getRefrigerant();
  if(btn) { btn.disabled = true; btn.textContent = "Готовлю шаблон…"; }
  if(st)  { st.innerHTML = `<span class="spinner"></span>Формирую Excel… <span style="opacity:.85">(каналы: ${codes.length}, диапазон: ${new Date(start_ms).toLocaleString()} → ${new Date(end_ms).toLocaleString()}, Z: ${includeExtra ? 'да' : 'нет'}, хладагент: ${refrigerant})</span>`; }

  const qs = new URLSearchParams();
  qs.set("start_ms", String(start_ms));
  qs.set("end_ms", String(end_ms));
  qs.set("channels", codes.join(","));
  qs.set("step", String(step));
  qs.set('include_extra', String(includeExtra));
  qs.set('refrigerant', refrigerant);

  let _tplServerTotalS = null;
  let _tplServerTiming = null;

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
      try {
        _tplServerTotalS = r.headers.get('X-Export-Total-S');
        _tplServerTiming = r.headers.get('X-Export-Timing');
      } catch(e) {}
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
      if(st) {
        const t = (_tplServerTotalS ? (' (сервер: ' + _tplServerTotalS + 'с)') : '');
        st.textContent = 'Готово. Файл скачивается…' + t;
      }
      toast('Шаблон готов', _tplServerTotalS ? (`${codes.length} каналов (сервер: ${_tplServerTotalS}с)`) : (`${codes.length} каналов`), 'ok');
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

  // 1) Row mark color (rmColor) — поле под колорпикером
  const rm = el('rmColor');
  if(rm && rm.getAttribute('data-hascode') !== '1'){
    // Если поле кода уже есть в HTML — используем его, иначе создаём.
    let ci = el('rmColorCode');
    if(!ci){
      ci = document.createElement('input');
      ci.type = 'text';
      ci.className = 'clrCodeInput';
      ci.id = 'rmColorCode';
      ci.autocomplete = 'off';
      ci.spellcheck = false;
      rm.insertAdjacentElement('afterend', ci);
    }
    rm.setAttribute('data-hascode', '1');
    rm.setAttribute('data-code-input', ci.id);
    bind(rm, ci);
  }

  // 1b) t нагнетания (tdColor) — поле кода рядом с колорпикером
  const td = el('tdColor');
  if(td && td.getAttribute('data-hascode') !== '1'){
    let ci = el('tdColorCode');
    if(!ci){
      ci = document.createElement('input');
      ci.type = 'text';
      ci.className = 'clrCodeInput';
      ci.id = 'tdColorCode';
      ci.autocomplete = 'off';
      ci.spellcheck = false;
      td.insertAdjacentElement('afterend', ci);
    }
    td.setAttribute('data-hascode', '1');
    td.setAttribute('data-code-input', ci.id);
    bind(td, ci);
  }

  // 1c) t всасывания (tsColor) — поле кода рядом с колорпикером
  const ts = el('tsColor');
  if(ts && ts.getAttribute('data-hascode') !== '1'){
    let ci = el('tsColorCode');
    if(!ci){
      ci = document.createElement('input');
      ci.type = 'text';
      ci.className = 'clrCodeInput';
      ci.id = 'tsColorCode';
      ci.autocomplete = 'off';
      ci.spellcheck = false;
      ts.insertAdjacentElement('afterend', ci);
    }
    ts.setAttribute('data-hascode', '1');
    ts.setAttribute('data-code-input', ci.id);
    bind(ts, ci);
  }


  // 2) Colors in W/X/Y scales — поле под цветом
  document.querySelectorAll('.scaleGrid input[type=color]').forEach(inp => {
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
  const _safeMax = (optVal, fallbackDelta=1) => {
    const o = _num(optVal, 0);
    // max не показываем в UI, но сервер ожидает opt < max
    return o + fallbackDelta;
  };
  const _numOrNull = (v) => {
    const t = String(v ?? '').trim();
    if(t === '') return null;
    const x = parseFloat(t);
    return Number.isFinite(x) ? x : null;
  };
  return {
    row_mark: {
      threshold_T: _num(el('rmThreshold')?.value, 150),
      color: (el('rmColor')?.value || '#EAD706'),
      intensity: Math.max(0, Math.min(100, parseInt(el('rmIntensity')?.value || '100', 10)))
    },
    discharge_mark: {
      threshold: _numOrNull(el('tdThreshold')?.value),
      color: (el('tdColor')?.value || '#FFC000')
    },
    suction_mark: {
      threshold: _numOrNull(el('tsThreshold')?.value),
      color: (el('tsColor')?.value || '#00B0F0')
    },
    scales: {
      W: {
        // min..opt = нормальный диапазон температуры
        min: _num(el('wMin')?.value, -1),
        opt: _num(el('wOpt')?.value, 1),
        max: _safeMax(el('wOpt')?.value, 1),
        colors: {
          min: (el('wCMin')?.value || '#1CBCF2'),
          opt: (el('wCOpt')?.value || '#00FF00'),
          max: (el('wCMax')?.value || '#F3919B'),
        }
      },
      X: {
        min: _num(el('xMin')?.value, -1),
        opt: _num(el('xOpt')?.value, 1),
        max: _safeMax(el('xOpt')?.value, 1),
        colors: {
          min: (el('xCMin')?.value || '#1CBCF2'),
          opt: (el('xCOpt')?.value || '#00FF00'),
          max: (el('xCMax')?.value || '#F3919B'),
        }
      },
      Y: {
        min: _num(el('yMin')?.value, -1),
        opt: _num(el('yOpt')?.value, 1),
        max: _safeMax(el('yOpt')?.value, 1),
        colors: {
          min: (el('yCMin')?.value || '#1CBCF2'),
          opt: (el('yCOpt')?.value || '#00FF00'),
          max: (el('yCMax')?.value || '#F3919B'),
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

  const dm = s.discharge_mark || {};
  _set('tdThreshold', dm.threshold);
  _set('tdColor', dm.color || '#FFC000');

  const sm = s.suction_mark || {};
  _set('tsThreshold', sm.threshold);
  _set('tsColor', sm.color || '#00B0F0');

  const sc = s.scales || {};
  const w = sc.W || {}; const x = sc.X || {}; const y = sc.Y || {};
  // min..opt = нормальный диапазон температуры
  _set('wMin', w.min ?? -1); _set('wOpt', w.opt ?? 1);
  _set('xMin', x.min ?? -1); _set('xOpt', x.opt ?? 1);
  _set('yMin', y.min ?? -1); _set('yOpt', y.opt ?? 1);

  const wc = (w.colors || {});
  const xc = (x.colors || {});
  const yc = (y.colors || {});
  _set('wCMin', wc.min || '#1CBCF2'); _set('wCOpt', wc.opt || '#00FF00'); _set('wCMax', wc.max || '#F3919B');
  _set('xCMin', xc.min || '#1CBCF2'); _set('xCOpt', xc.opt || '#00FF00'); _set('xCMax', xc.max || '#F3919B');
  _set('yCMin', yc.min || '#1CBCF2'); _set('yCOpt', yc.opt || '#00FF00'); _set('yCMax', yc.max || '#F3919B');

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
