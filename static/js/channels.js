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

      // Collect all selected items for dragging
      const list = el("channelList");
      const allItems = Array.from(list.querySelectorAll(".chanItem"));
      const selectedItems = allItems.filter(item => {
        const code = item.getAttribute("data-code");
        return selected.has(code);
      });

      // Store drag data
      dragData = {
        draggedCode: ch.code,
        selectedItems: selectedItems
      };

      li.classList.add("dragging");
      ev.dataTransfer.setData("text/plain", ch.code);
      ev.dataTransfer.effectAllowed = "move";
      if(ev.dataTransfer.setDragImage) {
        // If multiple items are selected, create a custom drag image indicating this
        if(selected.size > 1) {
          // Create a temporary element to show the number of selected items
          const dragPreview = document.createElement("div");
          dragPreview.style.position = "absolute";
          dragPreview.style.left = "-9999px";
          dragPreview.style.backgroundColor = "#eaf2ff";
          dragPreview.style.border = "2px solid #7aa7ff";
          dragPreview.style.borderRadius = "4px";
          dragPreview.style.padding = "8px";
          dragPreview.style.fontSize = "14px";
          dragPreview.style.fontFamily = "Arial, sans-serif";
          dragPreview.style.zIndex = "9999";
          dragPreview.innerHTML = `${selected.size} элементов`;

          document.body.appendChild(dragPreview);
          try {
            ev.dataTransfer.setDragImage(dragPreview, 10, 10);
            // Clean up the temporary element after a short delay
            setTimeout(() => {
              if(dragPreview.parentNode) {
                dragPreview.parentNode.removeChild(dragPreview);
              }
            }, 100);
          } catch(_) {
            // Fallback to original behavior
            try { ev.dataTransfer.setDragImage(li, 20, 12); } catch(_) {}
          }
        } else {
          try { ev.dataTransfer.setDragImage(li, 20, 12); } catch(_) {}
        }
      }
    });
    handle.addEventListener("dragend", () => {
      li.classList.remove("dragging");
      list.querySelectorAll(".chanItem").forEach(x => x.classList.remove("dragOver"));
      dragData = null; // Clear drag data
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

      // Check if we have multiple selected items to drag
      if(dragData && dragData.selectedItems && dragData.selectedItems.length > 1) {
        moveSelectedCodesBefore(draggedCode, ch.code);
      } else {
        moveCodeBefore(draggedCode, ch.code);
      }
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

// Move all selected codes before the target code
function moveSelectedCodesBefore(draggedCode, targetCode) {
  const list = el("channelList");
  const items = Array.from(list.querySelectorAll(".chanItem"));

  const target = items.find(li => li.getAttribute("data-code") === targetCode);
  if(!target) return;

  // Find all selected items
  const selectedItems = items.filter(item => {
    const code = item.getAttribute("data-code");
    return selected.has(code);
  });

  // If only one item is selected, use the original function
  if(selectedItems.length <= 1) {
    moveCodeBefore(draggedCode, targetCode);
    return;
  }

  // Sort selected items by their current position in the list to preserve their relative order
  const sortedSelectedItems = selectedItems.sort((a, b) => {
    return items.indexOf(a) - items.indexOf(b);
  });

  // Insert all selected items before the target
  sortedSelectedItems.forEach(item => {
    list.insertBefore(item, target);
  });

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
      // Новый тест загружен: на следующем построении графика диапазон должен
      // сброситься к полному диапазону текущего теста.
      _forceResetRangeOnNextPlot = true;
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


