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
  initRefrigerantUI();

  initChannelControlsHints();

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
  if(el('presetSelect') || el('btnSavePreset') || el('btnLoadPreset') || el('btnDeletePreset') || el('btnRefreshPresets')) {
    refreshPresets();
  }

  // init step UI
  updateStepUI();

  log("Ready. По умолчанию порядок каналов = как в файле. Сортировку можно менять, а перетаскивание сохраняет 'Мой порядок'.");
}

wire();
