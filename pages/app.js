(() => {
  "use strict";

  const SCALE_EXPONENT_KEY = "gdm-warehouse-scale-exponent-v1";
  const DECIMAL_PLACES_KEY = "gdm-warehouse-decimal-places-v1";
  const PAGE_SIZE = 100;
  const state = {
    database: null, sessions: [], session: null, plots: [], weights: [], selected: null,
    serialPort: null, serialReader: null, readLoop: null, keepReading: false, serialBuffer: "",
    serialFlushTimer: null, rawScaleWeight: null, scaleExponent: loadScaleExponent(), decimalPlaces: loadDecimalPlaces(), page: 1,
    exactPendingWeight: null, weightEdited: false,
    lotSite: "", lotStorage: "", carriedLotSite: "", carriedStorage: "", keepLotContext: true, hasLotContext: false,
    scanAlert: null, scanAlertOpenedAt: 0,
    selectedTrials: new Set(), sortKey: "", sortDirection: "asc", filteredPlots: [],
  };
  const byFeid = new Map();
  const byUuid = new Map();
  let toastTimer;
  let scanAudioContext;
  const $ = (id) => document.getElementById(id);
  const refs = {
    sessionSelect: $("session-select"), renameSession: $("rename-session"), deleteSession: $("delete-session"), newSession: $("new-session"),
    trialSlicerButton: $("trial-slicer-button"), trialSlicerMenu: $("trial-slicer-menu"), trialSlicerOptions: $("trial-slicer-options"), trialSelectAll: $("trial-select-all"), trialClearAll: $("trial-clear-all"),
    weighingView: $("weighing-view"), dashboardView: $("dashboard-view"),
    importData: $("import-data"), dataFile: $("data-file"), exportExcel: $("export-excel"), exportCsv: $("export-csv"), datasetNote: $("dataset-note"),
    scanForm: $("scan-form"), scanMode: $("scan-mode"), scanValue: $("scan-value"), scanError: $("scan-error"),
    plotCard: $("plot-card"), weightForm: $("weight-form"), weight: $("plot-weight"), lotSite: $("lot-site"), lotStorage: $("lot-storage"), keepLotContext: $("keep-lot-context"), saveButton: $("save-button"), existingBadge: $("existing-badge"),
    recentList: $("recent-list"), recentEmpty: $("recent-empty"), toast: $("toast"),
    scanAlert: $("scan-alert"), scanAlertTitle: $("scan-alert-title"), scanAlertMessage: $("scan-alert-message"), scanAlertAction: $("scan-alert-action"), scanAlertUpdate: $("scan-alert-update"), scanAlertHint: $("scan-alert-hint"),
    connectScale: $("connect-scale"), baudRate: $("baud-rate"), scaleFactor: $("scale-factor"), decimalPlaces: $("decimal-places"), scaleStatus: $("scale-status"), scaleWeight: $("scale-weight"), scaleReadingNote: $("scale-reading-note"), serialHelp: $("serial-help"),
    tableSearch: $("table-search"), trialFilter: $("trial-filter"), locationFilter: $("location-filter"), statusFilter: $("status-filter"), tableBody: $("plot-table-body"),
    pagePrev: $("page-prev"), pageNext: $("page-next"), pageNumber: $("page-number"), pageSummary: $("page-summary"), tableCount: $("table-count"),
    pagePrevTop: $("page-prev-top"), pageNextTop: $("page-next-top"), pageNumberTop: $("page-number-top"), pageSummaryTop: $("page-summary-top"),
    dashboardExportExcel: $("dashboard-export-excel"), dashboardExportCsv: $("dashboard-export-csv"), dashboardLotsExcel: $("dashboard-lots-excel"), dashboardLotsCsv: $("dashboard-lots-csv"),
  };

  function loadScaleExponent() {
    const value = Number(localStorage.getItem(SCALE_EXPONENT_KEY) || 0);
    return Number.isInteger(value) && value >= -10 && value <= 10 ? value : 0;
  }
  function loadDecimalPlaces() {
    const value = Number(localStorage.getItem(DECIMAL_PLACES_KEY) || 0);
    return Number.isInteger(value) && value >= 0 && value <= 6 ? value : 0;
  }
  function normalize(value) { return window.GdmWeighingUtils.normalize(value); }
  function parseWeight(value) { const text = String(value || "").trim(); return text ? Number(text.replace(",", ".")) : NaN; }
  function formatNumber(value, places = state.decimalPlaces) { return new Intl.NumberFormat("en-US", { minimumFractionDigits: places, maximumFractionDigits: places }).format(value); }
  function formatInputNumber(value, places = state.decimalPlaces) { return new Intl.NumberFormat("en-US", { useGrouping: false, minimumFractionDigits: places, maximumFractionDigits: places }).format(value); }
  function formatRawNumber(value) { return new Intl.NumberFormat("en-US", { maximumFractionDigits: 12 }).format(value); }
  function plotDisplayName(plot) { const name = plot.obsName || plot.feid || "—"; return /^plot\b/i.test(name) ? name : `Plot ${name}`; }
  function formatDateTime(value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? "Unavailable" : new Intl.DateTimeFormat("en-US", { dateStyle: "medium", timeStyle: "short" }).format(date);
  }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]); }
  function text(id, value) { $(id).textContent = value ?? "—"; }
  function showToast(message, error = false) {
    clearTimeout(toastTimer);
    refs.toast.textContent = message;
    refs.toast.className = error ? "toast toast--error" : "toast";
    refs.toast.hidden = false;
    toastTimer = setTimeout(() => { refs.toast.hidden = true; }, 5200);
  }
  function showScanError(message) {
    refs.scanError.textContent = `⚠ ${message}`;
    refs.scanError.hidden = false;
    refs.scanValue.focus(); refs.scanValue.select();
  }
  function closeScanAlert(focusTarget = "scan") {
    state.scanAlert = null;
    refs.scanAlert.hidden = true;
    setTimeout(() => {
      const target = focusTarget === "weight" ? refs.weight : refs.scanValue;
      target.focus(); target.select();
    }, 0);
  }
  function keepExistingWeight() {
    state.selected = null; state.exactPendingWeight = null; state.weightEdited = false;
    refs.plotCard.hidden = true; refs.scanValue.value = ""; refs.weight.value = "";
    closeScanAlert();
  }
  function updateExistingWeight() {
    closeScanAlert("weight");
  }
  function showScanAlert(kind, title, message, actionLabel) {
    state.scanAlert = kind;
    state.scanAlertOpenedAt = performance.now();
    refs.scanAlert.className = `scan-alert scan-alert--${kind}`;
    refs.scanAlertTitle.textContent = title;
    refs.scanAlertMessage.textContent = message;
    refs.scanAlertUpdate.hidden = kind !== "existing";
    refs.scanAlertAction.textContent = kind === "existing" ? "Keep current PW" : actionLabel;
    refs.scanAlertHint.textContent = kind === "existing" ? "Enter keeps the current PW · choose Update PW to replace it" : "Press Enter or click the button to continue";
    refs.scanAlert.hidden = false;
    setTimeout(() => refs.scanAlertAction.focus(), 0);
  }
  function weightsMap() { return window.GdmWeighingUtils.byUuid(state.weights); }
  function trialNames() { return [...new Set(state.plots.map((plot) => plot.entityName || "Unnamed trial"))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })); }
  function activePlots() { return state.plots.filter((plot) => state.selectedTrials.has(plot.entityName || "Unnamed trial")); }
  function superscript(value) { return String(value).replace(/\d/g, (digit) => "⁰¹²³⁴⁵⁶⁷⁸⁹"[Number(digit)]); }
  function scaleFactorLabel(exponent = state.scaleExponent) { return exponent < 0 ? `× 10${superscript(Math.abs(exponent))}` : exponent > 0 ? `÷ 10${superscript(exponent)}` : "raw value"; }
  function playScanTone(kind) {
    try {
      const AudioContextClass = window.AudioContext || window.webkitAudioContext;
      if (!AudioContextClass) return;
      const context = scanAudioContext || new AudioContextClass();
      scanAudioContext = context;
      if (context.state === "suspended") void context.resume();
      const patterns = kind === "found"
        ? [[720, 0, .32], [980, .36, .34]]
        : kind === "existing"
          ? [[520, 0, .45], [400, .58, .52]]
          : [[240, 0, .38], [165, .48, .4], [240, .98, .42]];
      for (const [frequency, offset, duration] of patterns) {
        const gain = context.createGain();
        const oscillator = context.createOscillator(); oscillator.type = "sine"; oscillator.frequency.value = frequency; oscillator.connect(gain); gain.connect(context.destination);
        gain.gain.setValueAtTime(.0001, context.currentTime + offset); gain.gain.exponentialRampToValueAtTime(.18, context.currentTime + offset + .025); gain.gain.exponentialRampToValueAtTime(.0001, context.currentTime + offset + duration);
        oscillator.start(context.currentTime + offset); oscillator.stop(context.currentTime + offset + duration + .01);
      }
    } catch { /* Audio feedback is optional. */ }
  }

  async function refreshSessionList() {
    state.sessions = await window.GdmWeighingStore.listSessions(state.database);
    refs.sessionSelect.innerHTML = state.sessions.length
      ? state.sessions.map((session) => `<option value="${escapeHtml(session.id)}">${escapeHtml(session.name)}</option>`).join("")
      : '<option value="">No active session</option>';
    refs.sessionSelect.value = state.session?.id || "";
    refs.renameSession.disabled = !state.session;
    refs.deleteSession.disabled = !state.session;
  }

  async function loadSession(id) {
    state.session = id ? await window.GdmWeighingStore.getSession(state.database, id) : null;
    state.plots = state.session?.plots || [];
    state.weights = state.session ? await window.GdmWeighingStore.getWeights(state.database, state.session.id) : [];
    const lotContext = state.session
      ? await window.GdmWeighingStore.getLotContext(state.database, state.session.id)
      : { keepForNext: true, hasValue: false, lotSite: "", storage: "" };
    state.keepLotContext = lotContext.keepForNext; state.hasLotContext = lotContext.hasValue;
    state.carriedLotSite = lotContext.lotSite; state.carriedStorage = lotContext.storage;
    state.selected = null;
    state.scanAlert = null;
    refs.scanAlert.hidden = true;
    state.exactPendingWeight = null;
    state.weightEdited = false;
    state.page = 1;
    state.sortKey = "";
    state.selectedTrials = new Set(trialNames());
    byFeid.clear(); byUuid.clear();
    for (const plot of state.plots) { byFeid.set(normalize(plot.feid), plot); byUuid.set(normalize(plot.uuid), plot); }
    refs.plotCard.hidden = true;
    refs.scanValue.value = "";
    refs.weight.value = "";
    refs.lotSite.value = ""; refs.lotStorage.value = ""; refs.keepLotContext.checked = state.keepLotContext;
    refs.weight.placeholder = formatInputNumber(0);
    await refreshSessionList();
    renderAll();
    refs.scanValue.focus();
  }

  async function setActiveSession(id) {
    await window.GdmWeighingStore.setActiveSession(state.database, id || null);
    await loadSession(id || null);
  }

  function renderAll() {
    renderTrialSlicer();
    const visiblePlots = activePlots();
    const overall = window.GdmWeighingUtils.overallProgress(visiblePlots, state.weights);
    text("header-count", `${overall.completed} of ${overall.total}`);
    text("header-percent", `${overall.percent}%`);
    refs.datasetNote.textContent = state.session
      ? `Active session: ${state.session.name} · ${visiblePlots.length} of ${state.plots.length} plots in the trial filter · last changed ${formatDateTime(state.session.updatedAt)}`
      : "No weighing session loaded. Import a workbook to begin.";
    refs.exportExcel.disabled = !state.session;
    refs.exportCsv.disabled = !state.session;
    renderRecent();
    renderDashboard();
  }

  function renderRecent() {
    const selected = state.selectedTrials;
    const recent = state.weights.filter((record) => selected.has(record.entityName || "Unnamed trial")).sort((a, b) => Date.parse(b.weighedAt || b.updatedAt) - Date.parse(a.weighedAt || a.updatedAt)).slice(0, 10);
    refs.recentEmpty.hidden = recent.length > 0;
    refs.recentList.hidden = recent.length === 0;
    refs.recentList.innerHTML = recent.map((record) => `
      <article class="recent-item"><div class="recent-item__plot"><strong>Plot ${escapeHtml(record.obsName || "—")}</strong><span>${escapeHtml(record.entityName || "Unnamed trial")} · FEID ${escapeHtml(record.feid || "—")}</span></div><div class="recent-item__weight"><small>PW</small><strong>${escapeHtml(formatNumber(Number(record.weight)))}</strong></div><time datetime="${escapeHtml(record.weighedAt || "")}">${escapeHtml(formatDateTime(record.weighedAt || record.updatedAt))}</time></article>`).join("");
  }

  function donutCards(items) {
    if (!items.length) return '<p class="table-empty">No data available for this session.</p>';
    return items.map((item) => `<article class="donut-card"><div class="donut" style="--percent:${item.percent}" role="img" aria-label="${escapeHtml(item.label)}: ${item.percent}% complete"><strong>${item.percent}%</strong></div><div><h3 title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</h3><p><strong>${item.completed}</strong> weighed · <strong>${item.remaining}</strong> pending</p></div></article>`).join("");
  }

  function fillFilter(select, values, label) {
    const previous = select.value;
    select.innerHTML = `<option value="">${label}</option>${values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}`;
    if (values.includes(previous)) select.value = previous;
  }

  function renderTrialSlicer() {
    const names = trialNames();
    for (const selected of [...state.selectedTrials]) if (!names.includes(selected)) state.selectedTrials.delete(selected);
    refs.trialSlicerOptions.innerHTML = names.length ? names.map((name) => `<label><input type="checkbox" value="${escapeHtml(name)}" ${state.selectedTrials.has(name) ? "checked" : ""} /><span>${escapeHtml(name)}</span></label>`).join("") : '<p class="table-empty">No trials loaded.</p>';
    const count = state.selectedTrials.size;
    refs.trialSlicerButton.textContent = !names.length ? "Trials: None" : count === names.length ? "Trials: All" : `Trials: ${count} of ${names.length}`;
    refs.trialSlicerButton.disabled = !names.length;
  }

  function renderDashboard() {
    const visiblePlots = activePlots();
    const overall = window.GdmWeighingUtils.overallProgress(visiblePlots, state.weights);
    text("summary-total", overall.total); text("summary-weighed", overall.completed); text("summary-pending", overall.remaining); text("summary-percent", `${overall.percent}%`);
    $("summary-progress").style.width = `${overall.percent}%`;
    const trials = window.GdmWeighingUtils.groupProgress(visiblePlots, state.weights, "trial");
    const locations = window.GdmWeighingUtils.groupProgress(visiblePlots, state.weights, "location");
    $("trial-donuts").innerHTML = donutCards(trials);
    $("location-donuts").innerHTML = donutCards(locations);
    fillFilter(refs.trialFilter, [...new Set(visiblePlots.map((plot) => plot.entityName).filter(Boolean))].sort(), "All trials");
    fillFilter(refs.locationFilter, [...new Set(visiblePlots.map((plot) => plot.location).filter(Boolean))].sort(), "All locations");
    renderTable();
  }

  function renderTable() {
    const records = weightsMap();
    const search = normalize(refs.tableSearch.value);
    const trial = refs.trialFilter.value;
    const location = refs.locationFilter.value;
    const status = refs.statusFilter.value;
    const filtered = activePlots().filter((plot) => {
      const weighed = records.has(normalize(plot.uuid));
      const haystack = normalize([plot.feid, plot.uuid, plot.seasonYear, plot.obsName, plot.entityName, plot.gerName, plot.location].join(" "));
      return (!search || haystack.includes(search)) && (!trial || plot.entityName === trial) && (!location || plot.location === location)
        && (!status || (status === "weighed" ? weighed : !weighed));
    });
    if (state.sortKey) {
      const collator = new Intl.Collator("en-US", { numeric: true, sensitivity: "base" });
      filtered.sort((left, right) => {
        const leftRecord = records.get(normalize(left.uuid));
        const rightRecord = records.get(normalize(right.uuid));
        const value = (plot, record) => state.sortKey === "status" ? (record ? "Weighed" : "Pending") : state.sortKey === "weight" ? (record?.weight ?? Number.POSITIVE_INFINITY) : state.sortKey === "weighedAt" ? (record?.weighedAt || record?.updatedAt || "") : (plot[state.sortKey] ?? "");
        const a = value(left, leftRecord); const b = value(right, rightRecord);
        const result = typeof a === "number" && typeof b === "number" ? a - b : collator.compare(String(a), String(b));
        return state.sortDirection === "desc" ? -result : result;
      });
    }
    state.filteredPlots = filtered;
    const pages = Math.max(Math.ceil(filtered.length / PAGE_SIZE), 1);
    state.page = Math.min(Math.max(state.page, 1), pages);
    const start = (state.page - 1) * PAGE_SIZE;
    const rows = filtered.slice(start, start + PAGE_SIZE);
    refs.tableBody.innerHTML = rows.length ? rows.map((plot) => {
      const record = records.get(normalize(plot.uuid));
      return `<tr><td><span class="status-pill ${record ? "status-pill--done" : "status-pill--pending"}">${record ? "Weighed" : "Pending"}</span></td><td>${escapeHtml(plot.seasonYear || "—")}</td><td title="${escapeHtml(plot.entityName)}">${escapeHtml(plot.entityName)}</td><td>${escapeHtml(plot.obsName)}</td><td>${escapeHtml(plot.feid)}</td><td title="${escapeHtml(plot.uuid)}">${escapeHtml(plot.uuid)}</td><td>${escapeHtml(plot.block)}</td><td>${escapeHtml(plot.entryCode)}</td><td>${escapeHtml(plot.row)}</td><td>${escapeHtml(plot.column)}</td><td title="${escapeHtml(plot.gerName)}">${escapeHtml(plot.gerName || "—")}</td><td>${record ? escapeHtml(formatNumber(Number(record.weight))) : "—"}</td><td>${record ? escapeHtml(formatDateTime(record.weighedAt || record.updatedAt)) : "—"}</td></tr>`;
    }).join("") : '<tr><td colspan="13" class="table-empty">No plots match the current filters.</td></tr>';
    refs.tableCount.textContent = `${filtered.length} record${filtered.length === 1 ? "" : "s"}`;
    refs.pageSummary.textContent = filtered.length ? `${start + 1}–${Math.min(start + PAGE_SIZE, filtered.length)} of ${filtered.length}` : "0–0 of 0";
    refs.pageNumber.textContent = `Page ${state.page} of ${pages}`;
    refs.pagePrev.disabled = state.page <= 1;
    refs.pageNext.disabled = state.page >= pages;
    refs.pageSummaryTop.textContent = refs.pageSummary.textContent;
    refs.pageNumberTop.textContent = refs.pageNumber.textContent;
    refs.pagePrevTop.disabled = refs.pagePrev.disabled;
    refs.pageNextTop.disabled = refs.pageNext.disabled;
    refs.dashboardExportExcel.disabled = !state.session || !filtered.length;
    refs.dashboardExportCsv.disabled = !state.session || !filtered.length;
    const hasWeighed = filtered.some((plot) => records.has(normalize(plot.uuid)));
    refs.dashboardLotsExcel.disabled = !state.session || !hasWeighed;
    refs.dashboardLotsCsv.disabled = !state.session || !hasWeighed;
    document.querySelectorAll("[data-sort]").forEach((button) => {
      if (button.dataset.sort === state.sortKey) button.dataset.direction = state.sortDirection;
      else delete button.dataset.direction;
    });
  }

  function selectPlot(plot) {
    state.selected = plot;
    refs.scanError.hidden = true; refs.plotCard.hidden = false; refs.scanValue.value = "";
    text("entity-name", plot.entityName); text("obs-name", plot.obsName); text("ger-name", plot.gerName || "—");
    text("location", `⌖ Location: ${plot.location || "Unspecified"} · Trial site: ${plot.site || "Unspecified"}`);
    text("season-year", plot.seasonYear || "—"); text("block", plot.block || "—"); text("entry-code", plot.entryCode || "—"); text("row", plot.row || "—"); text("column", plot.column || "—"); text("feid", plot.feid); text("uuid", plot.uuid);
    const existing = weightsMap().get(normalize(plot.uuid));
    state.exactPendingWeight = existing ? Number(existing.weight) : null;
    state.weightEdited = false;
    refs.weight.value = existing ? formatInputNumber(Number(existing.weight)) : "";
    if (existing) {
      refs.lotSite.value = String(existing.lotSite || "");
      refs.lotStorage.value = existing.storage || "";
    } else if (state.keepLotContext && state.hasLotContext) {
      refs.lotSite.value = state.carriedLotSite; refs.lotStorage.value = state.carriedStorage;
    } else {
      refs.lotSite.value = ""; refs.lotStorage.value = "";
    }
    refs.existingBadge.hidden = !existing;
    refs.existingBadge.textContent = existing ? `Already weighed: PW ${formatNumber(existing.weight)}` : "";
    refs.saveButton.textContent = existing ? "✓ Update PW" : "✓ Save PW";
    setTimeout(() => refs.scanValue.focus(), 0);
  }
  function clearSelection() {
    state.selected = null; state.exactPendingWeight = null; state.weightEdited = false; refs.plotCard.hidden = true; refs.scanValue.value = ""; refs.weight.value = ""; refs.lotSite.value = ""; refs.lotStorage.value = "";
    setTimeout(() => refs.scanValue.focus(), 0);
  }
  function refreshPrecisionDisplay() {
    refs.weight.placeholder = formatInputNumber(0);
    if (state.rawScaleWeight !== null) {
      const scaled = state.rawScaleWeight / (10 ** state.scaleExponent);
      refs.scaleWeight.textContent = formatNumber(scaled);
    }
    if (state.selected) {
      const existing = weightsMap().get(normalize(state.selected.uuid));
      refs.existingBadge.textContent = existing ? `Already weighed: PW ${formatNumber(Number(existing.weight))}` : "";
      if (!state.weightEdited && state.exactPendingWeight !== null) refs.weight.value = formatInputNumber(state.exactPendingWeight);
    }
    renderRecent();
    renderDashboard();
  }
  async function saveCurrentWeight() {
    if (!state.selected || !state.session) return;
    const weight = !state.weightEdited && state.exactPendingWeight !== null ? state.exactPendingWeight : parseWeight(refs.weight.value);
    if (!Number.isFinite(weight) || weight < 0) { showToast("Enter or wait for a valid non-negative scale weight.", true); refs.weight.focus(); refs.weight.select(); return; }
    refs.saveButton.disabled = true;
    try {
      const plot = state.selected;
      const saved = await window.GdmWeighingStore.saveWeight(
        state.database, state.session.id, plot, weight, state.serialPort ? "serial" : "manual", undefined,
        { lotSite: refs.lotSite.value, storage: refs.lotStorage.value, keepForNext: state.keepLotContext },
      );
      state.weights = [saved, ...state.weights.filter((item) => normalize(item.uuid) !== normalize(saved.uuid))];
      if (state.keepLotContext) {
        state.carriedLotSite = saved.lotSite || ""; state.carriedStorage = saved.storage || ""; state.hasLotContext = true;
      } else {
        state.carriedLotSite = ""; state.carriedStorage = ""; state.hasLotContext = false;
      }
      state.session.updatedAt = new Date().toISOString();
      renderAll();
      showToast(`PW ${formatNumber(weight)} saved for plot ${plot.obsName}.`);
      clearSelection();
    } catch (error) { showToast(error instanceof Error ? error.message : "Could not save the weight.", true); }
    finally { refs.saveButton.disabled = false; }
  }

  async function importFile(file) {
    refs.importData.disabled = true; refs.importData.textContent = "Importing…";
    try {
      const result = await window.GdmPlotImport.parseExcelFile(file);
      const invalid = result.invalidRows.length + result.invalidWeightRows.length;
      if (!state.session || !result.importedWeights.length) {
        const session = await window.GdmWeighingStore.createSession(state.database, result.plots, result.fileName, undefined, result.importedWeights);
        await setActiveSession(session.id);
        showToast(`Session created with ${result.plots.length} plots and ${result.importedWeights.length} weights${invalid ? `; ${invalid} invalid row(s) skipped` : ""}.`);
        return;
      }
      const merge = window.GdmWeighingUtils.prepareMerge(state.plots, state.weights, result.importedWeights);
      let includeUnresolved = false;
      if (merge.unresolved.length) {
        includeUnresolved = window.confirm(`${merge.unresolved.length} conflicting weight(s) do not have comparable timestamps. Select OK to use the imported values, or Cancel to keep the current session values.`);
      }
      const entries = [...merge.ready, ...(includeUnresolved ? merge.unresolved : [])];
      await window.GdmWeighingStore.saveWeights(state.database, state.session.id, entries);
      state.weights = await window.GdmWeighingStore.getWeights(state.database, state.session.id);
      state.session = await window.GdmWeighingStore.getSession(state.database, state.session.id);
      renderAll();
      const kept = merge.keptCurrent + (includeUnresolved ? 0 : merge.unresolved.length);
      showToast(`Partial results: ${entries.length} imported, ${kept} current kept, ${merge.unchanged} unchanged, ${merge.ignored} unmatched, ${invalid} invalid.`);
    } catch (error) { showToast(error instanceof Error ? error.message : "Could not import the file.", true); }
    finally { refs.importData.disabled = false; refs.importData.textContent = "⇧ Import Excel / CSV"; refs.dataFile.value = ""; }
  }

  function exportData(format, filteredOnly = false) {
    if (!state.session) { showToast("Load a weighing session before exporting.", true); return; }
    const exportPlots = filteredOnly ? state.filteredPlots : state.plots;
    if (!exportPlots.length) { showToast("No plot records match the current filters.", true); return; }
    try { const count = window.GdmWeighingUtils.exportSession(state.session, state.weights, format, exportPlots, state.decimalPlaces); showToast(`${count} ${filteredOnly ? "filtered " : ""}plot records exported to ${format === "xlsx" ? "Excel" : "CSV"} with ${state.decimalPlaces} decimal place${state.decimalPlaces === 1 ? "" : "s"}.`); }
    catch (error) { showToast(error instanceof Error ? error.message : "Could not export the session.", true); }
  }

  function exportLots(format) {
    if (!state.session) { showToast("Load a weighing session before exporting lots.", true); return; }
    try {
      const count = window.GdmWeighingUtils.exportLots(state.session, state.weights, format, state.filteredPlots, state.decimalPlaces);
      showToast(`${count} filtered lot record${count === 1 ? "" : "s"} exported to ${format === "xlsx" ? "Excel" : "CSV"}.`);
    } catch (error) { showToast(error instanceof Error ? error.message : "Could not export lots.", true); }
  }

  function parseScaleWeight(rawLine) {
    const matches = String(rawLine || "").replace(/\u0000/g, " ").trim().match(/[-+]?\d+(?:[.,]\d+)?/g);
    if (!matches?.length) return null;
    const value = Number(matches[matches.length - 1].replace(",", "."));
    return Number.isFinite(value) && value >= 0 ? value : null;
  }
  function applyScaleWeight(rawValue, rawLine) {
    state.rawScaleWeight = rawValue;
    const value = rawValue / (10 ** state.scaleExponent);
    refs.scaleWeight.textContent = formatNumber(value); refs.scaleWeight.classList.add("is-live");
    const factor = state.scaleExponent ? ` · ${scaleFactorLabel()}` : "";
    refs.scaleReadingNote.textContent = state.selected ? `PW filled automatically${factor}` : `Scan a plot to apply${factor}`;
    refs.scaleWeight.title = `Raw reading: ${formatRawNumber(rawValue)}${rawLine ? ` (${String(rawLine).trim()})` : ""}`;
    if (state.selected) { state.exactPendingWeight = value; state.weightEdited = false; refs.weight.value = formatInputNumber(value); }
  }
  function consumeSerialText(chunk) {
    state.serialBuffer += chunk;
    const lines = state.serialBuffer.split(/\r\n|\n|\r/); state.serialBuffer = lines.pop() || "";
    for (const line of lines) { const value = parseScaleWeight(line); if (value !== null) applyScaleWeight(value, line); }
    clearTimeout(state.serialFlushTimer);
    state.serialFlushTimer = setTimeout(() => { const value = parseScaleWeight(state.serialBuffer); if (value !== null) applyScaleWeight(value, state.serialBuffer); state.serialBuffer = ""; }, 180);
  }
  async function readFromScale() {
    const decoder = new TextDecoder();
    try {
      while (state.keepReading && state.serialPort?.readable) {
        state.serialReader = state.serialPort.readable.getReader();
        try { while (state.keepReading) { const { value, done } = await state.serialReader.read(); if (done) break; if (value) consumeSerialText(decoder.decode(value, { stream: true })); } }
        finally { state.serialReader.releaseLock(); state.serialReader = null; }
      }
    } catch (error) { if (state.keepReading) showToast(error instanceof Error ? error.message : "Scale reading stopped.", true); }
  }
  function setScaleConnected(connected, label = "Not connected") {
    refs.connectScale.dataset.connected = String(connected); refs.connectScale.textContent = connected ? "Disconnect" : "Connect scale"; refs.baudRate.disabled = connected; refs.scaleStatus.textContent = label;
    if (!connected) { state.rawScaleWeight = null; refs.scaleWeight.textContent = "—"; refs.scaleWeight.classList.remove("is-live"); refs.scaleReadingNote.textContent = "Waiting for connection"; }
  }
  async function disconnectScale(quiet = false) {
    state.keepReading = false; clearTimeout(state.serialFlushTimer);
    try { await state.serialReader?.cancel(); } catch { /* already closed */ }
    try { await state.readLoop; } catch { /* surfaced above */ }
    try { await state.serialPort?.close(); } catch { /* disconnected device */ }
    state.serialReader = null; state.serialPort = null; state.readLoop = null; state.serialBuffer = ""; setScaleConnected(false);
    if (!quiet) showToast("Scale disconnected.");
  }
  async function toggleScaleConnection() {
    if (state.serialPort) { await disconnectScale(); return; }
    if (!("serial" in navigator)) { showToast("Use Google Chrome or Microsoft Edge to connect through a COM port.", true); return; }
    try {
      const port = await navigator.serial.requestPort(); await port.open({ baudRate: Number(refs.baudRate.value) });
      state.serialPort = port; state.keepReading = true;
      const info = port.getInfo?.() || {};
      const identifiers = [info.usbVendorId ? `VID ${info.usbVendorId.toString(16).toUpperCase().padStart(4, "0")}` : "", info.usbProductId ? `PID ${info.usbProductId.toString(16).toUpperCase().padStart(4, "0")}` : ""].filter(Boolean).join(" · ");
      setScaleConnected(true, identifiers ? `Connected · ${identifiers}` : "Connected to selected port");
      refs.scaleReadingNote.textContent = "Waiting for scale weight"; state.readLoop = readFromScale(); showToast("Scale connected. Waiting for a weight reading.");
    } catch (error) { if (error?.name !== "NotFoundError") showToast(error instanceof Error ? error.message : "Could not connect to the scale.", true); await disconnectScale(true); }
  }

  refs.scanForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (state.scanAlert) return;
    if (!state.session) { showToast("Import a workbook before weighing plots.", true); return; }
    const scannedValue = refs.scanValue.value.trim();
    const code = normalize(refs.scanValue.value);
    const selectedCode = state.selected ? normalize(refs.scanMode.value === "uuid" ? state.selected.uuid : state.selected.feid) : "";
    if (state.selected && (!code || code === selectedCode)) { void saveCurrentWeight(); return; }
    const plot = refs.scanMode.value === "uuid" ? byUuid.get(code) : byFeid.get(code);
    if (!plot || !state.selectedTrials.has(plot.entityName || "Unnamed trial")) {
      const message = plot
        ? `${plotDisplayName(plot)} belongs to trial ${plot.entityName || "Unnamed trial"}, which is excluded by the current trial filter.`
        : `${refs.scanMode.value.toUpperCase()} ${scannedValue || "—"} was not found in this session.`;
      state.selected = null; state.exactPendingWeight = null; state.weightEdited = false; refs.plotCard.hidden = true; refs.weight.value = "";
      playScanTone("attention"); showScanError(message); showScanAlert("unavailable", "Plot unavailable", message, "Scan again"); return;
    }
    const record = weightsMap().get(normalize(plot.uuid));
    playScanTone(record ? "existing" : "found");
    selectPlot(plot);
    if (record) showScanAlert("existing", "Plot already weighed", `${plotDisplayName(plot)} already has PW ${formatNumber(Number(record.weight))}. Choose whether to update it or keep the current value.`, "Keep current PW");
  });
  refs.scanMode.addEventListener("change", () => { refs.scanValue.placeholder = refs.scanMode.value === "feid" ? "Scan or enter the FEID" : "Scan or enter the UUID"; refs.scanValue.value = ""; refs.scanError.hidden = true; if (state.scanAlert) closeScanAlert(); else refs.scanValue.focus(); });
  refs.scanAlertAction.addEventListener("click", () => { if (state.scanAlert === "existing") keepExistingWeight(); else closeScanAlert(); });
  refs.scanAlertUpdate.addEventListener("click", updateExistingWeight);
  document.addEventListener("keydown", (event) => {
    if (!state.scanAlert) return;
    if (event.key === "Tab") {
      event.preventDefault();
      if (state.scanAlert !== "existing") { refs.scanAlertAction.focus(); return; }
      (document.activeElement === refs.scanAlertAction ? refs.scanAlertUpdate : refs.scanAlertAction).focus();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault(); event.stopPropagation();
      if (performance.now() - state.scanAlertOpenedAt < 700) return;
      if (state.scanAlert === "existing") {
        if (document.activeElement === refs.scanAlertUpdate) updateExistingWeight(); else keepExistingWeight();
      } else closeScanAlert();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault(); event.stopPropagation();
      if (state.scanAlert === "existing") keepExistingWeight(); else closeScanAlert();
    }
  }, true);
  refs.weightForm.addEventListener("submit", (event) => { event.preventDefault(); void saveCurrentWeight(); });
  refs.weight.addEventListener("input", () => { state.weightEdited = true; state.exactPendingWeight = null; });
  refs.keepLotContext.addEventListener("change", async () => {
    state.keepLotContext = refs.keepLotContext.checked;
    if (!state.keepLotContext) { state.hasLotContext = false; state.carriedLotSite = ""; state.carriedStorage = ""; }
    if (!state.session) return;
    try {
      await window.GdmWeighingStore.setLotContext(state.database, state.session.id, {
        keepForNext: state.keepLotContext, hasValue: state.keepLotContext ? state.hasLotContext : false,
        lotSite: state.keepLotContext ? state.carriedLotSite : "", storage: state.keepLotContext ? state.carriedStorage : "",
      });
    } catch (error) { showToast(error instanceof Error ? error.message : "Could not save the lot field preference.", true); }
  });
  refs.scaleFactor.value = String(state.scaleExponent);
  refs.scaleFactor.addEventListener("change", () => { state.scaleExponent = Number(refs.scaleFactor.value); localStorage.setItem(SCALE_EXPONENT_KEY, String(state.scaleExponent)); if (state.rawScaleWeight !== null) applyScaleWeight(state.rawScaleWeight, ""); showToast(state.scaleExponent ? `Scale factor applied: reading ${scaleFactorLabel()}.` : "Raw scale value selected."); });
  refs.decimalPlaces.value = String(state.decimalPlaces);
  refs.decimalPlaces.addEventListener("change", () => {
    state.decimalPlaces = Number(refs.decimalPlaces.value);
    localStorage.setItem(DECIMAL_PLACES_KEY, String(state.decimalPlaces));
    refreshPrecisionDisplay();
    showToast(`Display and exports set to ${state.decimalPlaces} decimal place${state.decimalPlaces === 1 ? "" : "s"}.`);
  });
  refs.connectScale.addEventListener("click", () => void toggleScaleConnection());
  refs.importData.addEventListener("click", () => refs.dataFile.click());
  refs.dataFile.addEventListener("change", () => { const [file] = refs.dataFile.files || []; if (file) void importFile(file); });
  refs.exportExcel.addEventListener("click", () => exportData("xlsx"));
  refs.exportCsv.addEventListener("click", () => exportData("csv"));
  refs.dashboardExportExcel.addEventListener("click", () => exportData("xlsx", true));
  refs.dashboardExportCsv.addEventListener("click", () => exportData("csv", true));
  refs.dashboardLotsExcel.addEventListener("click", () => exportLots("xlsx"));
  refs.dashboardLotsCsv.addEventListener("click", () => exportLots("csv"));
  refs.sessionSelect.addEventListener("change", () => void setActiveSession(refs.sessionSelect.value));
  refs.trialSlicerButton.addEventListener("click", () => {
    refs.trialSlicerMenu.hidden = !refs.trialSlicerMenu.hidden;
    refs.trialSlicerButton.setAttribute("aria-expanded", String(!refs.trialSlicerMenu.hidden));
  });
  refs.trialSlicerOptions.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[type="checkbox"]');
    if (!checkbox) return;
    if (checkbox.checked) state.selectedTrials.add(checkbox.value); else state.selectedTrials.delete(checkbox.value);
    if (state.selected && !state.selectedTrials.has(state.selected.entityName || "Unnamed trial")) clearSelection();
    state.page = 1; renderAll();
  });
  refs.trialSelectAll.addEventListener("click", () => { state.selectedTrials = new Set(trialNames()); state.page = 1; renderAll(); });
  refs.trialClearAll.addEventListener("click", () => { state.selectedTrials.clear(); clearSelection(); state.page = 1; renderAll(); });
  document.addEventListener("click", (event) => {
    if (!event.target.closest(".trial-slicer")) { refs.trialSlicerMenu.hidden = true; refs.trialSlicerButton.setAttribute("aria-expanded", "false"); }
  });
  refs.newSession.addEventListener("click", async () => {
    try {
      if (!state.session) { refs.dataFile.click(); return; }
      const session = await window.GdmWeighingStore.createSession(state.database, state.plots, state.session.sourceFileName);
      await setActiveSession(session.id); showToast(`New session “${session.name}” started.`);
    } catch (error) { showToast(error instanceof Error ? error.message : "Could not start a new session.", true); }
  });
  refs.renameSession.addEventListener("click", async () => {
    if (!state.session) return;
    const name = window.prompt("Enter a new name for this session:", state.session.name);
    if (name === null || !name.trim() || name.trim() === state.session.name) return;
    try { state.session = await window.GdmWeighingStore.renameSession(state.database, state.session.id, name); await refreshSessionList(); renderAll(); showToast("Session renamed."); }
    catch (error) { showToast(error instanceof Error ? error.message : "Could not rename the session.", true); }
  });
  refs.deleteSession.addEventListener("click", async () => {
    if (!state.session || !window.confirm(`Delete “${state.session.name}” and all of its locally saved weights? This cannot be undone.`)) return;
    try { await window.GdmWeighingStore.deleteSession(state.database, state.session.id); const sessions = await window.GdmWeighingStore.listSessions(state.database); await setActiveSession(sessions[0]?.id || null); showToast("Session deleted."); }
    catch (error) { showToast(error instanceof Error ? error.message : "Could not delete the session.", true); }
  });
  document.querySelectorAll("[data-tab]").forEach((button) => button.addEventListener("click", () => {
    document.querySelectorAll("[data-tab]").forEach((item) => item.classList.toggle("is-active", item === button));
    const dashboard = button.dataset.tab === "dashboard";
    refs.weighingView.hidden = dashboard; refs.dashboardView.hidden = !dashboard;
    if (dashboard) renderDashboard(); else setTimeout(() => refs.scanValue.focus(), 0);
  }));
  [refs.tableSearch, refs.trialFilter, refs.locationFilter, refs.statusFilter].forEach((control) => control.addEventListener("input", () => { state.page = 1; renderTable(); }));
  refs.pagePrev.addEventListener("click", () => { state.page -= 1; renderTable(); });
  refs.pageNext.addEventListener("click", () => { state.page += 1; renderTable(); });
  refs.pagePrevTop.addEventListener("click", () => { state.page -= 1; renderTable(); });
  refs.pageNextTop.addEventListener("click", () => { state.page += 1; renderTable(); });
  document.querySelectorAll("[data-sort]").forEach((button) => button.addEventListener("click", () => {
    const key = button.dataset.sort;
    if (state.sortKey === key) state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
    else { state.sortKey = key; state.sortDirection = "asc"; }
    state.page = 1; renderTable();
  }));

  if (!("serial" in navigator)) refs.serialHelp.textContent = "COM connection is not available in this browser. Open the app in Google Chrome or Microsoft Edge.";
  navigator.serial?.addEventListener("disconnect", (event) => { if (event.target === state.serialPort || event.port === state.serialPort) { void disconnectScale(true); showToast("The scale was disconnected from the computer.", true); } });

  (async function initialize() {
    try {
      const initialized = await window.GdmWeighingStore.init();
      state.database = initialized.database;
      await loadSession(initialized.activeSessionId);
      if (state.session?.name === "Imported legacy session") showToast("Existing browser data was migrated to “Imported legacy session”.");
    } catch (error) { showToast(error instanceof Error ? error.message : "Could not initialize browser storage.", true); renderAll(); }
  })();
})();
