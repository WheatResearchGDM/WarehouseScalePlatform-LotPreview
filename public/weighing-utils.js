(function (global) {
  "use strict";

  const sourceColumns = [
    ["ID", "id"], ["FEID", "feid"], ["UUID", "uuid"], ["Season year", "seasonYear"], ["Entity name", "entityName"],
    ["Trial type", "trialType"], ["Location", "location"], ["Row", "row"],
    ["Column", "column"], ["Entry code", "entryCode"], ["Block", "block"], ["(OBS) Name", "obsName"],
    ["GID", "gid"], ["(GER) Name", "gerName"], ["Initial plot", "initialPlot"], ["Final plot", "finalPlot"],
  ];
  const weightVariables = {
    plotWeight: { label: "Plot weight", aliases: ["Plot weight", "PW"] },
    seedWeight: { label: "Seed weight", aliases: ["Seed weight", "SEED.W"] },
  };
  function weightVariable(value) { return value === "seedWeight" ? "seedWeight" : "plotWeight"; }
  function weightLabel(sessionOrVariable) {
    const value = typeof sessionOrVariable === "string" ? sessionOrVariable : sessionOrVariable?.weightVariable;
    return weightVariables[weightVariable(value)].label;
  }

  const lotFields = ["lotSite", "storage", "lotLocation", "firstOrder", "secondOrder"];
  const lotCatalog = Object.freeze({
    "Passo Fundo": Object.freeze(["Câmara fria_1", "Câmara fria_2", "Galpão externo"]),
    "Chacabuco": Object.freeze(["Cámara fría_1", "Cámara fría_2", "Galpón exterior"]),
  });
  function lotDetails(value) { return Object.fromEntries(lotFields.map((key) => [key, String(value?.[key] ?? "")])); }
  function hasLotDetails(value) { return lotFields.some((key) => String(value?.[key] ?? "").trim()); }
  function selectLotDetails(plot, record, context) {
    if (record) return lotDetails(record);
    if (hasLotDetails(plot?.lotMetadata)) return lotDetails(plot.lotMetadata);
    return lotDetails(context?.keepForNext && context?.hasValue ? context : null);
  }
  function lotOptions(site, current, isStorage = false) {
    const values = isStorage ? (lotCatalog[site] || []) : Object.keys(lotCatalog);
    return [...(current && !values.includes(current) ? [{ value: current, label: `${current} (existing)` }] : []),
      ...values.map((value) => ({ value, label: value }))];
  }
  function importVariable(result, currentVariable, hasSession) {
    const columns = result.weightColumns;
    if (!hasSession && Boolean(columns.plotWeight) !== Boolean(columns.seedWeight)) return columns.seedWeight ? "seedWeight" : "plotWeight";
    return weightVariable(currentVariable);
  }
  function lotName(plot) {
    const parts = [plot.seasonYear, plot.entityName, plot.block, plot.obsName].map((value) => String(value ?? "").trim());
    return parts.some(Boolean) ? parts.join(" | ") : "";
  }

  function normalize(value) { return String(value || "").trim().toUpperCase(); }
  function decimalPlaces(value) {
    const parsed = Number(value);
    return Number.isInteger(parsed) && parsed >= 0 && parsed <= 6 ? parsed : 0;
  }
  function fixedDecimal(value, places) {
    const digits = decimalPlaces(places);
    const factor = 10 ** digits;
    const rounded = Math.round((Number(value) + Number.EPSILON) * factor) / factor;
    return rounded.toFixed(digits);
  }
  function byUuid(weights) { return new Map((weights || []).map((item) => [normalize(item.uuid), item])); }
  function groupProgress(plots, weights, key) {
    const weightMap = byUuid(weights);
    if (key === "location") {
      const locations = new Map();
      for (const plot of plots || []) {
        const location = plot.location || "Unspecified";
        const current = locations.get(location) || { key: location, label: location, total: 0, completed: 0 };
        current.total += 1;
        if (weightMap.has(normalize(plot.uuid))) current.completed += 1;
        locations.set(location, current);
      }
      return Array.from(locations.values()).map((item) => ({
        ...item, remaining: Math.max(item.total - item.completed, 0),
        percent: item.total ? Math.min(Math.round((item.completed / item.total) * 100), 100) : 0,
      }));
    }
    const trials = new Map();
    for (const plot of plots || []) {
      const trialKey = plot.entityName || "Unnamed trial";
      if (!trials.has(trialKey)) trials.set(trialKey, []);
      trials.get(trialKey).push(plot);
    }
    const trialProgress = Array.from(trials.entries()).map(([entityName, items]) => {
      const total = items.length;
      const completed = items.filter((item) => weightMap.has(normalize(item.uuid))).length;
      return {
        key: entityName, label: entityName, trialType: items[0]?.trialType || "—", location: items[0]?.location || "Unspecified",
        total, completed, remaining: Math.max(total - completed, 0),
        percent: total ? Math.min(Math.round((completed / total) * 100), 100) : 0,
      };
    });
    return trialProgress;
  }

  function overallProgress(plots, weights) {
    const trials = groupProgress(plots, weights, "trial");
    const total = trials.reduce((sum, item) => sum + item.total, 0);
    const completed = trials.reduce((sum, item) => sum + item.completed, 0);
    return { total, completed, remaining: Math.max(total - completed, 0), percent: total ? Math.min(Math.round((completed / total) * 100), 100) : 0 };
  }

  function prepareMerge(activePlots, currentWeights, importedWeights) {
    const plots = new Map((activePlots || []).map((plot) => [normalize(plot.uuid), plot]));
    const current = byUuid(currentWeights);
    const ready = [];
    const unresolved = [];
    let ignored = 0;
    let unchanged = 0;
    let keptCurrent = 0;
    const withMissingMetadata = (existing, incoming) => ({
      uuid: existing.uuid, weight: Number(existing.weight), weighedAt: existing.weighedAt || existing.updatedAt || "",
      weightVariable: weightVariable(existing.weightVariable || incoming.weightVariable),
      lotSite: existing.lotSite || incoming.lotSite || "", lotLocation: existing.lotLocation || incoming.lotLocation || "",
      storage: existing.storage || incoming.storage || "", firstOrder: existing.firstOrder || incoming.firstOrder || "",
      secondOrder: existing.secondOrder || incoming.secondOrder || "", source: existing.source || "import",
    });
    const withPreservedMetadata = (existing, incoming) => ({
      ...incoming,
      lotSite: incoming.lotSite || existing.lotSite || "",
      lotLocation: incoming.lotLocation || existing.lotLocation || "",
      storage: incoming.storage || existing.storage || "",
      firstOrder: incoming.firstOrder || existing.firstOrder || "",
      secondOrder: incoming.secondOrder || existing.secondOrder || "",
    });
    const hasMetadataToFill = (existing, incoming) => (!existing.lotSite && incoming.lotSite)
      || (!existing.lotLocation && incoming.lotLocation) || (!existing.storage && incoming.storage)
      || (!existing.firstOrder && incoming.firstOrder) || (!existing.secondOrder && incoming.secondOrder);
    for (const incoming of importedWeights || []) {
      const key = normalize(incoming.uuid);
      if (!plots.has(key)) { ignored += 1; continue; }
      const existing = current.get(key);
      if (!existing) { ready.push({ ...withPreservedMetadata(plots.get(key).lotMetadata || {}, incoming), source: "import" }); continue; }
      if (Number(existing.weight) === Number(incoming.weight)) {
        if (hasMetadataToFill(existing, incoming)) ready.push(withMissingMetadata(existing, incoming));
        else unchanged += 1;
        continue;
      }
      const oldTime = Date.parse(existing.weighedAt || existing.updatedAt || "");
      const newTime = Date.parse(incoming.weighedAt || "");
      if (Number.isFinite(oldTime) && Number.isFinite(newTime) && oldTime !== newTime) {
        if (newTime > oldTime) ready.push({ ...withPreservedMetadata(existing, incoming), source: "import" });
        else {
          if (hasMetadataToFill(existing, incoming)) ready.push(withMissingMetadata(existing, incoming));
          keptCurrent += 1;
        }
      } else unresolved.push({ ...withPreservedMetadata(existing, incoming), source: "import", weighedAt: incoming.weighedAt || new Date().toISOString() });
    }
    return { ready, unresolved, ignored, unchanged, keptCurrent };
  }

  function exportRows(session, weights, plots = session.plots, exportedAt = new Date().toISOString()) {
    const records = byUuid(weights);
    return plots.map((plot) => {
      const record = records.get(normalize(plot.uuid));
      const details = lotDetails(record || plot.lotMetadata);
      const row = {
        "Lot name": lotName(plot), "Lot site": details.lotSite, Storage: details.storage,
        "Lot location": details.lotLocation, "First order": details.firstOrder, "Second order": details.secondOrder,
        "Lot weight": record ? Number(record.weight) : "",
      };
      for (const [label, key] of sourceColumns) row[label] = plot[key] ?? "";
      row[weightLabel(session)] = record ? Number(record.weight) : "";
      row["Weighing status"] = record ? "Weighed" : "Pending";
      row["Weighed at"] = record?.weighedAt || record?.updatedAt || "";
      row["Session ID"] = session.id;
      row["Session name"] = session.name;
      row["Exported at"] = exportedAt;
      return row;
    });
  }

  function fileStamp(date = new Date()) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}_${pad(date.getHours())}-${pad(date.getMinutes())}-${pad(date.getSeconds())}`;
  }
  function slug(value) {
    return String(value || "weighing-session").normalize("NFKD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]+/g, "-").replace(/^-|-$/g, "").toLowerCase() || "weighing-session";
  }
  function downloadBlob(blob, name) {
    const href = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = href;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(href), 0);
  }

  function workbookBlob(workbook) {
    const content = global.XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
    return new Blob([content], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  function prepareSessionExport(session, weights, format, plots = session.plots, places = 0, exportedOn = new Date()) {
    if (!global.XLSX) throw new Error("The spreadsheet writer is unavailable.");
    if (!plots?.length) throw new Error("No plot records match the current filters.");
    const exportedAt = exportedOn.toISOString();
    const rows = exportRows(session, weights, plots, exportedAt);
    const seen = new Set();
    const duplicateLotNames = [...new Set(rows.filter((row) => {
      const key = normalize(row["Lot name"]);
      if (!key) return false;
      if (seen.has(key)) return true;
      seen.add(key); return false;
    }).map((row) => row["Lot name"]))];
    const weightColumnLabel = weightLabel(session);
    const digits = decimalPlaces(places);
    const stamp = fileStamp(exportedOn);
    const name = `${slug(session.name)}_${stamp}`;
    const sheet = global.XLSX.utils.json_to_sheet(rows);
    if (format === "xlsx") {
      const numberFormat = digits ? `0.${"0".repeat(digits)}` : "0";
      for (const label of ["Lot weight", weightColumnLabel]) {
        const pwColumn = Object.keys(rows[0]).indexOf(label);
        for (let rowIndex = 1; rowIndex <= rows.length; rowIndex += 1) {
        const cell = sheet[global.XLSX.utils.encode_cell({ r: rowIndex, c: pwColumn })];
        if (cell && cell.t === "n") cell.z = numberFormat;
        }
      }
      const workbook = global.XLSX.utils.book_new();
      global.XLSX.utils.book_append_sheet(workbook, sheet, "Weighing Data");
      const info = global.XLSX.utils.json_to_sheet([
        { Field: "Session ID", Value: session.id }, { Field: "Session name", Value: session.name },
        { Field: "Weight variable", Value: weightColumnLabel },
        { Field: "Source file", Value: session.sourceFileName }, { Field: "Created at", Value: session.createdAt },
        { Field: "Exported at", Value: rows[0]?.["Exported at"] || new Date().toISOString() },
      ]);
      global.XLSX.utils.book_append_sheet(workbook, info, "Session Info");
      return { blob: workbookBlob(workbook), fileName: `${name}.xlsx`, count: rows.length, duplicateLotNames };
    } else {
      const csvRows = rows.map((row) => ({ ...row,
        [weightColumnLabel]: row[weightColumnLabel] === "" ? "" : fixedDecimal(row[weightColumnLabel], digits),
        "Lot weight": row["Lot weight"] === "" ? "" : fixedDecimal(row["Lot weight"], digits),
      }));
      const csvSheet = global.XLSX.utils.json_to_sheet(csvRows);
      const csv = global.XLSX.utils.sheet_to_csv(csvSheet, { FS: ",", RS: "\r\n" });
      return { blob: new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" }), fileName: `${name}.csv`, count: rows.length, duplicateLotNames };
    }
  }

  function exportFileName(value, suggestedName) {
    const extension = suggestedName.endsWith(".xlsx") ? ".xlsx" : ".csv";
    const name = String(value ?? "").trim().replace(/\.(xlsx|csv)$/i, "");
    if (!name.trim()) throw new Error("Enter a file name.");
    if (/[<>:"/\\|?*]/.test(name) || [...name].some((character) => character.charCodeAt(0) < 32)) {
      throw new Error('The file name cannot contain < > : " / \\ | ? * or control characters.');
    }
    if (/[. ]$/.test(name) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)) {
      throw new Error("Choose a different file name, without a trailing dot or space.");
    }
    return name + extension;
  }

  function downloadPreparedExport(artifact, name = artifact.fileName) {
    const fileName = exportFileName(name, artifact.fileName);
    downloadBlob(artifact.blob, fileName);
    return { count: artifact.count, duplicateLotNames: artifact.duplicateLotNames, fileName };
  }

  function exportSession(session, weights, format, plots = session.plots, places = 0) {
    return downloadPreparedExport(prepareSessionExport(session, weights, format, plots, places));
  }

  global.GdmWeighingUtils = { normalize, decimalPlaces, fixedDecimal, byUuid, groupProgress, overallProgress, prepareMerge,
    exportRows, prepareSessionExport, exportFileName, downloadPreparedExport, exportSession, weightVariables, weightVariable, weightLabel,
    lotFields, lotCatalog, lotDetails, hasLotDetails, selectLotDetails, lotOptions, importVariable, lotName };
})(window);
