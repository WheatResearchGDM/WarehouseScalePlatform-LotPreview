(function (global) {
  "use strict";

  const sourceColumns = [
    ["ID", "id"], ["FEID", "feid"], ["UUID", "uuid"], ["Season year", "seasonYear"], ["Entity name", "entityName"],
    ["Trial type", "trialType"], ["Location", "location"], ["Row", "row"],
    ["Column", "column"], ["Entry code", "entryCode"], ["Block", "block"], ["(OBS) Name", "obsName"],
    ["GID", "gid"], ["(GER) Name", "gerName"], ["Initial plot", "initialPlot"], ["Final plot", "finalPlot"],
  ];

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
      lotSite: existing.lotSite || incoming.lotSite || "", storage: existing.storage || incoming.storage || "", source: existing.source || "import",
    });
    const withPreservedMetadata = (existing, incoming) => ({
      ...incoming,
      lotSite: incoming.lotSite || existing.lotSite || "",
      storage: incoming.storage || existing.storage || "",
    });
    const hasMetadataToFill = (existing, incoming) => (!existing.lotSite && incoming.lotSite) || (!existing.storage && incoming.storage);
    for (const incoming of importedWeights || []) {
      const key = normalize(incoming.uuid);
      if (!plots.has(key)) { ignored += 1; continue; }
      const existing = current.get(key);
      if (!existing) { ready.push({ ...incoming, source: "import" }); continue; }
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

  function exportRows(session, weights, plots = session.plots) {
    const records = byUuid(weights);
    const exportedAt = new Date().toISOString();
    return plots.map((plot) => {
      const row = {};
      for (const [label, key] of sourceColumns) row[label] = plot[key] ?? "";
      const record = records.get(normalize(plot.uuid));
      row["Lot site"] = record?.lotSite ?? "";
      row.PW = record ? Number(record.weight) : "";
      row.Storage = record?.storage ?? "";
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
  function exportSession(session, weights, format, plots = session.plots, places = 0) {
    if (!global.XLSX) throw new Error("The spreadsheet writer is unavailable.");
    const rows = exportRows(session, weights, plots);
    const digits = decimalPlaces(places);
    const stamp = fileStamp();
    const name = `${slug(session.name)}_${stamp}`;
    const sheet = global.XLSX.utils.json_to_sheet(rows);
    if (format === "xlsx") {
      const pwColumn = Object.keys(rows[0]).indexOf("PW");
      const numberFormat = digits ? `0.${"0".repeat(digits)}` : "0";
      for (let rowIndex = 1; rowIndex <= rows.length; rowIndex += 1) {
        const cell = sheet[global.XLSX.utils.encode_cell({ r: rowIndex, c: pwColumn })];
        if (cell && cell.t === "n") cell.z = numberFormat;
      }
      const workbook = global.XLSX.utils.book_new();
      global.XLSX.utils.book_append_sheet(workbook, sheet, "Weighing Data");
      const info = global.XLSX.utils.json_to_sheet([
        { Field: "Session ID", Value: session.id }, { Field: "Session name", Value: session.name },
        { Field: "Source file", Value: session.sourceFileName }, { Field: "Created at", Value: session.createdAt },
        { Field: "Exported at", Value: rows[0]?.["Exported at"] || new Date().toISOString() },
      ]);
      global.XLSX.utils.book_append_sheet(workbook, info, "Session Info");
      global.XLSX.writeFile(workbook, `${name}.xlsx`, { compression: true });
    } else {
      const csvRows = rows.map((row) => ({ ...row, PW: row.PW === "" ? "" : fixedDecimal(row.PW, digits) }));
      const csvSheet = global.XLSX.utils.json_to_sheet(csvRows);
      const csv = global.XLSX.utils.sheet_to_csv(csvSheet, { FS: ",", RS: "\r\n" });
      downloadBlob(new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" }), `${name}.csv`);
    }
    return rows.length;
  }

  function exportLots(session, weights, format, plots = session.plots, places = 0) {
    if (!global.XLSX) throw new Error("The spreadsheet writer is unavailable.");
    const records = byUuid(weights);
    const candidates = (plots || []).filter((plot) => records.has(normalize(plot.uuid)));
    if (!candidates.length) throw new Error("No weighed plots match the current filters.");
    const invalid = [];
    const seen = new Map();
    const duplicates = new Set();
    const exportedAt = new Date().toISOString();
    const rows = candidates.map((plot) => {
      const parts = [plot.seasonYear, plot.entityName, plot.block, plot.obsName].map((value) => String(value ?? "").trim());
      if (parts.some((value) => !value)) invalid.push(plot.feid || plot.uuid || "Unknown plot");
      const lotName = parts.join(" | ");
      const duplicateKey = normalize(lotName);
      if (seen.has(duplicateKey)) duplicates.add(lotName);
      else seen.set(duplicateKey, plot.uuid);
      const record = records.get(normalize(plot.uuid));
      const row = { "Lot name": lotName };
      for (const [label, field] of sourceColumns) row[label] = plot[field] ?? "";
      row["Lot site"] = record.lotSite ?? "";
      row.Weight = Number(record.weight);
      row.Storage = record.storage ?? "";
      row["Weighing status"] = "Weighed";
      row["Weighed at"] = record.weighedAt || record.updatedAt || "";
      row["Session ID"] = session.id;
      row["Session name"] = session.name;
      row["Exported at"] = exportedAt;
      return row;
    });
    if (invalid.length) throw new Error(`Lot export requires Season year, Entity name, Block, and (OBS) Name. Check: ${invalid.slice(0, 5).join(", ")}${invalid.length > 5 ? "…" : ""}.`);
    if (duplicates.size) {
      const names = Array.from(duplicates);
      throw new Error(`Duplicate Lot name values found: ${names.slice(0, 5).join(", ")}${names.length > 5 ? "…" : ""}.`);
    }
    const digits = decimalPlaces(places);
    const name = `${slug(session.name)}_lots_${fileStamp()}`;
    const sheet = global.XLSX.utils.json_to_sheet(rows);
    if (format === "xlsx") {
      const weightColumn = Object.keys(rows[0]).indexOf("Weight");
      const numberFormat = digits ? `0.${"0".repeat(digits)}` : "0";
      for (let rowIndex = 1; rowIndex <= rows.length; rowIndex += 1) {
        const cell = sheet[global.XLSX.utils.encode_cell({ r: rowIndex, c: weightColumn })];
        if (cell && cell.t === "n") cell.z = numberFormat;
      }
      const workbook = global.XLSX.utils.book_new();
      global.XLSX.utils.book_append_sheet(workbook, sheet, "Lots");
      const info = global.XLSX.utils.json_to_sheet([
        { Field: "Session ID", Value: session.id }, { Field: "Session name", Value: session.name },
        { Field: "Source file", Value: session.sourceFileName }, { Field: "Created at", Value: session.createdAt },
        { Field: "Exported at", Value: exportedAt },
      ]);
      global.XLSX.utils.book_append_sheet(workbook, info, "Session Info");
      global.XLSX.writeFile(workbook, `${name}.xlsx`, { compression: true });
    } else {
      const csvRows = rows.map((row) => ({ ...row, Weight: fixedDecimal(row.Weight, digits) }));
      const csvSheet = global.XLSX.utils.json_to_sheet(csvRows);
      const csv = global.XLSX.utils.sheet_to_csv(csvSheet, { FS: ",", RS: "\r\n" });
      downloadBlob(new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" }), `${name}.csv`);
    }
    return rows.length;
  }

  global.GdmWeighingUtils = { normalize, decimalPlaces, fixedDecimal, byUuid, groupProgress, overallProgress, prepareMerge, exportRows, exportSession, exportLots };
})(window);
