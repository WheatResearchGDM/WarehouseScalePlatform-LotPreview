(function (global) {
  "use strict";

  const fields = [
    ["id", "ID", ["ID."]],
    ["feid", "FEID", ["ID.FE.", "ID FE"]],
    ["uuid", "UUID", []],
    ["seasonYear", "Season year", []],
    ["entityName", "Entity name", ["Nombre de la entidad", "Nome da entidade"]],
    ["trialType", "Trial type", ["Tipo de ensayo", "Tipo de ensaio"]],
    ["location", "Location", ["Ubicación", "Ubicaci�n", "Localização", "Localiza��o"]],
    ["row", "Row", ["Fila", "Linha"]],
    ["column", "Column", ["Columna", "Coluna"]],
    ["entryCode", "Entry code", ["Código de entrada", "C�digo de entrada"]],
    ["block", "Block", ["Bloque", "Bloco"]],
    ["obsName", "(OBS) Name", ["(OBS) Nombre", "(OBS) Nome"]],
    ["gid", "GID", []],
    ["gerName", "(GER) Name", ["(GER) Nombre", "(GER) Nome"]],
    ["initialPlot", "Initial plot", ["Parcela inicial"]],
    ["finalPlot", "Final plot", ["Parcela final"]],
    ["pw", "PW", []],
  ];
  const optionalFields = [
    ["site", "Site", ["Sitio", "Unidade"]],
    ["lotSite", "Lot site", []],
    ["storage", "Storage", []],
    ["weighingStatus", "Weighing status", []], ["weighedAt", "Weighed at", []],
    ["sessionId", "Session ID", []], ["sessionName", "Session name", []], ["exportedAt", "Exported at", []],
  ];

  function normalizeHeader(value) {
    return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .trim().replace(/[._-]+/g, " ").replace(/\s+/g, " ").toLowerCase();
  }
  function headerAliases(field) { return [field[1], ...(field[2] || [])].map(normalizeHeader); }
  function cellText(value) { return value === null || value === undefined ? "" : String(value).trim(); }
  function numericCell(value) {
    const text = cellText(value);
    if (!text) return NaN;
    const parsed = typeof value === "number" ? value : Number(text.replace(/,/g, "."));
    return Number.isFinite(parsed) ? parsed : NaN;
  }
  function validDateText(value) {
    const text = cellText(value);
    if (!text) return "";
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString();
  }

  async function parseExcelFile(file) {
    if (!global.XLSX) throw new Error("The spreadsheet reader is unavailable. Refresh the page and try again.");
    if (!file) throw new Error("Select an Excel or CSV file.");
    const workbook = global.XLSX.read(await file.arrayBuffer(), { type: "array", raw: true, cellDates: true });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error("The workbook does not contain any sheets.");
    const rows = global.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    if (!rows.length) throw new Error("The first sheet is empty.");

    const requiredNames = fields.map(headerAliases);
    const headerRowIndex = rows.findIndex((row) => {
      const values = new Set(row.map(normalizeHeader));
      return requiredNames.every((aliases) => aliases.some((name) => values.has(name)));
    });
    if (headerRowIndex < 0) {
      const available = new Set(rows.slice(0, 20).flat().map(normalizeHeader));
      const missing = fields.filter((field) => !headerAliases(field).some((alias) => available.has(alias))).map(([, label]) => label);
      throw new Error(`Missing headers: ${missing.join(", ")}.`);
    }

    const headers = rows[headerRowIndex].map(normalizeHeader);
    const indexes = Object.fromEntries([
      ...fields.map((field) => [field[0], headerAliases(field).map((alias) => headers.indexOf(alias)).find((index) => index >= 0) ?? -1]),
      ...optionalFields.map((field) => [field[0], headerAliases(field).map((alias) => headers.indexOf(alias)).find((index) => index >= 0) ?? -1]),
    ]);
    const plots = [];
    const importedWeights = [];
    const invalidRows = [];
    const invalidWeightRows = [];
    const feids = new Set();
    const uuids = new Set();
    for (let index = headerRowIndex + 1; index < rows.length; index += 1) {
      const row = rows[index];
      if (!row.some((value) => cellText(value))) continue;
      const plot = {
        id: cellText(row[indexes.id]), feid: cellText(row[indexes.feid]), uuid: cellText(row[indexes.uuid]).toUpperCase(),
        seasonYear: cellText(row[indexes.seasonYear]),
        entityName: cellText(row[indexes.entityName]), trialType: cellText(row[indexes.trialType]), site: indexes.site >= 0 ? cellText(row[indexes.site]) : "",
        location: cellText(row[indexes.location]), row: cellText(row[indexes.row]), column: cellText(row[indexes.column]),
        entryCode: cellText(row[indexes.entryCode]), block: cellText(row[indexes.block]), obsName: cellText(row[indexes.obsName]),
        gid: cellText(row[indexes.gid]), gerName: cellText(row[indexes.gerName]), initialPlot: numericCell(row[indexes.initialPlot]),
        finalPlot: numericCell(row[indexes.finalPlot]),
      };
      const invalid = !plot.feid || !plot.uuid || !plot.entityName || !plot.obsName
        || !plot.seasonYear
        || !Number.isFinite(plot.initialPlot) || !Number.isFinite(plot.finalPlot)
        || plot.finalPlot < plot.initialPlot || feids.has(plot.feid.toUpperCase()) || uuids.has(plot.uuid);
      if (invalid) { invalidRows.push(index + 1); continue; }
      feids.add(plot.feid.toUpperCase());
      uuids.add(plot.uuid);
      plots.push(plot);
      const pwText = cellText(row[indexes.pw]);
      if (pwText) {
        const weight = numericCell(row[indexes.pw]);
        if (Number.isFinite(weight) && weight >= 0) {
          importedWeights.push({
            uuid: plot.uuid, weight, lotSite: indexes.lotSite >= 0 ? cellText(row[indexes.lotSite]) : "",
            storage: indexes.storage >= 0 ? cellText(row[indexes.storage]) : "",
            weighedAt: indexes.weighedAt >= 0 ? validDateText(row[indexes.weighedAt]) : "",
          });
        } else invalidWeightRows.push(index + 1);
      }
    }
    if (!plots.length) {
      const detail = invalidRows.length ? ` Invalid rows: ${invalidRows.slice(0, 8).join(", ")}.` : "";
      throw new Error(`No valid plots were found in the first sheet.${detail}`);
    }
    const firstDataRow = rows[headerRowIndex + 1] || [];
    return {
      plots, importedWeights, invalidRows, invalidWeightRows, fileName: file.name, sheetName,
      metadata: {
        sessionId: indexes.sessionId >= 0 ? cellText(firstDataRow[indexes.sessionId]) : "",
        sessionName: indexes.sessionName >= 0 ? cellText(firstDataRow[indexes.sessionName]) : "",
        exportedAt: indexes.exportedAt >= 0 ? validDateText(firstDataRow[indexes.exportedAt]) : "",
      },
    };
  }

  global.GdmPlotImport = {
    parseExcelFile,
    requiredHeaders: fields.map(([, label]) => label),
    exportHeaders: [...fields.map(([, label]) => label), ...optionalFields.map(([, label]) => label)],
  };
})(window);
