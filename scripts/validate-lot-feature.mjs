import assert from "node:assert/strict";
import fs from "node:fs/promises";
import vm from "node:vm";

globalThis.window = globalThis;
for (const path of ["public/vendor/xlsx.full.min.js", "public/weighing-utils.js", "public/plot-import.js"]) {
  vm.runInThisContext(await fs.readFile(path, "utf8"), { filename: path });
}
const utils = globalThis.GdmWeighingUtils;
const importer = globalThis.GdmPlotImport;
const XLSX = globalThis.XLSX;
const fields = { lotSite: "Legacy site", storage: "Legacy storage", lotLocation: "Dispatch", firstOrder: "001", secondOrder: "Order B" };
const plots = [
  { id: "1", feid: "10", uuid: "A", seasonYear: "2026", entityName: "Trial A", trialType: "EYT", site: "Trial site", location: "North", row: "1", column: "1", entryCode: "1", block: "1", obsName: "1001", gid: "G1", gerName: "R1", initialPlot: 1, finalPlot: 100 },
  { id: "2", feid: "20", uuid: "B", seasonYear: "", entityName: "", trialType: "", location: "South", row: "", column: "", entryCode: "", block: "", obsName: "", gid: "", gerName: "", lotMetadata: fields },
  { id: "3", feid: "30", uuid: "C", seasonYear: "2026", entityName: "Trial A", trialType: "EYT", location: "South", row: "", column: "", entryCode: "", block: "1", obsName: "1001", gid: "", gerName: "" },
];
const session = { id: "S1", name: "Test session", sourceFileName: "source.xlsx", createdAt: "2026-10-08T11:00:00.000Z", weightVariable: "plotWeight", plots };
const weights = [
  { uuid: "A", weight: 0, weightVariable: "plotWeight", ...fields, weighedAt: "2026-10-08T12:00:00.000Z" },
  { uuid: "C", weight: 3689.0007, weightVariable: "plotWeight", ...fields, weighedAt: "2026-10-08T13:00:00.000Z" },
];
const asFile = (blob, name) => ({ name, arrayBuffer: () => blob.arrayBuffer() });
const headers = ["Lot name", "Lot site", "Storage", "Lot location", "First order", "Second order", "Lot weight"];
const stamp = new Date("2026-10-24T15:32:18.000Z");
let downloads = 0;
let lastDownloadName = "";
globalThis.document = { body: { appendChild() {} }, createElement() { return { click() { downloads++; lastDownloadName = this.download; }, remove() {} }; } };
const originalCreate = URL.createObjectURL;
URL.createObjectURL = () => "blob:test";
URL.revokeObjectURL = () => {};

for (const variable of ["plotWeight", "seedWeight"]) {
  const current = { ...session, weightVariable: variable };
  const label = utils.weightLabel(variable);
  const records = weights.map((item) => ({ ...item, weightVariable: variable }));
  for (const precision of [0, 2, 4, 6]) for (const format of ["xlsx", "csv"]) {
    const artifact = utils.prepareSessionExport(current, records, format, plots, precision, stamp);
    assert.equal(artifact.count, 3);
    assert.deepEqual(artifact.duplicateLotNames, ["2026 | Trial A | 1 | 1001"]);
    assert.match(artifact.fileName, /test-session_2026-10-24_\d{2}-32-18\.(xlsx|csv)$/);
    const book = XLSX.read(await artifact.blob.arrayBuffer(), { type: "array", raw: true, cellNF: true });
    const sheet = book.Sheets[book.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    assert.deepEqual(rows[0].slice(0, 7), headers);
    assert.ok(rows[0].includes(label));
    assert.equal(rows[0].includes(variable === "plotWeight" ? "Seed weight" : "Plot weight"), false);
    for (const forbidden of ["PW", "SEED.W", "Weight", "Site"]) assert.equal(rows[0].includes(forbidden), false);
    const weightColumn = rows[0].indexOf(label);
    const lotColumn = rows[0].indexOf("Lot weight");
    assert.equal(rows[2][lotColumn], "");
    assert.equal(rows[2][weightColumn], "");
    for (const r of [1, 3]) assert.equal(rows[r][weightColumn], rows[r][lotColumn]);
    if (format === "xlsx") {
      for (const c of [lotColumn, weightColumn]) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 3, c })];
        assert.equal(cell.t, "n"); assert.equal(cell.v, 3689.0007);
        assert.equal(cell.z, precision ? "0." + "0".repeat(precision) : "0");
      }
    } else assert.ok((await artifact.blob.text()).includes(utils.fixedDecimal(3689.0007, precision)));
    const restored = await importer.parseExcelFile(asFile(artifact.blob, artifact.fileName));
    assert.equal(restored.plots.length, 3);
    assert.equal(restored.importedWeights.length, 2);
    assert.equal(restored.importedWeights[0].weight, 0);
    assert.equal(restored.importedWeights[1].weight, format === "xlsx" ? 3689.0007 : Number(utils.fixedDecimal(3689.0007, precision)));
    assert.equal(restored.importedWeights[0].weighedAt, records[0].weighedAt);
    assert.deepEqual(restored.plots[1].lotMetadata, fields);
    assert.equal(restored.plots[1].entityName, "");
    assert.equal(restored.plots[1].seasonYear, "");
    assert.equal(restored.plots[1].obsName, "");
    assert.equal(utils.importVariable(restored, "plotWeight", false), variable);
    assert.equal(utils.importVariable(restored, "plotWeight", true), "plotWeight");
    const before = downloads;
    const downloaded = utils.exportSession(current, records, format, plots, precision);
    assert.equal(downloads - before, 1);
    assert.equal(downloaded.count, 3);
  }
}
for (const format of ["xlsx", "csv"]) {
  const before = downloads;
  const artifact = utils.prepareSessionExport(session, weights, format, plots, 4, stamp);
  assert.equal(downloads, before, "Preparing the naming dialog must not start a download");
  const originalBytes = await artifact.blob.arrayBuffer();
  const defaultResult = utils.downloadPreparedExport(artifact);
  assert.equal(lastDownloadName, artifact.fileName, "Keep the suggested name unchanged by default");
  assert.equal(defaultResult.fileName, artifact.fileName);
  const customResult = utils.downloadPreparedExport(artifact, "Pesagem São Paulo - lote 01");
  assert.equal(lastDownloadName, `Pesagem São Paulo - lote 01.${format}`);
  assert.equal(customResult.count, 3);
  assert.deepEqual(await artifact.blob.arrayBuffer(), originalBytes, "Renaming must not regenerate or modify the data/timestamps");
  assert.equal(downloads - before, 2);
  assert.equal(utils.exportFileName(`custom.${format.toUpperCase()}`, artifact.fileName), `custom.${format}`);
  assert.equal(utils.exportFileName(`custom.${format === "csv" ? "xlsx" : "csv"}`, artifact.fileName), `custom.${format}`);
  const countBeforeErrors = downloads;
  for (const invalid of ["", "  ", `.${format}`, "../file", "folder\\file", "bad:name", "bad*name", "bad?name", "bad|name", "bad<name", 'bad"name', "bad\u0000name", "CON", "nul.txt", "LPT1", "name."]) {
    assert.throws(() => utils.downloadPreparedExport(artifact, invalid));
  }
  assert.equal(downloads, countBeforeErrors, "Invalid names must not download a file");
}
URL.createObjectURL = originalCreate;
assert.equal(weights[1].weight, 3689.0007, "Formatting must not mutate stored precision");
assert.equal(utils.lotName(plots[1]), "");
assert.equal(utils.lotName({ ...plots[0], seasonYear: "", block: "" }), " | Trial A |  | 1001");
assert.deepEqual(utils.overallProgress(plots, weights), { total: 3, completed: 2, remaining: 1, percent: 67 });
assert.deepEqual(utils.groupProgress(plots, weights, "location").map(({ label, total, completed }) => ({ label, total, completed })), [
  { label: "North", total: 1, completed: 1 }, { label: "South", total: 2, completed: 1 },
]);
for (const count of [99, 100, 101, 205]) {
  const many = Array.from({ length: count }, (_, i) => ({ ...plots[0], feid: String(i), uuid: "M" + i }));
  assert.equal(utils.prepareSessionExport(session, [], "csv", many).count, count);
}
assert.deepEqual(utils.lotCatalog["Chacabuco"], ["Cámara fría_1", "Cámara fría_2", "Galpón exterior"]);
assert.deepEqual(utils.lotCatalog["Passo Fundo"], ["Câmara fria_1", "Câmara fria_2", "Galpão externo"]);
assert.equal(utils.lotOptions("Chacabuco", "Old storage", true)[0].label, "Old storage (existing)");
assert.equal(utils.lotOptions("Chacabuco", "", true).length, 3);
assert.equal(utils.lotOptions("", "", true).length, 0);
const context = { ...fields, keepForNext: true, hasValue: true };
assert.deepEqual(utils.selectLotDetails(plots[1], undefined, { ...context, lotSite: "Passo Fundo" }), fields);
assert.deepEqual(utils.selectLotDetails(plots[0], undefined, context), fields);
assert.deepEqual(utils.selectLotDetails(plots[0], { ...fields, lotSite: "" }, context), { ...fields, lotSite: "" });
assert.equal(utils.selectLotDetails(plots[0], undefined, { ...context, keepForNext: false }).lotSite, "");
const base = { uuid: "A", weight: 1, weightVariable: "plotWeight", weighedAt: "2026-10-08T12:00:00.000Z", ...fields };
assert.equal(utils.prepareMerge(plots, [base], [{ ...base, weight: 2, weighedAt: "2026-10-08T13:00:00.000Z" }]).ready[0].weight, 2);
assert.equal(utils.prepareMerge(plots, [base], [{ ...base, weight: 2, weighedAt: "2026-10-08T11:00:00.000Z" }]).keptCurrent, 1);
assert.equal(utils.prepareMerge(plots, [base], [{ ...base, weight: 2, weighedAt: "" }]).unresolved.length, 1);
assert.equal(utils.prepareMerge(plots, [{ ...base, firstOrder: "" }], [{ ...base, firstOrder: "fill" }]).ready[0].firstOrder, "fill");
assert.equal(utils.prepareMerge(plots, [], [{ ...base, uuid: "unknown" }]).ignored, 1);

const sourceRow = utils.exportRows(session, weights)[0];
async function parseRows(rows, extra = []) {
  const sheet = XLSX.utils.json_to_sheet(rows, { header: [...importer.requiredHeaders, ...extra] });
  const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, "Data");
  return importer.parseExcelFile({ name: "input.xlsx", arrayBuffer: async () => XLSX.write(book, { type: "array", bookType: "xlsx" }) });
}
for (const [header, variable] of [["Plot weight", "plotWeight"], ["PW", "plotWeight"], ["Seed weight", "seedWeight"], ["SEED.W", "seedWeight"]]) {
  const row = { ...sourceRow }; delete row["Plot weight"]; row[header] = 0;
  const result = await parseRows([row]);
  assert.equal(result.importedWeights[0].weight, 0);
  assert.equal(result.importedWeights[0].weightVariable, variable);
}
let result = await parseRows([{ ...sourceRow, PW: 9 }]);
assert.equal(result.weightConflicts.length, 1); assert.equal(result.importedWeights.length, 0);
result = await parseRows([{ ...sourceRow, PW: 0 }]); assert.equal(result.importedWeights.length, 1);
result = await parseRows([{ ...sourceRow, "Plot weight": "", PW: 10 }]); assert.equal(result.importedWeights[0].weight, 10);
result = await parseRows([{ ...sourceRow, "Plot weight": "", "Lot weight": 999 }]); assert.equal(result.importedWeights.length, 0);
result = await parseRows([{ ...sourceRow, "Plot weight": "", "Seed weight": 20 }]);
assert.equal(utils.importVariable(result, "plotWeight", false), "plotWeight", "Both columns retain operator selection");
result = await parseRows([{ ...sourceRow, "Plot weight": "invalid" }]); assert.equal(result.invalidWeightRows.length, 1);
result = await parseRows([sourceRow, { ...sourceRow, UUID: "new" }, { ...sourceRow, UUID: "new2", FEID: "" }]);
assert.equal(result.plots.length, 1); assert.equal(result.invalidRows.length, 2);
result = await parseRows([{ ...sourceRow, "Initial plot": 100, "Final plot": 1 }]);
assert.equal(result.plots[0].finalPlot, 1);
for (const key of ["PW", "Plot weight", "Initial plot", "Final plot"]) assert.equal(importer.requiredHeaders.includes(key), false);
const html = await fs.readFile("index.html", "utf8");
assert.equal(html.includes('id="export-prompt"'), false);
assert.equal(html.includes("Export lots"), false);
assert.ok(html.includes('id="new-session-prompt"'));
assert.ok(html.includes('id="export-name-prompt"'));
assert.ok(html.includes('id="export-file-name"'));
assert.ok(html.includes('<select id="lot-site">'));
assert.ok(html.includes('<select id="lot-storage" disabled>'));
assert.ok(html.indexOf('class="hero-details"') < html.indexOf('id="weight-form"'));
const utf8 = await importer.parseExcelFile(asFile(new Blob([await fs.readFile("test-fixtures/unified-seed-results.csv")]), "utf8.csv"));
assert.equal(utf8.importedWeights[0].storage, "Câmara fria_1", "UTF-8 CSV without BOM preserves accents");
console.log("PASS: real XLSX/CSV round trips, precision, metadata, variables, filtering, catalogue and merge rules.");
