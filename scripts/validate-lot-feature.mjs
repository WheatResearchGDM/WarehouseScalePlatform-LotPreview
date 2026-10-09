import assert from "node:assert/strict";
import fs from "node:fs/promises";
import vm from "node:vm";

globalThis.window = globalThis;

const loadBrowserScript = async (path) => vm.runInThisContext(await fs.readFile(path, "utf8"), { filename: path });

await loadBrowserScript("public/weighing-utils.js");

const plots = [
  { id: "1", feid: "10", uuid: "A", seasonYear: "2026", entityName: "Trial A", trialType: "EYT", site: "Site 1", location: "North", row: "1", column: "1", entryCode: "1", block: "1", obsName: "1001", gid: "G1", gerName: "R1", initialPlot: 1, finalPlot: 100 },
  { id: "2", feid: "20", uuid: "B", seasonYear: "2026", entityName: "Trial A", trialType: "EYT", site: "Site 1", location: "North", row: "1", column: "2", entryCode: "2", block: "1", obsName: "1002", gid: "G2", gerName: "R2", initialPlot: 1, finalPlot: 100 },
  { id: "3", feid: "30", uuid: "C", seasonYear: "2026", entityName: "Trial A", trialType: "EYT", site: "Site 2", location: "South", row: "1", column: "3", entryCode: "3", block: "1", obsName: "1003", gid: "G3", gerName: "R3", initialPlot: 1, finalPlot: 100 },
];
const weights = [{ uuid: "A", weight: 0, lotSite: "Operational site", storage: "Cold room", weighedAt: "2026-10-08T12:00:00.000Z" }];

assert.deepEqual(globalThis.GdmWeighingUtils.overallProgress(plots, weights), { total: 3, completed: 1, remaining: 2, percent: 33 });
assert.deepEqual(globalThis.GdmWeighingUtils.groupProgress(plots, weights, "location").map(({ label, total, completed }) => ({ label, total, completed })), [
  { label: "North", total: 2, completed: 1 },
  { label: "South", total: 1, completed: 0 },
]);
const metadataMerge = globalThis.GdmWeighingUtils.prepareMerge(
  plots,
  [{ uuid: "A", weight: 0, weighedAt: "2026-10-08T12:00:00.000Z", lotSite: "", storage: "" }],
  [{ uuid: "A", weight: 0, weighedAt: "2026-10-08T12:00:00.000Z", lotSite: "Imported site", storage: "Imported storage" }],
);
assert.equal(metadataMerge.ready.length, 1);
assert.equal(metadataMerge.ready[0].lotSite, "Imported site");
assert.equal(metadataMerge.ready[0].storage, "Imported storage");
const newerWeightMerge = globalThis.GdmWeighingUtils.prepareMerge(
  plots,
  [{ uuid: "A", weight: 1, weighedAt: "2026-10-08T12:00:00.000Z", lotSite: "Current lot site", storage: "Current storage" }],
  [{ uuid: "A", weight: 2, weighedAt: "2026-10-08T13:00:00.000Z", lotSite: "", storage: "" }],
);
assert.equal(newerWeightMerge.ready[0].weight, 2);
assert.equal(newerWeightMerge.ready[0].lotSite, "Current lot site");
assert.equal(newerWeightMerge.ready[0].storage, "Current storage");

const capturedSheets = [];
globalThis.XLSX = {
  utils: {
    json_to_sheet(rows) { capturedSheets.push(rows); return {}; },
    encode_cell({ r, c }) { return `${r}:${c}`; },
    book_new() { return {}; },
    book_append_sheet() {},
  },
  writeFile() {},
};
const session = { id: "S1", name: "Session", sourceFileName: "source.xlsx", createdAt: "2026-10-08T11:00:00.000Z", plots };
assert.equal(globalThis.GdmWeighingUtils.exportSession(session, weights, "xlsx", plots, 2), 3);
const sessionRows = capturedSheets[0];
assert.equal(sessionRows[0]["Lot site"], "Operational site");
assert.equal("Site" in sessionRows[0], false);
assert.equal(sessionRows[1]["Lot site"], "");
capturedSheets.length = 0;
assert.equal(globalThis.GdmWeighingUtils.exportLots(session, weights, "xlsx", plots, 2), 1);
const capturedRows = capturedSheets[0];
assert.equal(capturedRows[0]["Lot name"], "2026 | Trial A | 1 | 1001");
assert.equal(capturedRows[0].Weight, 0);
assert.equal(capturedRows[0]["Lot site"], "Operational site");
assert.equal("Site" in capturedRows[0], false);
assert.equal(capturedRows[0].Storage, "Cold room");
assert.equal("PW" in capturedRows[0], false);

const duplicatePlots = [plots[0], { ...plots[0], uuid: "D", feid: "40" }];
const duplicateWeights = [{ ...weights[0] }, { ...weights[0], uuid: "D" }];
assert.throws(() => globalThis.GdmWeighingUtils.exportLots(session, duplicateWeights, "xlsx", duplicatePlots, 0), /Duplicate Lot name/);
assert.throws(() => globalThis.GdmWeighingUtils.exportLots(session, weights, "xlsx", [{ ...plots[0], seasonYear: "" }], 0), /requires Season year/);

await loadBrowserScript("public/plot-import.js");
const requiredHeaders = globalThis.GdmPlotImport.requiredHeaders;
const headers = [...requiredHeaders, "Site", "Lot site", "Storage"];
const row = headers.map((header) => ({
  ID: "1", FEID: "10", UUID: "A", "Season year": "2026", "Entity name": "Trial A", "Trial type": "EYT",
  Site: "Site 1", Location: "North", Row: "1", Column: "1", "Entry code": "1", Block: "1", "(OBS) Name": "1001",
  GID: "G1", "(GER) Name": "R1", "Initial plot": 1, "Final plot": 100, PW: 0,
  "Lot site": "Lot Site 1", Storage: "Cold room",
}[header]));
let importRows = [headers, row];
globalThis.XLSX.read = () => ({ SheetNames: ["Data"], Sheets: { Data: {} } });
globalThis.XLSX.utils.sheet_to_json = () => importRows;
const imported = await globalThis.GdmPlotImport.parseExcelFile({ name: "source.xlsx", async arrayBuffer() { return new ArrayBuffer(0); } });
assert.equal(imported.plots[0].seasonYear, "2026");
assert.equal(imported.plots[0].site, "Site 1");
assert.equal(imported.importedWeights[0].weight, 0);
assert.equal(imported.importedWeights[0].lotSite, "Lot Site 1");

const headersWithoutSite = [...requiredHeaders, "Lot site", "Storage"];
const rowWithoutSite = headersWithoutSite.map((header) => ({
  ID: "1", FEID: "10", UUID: "A", "Season year": "2026", "Entity name": "Trial A", "Trial type": "EYT",
  Location: "North", Row: "1", Column: "1", "Entry code": "1", Block: "1", "(OBS) Name": "1001",
  GID: "G1", "(GER) Name": "R1", "Initial plot": 1, "Final plot": 100, PW: 0,
  "Lot site": "Lot Site 2", Storage: "Dry room",
}[header]));
importRows = [headersWithoutSite, rowWithoutSite];
const reimported = await globalThis.GdmPlotImport.parseExcelFile({ name: "export.xlsx", async arrayBuffer() { return new ArrayBuffer(0); } });
assert.equal(reimported.plots[0].site, "");
assert.equal(reimported.importedWeights[0].lotSite, "Lot Site 2");

importRows = [headers.filter((header) => header !== "Season year"), row.filter((_, index) => headers[index] !== "Season year")];
await assert.rejects(() => globalThis.GdmPlotImport.parseExcelFile({ name: "missing.xlsx", async arrayBuffer() { return new ArrayBuffer(0); } }), /Missing headers: Season year/);

console.log("Lot feature validation passed.");
