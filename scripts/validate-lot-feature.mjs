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
const weights = [{ uuid: "A", weight: 0, weightVariable: "plotWeight", lotSite: "Operational site", lotLocation: "Dispatch area", storage: "Cold room", weighedAt: "2026-10-08T12:00:00.000Z" }];

assert.deepEqual(globalThis.GdmWeighingUtils.overallProgress(plots, weights), { total: 3, completed: 1, remaining: 2, percent: 33 });
assert.deepEqual(globalThis.GdmWeighingUtils.groupProgress(plots, weights, "location").map(({ label, total, completed }) => ({ label, total, completed })), [
  { label: "North", total: 2, completed: 1 },
  { label: "South", total: 1, completed: 0 },
]);
const metadataMerge = globalThis.GdmWeighingUtils.prepareMerge(
  plots,
  [{ uuid: "A", weight: 0, weightVariable: "plotWeight", weighedAt: "2026-10-08T12:00:00.000Z", lotSite: "", lotLocation: "", storage: "" }],
  [{ uuid: "A", weight: 0, weightVariable: "plotWeight", weighedAt: "2026-10-08T12:00:00.000Z", lotSite: "Imported site", lotLocation: "Imported destination", storage: "Imported storage" }],
);
assert.equal(metadataMerge.ready.length, 1);
assert.equal(metadataMerge.ready[0].lotSite, "Imported site");
assert.equal(metadataMerge.ready[0].lotLocation, "Imported destination");
assert.equal(metadataMerge.ready[0].storage, "Imported storage");

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
const session = { id: "S1", name: "Session", sourceFileName: "source.xlsx", createdAt: "2026-10-08T11:00:00.000Z", weightVariable: "plotWeight", plots };
assert.equal(globalThis.GdmWeighingUtils.exportSession(session, weights, "xlsx", plots, 2), 3);
const sessionRows = capturedSheets[0];
assert.equal(sessionRows[0]["Plot weight"], 0);
assert.equal("PW" in sessionRows[0], false);
assert.equal("Seed weight" in sessionRows[0], false);
assert.equal(sessionRows[0]["Lot location"], "Dispatch area");
capturedSheets.length = 0;
assert.equal(globalThis.GdmWeighingUtils.exportLots(session, weights, "xlsx", plots, 2), 1);
const lotRows = capturedSheets[0];
assert.equal(lotRows[0]["Lot name"], "2026 | Trial A | 1 | 1001");
assert.equal(lotRows[0]["Plot weight"], 0);
assert.equal("Weight" in lotRows[0], false);
assert.equal("PW" in lotRows[0], false);

capturedSheets.length = 0;
const seedSession = { ...session, id: "S2", weightVariable: "seedWeight" };
const seedWeights = [{ ...weights[0], weight: 12.5, weightVariable: "seedWeight" }];
globalThis.GdmWeighingUtils.exportSession(seedSession, seedWeights, "xlsx", plots, 2);
assert.equal(capturedSheets[0][0]["Seed weight"], 12.5);
assert.equal("Plot weight" in capturedSheets[0][0], false);
capturedSheets.length = 0;
globalThis.GdmWeighingUtils.exportLots(seedSession, seedWeights, "xlsx", plots, 2);
assert.equal(capturedSheets[0][0]["Seed weight"], 12.5);
assert.equal("Weight" in capturedSheets[0][0], false);

const duplicatePlots = [plots[0], { ...plots[0], uuid: "D", feid: "40" }];
const duplicateWeights = [{ ...weights[0] }, { ...weights[0], uuid: "D" }];
assert.throws(() => globalThis.GdmWeighingUtils.exportLots(session, duplicateWeights, "xlsx", duplicatePlots, 0), /Duplicate Lot name/);
assert.throws(() => globalThis.GdmWeighingUtils.exportLots(session, weights, "xlsx", [{ ...plots[0], seasonYear: "" }], 0), /requires Season year/);

await loadBrowserScript("public/plot-import.js");
const requiredHeaders = globalThis.GdmPlotImport.requiredHeaders;
assert.equal(requiredHeaders.includes("PW"), false);
assert.equal(requiredHeaders.includes("Plot weight"), false);

const baseValues = {
  ID: "1", FEID: "10", UUID: "A", "Season year": "2026", "Entity name": "Trial A", "Trial type": "EYT",
  Site: "Site 1", Location: "North", Row: "1", Column: "1", "Entry code": "1", Block: "1", "(OBS) Name": "1001",
  GID: "G1", "(GER) Name": "R1", "Initial plot": 1, "Final plot": 100,
  "Lot site": "Lot Site 1", "Lot location": "Dispatch area", Storage: "Cold room", "Weighed at": "2026-10-08T12:00:00.000Z",
};
let importRows = [];
globalThis.XLSX.read = () => ({ SheetNames: ["Data"], Sheets: { Data: {} } });
globalThis.XLSX.utils.sheet_to_json = () => importRows;
const parse = async (extraHeaders = [], overrides = {}) => {
  const headers = [...requiredHeaders, "Site", "Lot site", "Lot location", "Storage", "Weighed at", ...extraHeaders];
  const values = { ...baseValues, ...overrides };
  importRows = [headers, headers.map((header) => values[header] ?? "")];
  return globalThis.GdmPlotImport.parseExcelFile({ name: "source.xlsx", async arrayBuffer() { return new ArrayBuffer(0); } });
};

const noWeights = await parse();
assert.equal(noWeights.importedWeights.length, 0);
assert.equal(noWeights.plots.length, 1);

for (const [header, variable] of [["Plot weight", "plotWeight"], ["PW", "plotWeight"], ["Seed weight", "seedWeight"], ["SEED.W", "seedWeight"]]) {
  const imported = await parse([header], { [header]: 0 });
  assert.equal(imported.importedWeights.length, 1, `${header} should import zero`);
  assert.equal(imported.importedWeights[0].weight, 0);
  assert.equal(imported.importedWeights[0].weightVariable, variable);
  assert.equal(imported.importedWeights[0].lotSite, "Lot Site 1");
}

const bothVariables = await parse(["Plot weight", "Seed weight"], { "Plot weight": 10, "Seed weight": 20 });
assert.deepEqual(bothVariables.weightValueCounts, { plotWeight: 1, seedWeight: 1 });
assert.equal(bothVariables.importedWeights.length, 2);
const equalAliases = await parse(["Plot weight", "PW"], { "Plot weight": 10, PW: 10 });
assert.equal(equalAliases.importedWeights.length, 1);
assert.equal(equalAliases.weightConflicts.length, 0);
const oneAliasFilled = await parse(["Plot weight", "PW"], { "Plot weight": "", PW: 11 });
assert.equal(oneAliasFilled.importedWeights[0].weight, 11);
const conflictingAliases = await parse(["Plot weight", "PW"], { "Plot weight": 10, PW: 11 });
assert.equal(conflictingAliases.importedWeights.length, 0);
assert.equal(conflictingAliases.weightConflicts[0].row, 2);
assert.equal(conflictingAliases.weightConflicts[0].weightVariable, "plotWeight");
const invalidWeight = await parse(["Seed weight"], { "Seed weight": "not-a-number" });
assert.equal(invalidWeight.importedWeights.length, 0);
assert.equal(invalidWeight.invalidWeightRows[0].weightVariable, "seedWeight");

const missingHeaders = requiredHeaders.filter((header) => header !== "Season year");
importRows = [missingHeaders, missingHeaders.map((header) => baseValues[header] ?? "")];
await assert.rejects(() => globalThis.GdmPlotImport.parseExcelFile({ name: "missing.xlsx", async arrayBuffer() { return new ArrayBuffer(0); } }), /Missing headers: Season year/);

console.log("Weight variable and lot feature validation passed.");
