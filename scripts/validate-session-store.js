/* Browser integration tests: real IndexedDB, isolated from operational data. */
document.getElementById("run-tests").addEventListener("click", async () => {
  const button = document.getElementById("run-tests");
  const output = document.getElementById("test-results");
  button.disabled = true;
  const databaseName = "gdm-regression-" + crypto.randomUUID();
  const store = window.GdmWeighingStore;
  const opened = [];
  let checks = 0;
  const check = (condition, message) => { if (!condition) throw new Error(message); checks++; output.textContent += "\nPASS " + message; };
  const init = async () => { const result = await store.init({ databaseName }); opened.push(result.database); return result; };
  output.textContent = "Running…";
  try {
    const { database } = await init();
    const metadata = { lotSite: "Old site", storage: "Old storage", lotLocation: "Destination", firstOrder: "001", secondOrder: "B" };
    const plots = [
      { uuid: "A", feid: "1", entityName: "", obsName: "", seasonYear: "", lotMetadata: metadata },
      { uuid: "B", feid: "2", entityName: "Trial", obsName: "2", seasonYear: "2026" },
    ];
    const first = await store.createSession(database, plots, "first.csv");
    check((await store.getWeights(database, first.id)).length === 0, "Pending metadata does not create a weight");
    check((await store.getSession(database, first.id)).plots[0].lotMetadata.firstOrder === "001", "Pending metadata persisted");
    await store.saveWeight(database, first.id, plots[0], 0, "manual", undefined, { ...metadata, keepForNext: true });
    const context = await store.getLotContext(database, first.id);
    check(context.keepForNext && context.storage === "Old storage" && context.firstOrder === "001", "Keep persists all five fields");
    let restored = await init();
    check(restored.activeSessionId === first.id, "Active session restores on reopen");
    check((await store.getWeights(restored.database, first.id))[0].weight === 0, "Zero restores as weighed");
    const before = await store.getSession(database, first.id);
    await store.beginNewSession(database);
    restored = await init();
    check(restored.activeSessionId === null, "Waiting for import persists on reopen");
    check(restored.sessions.length === 1, "Cancelled import creates no empty session");
    check(JSON.stringify(await store.getSession(database, first.id)) === JSON.stringify(before), "Previous session remains unchanged");
    const second = await store.createSession(database, [plots[1]], "second.csv", undefined, "seedWeight");
    check(second.id !== first.id && second.plots.length === 1, "New session has its own ID and dataset");
    check((await store.getWeights(database, second.id)).length === 0, "New session copies no weights");
    const blank = await store.getLotContext(database, second.id);
    check(blank.keepForNext && !blank.hasValue && blank.lotSite === "", "New session has empty continuity with Keep on");
    await store.setActiveSession(database, first.id);
    restored = await init();
    check(restored.activeSessionId === first.id, "Previous session can be selected again");
    await store.saveWeights(database, first.id, [], [{ ...plots[1], lotMetadata: metadata }]);
    check((await store.getWeights(database, first.id)).length === 1, "Metadata-only partial import creates no fake weight");
    check((await store.getSession(database, first.id)).plots[1].lotMetadata.secondOrder === "B", "Pending metadata merges by UUID");
    await store.saveWeights(database, first.id, [], [{ ...plots[1], lotMetadata: { ...metadata, storage: "Replacement" } }]);
    check((await store.getSession(database, first.id)).plots[1].lotMetadata.storage === "Old storage", "Partial metadata cannot overwrite existing values");
    await store.saveWeights(database, first.id, [{ uuid: "B", weight: 3689.0007, weightVariable: "plotWeight", ...metadata, weighedAt: "2026-10-08T12:00:00.000Z" }]);
    check((await store.getWeights(database, first.id)).find((r) => r.uuid === "B").weight === 3689.0007, "Exact imported weight persists");
    await store.saveWeight(database, first.id, plots[1], 12, "manual", undefined, { ...metadata, keepForNext: false });
    check(!(await store.getLotContext(database, first.id)).hasValue, "Keep off clears continuity");
    const saving = store.saveWeight(database, first.id, plots[1], 13, "manual", undefined, metadata);
    await saving; await store.beginNewSession(database);
    check((await store.getWeights(database, first.id)).find((r) => r.uuid === "B").weight === 13, "In-flight save completes before blank state");
    restored = await init();
    check(restored.activeSessionId === null, "Blank state remains after the completed save");
    output.textContent += "\n\nPASS: " + checks + " IndexedDB integration checks.";
  } catch (error) { output.textContent += "\nFAIL: " + error.message; }
  finally {
    opened.forEach((database) => database.close());
    // Only delete the unique database created above for this test run.
    indexedDB.deleteDatabase(databaseName);
    button.disabled = false;
  }
});
