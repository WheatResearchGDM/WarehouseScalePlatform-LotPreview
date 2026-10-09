(function (global) {
  "use strict";
  const DB_NAME = "gdm-warehouse-scale";
  const DB_VERSION = 1;
  const ACTIVE_SESSION_KEY = "active-session-id";
  const LOT_SITE_MIGRATION_KEY = "lot-site-v1-migrated";
  const LEGACY_PLOTS_KEY = "gdm-warehouse-scale-plots-v1";
  const LEGACY_WEIGHTS_KEY = "gdm-warehouse-scale-weights-v1";

  function requestResult(request) {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Browser storage request failed."));
    });
  }
  function transactionDone(transaction) {
    return new Promise((resolve, reject) => {
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("Browser storage transaction failed."));
      transaction.onabort = () => reject(transaction.error || new Error("Browser storage transaction was cancelled."));
    });
  }
  function openDatabase() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in global)) { reject(new Error("This browser does not support persistent session storage.")); return; }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains("sessions")) database.createObjectStore("sessions", { keyPath: "id" });
        if (!database.objectStoreNames.contains("weights")) {
          const weights = database.createObjectStore("weights", { keyPath: "key" });
          weights.createIndex("sessionId", "sessionId", { unique: false });
        }
        if (!database.objectStoreNames.contains("settings")) database.createObjectStore("settings", { keyPath: "key" });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Could not open persistent browser storage."));
    });
  }
  function uuid() { return global.crypto?.randomUUID?.() || `session-${Date.now()}-${Math.random().toString(16).slice(2)}`; }
  function baseName(fileName) { return String(fileName || "Weighing").replace(/\.[^.]+$/, "").trim() || "Weighing"; }
  function localName(fileName, date = new Date()) {
    const parts = new Intl.DateTimeFormat("en-CA", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).formatToParts(date);
    const get = (type) => parts.find((part) => part.type === type)?.value || "00";
    return `${baseName(fileName)} – ${get("year")}-${get("month")}-${get("day")} ${get("hour")}:${get("minute")}`;
  }
  function normalize(value) { return String(value || "").trim().toUpperCase(); }
  function lotContextKey(sessionId) { return `lot-context::${sessionId}`; }
  async function getSetting(database, key) {
    const tx = database.transaction("settings", "readonly");
    const result = await requestResult(tx.objectStore("settings").get(key));
    return result?.value ?? null;
  }
  async function setSetting(database, key, value) {
    const tx = database.transaction("settings", "readwrite");
    tx.objectStore("settings").put({ key, value });
    await transactionDone(tx);
  }
  async function listSessions(database) {
    const tx = database.transaction("sessions", "readonly");
    const sessions = await requestResult(tx.objectStore("sessions").getAll());
    return sessions.sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
  }
  async function getSession(database, id) {
    if (!id) return null;
    const tx = database.transaction("sessions", "readonly");
    return (await requestResult(tx.objectStore("sessions").get(id))) || null;
  }
  async function createSession(database, plots, fileName, name, initialWeights = []) {
    const now = new Date().toISOString();
    const session = { id: uuid(), name: name || localName(fileName), sourceFileName: fileName || "Imported workbook", createdAt: now, updatedAt: now, version: 2, plots };
    const tx = database.transaction(["sessions", "weights", "settings"], "readwrite");
    tx.objectStore("sessions").put(session);
    const weightsStore = tx.objectStore("weights");
    const byUuid = new Map(plots.map((plot) => [normalize(plot.uuid), plot]));
    for (const item of initialWeights) {
      const plot = byUuid.get(normalize(item.uuid));
      if (!plot || !Number.isFinite(Number(item.weight)) || Number(item.weight) < 0) continue;
      const timestamp = item.weighedAt || now;
      weightsStore.put({
        key: `${session.id}::${normalize(plot.uuid)}`, sessionId: session.id, uuid: plot.uuid, feid: plot.feid,
        entityName: plot.entityName, obsName: plot.obsName, weight: Number(item.weight), weighedAt: timestamp,
        updatedAt: timestamp, source: item.source || "import", lotSite: item.lotSite ?? "",
        lotLocation: item.lotLocation ?? "", storage: item.storage ?? "",
      });
    }
    tx.objectStore("settings").put({ key: ACTIVE_SESSION_KEY, value: session.id });
    await transactionDone(tx);
    return session;
  }
  async function getWeights(database, sessionId) {
    if (!sessionId) return [];
    const tx = database.transaction("weights", "readonly");
    return requestResult(tx.objectStore("weights").index("sessionId").getAll(sessionId));
  }
  async function saveWeight(database, sessionId, plot, weight, source = "manual", weighedAt = new Date().toISOString(), lot = {}) {
    const numericWeight = Number(weight);
    if (!sessionId || !plot?.uuid || !Number.isFinite(numericWeight) || numericWeight < 0) throw new Error("A valid plot and non-negative weight are required.");
    const parsed = new Date(weighedAt);
    const iso = Number.isNaN(parsed.getTime()) ? new Date().toISOString() : parsed.toISOString();
    const record = {
      key: `${sessionId}::${normalize(plot.uuid)}`, sessionId, uuid: plot.uuid, feid: plot.feid,
      entityName: plot.entityName, obsName: plot.obsName, weight: numericWeight, weighedAt: iso,
      updatedAt: iso, source, lotSite: String(lot.lotSite ?? "").trim(),
      lotLocation: String(lot.lotLocation ?? "").trim(), storage: String(lot.storage ?? "").trim(),
    };
    const tx = database.transaction(["weights", "sessions", "settings"], "readwrite");
    tx.objectStore("weights").put(record);
    const sessionsStore = tx.objectStore("sessions");
    const session = await requestResult(sessionsStore.get(sessionId));
    if (!session) { tx.abort(); throw new Error("The active weighing session no longer exists."); }
    session.updatedAt = new Date().toISOString();
    sessionsStore.put(session);
    if (Object.prototype.hasOwnProperty.call(lot, "keepForNext")) {
      tx.objectStore("settings").put({
        key: lotContextKey(sessionId),
        value: lot.keepForNext
          ? { keepForNext: true, hasValue: true, lotSite: record.lotSite, lotLocation: record.lotLocation, storage: record.storage }
          : { keepForNext: false, hasValue: false, lotSite: "", lotLocation: "", storage: "" },
      });
    }
    await transactionDone(tx);
    return record;
  }
  async function saveWeights(database, sessionId, entries) {
    if (!entries.length) return [];
    const session = await getSession(database, sessionId);
    if (!session) throw new Error("The active weighing session no longer exists.");
    const plotMap = new Map(session.plots.map((plot) => [normalize(plot.uuid), plot]));
    const now = new Date().toISOString();
    const records = [];
    const tx = database.transaction(["weights", "sessions"], "readwrite");
    const store = tx.objectStore("weights");
    for (const entry of entries) {
      const plot = plotMap.get(normalize(entry.uuid));
      if (!plot || !Number.isFinite(Number(entry.weight)) || Number(entry.weight) < 0) continue;
      const parsed = new Date(entry.weighedAt || now);
      const iso = Number.isNaN(parsed.getTime()) ? now : parsed.toISOString();
      const record = {
        key: `${sessionId}::${normalize(plot.uuid)}`, sessionId, uuid: plot.uuid, feid: plot.feid,
        entityName: plot.entityName, obsName: plot.obsName, weight: Number(entry.weight), weighedAt: iso,
        updatedAt: iso, source: entry.source || "import", lotSite: entry.lotSite ?? "",
        lotLocation: entry.lotLocation ?? "", storage: entry.storage ?? "",
      };
      store.put(record);
      records.push(record);
    }
    session.updatedAt = now;
    tx.objectStore("sessions").put(session);
    await transactionDone(tx);
    return records;
  }
  async function renameSession(database, id, name) {
    const cleanName = String(name || "").trim();
    if (!cleanName) throw new Error("Session name cannot be empty.");
    const session = await getSession(database, id);
    if (!session) throw new Error("Session not found.");
    session.name = cleanName;
    session.updatedAt = new Date().toISOString();
    const tx = database.transaction("sessions", "readwrite");
    tx.objectStore("sessions").put(session);
    await transactionDone(tx);
    return session;
  }
  async function deleteSession(database, id) {
    const tx = database.transaction(["sessions", "weights", "settings"], "readwrite");
    tx.objectStore("sessions").delete(id);
    const store = tx.objectStore("weights");
    const keys = await requestResult(store.index("sessionId").getAllKeys(id));
    for (const key of keys) store.delete(key);
    const active = await requestResult(tx.objectStore("settings").get(ACTIVE_SESSION_KEY));
    if (active?.value === id) tx.objectStore("settings").put({ key: ACTIVE_SESSION_KEY, value: null });
    tx.objectStore("settings").delete(lotContextKey(id));
    await transactionDone(tx);
  }
  async function migrateLegacy(database) {
    if ((await listSessions(database)).length) return;
    let dataset = null;
    let weights = {};
    try { dataset = JSON.parse(localStorage.getItem(LEGACY_PLOTS_KEY) || "null"); } catch { dataset = null; }
    try { weights = JSON.parse(localStorage.getItem(LEGACY_WEIGHTS_KEY) || "{}"); } catch { weights = {}; }
    if (!dataset?.plots?.length) return;
    const initial = Object.values(weights || {}).map((item) => ({ ...item, weighedAt: item.weighedAt || item.updatedAt || new Date().toISOString(), source: "legacy" }));
    await createSession(database, dataset.plots, dataset.fileName || "Legacy workbook", "Imported legacy session", initial);
  }
  async function migrateLotSite(database) {
    if (await getSetting(database, LOT_SITE_MIGRATION_KEY)) return;
    const tx = database.transaction(["weights", "settings"], "readwrite");
    const weightsStore = tx.objectStore("weights");
    const settingsStore = tx.objectStore("settings");
    const [weights, settings] = await Promise.all([
      requestResult(weightsStore.getAll()),
      requestResult(settingsStore.getAll()),
    ]);
    for (const record of weights) {
      delete record.site;
      record.lotSite = "";
      weightsStore.put(record);
    }
    for (const setting of settings) {
      if (!String(setting.key || "").startsWith("lot-context::")) continue;
      const value = setting.value && typeof setting.value === "object" ? setting.value : {};
      const storage = String(value.storage || "");
      setting.value = {
        keepForNext: value.keepForNext !== false,
        hasValue: value.keepForNext !== false && Boolean(storage),
        lotSite: "", lotLocation: "",
        storage,
      };
      settingsStore.put(setting);
    }
    settingsStore.put({ key: LOT_SITE_MIGRATION_KEY, value: true });
    await transactionDone(tx);
  }
  async function init() {
    const database = await openDatabase();
    await migrateLegacy(database);
    await migrateLotSite(database);
    const sessions = await listSessions(database);
    let activeSessionId = await getSetting(database, ACTIVE_SESSION_KEY);
    if (!sessions.some((session) => session.id === activeSessionId)) {
      activeSessionId = sessions[0]?.id || null;
      await setSetting(database, ACTIVE_SESSION_KEY, activeSessionId);
    }
    return { database, sessions: await listSessions(database), activeSessionId };
  }

  global.GdmWeighingStore = {
    init, listSessions, getSession, createSession, getWeights, saveWeight, saveWeights, renameSession, deleteSession,
    async getLotContext(database, sessionId) {
      const value = await getSetting(database, lotContextKey(sessionId));
      return value && typeof value === "object"
        ? { keepForNext: value.keepForNext !== false, hasValue: value.hasValue === true, lotSite: String(value.lotSite || ""), lotLocation: String(value.lotLocation || ""), storage: String(value.storage || "") }
        : { keepForNext: true, hasValue: false, lotSite: "", lotLocation: "", storage: "" };
    },
    setLotContext(database, sessionId, context) {
      return setSetting(database, lotContextKey(sessionId), {
        keepForNext: context?.keepForNext !== false,
        hasValue: context?.hasValue === true,
        lotSite: String(context?.lotSite || ""), lotLocation: String(context?.lotLocation || ""), storage: String(context?.storage || ""),
      });
    },
    setActiveSession(database, id) { return setSetting(database, ACTIVE_SESSION_KEY, id); },
    sessionName: localName, normalize,
  };
})(window);
