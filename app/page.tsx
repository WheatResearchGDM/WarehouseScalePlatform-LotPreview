"use client";
/* eslint-disable @next/next/no-img-element */

import {
  Barcode, Cable, Check, CircleAlert, Download, Gauge, LayoutDashboard, LoaderCircle,
  MapPin, Pencil, Plus, Scale, ScanLine, Search, Trash2, Upload,
} from "lucide-react";
import { ChangeEvent, FormEvent, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { NativeSelect, NativeSelectOption } from "@/components/ui/native-select";
import { Toaster } from "@/components/ui/sonner";

type Plot = {
  id: string; feid: string; uuid: string; seasonYear: string; entityName: string; trialType: string; site: string;
  location: string; row: string; column: string; entryCode: string; block: string; obsName: string;
  gid: string; gerName: string; initialPlot: number; finalPlot: number;
};
type WeightVariable = "plotWeight" | "seedWeight";
type WeightRecord = {
  key: string; sessionId: string; uuid: string; feid: string; entityName: string; obsName: string;
  weight: number; weightVariable: WeightVariable; weighedAt: string; updatedAt: string; source: string; lotSite?: string; lotLocation?: string; storage?: string;
};
type WeighingSession = {
  id: string; name: string; sourceFileName: string; createdAt: string; updatedAt: string; version: number; weightVariable: WeightVariable; plots: Plot[];
};
type ImportedWeight = { uuid: string; weight: number; weightVariable: WeightVariable; weighedAt: string; source?: string; lotSite?: string; lotLocation?: string; storage?: string };
type ImportResult = {
  plots: Plot[]; importedWeights: ImportedWeight[]; invalidRows: number[];
  invalidWeightRows: { row: number; weightVariable: WeightVariable; label: string }[];
  weightConflicts: { row: number; weightVariable: WeightVariable; label: string; values: number[] }[];
  weightValueCounts: Record<WeightVariable, number>; weightColumns: Record<WeightVariable, boolean>;
  fileName: string; sheetName: string;
};
type ProgressItem = {
  key: string; label: string; total: number; completed: number; remaining: number; percent: number;
  trialType?: string; location?: string;
};
type OverallProgress = { total: number; completed: number; remaining: number; percent: number };
type MergePlan = { ready: ImportedWeight[]; unresolved: ImportedWeight[]; ignored: number; unchanged: number; keptCurrent: number };
type ScanAlert = { kind: "existing" | "unavailable"; title: string; message: string; actionLabel: string };
type StoreApi = {
  init(): Promise<{ database: IDBDatabase; sessions: WeighingSession[]; activeSessionId: string | null }>;
  listSessions(database: IDBDatabase): Promise<WeighingSession[]>;
  getSession(database: IDBDatabase, id: string | null): Promise<WeighingSession | null>;
  createSession(database: IDBDatabase, plots: Plot[], fileName: string, name?: string, weightVariable?: WeightVariable, initialWeights?: ImportedWeight[]): Promise<WeighingSession>;
  getWeights(database: IDBDatabase, sessionId: string): Promise<WeightRecord[]>;
  saveWeight(database: IDBDatabase, sessionId: string, plot: Plot, weight: number, source?: string, weighedAt?: string, lot?: { lotSite: string; lotLocation: string; storage: string; keepForNext: boolean }): Promise<WeightRecord>;
  saveWeights(database: IDBDatabase, sessionId: string, entries: ImportedWeight[]): Promise<WeightRecord[]>;
  getLotContext(database: IDBDatabase, sessionId: string): Promise<{ keepForNext: boolean; hasValue: boolean; lotSite: string; lotLocation: string; storage: string }>;
  setLotContext(database: IDBDatabase, sessionId: string, context: { keepForNext: boolean; hasValue: boolean; lotSite: string; lotLocation: string; storage: string }): Promise<void>;
  renameSession(database: IDBDatabase, id: string, name: string): Promise<WeighingSession>;
  setWeightVariable(database: IDBDatabase, id: string, weightVariable: WeightVariable): Promise<WeighingSession>;
  deleteSession(database: IDBDatabase, id: string): Promise<void>;
  setActiveSession(database: IDBDatabase, id: string | null): Promise<void>;
};
type UtilsApi = {
  normalize(value: unknown): string;
  groupProgress(plots: Plot[], weights: WeightRecord[], key: "trial" | "location"): ProgressItem[];
  overallProgress(plots: Plot[], weights: WeightRecord[]): OverallProgress;
  prepareMerge(plots: Plot[], current: WeightRecord[], imported: ImportedWeight[]): MergePlan;
  exportSession(session: WeighingSession, weights: WeightRecord[], format: "xlsx" | "csv", plots?: Plot[], decimalPlaces?: number): number;
  exportLots(session: WeighingSession, weights: WeightRecord[], format: "xlsx" | "csv", plots?: Plot[], decimalPlaces?: number): number;
  weightLabel(sessionOrVariable: WeighingSession | WeightVariable): string;
};
type ImportApi = { parseExcelFile(file: File): Promise<ImportResult> };
type SerialPortLike = {
  readable: ReadableStream<Uint8Array> | null; open(options: { baudRate: number }): Promise<void>;
  close(): Promise<void>; getInfo?(): { usbVendorId?: number; usbProductId?: number };
};
type SerialLike = {
  requestPort(): Promise<SerialPortLike>;
  addEventListener(type: "disconnect", listener: (event: Event) => void): void;
  removeEventListener(type: "disconnect", listener: (event: Event) => void): void;
};

declare global {
  interface Window { GdmWeighingStore?: StoreApi; GdmWeighingUtils?: UtilsApi; GdmPlotImport?: ImportApi; }
}

const PAGE_SIZE = 100;
const SCALE_EXPONENT_KEY = "gdm-warehouse-scale-exponent-v1";
const DECIMAL_PLACES_KEY = "gdm-warehouse-decimal-places-v1";
const WEIGHT_VARIABLE_LABELS: Record<WeightVariable, string> = { plotWeight: "Plot weight", seedWeight: "Seed weight" };
function weightLabel(variable: WeightVariable | undefined) { return WEIGHT_VARIABLE_LABELS[variable === "seedWeight" ? "seedWeight" : "plotWeight"]; }

function loadBrowserScript(src: string, id: string) {
  return new Promise<void>((resolve, reject) => {
    const existing = document.getElementById(id) as HTMLScriptElement | null;
    if (existing?.dataset.loaded === "true") { resolve(); return; }
    const script = existing ?? document.createElement("script");
    script.id = id;
    script.addEventListener("load", () => { script.dataset.loaded = "true"; resolve(); }, { once: true });
    script.addEventListener("error", () => reject(new Error(`Could not load ${src}.`)), { once: true });
    if (!existing) { script.src = src; document.head.appendChild(script); }
  });
}

async function loadRuntime() {
  await loadBrowserScript("/vendor/xlsx.full.min.js", "gdm-xlsx-runtime");
  await loadBrowserScript("/plot-import.js", "gdm-import-runtime");
  await loadBrowserScript("/weighing-store.js", "gdm-store-runtime");
  await loadBrowserScript("/weighing-utils.js", "gdm-utils-runtime");
  if (!window.GdmPlotImport || !window.GdmWeighingStore || !window.GdmWeighingUtils) throw new Error("The weighing runtime could not be initialized.");
  return { importer: window.GdmPlotImport, store: window.GdmWeighingStore, utils: window.GdmWeighingUtils };
}

function normalize(value: unknown) { return String(value ?? "").trim().toUpperCase(); }
function parseWeight(value: string) { const clean = value.trim(); return clean ? Number(clean.replace(",", ".")) : Number.NaN; }
function formatNumber(value: number, decimalPlaces = 0) { return new Intl.NumberFormat("en-US", { minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces }).format(value); }
function formatInputNumber(value: number, decimalPlaces = 0) { return new Intl.NumberFormat("en-US", { useGrouping: false, minimumFractionDigits: decimalPlaces, maximumFractionDigits: decimalPlaces }).format(value); }
function formatRawNumber(value: number) { return new Intl.NumberFormat("en-US", { maximumFractionDigits: 12 }).format(value); }
function plotDisplayName(plot: Plot) { const name = plot.obsName || plot.feid || "—"; return /^plot\b/i.test(name) ? name : `Plot ${name}`; }
function formatDateTime(value: string) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "Unavailable" : new Intl.DateTimeFormat("en-US", { dateStyle: "medium", timeStyle: "short" }).format(date);
}
function getSerial() { return (navigator as Navigator & { serial?: SerialLike }).serial; }
function parseScaleWeight(rawLine: string) {
  const matches = rawLine.replace(/\u0000/g, " ").trim().match(/[-+]?\d+(?:[.,]\d+)?/g);
  if (!matches?.length) return null;
  const value = Number(matches[matches.length - 1].replace(",", "."));
  return Number.isFinite(value) && value >= 0 ? value : null;
}
function superscript(value: number) { return String(value).replace(/\d/g, (digit) => "⁰¹²³⁴⁵⁶⁷⁸⁹"[Number(digit)]); }
function scaleFactorLabel(exponent: number) { return exponent < 0 ? `× 10${superscript(Math.abs(exponent))}` : exponent > 0 ? `÷ 10${superscript(exponent)}` : "raw value"; }
let scanAudioContext: AudioContext | null = null;
function playScanTone(kind: "found" | "existing" | "attention") {
  try {
    const AudioContextClass = window.AudioContext || (window as typeof window & { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
    if (!AudioContextClass) return;
    const context = scanAudioContext ?? new AudioContextClass();
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

function calculateProgress(plots: Plot[], weights: WeightRecord[], grouping: "trial" | "location") {
  const records = new Set(weights.map((item) => normalize(item.uuid)));
  if (grouping === "location") {
    const locations = new Map<string, ProgressItem>();
    for (const plot of plots) {
      const key = plot.location || "Unspecified";
      const current = locations.get(key) ?? { key, label: key, total: 0, completed: 0, remaining: 0, percent: 0 };
      current.total += 1;
      if (records.has(normalize(plot.uuid))) current.completed += 1;
      current.remaining = Math.max(current.total - current.completed, 0);
      current.percent = current.total ? Math.min(Math.round((current.completed / current.total) * 100), 100) : 0;
      locations.set(key, current);
    }
    return Array.from(locations.values());
  }
  const trials = new Map<string, Plot[]>();
  for (const plot of plots) {
    const key = plot.entityName || "Unnamed trial";
    trials.set(key, [...(trials.get(key) ?? []), plot]);
  }
  const trialItems: ProgressItem[] = Array.from(trials.entries()).map(([entityName, items]) => {
    const total = items.length;
    const completed = items.filter((item) => records.has(normalize(item.uuid))).length;
    return {
      key: entityName, label: entityName, trialType: items[0]?.trialType || "—",
      location: items[0]?.location || "Unspecified", total, completed,
      remaining: Math.max(total - completed, 0),
      percent: total ? Math.min(Math.round((completed / total) * 100), 100) : 0,
    };
  });
  return trialItems;
}

function calculateOverall(plots: Plot[], weights: WeightRecord[]): OverallProgress {
  const trials = calculateProgress(plots, weights, "trial");
  const total = trials.reduce((sum, item) => sum + item.total, 0);
  const completed = trials.reduce((sum, item) => sum + item.completed, 0);
  return { total, completed, remaining: Math.max(total - completed, 0), percent: total ? Math.min(Math.round((completed / total) * 100), 100) : 0 };
}

function DonutCard({ item }: { item: ProgressItem }) {
  return (
    <article className="grid grid-cols-[92px_minmax(0,1fr)] items-center gap-4 rounded-lg border border-[#d3e0ec] bg-[#f8fbfe] p-4">
      <div className="relative grid size-[92px] place-items-center rounded-full" style={{ background: `conic-gradient(#4f8fc9 ${item.percent}%, #dfe8ef 0)` }} role="img" aria-label={`${item.label}: ${item.percent}% complete`}>
        <div className="absolute inset-3 rounded-full bg-white" />
        <strong className="relative z-10 text-lg font-black text-[#173a61]">{item.percent}%</strong>
      </div>
      <div className="min-w-0">
        <h3 className="break-words text-sm font-extrabold leading-tight text-[#173a61]">{item.label}</h3>
        <p className="mt-2 text-xs text-[#657b90]"><strong>{item.completed}</strong> weighed · <strong>{item.remaining}</strong> pending</p>
      </div>
    </article>
  );
}

export default function Home() {
  const [database, setDatabase] = useState<IDBDatabase | null>(null);
  const [sessions, setSessions] = useState<WeighingSession[]>([]);
  const [session, setSession] = useState<WeighingSession | null>(null);
  const [plots, setPlots] = useState<Plot[]>([]);
  const [weights, setWeights] = useState<WeightRecord[]>([]);
  const [activeTab, setActiveTab] = useState<"weighing" | "dashboard">("weighing");
  const [loading, setLoading] = useState(true);
  const [importing, setImporting] = useState(false);
  const [scanMode, setScanMode] = useState<"feid" | "uuid">("feid");
  const [weightVariable, setWeightVariable] = useState<WeightVariable>("plotWeight");
  const [scanValue, setScanValue] = useState("");
  const [scanError, setScanError] = useState("");
  const [scanAlert, setScanAlert] = useState<ScanAlert | null>(null);
  const [selected, setSelected] = useState<Plot | null>(null);
  const [weightValue, setWeightValue] = useState("");
  const [lotSite, setLotSite] = useState("");
  const [lotLocation, setLotLocation] = useState("");
  const [lotStorage, setLotStorage] = useState("");
  const [carriedLotSite, setCarriedLotSite] = useState("");
  const [carriedLotLocation, setCarriedLotLocation] = useState("");
  const [carriedStorage, setCarriedStorage] = useState("");
  const [keepLotContext, setKeepLotContext] = useState(true);
  const [hasLotContext, setHasLotContext] = useState(false);
  const [saving, setSaving] = useState(false);
  const [search, setSearch] = useState("");
  const [trialFilter, setTrialFilter] = useState("");
  const [locationFilter, setLocationFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [page, setPage] = useState(1);
  const [selectedTrials, setSelectedTrials] = useState<string[]>([]);
  const [sort, setSort] = useState<{ key: string; direction: "asc" | "desc" }>({ key: "", direction: "asc" });
  const [baudRate, setBaudRate] = useState("9600");
  const [scaleExponent, setScaleExponent] = useState("0");
  const [decimalPlaces, setDecimalPlaces] = useState(0);
  const [scaleConnected, setScaleConnected] = useState(false);
  const [scaleStatus, setScaleStatus] = useState("Not connected");
  const [scaleWeight, setScaleWeight] = useState<number | null>(null);
  const [rawScaleWeight, setRawScaleWeight] = useState<number | null>(null);

  const fileRef = useRef<HTMLInputElement>(null);
  const scanRef = useRef<HTMLInputElement>(null);
  const scanAlertButtonRef = useRef<HTMLButtonElement>(null);
  const scanAlertUpdateButtonRef = useRef<HTMLButtonElement>(null);
  const scanAlertRef = useRef<ScanAlert | null>(null);
  const scanAlertOpenedAtRef = useRef(0);
  const weightRef = useRef<HTMLInputElement>(null);
  const selectedRef = useRef<Plot | null>(null);
  const serialPortRef = useRef<SerialPortLike | null>(null);
  const serialReaderRef = useRef<ReadableStreamDefaultReader<Uint8Array> | null>(null);
  const serialTaskRef = useRef<Promise<void> | null>(null);
  const keepReadingRef = useRef(false);
  const serialBufferRef = useRef("");
  const serialTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const exponentRef = useRef(0);
  const decimalPlacesRef = useRef(0);
  const exactPendingWeightRef = useRef<number | null>(null);
  const weightEditedRef = useRef(false);

  useEffect(() => { selectedRef.current = selected; }, [selected]);

  const closeScanAlert = useCallback((focusTarget: "scan" | "weight" = "scan") => {
    scanAlertRef.current = null;
    setScanAlert(null);
    window.setTimeout(() => {
      const target = focusTarget === "weight" ? weightRef.current : scanRef.current;
      target?.focus(); target?.select();
    }, 0);
  }, []);

  const keepExistingWeight = useCallback(() => {
    exactPendingWeightRef.current = null; weightEditedRef.current = false;
    setSelected(null); setScanValue(""); setWeightValue("");
    closeScanAlert();
  }, [closeScanAlert]);

  const updateExistingWeight = useCallback(() => closeScanAlert("weight"), [closeScanAlert]);

  const openScanAlert = useCallback((alert: ScanAlert) => {
    scanAlertRef.current = alert;
    scanAlertOpenedAtRef.current = performance.now();
    setScanAlert(alert);
  }, []);

  useEffect(() => {
    if (!scanAlert) return;
    window.setTimeout(() => scanAlertButtonRef.current?.focus(), 0);
    const closeOnKeyboard = (event: KeyboardEvent) => {
      if (event.key === "Tab") {
        event.preventDefault();
        if (scanAlert.kind !== "existing") { scanAlertButtonRef.current?.focus(); return; }
        const updateButton = scanAlertUpdateButtonRef.current; const keepButton = scanAlertButtonRef.current;
        (document.activeElement === keepButton ? updateButton : keepButton)?.focus();
        return;
      }
      if (event.key === "Enter") {
        event.preventDefault(); event.stopPropagation();
        if (performance.now() - scanAlertOpenedAtRef.current < 700) return;
        if (scanAlert.kind === "existing") {
          if (document.activeElement === scanAlertUpdateButtonRef.current) updateExistingWeight(); else keepExistingWeight();
        } else closeScanAlert();
        return;
      }
      if (event.key === "Escape") {
        event.preventDefault(); event.stopPropagation();
        if (scanAlert.kind === "existing") keepExistingWeight(); else closeScanAlert();
      }
    };
    window.addEventListener("keydown", closeOnKeyboard, true);
    return () => window.removeEventListener("keydown", closeOnKeyboard, true);
  }, [scanAlert, closeScanAlert, keepExistingWeight, updateExistingWeight]);

  const byFeid = useMemo(() => new Map(plots.map((plot) => [normalize(plot.feid), plot])), [plots]);
  const byUuid = useMemo(() => new Map(plots.map((plot) => [normalize(plot.uuid), plot])), [plots]);
  const weightsByUuid = useMemo(() => new Map(weights.map((record) => [normalize(record.uuid), record])), [weights]);
  const trials = useMemo(() => [...new Set(plots.map((plot) => plot.entityName || "Unnamed trial"))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })), [plots]);
  const selectedTrialSet = useMemo(() => new Set(selectedTrials), [selectedTrials]);
  const activePlots = useMemo(() => plots.filter((plot) => selectedTrialSet.has(plot.entityName || "Unnamed trial")), [plots, selectedTrialSet]);
  const overall = useMemo<OverallProgress>(() => calculateOverall(activePlots, weights), [activePlots, weights]);
  const trialProgress = useMemo(() => calculateProgress(activePlots, weights, "trial"), [activePlots, weights]);
  const locationProgress = useMemo(() => calculateProgress(activePlots, weights, "location"), [activePlots, weights]);
  const recent = useMemo(() => weights.filter((record) => selectedTrialSet.has(record.entityName || "Unnamed trial")).sort((a, b) => Date.parse(b.weighedAt || b.updatedAt) - Date.parse(a.weighedAt || a.updatedAt)).slice(0, 10), [weights, selectedTrialSet]);
  const locations = useMemo(() => [...new Set(activePlots.map((plot) => plot.location).filter(Boolean))].sort(), [activePlots]);
  const filteredPlots = useMemo(() => {
    const needle = normalize(search);
    const rows = activePlots.filter((plot) => {
      const weighed = weightsByUuid.has(normalize(plot.uuid));
      const haystack = normalize([plot.feid, plot.uuid, plot.seasonYear, plot.obsName, plot.entityName, plot.gerName, plot.location].join(" "));
      return (!needle || haystack.includes(needle)) && (!trialFilter || plot.entityName === trialFilter)
        && (!locationFilter || plot.location === locationFilter) && (!statusFilter || (statusFilter === "weighed" ? weighed : !weighed));
    });
    if (sort.key) {
      const collator = new Intl.Collator("en-US", { numeric: true, sensitivity: "base" });
      rows.sort((left, right) => {
        const leftRecord = weightsByUuid.get(normalize(left.uuid)); const rightRecord = weightsByUuid.get(normalize(right.uuid));
        const value = (plot: Plot, record?: WeightRecord): string | number => sort.key === "status" ? (record ? "Weighed" : "Pending") : sort.key === "weight" ? (record?.weight ?? Number.POSITIVE_INFINITY) : sort.key === "weighedAt" ? (record?.weighedAt || record?.updatedAt || "") : String(plot[sort.key as keyof Plot] ?? "");
        const a = value(left, leftRecord); const b = value(right, rightRecord);
        const result = typeof a === "number" && typeof b === "number" ? a - b : collator.compare(String(a), String(b));
        return sort.direction === "desc" ? -result : result;
      });
    }
    return rows;
  }, [activePlots, weightsByUuid, search, trialFilter, locationFilter, statusFilter, sort]);
  const pageCount = Math.max(Math.ceil(filteredPlots.length / PAGE_SIZE), 1);
  const safePage = Math.min(page, pageCount);
  const pageRows = filteredPlots.slice((safePage - 1) * PAGE_SIZE, safePage * PAGE_SIZE);
  const existingWeight = selected ? weightsByUuid.get(normalize(selected.uuid)) : undefined;
  const currentWeightLabel = weightLabel(weightVariable);

  const applySession = useCallback(async (db: IDBDatabase, id: string | null) => {
    const store = window.GdmWeighingStore;
    if (!store) throw new Error("Session storage is unavailable.");
    const nextSession = await store.getSession(db, id);
    const nextWeights = nextSession ? await store.getWeights(db, nextSession.id) : [];
    const lotContext = nextSession ? await store.getLotContext(db, nextSession.id) : { keepForNext: true, hasValue: false, lotSite: "", lotLocation: "", storage: "" };
    const nextPlots = nextSession?.plots ?? [];
    exactPendingWeightRef.current = null; weightEditedRef.current = false;
    scanAlertRef.current = null;
    setSession(nextSession); setPlots(nextPlots); setWeights(nextWeights); setWeightVariable(nextSession?.weightVariable || "plotWeight"); setSelected(null); setScanAlert(null);
    setKeepLotContext(lotContext.keepForNext); setHasLotContext(lotContext.hasValue);
    setCarriedLotSite(lotContext.lotSite); setCarriedLotLocation(lotContext.lotLocation); setCarriedStorage(lotContext.storage);
    setLotSite(""); setLotLocation(""); setLotStorage("");
    setSelectedTrials([...new Set(nextPlots.map((plot) => plot.entityName || "Unnamed trial"))]); setSort({ key: "", direction: "asc" });
    setScanValue(""); setWeightValue(""); setPage(1); setSessions(await store.listSessions(db));
    window.setTimeout(() => scanRef.current?.focus(), 0);
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const { store } = await loadRuntime();
        const exponent = Number(localStorage.getItem(SCALE_EXPONENT_KEY) || 0);
        if (Number.isInteger(exponent) && exponent >= -10 && exponent <= 10) { exponentRef.current = exponent; setScaleExponent(String(exponent)); }
        const savedDecimalPlaces = Number(localStorage.getItem(DECIMAL_PLACES_KEY) || 0);
        if (Number.isInteger(savedDecimalPlaces) && savedDecimalPlaces >= 0 && savedDecimalPlaces <= 6) { decimalPlacesRef.current = savedDecimalPlaces; setDecimalPlaces(savedDecimalPlaces); }
        const initialized = await store.init();
        if (cancelled) return;
        setDatabase(initialized.database);
        await applySession(initialized.database, initialized.activeSessionId);
        if (initialized.sessions.some((item) => item.name === "Imported legacy session")) toast.success("Existing browser data was migrated to “Imported legacy session”.");
      } catch (error) { toast.error(error instanceof Error ? error.message : "Could not initialize browser storage."); }
      finally { if (!cancelled) setLoading(false); }
    })();
    return () => { cancelled = true; };
  }, [applySession]);

  async function changeSession(id: string) {
    if (!database || !window.GdmWeighingStore) return;
    await window.GdmWeighingStore.setActiveSession(database, id || null);
    await applySession(database, id || null);
  }

  async function changeWeightVariable(next: WeightVariable) {
    if (weights.length) { toast.error("The weight variable cannot be changed after weights exist. Use Start New Weighing."); return; }
    if (!session) { setWeightVariable(next); return; }
    if (!database || !window.GdmWeighingStore) return;
    try {
      await window.GdmWeighingStore.setWeightVariable(database, session.id, next);
      await applySession(database, session.id);
      toast.success(`${weightLabel(next)} selected for this session.`);
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not change the weight variable."); }
  }

  async function handleImport(event: ChangeEvent<HTMLInputElement>) {
    const [file] = Array.from(event.target.files ?? []); event.target.value = "";
    if (!file || !database || !window.GdmPlotImport || !window.GdmWeighingStore || !window.GdmWeighingUtils) return;
    setImporting(true);
    try {
      const result = await window.GdmPlotImport.parseExcelFile(file);
      const targetVariable = session?.weightVariable || weightVariable;
      const targetLabel = weightLabel(targetVariable);
      const otherVariable: WeightVariable = targetVariable === "plotWeight" ? "seedWeight" : "plotWeight";
      const importedWeights = result.importedWeights.filter((item) => item.weightVariable === targetVariable);
      const ignoredOtherWeights = result.importedWeights.filter((item) => item.weightVariable === otherVariable).length;
      const relevantInvalid = result.invalidWeightRows.filter((item) => item.weightVariable === targetVariable).length;
      const targetConflicts = result.weightConflicts.filter((item) => item.weightVariable === targetVariable);
      const relevantConflicts = targetConflicts.length;
      const invalid = result.invalidRows.length + relevantInvalid;
      if (!importedWeights.length && !relevantInvalid && !relevantConflicts && ignoredOtherWeights > 0) {
        throw new Error(`This file contains ${weightLabel(otherVariable)} values but the current session uses ${targetLabel}. Select the matching variable in an empty session or use Start New Weighing.`);
      }
      const importNotes = [
        ignoredOtherWeights ? `${ignoredOtherWeights} ${weightLabel(otherVariable)} value(s) ignored` : "",
        relevantConflicts ? `${relevantConflicts} conflicting ${targetLabel} row(s) skipped (${targetConflicts.slice(0, 5).map((item) => item.row).join(", ")}${relevantConflicts > 5 ? ", …" : ""})` : "",
      ].filter(Boolean);
      if (!session || !importedWeights.length) {
        const created = await window.GdmWeighingStore.createSession(database, result.plots, result.fileName, undefined, targetVariable, importedWeights);
        await window.GdmWeighingStore.setActiveSession(database, created.id);
        await applySession(database, created.id);
        toast.success(`Session created for ${targetLabel} with ${result.plots.length} plots and ${importedWeights.length} existing weight${importedWeights.length === 1 ? "" : "s"}${invalid ? `; ${invalid} invalid row(s) skipped` : ""}${importNotes.length ? `; ${importNotes.join("; ")}` : ""}.`);
      } else {
        const merge = window.GdmWeighingUtils.prepareMerge(plots, weights, importedWeights);
        const overwrite = merge.unresolved.length
          ? window.confirm(`${merge.unresolved.length} conflicting weight(s) do not have comparable timestamps. Select OK to use the imported values, or Cancel to keep the current values.`)
          : false;
        const entries = [...merge.ready, ...(overwrite ? merge.unresolved : [])];
        await window.GdmWeighingStore.saveWeights(database, session.id, entries);
        await applySession(database, session.id);
        const kept = merge.keptCurrent + (overwrite ? 0 : merge.unresolved.length);
        toast.success(`Partial ${targetLabel} results: ${entries.length} imported, ${kept} current kept, ${merge.unchanged} unchanged, ${merge.ignored} unmatched, ${invalid} invalid${importNotes.length ? `; ${importNotes.join("; ")}` : ""}.`);
      }
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not import the file."); }
    finally { setImporting(false); }
  }

  async function startNewSession() {
    if (!database || !window.GdmWeighingStore) return;
    if (!session) { fileRef.current?.click(); return; }
    try {
      const created = await window.GdmWeighingStore.createSession(database, plots, session.sourceFileName, undefined, "plotWeight");
      await window.GdmWeighingStore.setActiveSession(database, created.id);
      await applySession(database, created.id);
      toast.success(`New session “${created.name}” started.`);
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not start a new session."); }
  }

  async function renameActiveSession() {
    if (!database || !session || !window.GdmWeighingStore) return;
    const name = window.prompt("Enter a new name for this session:", session.name);
    if (name === null || !name.trim() || name.trim() === session.name) return;
    try { await window.GdmWeighingStore.renameSession(database, session.id, name); await applySession(database, session.id); toast.success("Session renamed."); }
    catch (error) { toast.error(error instanceof Error ? error.message : "Could not rename the session."); }
  }

  async function deleteActiveSession() {
    if (!database || !session || !window.GdmWeighingStore || !window.confirm(`Delete “${session.name}” and all of its locally saved weights? This cannot be undone.`)) return;
    try {
      await window.GdmWeighingStore.deleteSession(database, session.id);
      const remaining = await window.GdmWeighingStore.listSessions(database);
      await window.GdmWeighingStore.setActiveSession(database, remaining[0]?.id || null);
      await applySession(database, remaining[0]?.id || null);
      toast.success("Session deleted.");
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not delete the session."); }
  }

  function selectPlot(plot: Plot) {
    setSelected(plot); setScanError(""); setScanValue("");
    const record = weightsByUuid.get(normalize(plot.uuid));
    exactPendingWeightRef.current = record ? Number(record.weight) : null;
    weightEditedRef.current = false;
    setWeightValue(record ? formatInputNumber(Number(record.weight), decimalPlaces) : "");
    if (record) {
      setLotSite(String(record.lotSite || ""));
      setLotLocation(String(record.lotLocation || ""));
      setLotStorage(record.storage || "");
    } else if (!keepLotContext || !hasLotContext) {
      setLotSite(""); setLotLocation(""); setLotStorage("");
    } else {
      setLotSite(carriedLotSite); setLotLocation(carriedLotLocation); setLotStorage(carriedStorage);
    }
    window.setTimeout(() => scanRef.current?.focus(), 0);
  }
  function handleScan(event: FormEvent) {
    event.preventDefault();
    if (scanAlertRef.current) return;
    if (!session) { toast.error("Import a workbook before weighing plots."); return; }
    const code = normalize(scanValue);
    const currentCode = selected ? normalize(scanMode === "feid" ? selected.feid : selected.uuid) : "";
    if (selected && (!code || code === currentCode)) { void saveCurrentWeight(); return; }
    const plot = scanMode === "feid" ? byFeid.get(code) : byUuid.get(code);
    if (!plot || !selectedTrialSet.has(plot.entityName || "Unnamed trial")) {
      const message = plot
        ? `${plotDisplayName(plot)} belongs to trial ${plot.entityName || "Unnamed trial"}, which is excluded by the current trial filter.`
        : `${scanMode.toUpperCase()} ${scanValue.trim() || "—"} was not found in this session.`;
      setSelected(null); setWeightValue(""); exactPendingWeightRef.current = null; weightEditedRef.current = false;
      playScanTone("attention"); setScanError(message); openScanAlert({ kind: "unavailable", title: "Plot unavailable", message, actionLabel: "Scan again" }); return;
    }
    const record = weightsByUuid.get(normalize(plot.uuid));
    playScanTone(record ? "existing" : "found");
    selectPlot(plot);
    if (record) openScanAlert({
      kind: "existing", title: "Plot already weighed",
      message: `${plotDisplayName(plot)} already has ${currentWeightLabel} ${formatNumber(record.weight, decimalPlaces)}. Choose whether to update it or keep the current value.`,
      actionLabel: `Keep current ${currentWeightLabel}`,
    });
  }
  async function saveCurrentWeight() {
    if (!database || !session || !selected || !window.GdmWeighingStore || saving) return;
    const value = !weightEditedRef.current && exactPendingWeightRef.current !== null ? exactPendingWeightRef.current : parseWeight(weightValue);
    if (!Number.isFinite(value) || value < 0) { toast.error("Enter or wait for a valid non-negative scale weight."); weightRef.current?.focus(); return; }
    setSaving(true);
    try {
      const saved = await window.GdmWeighingStore.saveWeight(
        database, session.id, selected, value, scaleConnected ? "serial" : "manual", undefined,
        { lotSite, lotLocation, storage: lotStorage, keepForNext: keepLotContext },
      );
      setWeights((current) => [saved, ...current.filter((item) => normalize(item.uuid) !== normalize(saved.uuid))]);
      setSession((current) => current ? { ...current, updatedAt: new Date().toISOString() } : current);
      if (keepLotContext) {
        setCarriedLotSite(saved.lotSite || ""); setCarriedLotLocation(saved.lotLocation || ""); setCarriedStorage(saved.storage || ""); setHasLotContext(true);
      } else {
        setCarriedLotSite(""); setCarriedLotLocation(""); setCarriedStorage(""); setHasLotContext(false);
      }
      toast.success(`${currentWeightLabel} ${formatNumber(value, decimalPlaces)} saved for plot ${selected.obsName}.`);
      exactPendingWeightRef.current = null; weightEditedRef.current = false;
      setSelected(null); setScanValue(""); setWeightValue(""); setLotSite(""); setLotLocation(""); setLotStorage(""); window.setTimeout(() => scanRef.current?.focus(), 0);
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not save the weight."); }
    finally { setSaving(false); }
  }

  function exportSession(format: "xlsx" | "csv", filteredOnly = false) {
    if (!session || !window.GdmWeighingUtils) { toast.error("Load a weighing session before exporting."); return; }
    const exportPlots = filteredOnly ? filteredPlots : plots;
    if (!exportPlots.length) { toast.error("No plot records match the current filters."); return; }
    try { const count = window.GdmWeighingUtils.exportSession(session, weights, format, exportPlots, decimalPlaces); toast.success(`${count} ${filteredOnly ? "filtered " : ""}plot records exported to ${format === "xlsx" ? "Excel" : "CSV"} with ${decimalPlaces} decimal place${decimalPlaces === 1 ? "" : "s"}.`); }
    catch (error) { toast.error(error instanceof Error ? error.message : "Could not export the session."); }
  }

  function exportLots(format: "xlsx" | "csv") {
    if (!session || !window.GdmWeighingUtils) { toast.error("Load a weighing session before exporting lots."); return; }
    try {
      const count = window.GdmWeighingUtils.exportLots(session, weights, format, filteredPlots, decimalPlaces);
      toast.success(`${count} filtered lot record${count === 1 ? "" : "s"} exported to ${format === "xlsx" ? "Excel" : "CSV"}.`);
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not export lots."); }
  }

  async function changeKeepLotContext(keep: boolean) {
    setKeepLotContext(keep);
    if (!keep) { setCarriedLotSite(""); setCarriedLotLocation(""); setCarriedStorage(""); setHasLotContext(false); }
    if (!database || !session || !window.GdmWeighingStore) return;
    try {
      await window.GdmWeighingStore.setLotContext(database, session.id, {
        keepForNext: keep, hasValue: keep ? hasLotContext : false,
        lotSite: keep ? carriedLotSite : "", lotLocation: keep ? carriedLotLocation : "", storage: keep ? carriedStorage : "",
      });
    } catch (error) { toast.error(error instanceof Error ? error.message : "Could not save the lot field preference."); }
  }

  const applyScaleWeight = useCallback((raw: number) => {
    const value = raw / (10 ** exponentRef.current);
    setRawScaleWeight(raw); setScaleWeight(value);
    if (selectedRef.current) { exactPendingWeightRef.current = value; weightEditedRef.current = false; setWeightValue(formatInputNumber(value, decimalPlacesRef.current)); }
  }, []);
  const consumeSerial = useCallback((chunk: string) => {
    serialBufferRef.current += chunk;
    const lines = serialBufferRef.current.split(/\r\n|\n|\r/); serialBufferRef.current = lines.pop() ?? "";
    for (const line of lines) { const value = parseScaleWeight(line); if (value !== null) applyScaleWeight(value); }
    if (serialTimerRef.current) clearTimeout(serialTimerRef.current);
    serialTimerRef.current = setTimeout(() => { const value = parseScaleWeight(serialBufferRef.current); if (value !== null) applyScaleWeight(value); serialBufferRef.current = ""; }, 180);
  }, [applyScaleWeight]);
  const disconnectScale = useCallback(async (quiet = false) => {
    keepReadingRef.current = false;
    if (serialTimerRef.current) clearTimeout(serialTimerRef.current);
    try { await serialReaderRef.current?.cancel(); } catch { /* already closed */ }
    try { await serialTaskRef.current; } catch { /* surfaced by reader */ }
    try { await serialPortRef.current?.close(); } catch { /* device disconnected */ }
    serialReaderRef.current = null; serialTaskRef.current = null; serialPortRef.current = null; serialBufferRef.current = "";
    setScaleConnected(false); setScaleStatus("Not connected"); setScaleWeight(null); setRawScaleWeight(null);
    if (!quiet) toast.success("Scale disconnected.");
  }, []);
  const toggleScale = useCallback(async () => {
    if (serialPortRef.current) { await disconnectScale(); return; }
    const serial = getSerial();
    if (!serial) { toast.error("Use Google Chrome or Microsoft Edge to connect through a COM port."); return; }
    try {
      const port = await serial.requestPort(); await port.open({ baudRate: Number(baudRate) });
      serialPortRef.current = port; keepReadingRef.current = true; setScaleConnected(true);
      const info = port.getInfo?.() ?? {};
      const identifiers = [info.usbVendorId ? `VID ${info.usbVendorId.toString(16).toUpperCase().padStart(4, "0")}` : "", info.usbProductId ? `PID ${info.usbProductId.toString(16).toUpperCase().padStart(4, "0")}` : ""].filter(Boolean).join(" · ");
      setScaleStatus(identifiers ? `Connected · ${identifiers}` : "Connected to selected port");
      serialTaskRef.current = (async () => {
        const decoder = new TextDecoder();
        try {
          while (keepReadingRef.current && port.readable) {
            const reader = port.readable.getReader(); serialReaderRef.current = reader;
            try { while (keepReadingRef.current) { const { value, done } = await reader.read(); if (done) break; if (value) consumeSerial(decoder.decode(value, { stream: true })); } }
            finally { reader.releaseLock(); serialReaderRef.current = null; }
          }
        } catch (error) { if (keepReadingRef.current) toast.error(error instanceof Error ? error.message : "Scale reading stopped."); }
      })();
      toast.success("Scale connected. Waiting for a weight reading.");
    } catch (error) {
      if (!(error instanceof DOMException && error.name === "NotFoundError")) toast.error(error instanceof Error ? error.message : "Could not connect to the scale.");
      await disconnectScale(true);
    }
  }, [baudRate, consumeSerial, disconnectScale]);
  useEffect(() => {
    const serial = getSerial(); if (!serial) return;
    const listener = (event: Event) => { if (event.target === serialPortRef.current as unknown as EventTarget) { void disconnectScale(true); toast.error("The scale was disconnected from the computer."); } };
    serial.addEventListener("disconnect", listener);
    return () => serial.removeEventListener("disconnect", listener);
  }, [disconnectScale]);
  function changeExponent(value: string) {
    exponentRef.current = Number(value); setScaleExponent(value); localStorage.setItem(SCALE_EXPONENT_KEY, value);
    if (rawScaleWeight !== null) applyScaleWeight(rawScaleWeight);
    toast.success(value === "0" ? "Raw scale value selected." : `Scale factor applied: reading ${scaleFactorLabel(Number(value))}.`);
  }
  function changeDecimalPlaces(value: string) {
    const next = Number(value);
    decimalPlacesRef.current = next; setDecimalPlaces(next); localStorage.setItem(DECIMAL_PLACES_KEY, String(next));
    if (selected && !weightEditedRef.current && exactPendingWeightRef.current !== null) setWeightValue(formatInputNumber(exactPendingWeightRef.current, next));
    toast.success(`Display and exports set to ${next} decimal place${next === 1 ? "" : "s"}.`);
  }
  function changeSort(key: string) {
    setSort((current) => current.key === key ? { key, direction: current.direction === "asc" ? "desc" : "asc" } : { key, direction: "asc" });
    setPage(1);
  }
  const paginationSummary = filteredPlots.length ? `${(safePage - 1) * PAGE_SIZE + 1}–${Math.min(safePage * PAGE_SIZE, filteredPlots.length)} of ${filteredPlots.length}` : "0–0 of 0";
  const tableColumns = [
    ["Status", "status"], ["Season year", "seasonYear"], ["Entity name", "entityName"], ["(OBS) Name", "obsName"], ["FEID", "feid"], ["UUID", "uuid"], ["Block", "block"],
    ["Entry code", "entryCode"], ["Row", "row"], ["Column", "column"], ["(GER) Name", "gerName"], [currentWeightLabel, "weight"], ["Weighed at", "weighedAt"],
  ];

  return (
    <main className="min-h-screen bg-[#edf3f8] pb-10 text-[#17365a]">
      <Toaster position="top-center" richColors />
      {scanAlert && <div className="fixed inset-0 z-[100] grid place-items-center bg-[#102840]/65 p-4" aria-hidden="false">
        <section role="alertdialog" aria-modal="true" aria-labelledby="scan-alert-title" aria-describedby="scan-alert-message" className={`w-full max-w-lg overflow-hidden rounded-xl border-2 bg-white shadow-[0_24px_70px_rgba(10,31,53,.4)] ${scanAlert.kind === "existing" ? "border-[#d18b16]" : "border-[#b42318]"}`}>
          <div className={`grid justify-items-center px-6 py-7 text-center ${scanAlert.kind === "existing" ? "bg-[#fff4d6] text-[#754c00]" : "bg-[#fff0ee] text-[#962b20]"}`}>
            <CircleAlert className="size-16" strokeWidth={2.4} />
            <h2 id="scan-alert-title" className="mt-3 text-3xl font-black">{scanAlert.title}</h2>
          </div>
          <div className="px-6 py-6 text-center">
            <p id="scan-alert-message" className="text-lg font-semibold leading-relaxed text-[#284966]">{scanAlert.message}</p>
            <div className="mt-6 flex flex-col-reverse justify-center gap-3 sm:flex-row">
              {scanAlert.kind === "existing" && <Button ref={scanAlertUpdateButtonRef} type="button" variant="outline" onClick={updateExistingWeight} className="h-13 min-w-44 rounded-md border-2 border-[#c98212] px-6 text-base font-black text-[#8a5708] hover:bg-[#fff4d6]">Update {currentWeightLabel}</Button>}
              <Button ref={scanAlertButtonRef} type="button" onClick={scanAlert.kind === "existing" ? keepExistingWeight : () => closeScanAlert()} className={`h-13 min-w-44 rounded-md px-6 text-base font-black text-white ${scanAlert.kind === "existing" ? "bg-[#c98212] hover:bg-[#ac6d0d]" : "bg-[#b42318] hover:bg-[#921f16]"}`}>{scanAlert.actionLabel}</Button>
            </div>
            <p className="mt-3 text-xs font-bold uppercase tracking-[.08em] text-[#728398]">{scanAlert.kind === "existing" ? `Enter keeps the current ${currentWeightLabel} · choose Update ${currentWeightLabel} to replace it` : "Press Enter or click the button to continue"}</p>
          </div>
        </section>
      </div>}
      <header className="mx-3 mt-2 rounded-lg bg-[#1f4269] text-white shadow-[0_14px_28px_rgba(24,55,88,0.17)]">
        <div className="mx-auto flex max-w-[1500px] items-center justify-between gap-6 px-5 py-[18px] sm:px-9">
          <div className="flex items-center gap-3"><img src="/gdm-logo.svg" alt="GDM" className="h-12 w-[68px] object-contain" /><div><p className="text-[22px] font-black tracking-[-0.02em]">Trial Weighing</p><p className="text-sm text-white/75">GDM Field Operations · Wheat</p></div></div>
          <div className="hidden items-center gap-4 sm:flex"><div className="text-right"><p className="text-xs font-bold uppercase tracking-[.12em] text-white/60">Overall progress</p><p className="text-lg font-bold">{overall.completed} of {overall.total}</p></div><div className="grid size-13 place-items-center rounded-full border-4 border-[#8bb7df] text-sm font-black">{overall.percent}%</div></div>
        </div>
      </header>

      <div className="mx-auto mt-4 flex max-w-[1500px] flex-col justify-between gap-4 px-[18px] lg:flex-row lg:items-center">
        <nav className="flex rounded-lg border border-[#cbdcec] bg-white p-1" aria-label="Main navigation">
          <button type="button" onClick={() => setActiveTab("weighing")} className={`flex h-10 flex-1 items-center justify-center gap-2 rounded-[5px] px-5 font-bold ${activeTab === "weighing" ? "bg-[#1f4269] text-white" : "text-[#58708a]"}`}><Scale className="size-4" /> Weighing</button>
          <button type="button" onClick={() => setActiveTab("dashboard")} className={`flex h-10 flex-1 items-center justify-center gap-2 rounded-[5px] px-5 font-bold ${activeTab === "dashboard" ? "bg-[#1f4269] text-white" : "text-[#58708a]"}`}><LayoutDashboard className="size-4" /> Dashboard</button>
        </nav>
        <div className="flex flex-wrap items-center justify-end gap-2">
          <label className="text-xs font-extrabold uppercase tracking-[.08em] text-[#58708a]">Session</label>
          <NativeSelect value={session?.id || ""} onChange={(event) => void changeSession(event.target.value)} className="h-10 max-w-[310px] rounded-[5px] border-[#bfd1e2] bg-white px-2 font-bold text-[#17365a]">
            {!sessions.length && <NativeSelectOption value="">No active session</NativeSelectOption>}
            {sessions.map((item) => <NativeSelectOption key={item.id} value={item.id}>{item.name}</NativeSelectOption>)}
          </NativeSelect>
          <details className="group relative">
            <summary className="flex h-10 min-w-36 cursor-pointer list-none items-center rounded-[5px] border border-[#bfd1e2] bg-white px-3 text-sm font-bold text-[#315b86]">{!trials.length ? "Trials: None" : selectedTrials.length === trials.length ? "Trials: All" : `Trials: ${selectedTrials.length} of ${trials.length}`}</summary>
            <div className="absolute right-0 z-20 mt-1 w-[min(360px,calc(100vw-36px))] rounded-lg border border-[#bfd1e2] bg-white p-2 shadow-xl">
              <div className="flex gap-2 border-b border-[#dce6ef] pb-2"><Button type="button" variant="outline" className="h-8 text-xs" onClick={() => { setSelectedTrials(trials); setTrialFilter(""); setLocationFilter(""); setPage(1); }}>Select all</Button><Button type="button" variant="outline" className="h-8 text-xs" onClick={() => { setSelectedTrials([]); setSelected(null); setTrialFilter(""); setLocationFilter(""); setPage(1); }}>Clear</Button></div>
              <div className="max-h-64 overflow-auto pt-1">{trials.length ? trials.map((trial) => <label key={trial} className="flex cursor-pointer items-start gap-2 px-1 py-2 text-sm font-semibold"><input type="checkbox" checked={selectedTrialSet.has(trial)} onChange={(event) => { setSelectedTrials((current) => event.target.checked ? [...current, trial] : current.filter((item) => item !== trial)); if (!event.target.checked && (selected?.entityName || "Unnamed trial") === trial) setSelected(null); setTrialFilter(""); setLocationFilter(""); setPage(1); }} className="mt-0.5 accent-[#1f4269]" /><span>{trial}</span></label>) : <p className="p-3 text-sm text-[#657b90]">No trials loaded.</p>}</div>
            </div>
          </details>
          <Button type="button" variant="outline" disabled={!session} onClick={() => void renameActiveSession()} className="h-10 rounded-[5px]"><Pencil className="size-4" /> Rename</Button>
          <Button type="button" variant="outline" disabled={!session} onClick={() => void deleteActiveSession()} className="h-10 rounded-[5px] text-red-700"><Trash2 className="size-4" /> Delete</Button>
          <Button type="button" onClick={() => void startNewSession()} className="h-10 rounded-[5px] bg-[#1f4269] text-white"><Plus className="size-4" /> Start New Weighing</Button>
        </div>
      </div>

      {activeTab === "weighing" ? (
        <section className="mx-auto max-w-[1180px] space-y-5 px-[18px] pt-5">
          <article className="rounded-lg border border-[#cbdcec] bg-white p-5 shadow-[0_10px_28px_rgba(26,59,93,0.08)] sm:p-6">
            <div className="flex flex-wrap items-start justify-between gap-4 border-b-2 border-[#d6e3ef] pb-3">
              <div><p className="mb-1 flex items-center gap-2 text-sm font-bold uppercase tracking-[.1em] text-[#315b86]"><ScanLine className="size-4" /> Weighing operation</p><h1 className="text-2xl font-extrabold text-[#173a61] sm:text-[29px]">Plot reading</h1></div>
              <div className="flex flex-wrap items-center gap-2">
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={(event) => void handleImport(event)} className="hidden" />
                <Button type="button" disabled={importing || loading} onClick={() => fileRef.current?.click()} className="h-9 rounded-[5px] border border-[#547fa6] bg-[#eaf3fb] px-3 font-bold text-[#285882] hover:bg-[#dceaf6]">{importing ? <LoaderCircle className="size-4 animate-spin" /> : <Upload className="size-4" />} {importing ? "Importing…" : "Import Excel / CSV"}</Button>
                <Button type="button" disabled={!session} onClick={() => exportSession("xlsx")} className="h-9 rounded-[5px] bg-[#c88918] px-3 font-bold text-white hover:bg-[#b77710]"><Download className="size-4" /> Export Excel</Button>
                <Button type="button" disabled={!session} onClick={() => exportSession("csv")} className="h-9 rounded-[5px] bg-[#c88918] px-3 font-bold text-white hover:bg-[#b77710]"><Download className="size-4" /> Export CSV</Button>
              </div>
            </div>
            <p className="mt-3 text-xs font-semibold text-[#60768d]">{session ? `Active session: ${session.name} · ${activePlots.length} of ${plots.length} plots in the trial filter · last changed ${formatDateTime(session.updatedAt)}` : "No weighing session loaded. Import a workbook to begin."}</p>

            <section className="mt-3 grid gap-3 rounded-[5px] border border-[#cbdcec] border-l-4 border-l-[#78a9d8] bg-[#eaf2f9] p-3.5 xl:grid-cols-[minmax(150px,.7fr)_minmax(330px,1.6fr)_minmax(110px,.5fr)] xl:items-center">
              <div className="flex items-center gap-3"><div className="grid size-11 place-items-center rounded-[5px] bg-[#d6e6f3] text-[#25537f]"><Cable className="size-5" /></div><div><p className="font-extrabold text-[#173a61]">Serial scale</p><p className="text-xs text-[#60768d]">{scaleStatus}</p></div></div>
              <div className="grid items-end gap-2 sm:grid-cols-2 xl:grid-cols-[minmax(85px,.58fr)_minmax(105px,.72fr)_minmax(105px,.72fr)_minmax(125px,1fr)]">
                <label className="text-[11px] font-extrabold uppercase tracking-[.08em] text-[#587064]">Baud rate<NativeSelect value={baudRate} onChange={(event) => setBaudRate(event.target.value)} disabled={scaleConnected} className="mt-1 h-10 w-full bg-white px-2 text-sm font-bold normal-case tracking-normal">{[1200,2400,4800,9600,19200,38400,57600,115200].map((rate) => <NativeSelectOption key={rate} value={String(rate)}>{rate} baud</NativeSelectOption>)}</NativeSelect></label>
                <label className="text-[11px] font-extrabold uppercase tracking-[.08em] text-[#587064]">Scale factor<NativeSelect value={scaleExponent} onChange={(event) => changeExponent(event.target.value)} className="mt-1 h-10 w-full bg-white px-2 text-sm font-bold normal-case tracking-normal">{[-10,-9,-8,-7,-6,-5,-4,-3,-2,-1,0,1,2,3,4,5,6,7,8,9,10].map((value) => <NativeSelectOption key={value} value={String(value)}>{value < 0 ? `× 10${superscript(Math.abs(value))}` : value > 0 ? `÷ 10${superscript(value)}` : "Raw value"}</NativeSelectOption>)}</NativeSelect></label>
                <label className="text-[11px] font-extrabold uppercase tracking-[.08em] text-[#587064]">Decimal places<NativeSelect value={String(decimalPlaces)} onChange={(event) => changeDecimalPlaces(event.target.value)} className="mt-1 h-10 w-full bg-white px-2 text-sm font-bold normal-case tracking-normal">{[0,1,2,3,4,5,6].map((value) => <NativeSelectOption key={value} value={String(value)}>{value}</NativeSelectOption>)}</NativeSelect></label>
                <Button type="button" onClick={() => void toggleScale()} className={`h-10 rounded-[5px] font-extrabold sm:col-span-2 xl:col-span-1 ${scaleConnected ? "bg-[#d63b38]" : "bg-[#1f4269]"}`}>{scaleConnected ? "Disconnect" : "Connect scale"}</Button>
              </div>
              <div className="border-t border-[#d5dfd0] pt-2 xl:border-l xl:border-t-0 xl:pl-4 xl:pt-0"><p className="text-[10px] font-bold uppercase tracking-[.08em] text-[#75867c]">Current reading</p><p className={`text-3xl font-black ${scaleWeight === null ? "text-[#173a61]" : "text-[#237a63]"}`} title={rawScaleWeight === null ? undefined : `Raw reading: ${formatRawNumber(rawScaleWeight)}`}>{scaleWeight === null ? "—" : formatNumber(scaleWeight, decimalPlaces)}</p><p className="mt-1 text-[10px] font-bold uppercase tracking-[.08em] text-[#75867c]">{scaleConnected ? `${selected ? `${currentWeightLabel} filled automatically` : "Scan a plot to apply"}${scaleExponent === "0" ? "" : ` · ${scaleFactorLabel(Number(scaleExponent))}`}` : "Waiting for connection"}</p></div>
            </section>
            <p className="mt-2 text-xs text-[#647a90]">Connect the scale, choose the weight variable and identifier, then scan a plot to fill {currentWeightLabel} automatically.</p>

            <form onSubmit={handleScan} className="mt-4 grid gap-3 rounded-[5px] border border-[#cbdcec] bg-[#f8fbfe] p-3.5 md:grid-cols-[190px_190px_minmax(0,1fr)]">
              <label><span className="mb-2 block text-sm font-bold text-[#365b80]">Weight variable</span><NativeSelect value={weightVariable} onChange={(event) => void changeWeightVariable(event.target.value as WeightVariable)} disabled={weights.length > 0} title={weights.length ? "Use Start New Weighing to change the variable after weights exist." : undefined} className="h-14 w-full rounded-[5px] border-2 border-[#d1dfed] bg-white px-4 font-bold"><NativeSelectOption value="plotWeight">Plot weight</NativeSelectOption><NativeSelectOption value="seedWeight">Seed weight</NativeSelectOption></NativeSelect></label>
              <label><span className="mb-2 block text-sm font-bold text-[#365b80]">Identifier</span><NativeSelect value={scanMode} onChange={(event) => { scanAlertRef.current = null; setScanMode(event.target.value as "feid" | "uuid"); setSelected(null); setScanAlert(null); setScanError(""); setScanValue(""); scanRef.current?.focus(); }} className="h-14 w-full rounded-[5px] border-2 border-[#d1dfed] bg-white px-4 font-bold"><NativeSelectOption value="feid">Plot FEID</NativeSelectOption><NativeSelectOption value="uuid">Plot UUID</NativeSelectOption></NativeSelect></label>
              <label><span className="mb-2 block text-sm font-bold text-[#365b80]">Scanned code</span><div className="relative"><Barcode className="absolute left-4 top-1/2 size-6 -translate-y-1/2 text-[#6e94b9]" /><Input ref={scanRef} autoFocus value={scanValue} onChange={(event) => setScanValue(event.target.value)} className="h-14 rounded-[5px] border-2 border-[#d1dfed] bg-white pl-13 font-mono text-lg font-semibold" placeholder={scanMode === "feid" ? "Scan or enter the FEID" : "Scan or enter the UUID"} autoComplete="off" spellCheck={false} /></div></label>
            </form>
            <p className="mt-3 text-sm text-[#657b90]"><strong>Quick flow:</strong> scan to load a plot. When {currentWeightLabel} is filled, press Enter or scan the same plot again to save.</p>
            {scanError && <div role="alert" className="mt-4 flex items-center gap-3 rounded-lg border border-[#f2c8be] bg-[#fff4f1] px-4 py-3 text-[#963827]"><CircleAlert className="size-5" /><strong>{scanError}</strong></div>}
          </article>

          {selected && <article className="overflow-hidden rounded-lg border border-[#cbdcec] bg-white shadow-[0_10px_28px_rgba(26,59,93,0.08)]">
            <div className="bg-[#1f4269] p-5 text-white sm:p-6"><div className="mb-5 flex justify-between gap-3"><span className="rounded-full bg-[#d9e9f6] px-3 py-1.5 text-sm font-black uppercase text-[#173f66]">✓ Plot found</span>{existingWeight && <span className="rounded-full bg-[#f5cf77] px-3 py-1.5 text-sm font-bold text-[#6a4700]">Already weighed: {currentWeightLabel} {formatNumber(existingWeight.weight, decimalPlaces)}</span>}</div><div className="grid gap-5 md:grid-cols-[1.45fr_.55fr]"><div><p className="text-xs font-bold uppercase tracking-[.14em] text-[#bbcee1]">Entity name</p><h2 className="text-3xl font-extrabold">{selected.entityName}</h2><p className="mt-5 text-xs font-bold uppercase tracking-[.14em] text-[#bbcee1]">(OBS) Name</p><p className="text-5xl font-black">{selected.obsName}</p></div><div className="rounded-md border border-white/20 bg-white/10 p-4"><p className="text-xs font-bold uppercase tracking-[.12em] text-[#bbcee1]">(GER) Name</p><p className="mt-2 text-xl font-extrabold">{selected.gerName || "—"}</p><p className="mt-4 flex items-center gap-2 text-sm text-[#c7d7e7]"><MapPin className="size-4" /> Location: {selected.location || "Unspecified"}</p><p className="mt-2 text-sm text-[#c7d7e7]">Trial site: {selected.site || "Unspecified"}</p></div></div></div>
            <div className="grid gap-5 p-5 sm:p-6 xl:grid-cols-[minmax(0,1fr)_430px]">
              <dl className="self-start overflow-hidden rounded-md border border-[#cfdeeb] bg-[#f7fafe]"><div className="grid grid-cols-2 sm:grid-cols-5">{[["Season year",selected.seasonYear],["Block",selected.block],["Entry code",selected.entryCode],["Row",selected.row],["Column",selected.column]].map(([label,value]) => <div key={label} className="border-b border-r border-[#d8e4ee] px-3 py-2.5 last:border-r-0 sm:border-b-0"><dt className="text-[10px] font-bold uppercase tracking-[.08em] text-[#6d8195]">{label}</dt><dd className="mt-0.5 text-xl font-black leading-tight">{value || "—"}</dd></div>)}</div><div className="px-3 py-2.5"><dt className="text-[10px] font-bold uppercase tracking-[.08em] text-[#6d8195]">Identifiers</dt><dd className="mt-1 grid gap-x-4 gap-y-1 text-xs sm:grid-cols-[minmax(120px,.45fr)_1fr]"><span><strong>FEID:</strong> {selected.feid}</span><span className="break-all"><strong>UUID:</strong> {selected.uuid}</span></dd></div></dl>
              <form onSubmit={(event) => { event.preventDefault(); void saveCurrentWeight(); }} className="rounded-md border border-[#cbdcec] bg-[#eaf2f9] p-4 sm:p-5">
                <label htmlFor="plot-weight" className="flex items-center gap-2 text-sm font-bold uppercase tracking-[.1em]"><Scale className="size-4" /> {currentWeightLabel}</label>
                <Input id="plot-weight" ref={weightRef} inputMode="decimal" value={weightValue} onChange={(event) => { weightEditedRef.current = true; exactPendingWeightRef.current = null; setWeightValue(event.target.value); }} className="mt-3 h-16 border-2 bg-white px-4 text-3xl font-black" placeholder={formatInputNumber(0, decimalPlaces)} />
                <div className="mt-3 grid gap-2 sm:grid-cols-2">
                  <label className="text-xs font-bold uppercase tracking-[.08em] text-[#526f89]">Lot site<Input value={lotSite} onChange={(event) => setLotSite(event.target.value)} className="mt-1 h-10 bg-white text-sm font-semibold normal-case tracking-normal" placeholder="Lot site" /></label>
                  <label className="text-xs font-bold uppercase tracking-[.08em] text-[#526f89]">Lot location<Input value={lotLocation} onChange={(event) => setLotLocation(event.target.value)} className="mt-1 h-10 bg-white text-sm font-semibold normal-case tracking-normal" placeholder="Lot location" /></label>
                  <label className="text-xs font-bold uppercase tracking-[.08em] text-[#526f89] sm:col-span-2">Storage<Input value={lotStorage} onChange={(event) => setLotStorage(event.target.value)} className="mt-1 h-10 bg-white text-sm font-semibold normal-case tracking-normal" placeholder="Storage" /></label>
                </div>
                <label className="mt-3 flex cursor-pointer items-start gap-2 text-xs font-bold text-[#365b80]"><input type="checkbox" checked={keepLotContext} onChange={(event) => void changeKeepLotContext(event.target.checked)} className="mt-0.5 accent-[#1f4269]" /><span>Keep Lot site, Lot location and Storage for next plot</span></label>
                <Button type="submit" disabled={saving} className="mt-3 h-12 w-full rounded-[5px] bg-[#1f4269] text-base font-black">{saving ? <LoaderCircle className="animate-spin" /> : <Check />} {existingWeight ? `Update ${currentWeightLabel}` : `Save ${currentWeightLabel}`}</Button>
              </form>
            </div>
          </article>}

          <section className="rounded-lg border border-[#cbdcec] bg-white p-5 shadow-[0_10px_28px_rgba(26,59,93,0.08)] sm:p-6"><div className="flex items-center justify-between border-b-2 border-[#d6e3ef] pb-3"><div><p className="text-sm font-bold uppercase tracking-[.1em] text-[#315b86]">◷ Recent history</p><h2 className="text-2xl font-extrabold">Latest weighings</h2></div><span className="rounded bg-[#eaf3fb] px-2.5 py-1 text-xs font-bold">10 most recent</span></div>{recent.length ? <div className="divide-y divide-[#dbe6f0]">{recent.map((record) => <article key={record.key} className="grid items-center gap-3 py-3.5 sm:grid-cols-[1fr_auto_auto]"><div><p className="text-lg font-extrabold">Plot {record.obsName || "—"}</p><p className="text-xs text-[#657b90]">{record.entityName || "Unnamed trial"} · FEID {record.feid || "—"}</p></div><div className="sm:text-right"><small className="font-bold">{currentWeightLabel}</small><p className="text-xl font-black">{formatNumber(record.weight, decimalPlaces)}</p></div><time className="text-xs text-[#657b90]">{formatDateTime(record.weighedAt || record.updatedAt)}</time></article>)}</div> : <div className="grid min-h-48 place-content-center justify-items-center text-center"><Scale className="size-12 rounded-lg bg-[#dceaf6] p-3" /><strong className="mt-3">No weighings recorded</strong><p className="text-sm text-[#647a90]">Saved weighings will appear here automatically.</p></div>}</section>
        </section>
      ) : (
        <section className="mx-auto max-w-[1500px] space-y-5 px-[18px] pt-5">
          <div className="flex items-start justify-between border-b-2 border-[#d6e3ef] pb-3"><div><p className="text-sm font-bold uppercase tracking-[.1em] text-[#315b86]">◴ Real-time report</p><h1 className="text-[29px] font-extrabold">Weighing dashboard</h1></div><span className="rounded-full bg-[#eaf3fb] px-3 py-1.5 text-sm font-bold text-[#315f8b]">● Live local data</span></div>
          <div className="grid gap-3 md:grid-cols-3 xl:grid-cols-[.55fr_.55fr_.55fr_1.4fr]">{[["Total expected",overall.total],["Weighed",overall.completed],["Pending",overall.remaining]].map(([label,value]) => <article key={label} className="rounded-lg bg-[#1f4269] p-5 text-white"><p className="text-xs font-bold uppercase tracking-[.08em] text-white/65">{label}</p><strong className="text-3xl">{value}</strong></article>)}<article className="rounded-lg bg-[#1f4269] p-5 text-white md:col-span-3 xl:col-span-1"><p className="text-xs font-bold uppercase tracking-[.08em] text-white/65">Completion</p><strong className="text-3xl">{overall.percent}%</strong><div className="mt-3 h-2 overflow-hidden rounded-full bg-white/15"><div className="h-full bg-[#8bb7df]" style={{ width: `${overall.percent}%` }} /></div></article></div>
          {[{ title: "Progress by trial", items: trialProgress, icon: "Trials" }, { title: "Progress by location", items: locationProgress, icon: "Locations" }].map((group) => <section key={group.title} className="rounded-lg border border-[#cbdcec] bg-white p-5 shadow-[0_10px_28px_rgba(26,59,93,.08)] sm:p-6"><div className="border-b-2 border-[#d6e3ef] pb-3"><p className="text-sm font-bold uppercase tracking-[.1em] text-[#315b86]">{group.icon}</p><h2 className="text-2xl font-extrabold">{group.title}</h2></div><div className="mt-4 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">{group.items.length ? group.items.map((item) => <DonutCard key={item.key} item={item} />) : <p className="text-sm text-[#657b90]">No data available for this session.</p>}</div></section>)}
          <section className="rounded-lg border border-[#cbdcec] bg-white p-5 shadow-[0_10px_28px_rgba(26,59,93,.08)] sm:p-6"><div className="flex flex-wrap items-end justify-between gap-3 border-b-2 border-[#d6e3ef] pb-3"><div><p className="text-sm font-bold uppercase tracking-[.1em] text-[#315b86]"><Gauge className="mr-1 inline size-4" /> Plot records</p><h2 className="text-2xl font-extrabold">Weighing data</h2></div><div className="flex flex-wrap items-center justify-end gap-2"><strong className="mr-1 text-sm text-[#60768d]">{filteredPlots.length} record{filteredPlots.length === 1 ? "" : "s"}</strong><Button type="button" disabled={!filteredPlots.length} onClick={() => exportSession("xlsx", true)} className="h-9 bg-[#c88918] text-white hover:bg-[#b77710]"><Download className="size-4" /> Export filtered Excel</Button><Button type="button" disabled={!filteredPlots.length} onClick={() => exportSession("csv", true)} className="h-9 bg-[#c88918] text-white hover:bg-[#b77710]"><Download className="size-4" /> Export filtered CSV</Button><Button type="button" disabled={!filteredPlots.some((plot) => weightsByUuid.has(normalize(plot.uuid)))} onClick={() => exportLots("xlsx")} className="h-9 bg-[#315f8b] text-white hover:bg-[#244b70]"><Download className="size-4" /> Export lots Excel</Button><Button type="button" disabled={!filteredPlots.some((plot) => weightsByUuid.has(normalize(plot.uuid)))} onClick={() => exportLots("csv")} className="h-9 bg-[#315f8b] text-white hover:bg-[#244b70]"><Download className="size-4" /> Export lots CSV</Button></div></div>
            <div className="my-4 grid gap-2 md:grid-cols-2 xl:grid-cols-[1.5fr_.7fr_.7fr_.7fr]"><div className="relative"><Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[#6e94b9]" /><Input value={search} onChange={(event) => { setSearch(event.target.value); setPage(1); }} className="h-11 pl-9" placeholder="Search FEID, UUID, OBS or names…" /></div><NativeSelect value={trialFilter} onChange={(event) => { setTrialFilter(event.target.value); setPage(1); }} className="h-11 bg-white px-2"><NativeSelectOption value="">All trials</NativeSelectOption>{trials.filter((value) => selectedTrialSet.has(value)).map((value) => <NativeSelectOption key={value} value={value}>{value}</NativeSelectOption>)}</NativeSelect><NativeSelect value={locationFilter} onChange={(event) => { setLocationFilter(event.target.value); setPage(1); }} className="h-11 bg-white px-2"><NativeSelectOption value="">All locations</NativeSelectOption>{locations.map((value) => <NativeSelectOption key={value} value={value}>{value}</NativeSelectOption>)}</NativeSelect><NativeSelect value={statusFilter} onChange={(event) => { setStatusFilter(event.target.value); setPage(1); }} className="h-11 bg-white px-2"><NativeSelectOption value="">All statuses</NativeSelectOption><NativeSelectOption value="weighed">Weighed</NativeSelectOption><NativeSelectOption value="pending">Pending</NativeSelectOption></NativeSelect></div>
            <div className="mb-3 flex flex-col justify-between gap-3 text-sm text-[#60768d] sm:flex-row sm:items-center"><span>{paginationSummary}</span><div className="flex items-center gap-2"><Button variant="outline" disabled={safePage <= 1} onClick={() => setPage((value) => Math.max(value - 1, 1))}>← Previous</Button><span>Page {safePage} of {pageCount}</span><Button variant="outline" disabled={safePage >= pageCount} onClick={() => setPage((value) => Math.min(value + 1, pageCount))}>Next →</Button></div></div>
            <div className="overflow-auto rounded-lg border border-[#d4e0eb]"><table className="w-full whitespace-nowrap text-left text-xs"><thead className="bg-[#1f4269] text-white"><tr>{tableColumns.map(([label,key]) => <th key={key} className="p-0 text-[11px] uppercase tracking-wide"><button type="button" onClick={() => changeSort(key)} className="w-full p-3 text-left font-bold">{label} <span className={sort.key === key ? "text-white" : "text-[#9fc0df]"}>{sort.key === key ? (sort.direction === "asc" ? "↑" : "↓") : "↕"}</span></button></th>)}</tr></thead><tbody>{pageRows.length ? pageRows.map((plot,index) => { const record = weightsByUuid.get(normalize(plot.uuid)); return <tr key={plot.uuid} className={index % 2 ? "bg-[#f7fafd]" : "bg-white"}><td className="p-2.5"><span className={`rounded-full px-2 py-1 font-bold ${record ? "bg-[#daf0e9] text-[#1d6e59]" : "bg-[#eef1f4] text-[#6b7b88]"}`}>{record ? "Weighed" : "Pending"}</span></td>{[plot.seasonYear,plot.entityName,plot.obsName,plot.feid,plot.uuid,plot.block,plot.entryCode,plot.row,plot.column,plot.gerName || "—"].map((value,i) => <td key={i} className="max-w-64 overflow-hidden text-ellipsis border-b border-[#e0e8ef] p-2.5" title={value}>{value}</td>)}<td className="border-b p-2.5">{record ? formatNumber(record.weight, decimalPlaces) : "—"}</td><td className="border-b p-2.5">{record ? formatDateTime(record.weighedAt || record.updatedAt) : "—"}</td></tr>; }) : <tr><td colSpan={13} className="p-8 text-center text-[#657b90]">No plots match the current filters.</td></tr>}</tbody></table></div>
            <div className="mt-4 flex flex-col justify-between gap-3 text-sm text-[#60768d] sm:flex-row sm:items-center"><span>{paginationSummary}</span><div className="flex items-center gap-2"><Button variant="outline" disabled={safePage <= 1} onClick={() => setPage((value) => Math.max(value - 1, 1))}>← Previous</Button><span>Page {safePage} of {pageCount}</span><Button variant="outline" disabled={safePage >= pageCount} onClick={() => setPage((value) => Math.min(value + 1, pageCount))}>Next →</Button></div></div>
          </section>
        </section>
      )}
    </main>
  );
}
