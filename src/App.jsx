import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Legend
} from "recharts";
import {
  LayoutDashboard, UploadCloud, Users, Building2, Search, Database,
  Trash2, Download, AlertTriangle, CheckCircle2, FileSpreadsheet, X,
  Settings, Cloud, CloudOff, Contact, Link2, UserPlus
} from "lucide-react";

// Local browser storage (used only as a fallback before Supabase is connected).
// This site is standalone, so it uses the browser's own storage rather than
// Claude's artifact storage API.
const storage = {
  async get(key) {
    const v = window.localStorage.getItem(key);
    if (v === null) throw new Error("not found");
    return { key, value: v };
  },
  async set(key, value) {
    window.localStorage.setItem(key, value);
    return { key, value };
  },
  async delete(key) {
    window.localStorage.removeItem(key);
    return { key, deleted: true };
  },
};

const PLAN_TYPES = ["D-SNP", "C-SNP", "I-SNP", "MA-HMO", "MA-PPO", "Med Supp", "PDP", "Other"];
const CHART_COLORS = ["#B8863B", "#12203B", "#2E7D8C", "#1F7A4D", "#6B4C7A", "#B4433D", "#8A95A5"];

const MAPPING_FIELDS = [
  { key: "agent", label: "Agent name", required: true },
  { key: "agentNpn", label: "Agent NPN (National Producer Number)", required: false },
  { key: "agentCarrierId", label: "Agent ID (this carrier's own agent #, e.g. Humana ID)", required: false },
  { key: "product", label: "Plan / product name", required: true },
  { key: "commissionAmount", label: "Commission amount", required: true },
  { key: "clientName", label: "Client name", required: false },
  { key: "saleDate", label: "Sale / write date", required: false },
  { key: "effectiveDate", label: "Effective date", required: false },
  { key: "status", label: "Status (active / termed)", required: false },
  { key: "commissionType", label: "Commission type (initial / renewal)", required: false },
  { key: "paymentDate", label: "Payment date", required: false },
];
function normalizeNameKey(s) {
  return String(s || "").trim().toUpperCase().replace(/\s+/g, " ");
}

const FIELD_TO_COLUMN = {
  carrier: "carrier", agent: "agent", planType: "plan_type", product: "product",
  clientName: "client_name", saleDate: "sale_date", effectiveDate: "effective_date",
  status: "status", commissionAmount: "commission_amount", commissionType: "commission_type",
  paymentDate: "payment_date", uploadBatchId: "upload_batch_id", sourceFile: "source_file", importedAt: "imported_at",
  agentNpn: "agent_npn", agentCarrierId: "carrier_agent_id",
};
function toSnakeRow(rec) {
  const out = {};
  Object.keys(rec).forEach((k) => { if (FIELD_TO_COLUMN[k]) out[FIELD_TO_COLUMN[k]] = rec[k]; });
  ["sale_date", "effective_date", "payment_date"].forEach((col) => { if (out[col] === "") out[col] = null; });
  return out;
}
function toCamelRow(row) {
  return {
    id: row.id, carrier: row.carrier, agent: row.agent, planType: row.plan_type || "Unspecified",
    product: row.product, clientName: row.client_name || "", saleDate: row.sale_date || "",
    effectiveDate: row.effective_date || "", status: row.status || "Active",
    commissionAmount: Number(row.commission_amount) || 0, commissionType: row.commission_type || "",
    paymentDate: row.payment_date || "", uploadBatchId: row.upload_batch_id,
    sourceFile: row.source_file, importedAt: row.imported_at,
    agentNpn: row.agent_npn || "", agentCarrierId: row.carrier_agent_id || "",
  };
}
function toSnakeBatch(b) { return { id: b.id, carrier: b.carrier, file_name: b.fileName, uploaded_at: b.uploadedAt, row_count: b.rowCount }; }
function toCamelBatch(b) { return { id: b.id, carrier: b.carrier, fileName: b.file_name, uploadedAt: b.uploaded_at, rowCount: b.row_count }; }
async function sbFetch(cfg, path, options = {}) {
  const res = await fetch(cfg.url.replace(/\/$/, "") + "/rest/v1/" + path, {
    ...options,
    headers: { apikey: cfg.key, Authorization: "Bearer " + cfg.key, "Content-Type": "application/json", ...(options.headers || {}) },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error("Supabase error " + res.status + ": " + text.slice(0, 200));
  }
  const ct = res.headers.get("content-type") || "";
  return ct.includes("application/json") ? res.json() : null;
}
const SETUP_SQL = `create table policies (
  id uuid primary key default gen_random_uuid(),
  carrier text not null,
  agent text not null,
  plan_type text,
  product text,
  client_name text,
  sale_date date,
  effective_date date,
  status text default 'Active',
  commission_amount numeric not null default 0,
  commission_type text,
  payment_date date,
  upload_batch_id text,
  source_file text,
  imported_at timestamptz default now(),
  agent_npn text,
  carrier_agent_id text
);

create table upload_batches (
  id text primary key,
  carrier text,
  file_name text,
  uploaded_at timestamptz default now(),
  row_count int
);

create table carrier_mappings (
  carrier text primary key,
  mapping jsonb,
  plan_type_mode text,
  plan_type_column text,
  plan_type_fixed text
);

create table agent_directory (
  id uuid primary key default gen_random_uuid(),
  canonical_name text not null,
  npn text,
  created_at timestamptz default now()
);

create table agent_aliases (
  id uuid primary key default gen_random_uuid(),
  agent_id uuid references agent_directory(id) on delete cascade,
  alias_type text not null,
  carrier text,
  alias_value text not null,
  created_at timestamptz default now()
);

alter table policies enable row level security;
alter table upload_batches enable row level security;
alter table carrier_mappings enable row level security;
alter table agent_directory enable row level security;
alter table agent_aliases enable row level security;

create policy "allow all - solo use" on policies for all using (true) with check (true);
create policy "allow all - solo use" on upload_batches for all using (true) with check (true);
create policy "allow all - solo use" on carrier_mappings for all using (true) with check (true);
create policy "allow all - solo use" on agent_directory for all using (true) with check (true);
create policy "allow all - solo use" on agent_aliases for all using (true) with check (true);`;

const MIGRATION_SQL = `alter table policies add column if not exists agent_npn text;
alter table policies add column if not exists carrier_agent_id text;

create table if not exists agent_directory (
  id uuid primary key default gen_random_uuid(),
  canonical_name text not null,
  npn text,
  created_at timestamptz default now()
);

create table if not exists agent_aliases (
  id uuid primary key default gen_random_uuid(),
  agent_id uuid references agent_directory(id) on delete cascade,
  alias_type text not null,
  carrier text,
  alias_value text not null,
  created_at timestamptz default now()
);

alter table agent_directory enable row level security;
alter table agent_aliases enable row level security;

drop policy if exists "allow all - solo use" on agent_directory;
drop policy if exists "allow all - solo use" on agent_aliases;
create policy "allow all - solo use" on agent_directory for all using (true) with check (true);
create policy "allow all - solo use" on agent_aliases for all using (true) with check (true);`;

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}
function parseDateValue(v) {
  if (v === undefined || v === null || v === "") return "";
  if (typeof v === "number") {
    const d = excelSerialToDate(v);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  const d = new Date(v);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return "";
}
function parseMoney(raw) {
  if (raw === undefined || raw === null || raw === "") return 0;
  let s = String(raw).trim();
  let negative = false;
  if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
  s = s.replace(/[$,]/g, "");
  let n = parseFloat(s);
  if (isNaN(n)) return 0;
  return negative ? -Math.abs(n) : n;
}
function normalizePlanType(raw) {
  if (!raw) return "Unspecified";
  const s = String(raw).toLowerCase();
  if (s.includes("d-snp") || s.includes("dsnp")) return "D-SNP";
  if (s.includes("c-snp") || s.includes("csnp")) return "C-SNP";
  if (s.includes("i-snp") || s.includes("isnp")) return "I-SNP";
  if (s.includes("ppo")) return "MA-PPO";
  if (s.includes("hmo")) return "MA-HMO";
  if (s.includes("supp")) return "Med Supp";
  if (s.includes("pdp") || s.includes("part d") || s.includes(" rx")) return "PDP";
  return String(raw).trim() || "Unspecified";
}
function fmtMoney(n) {
  const v = Number(n) || 0;
  const sign = v < 0 ? "-" : "";
  return sign + "$" + Math.abs(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtMoneyShort(n) {
  const v = Number(n) || 0;
  const sign = v < 0 ? "-" : "";
  return sign + "$" + Math.abs(v).toLocaleString("en-US", { maximumFractionDigits: 0 });
}
function fmtDate(iso) {
  if (!iso) return "\u2014";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}
function groupBy(arr, keyFn) {
  const map = new Map();
  arr.forEach((item) => {
    const k = keyFn(item) || "Unspecified";
    if (!map.has(k)) map.set(k, { key: k, revenue: 0, count: 0 });
    const g = map.get(k);
    g.revenue += item.commissionAmount;
    g.count += 1;
  });
  return Array.from(map.values());
}
function exportCSV(rows, filename) {
  const headers = ["Carrier", "Agent", "Plan Type", "Product", "Client Name", "Sale Date", "Effective Date", "Status", "Commission Amount", "Commission Type", "Payment Date", "Source File"];
  const csvRows = rows.map((r) =>
    [r.carrier, r.agent, r.planType, r.product, r.clientName, r.saleDate, r.effectiveDate, r.status, r.commissionAmount, r.commissionType, r.paymentDate, r.sourceFile]
      .map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")
  );
  const csv = [headers.join(","), ...csvRows].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export default function App() {
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState("dashboard");
  const [records, setRecords] = useState([]);
  const [batches, setBatches] = useState([]);
  const [carrierMappings, setCarrierMappings] = useState({});
  const [toast, setToast] = useState(null);
  const [cloudCfg, setCloudCfg] = useState(null);
  const [cloudTestUrl, setCloudTestUrl] = useState("");
  const [cloudTestKey, setCloudTestKey] = useState("");
  const [cloudStatus, setCloudStatus] = useState("idle");
  const [cloudError, setCloudError] = useState("");
  const [agentDirectory, setAgentDirectory] = useState([]);
  const [agentAliases, setAgentAliases] = useState([]);
  const [directoryAvailable, setDirectoryAvailable] = useState(false);

  async function loadDirectory(cfg) {
    try {
      const dir = await sbFetch(cfg, "agent_directory?select=*&order=canonical_name.asc");
      const al = await sbFetch(cfg, "agent_aliases?select=*");
      setAgentDirectory((dir || []).map((r) => ({ id: r.id, canonicalName: r.canonical_name, npn: r.npn || "" })));
      setAgentAliases((al || []).map((r) => ({ id: r.id, agentId: r.agent_id, aliasType: r.alias_type, carrier: r.carrier || "", aliasValue: r.alias_value })));
      setDirectoryAvailable(true);
    } catch (e) {
      setDirectoryAvailable(false);
    }
  }

  async function loadLocal() {
    try { const r = await storage.get("records", false); setRecords(r ? JSON.parse(r.value) : []); } catch (e) { setRecords([]); }
    try { const b = await storage.get("upload-batches", false); setBatches(b ? JSON.parse(b.value) : []); } catch (e) { setBatches([]); }
    try { const m = await storage.get("carrier-mappings", false); setCarrierMappings(m ? JSON.parse(m.value) : {}); } catch (e) { setCarrierMappings({}); }
  }
  async function loadFromCloud(cfg) {
    const pol = await sbFetch(cfg, "policies?select=*&order=imported_at.desc");
    setRecords((pol || []).map(toCamelRow));
    const bat = await sbFetch(cfg, "upload_batches?select=*&order=uploaded_at.desc");
    setBatches((bat || []).map(toCamelBatch));
    const maps = await sbFetch(cfg, "carrier_mappings?select=*");
    const mObj = {};
    (maps || []).forEach((row) => { mObj[row.carrier] = { mapping: row.mapping, planTypeMode: row.plan_type_mode, planTypeColumn: row.plan_type_column, planTypeFixed: row.plan_type_fixed }; });
    setCarrierMappings(mObj);
    await loadDirectory(cfg);
  }

  useEffect(() => {
    (async () => {
      let cfg = null;
      try {
        const saved = await storage.get("supabase-config", false);
        if (saved) cfg = JSON.parse(saved.value);
      } catch (e) {}
      if (cfg && cfg.url && cfg.key) {
        setCloudTestUrl(cfg.url); setCloudTestKey(cfg.key);
        try {
          await sbFetch(cfg, "policies?select=id&limit=1");
          setCloudCfg(cfg); setCloudStatus("connected");
          await loadFromCloud(cfg);
          setLoading(false);
          return;
        } catch (e) {
          setCloudStatus("error"); setCloudError("Saved connection failed: " + e.message);
        }
      }
      await loadLocal();
      setLoading(false);
    })();
  }, []);

  function showToast(msg, type = "success") {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }

  async function connectCloud() {
    setCloudStatus("connecting"); setCloudError("");
    const cfg = { url: cloudTestUrl.trim(), key: cloudTestKey.trim() };
    try {
      await sbFetch(cfg, "policies?select=id&limit=1");
      setCloudCfg(cfg); setCloudStatus("connected");
      try { await storage.set("supabase-config", JSON.stringify(cfg), false); } catch (e) {}
      await loadFromCloud(cfg);
      showToast("Connected to your database.");
    } catch (e) {
      setCloudStatus("error");
      setCloudError("Could not connect. Check your URL, key, and that you ran the setup SQL. (" + e.message + ")");
    }
  }
  async function disconnectCloud() {
    setCloudCfg(null); setCloudStatus("idle");
    try { await storage.delete("supabase-config", false); } catch (e) {}
    await loadLocal();
    showToast("Switched to local (browser-only) storage.");
  }

  const carriersList = useMemo(() => [...new Set(records.map((r) => r.carrier))].sort(), [records]);
  const agentsListAll = useMemo(() => [...new Set(records.map((r) => r.agent))].sort(), [records]);
  const planTypesListAll = useMemo(() => [...new Set(records.map((r) => r.planType))].sort(), [records]);
  const totalAllRevenue = useMemo(() => records.reduce((s, r) => s + r.commissionAmount, 0), [records]);

  // ---------- IMPORT STATE ----------
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [carrierInput, setCarrierInput] = useState("");
  const [mapping, setMapping] = useState({ agent: "", agentNpn: "", agentCarrierId: "", product: "", clientName: "", saleDate: "", effectiveDate: "", status: "", commissionAmount: "", commissionType: "", paymentDate: "" });
  const [planTypeMode, setPlanTypeMode] = useState("column");
  const [planTypeColumn, setPlanTypeColumn] = useState("");
  const [planTypeFixed, setPlanTypeFixed] = useState("D-SNP");
  const [importError, setImportError] = useState("");

  function resetImportStaging() {
    setFileName(""); setHeaders([]); setRawRows([]); setCarrierInput("");
    setMapping({ agent: "", agentNpn: "", agentCarrierId: "", product: "", clientName: "", saleDate: "", effectiveDate: "", status: "", commissionAmount: "", commissionType: "", paymentDate: "" });
    setPlanTypeMode("column"); setPlanTypeColumn(""); setPlanTypeFixed("D-SNP"); setImportError("");
  }

  function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setImportError("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!json.length) { setImportError("That file doesn't have any data rows we could read."); return; }
        const hdrs = Object.keys(json[0]);
        setHeaders(hdrs);
        setRawRows(json);
        setFileName(file.name);
      } catch (err) {
        setImportError("Couldn't read that file. Make sure it's a .xlsx, .xls, or .csv export.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function applyCarrierPreset(name) {
    setCarrierInput(name);
    const preset = carrierMappings[name];
    if (preset) {
      setMapping(preset.mapping || mapping);
      setPlanTypeMode(preset.planTypeMode || "column");
      setPlanTypeColumn(preset.planTypeColumn || "");
      setPlanTypeFixed(preset.planTypeFixed || "D-SNP");
    }
  }

  const mappingValid = carrierInput.trim() && mapping.agent && mapping.product && mapping.commissionAmount;

  const agentLookupMaps = useMemo(() => {
    const npnMap = {}, carrierIdMap = {}, nameTextMap = {}, agentById = {};
    agentDirectory.forEach((a) => { agentById[a.id] = a.canonicalName; });
    agentAliases.forEach((a) => {
      if (a.aliasType === "npn") npnMap[a.aliasValue] = a.agentId;
      else if (a.aliasType === "carrier_id") carrierIdMap[(a.carrier || "") + "::" + a.aliasValue] = a.agentId;
      else if (a.aliasType === "name_text") nameTextMap[a.aliasValue] = a.agentId;
    });
    agentDirectory.forEach((a) => { nameTextMap[normalizeNameKey(a.canonicalName)] = a.id; });
    return { npnMap, carrierIdMap, nameTextMap, agentById };
  }, [agentDirectory, agentAliases]);

  function resolveAgentName(rawName, npn, carrierId, carrier) {
    const { npnMap, carrierIdMap, nameTextMap, agentById } = agentLookupMaps;
    if (npn && npnMap[npn]) return agentById[npnMap[npn]] || rawName;
    if (carrierId && carrierIdMap[(carrier || "") + "::" + carrierId]) return agentById[carrierIdMap[(carrier || "") + "::" + carrierId]] || rawName;
    const key = normalizeNameKey(rawName);
    if (nameTextMap[key]) return agentById[nameTextMap[key]] || rawName;
    return rawName;
  }

  async function commitImport() {
    if (!mappingValid) return;
    const batchId = "b_" + Date.now();
    const carrier = carrierInput.trim();
    const newRecords = rawRows.map((r, i) => {
      const rawAgent = String(r[mapping.agent] ?? "").trim() || "Unassigned";
      const agentNpn = mapping.agentNpn ? String(r[mapping.agentNpn] ?? "").trim() : "";
      const agentCarrierId = mapping.agentCarrierId ? String(r[mapping.agentCarrierId] ?? "").trim() : "";
      return {
        id: batchId + "_" + i,
        carrier,
        agent: resolveAgentName(rawAgent, agentNpn, agentCarrierId, carrier),
        agentNpn,
        agentCarrierId,
        planType: planTypeMode === "fixed" ? (planTypeFixed || "Unspecified") : normalizePlanType(r[planTypeColumn]),
        product: String(r[mapping.product] ?? "").trim() || "Unspecified",
        clientName: mapping.clientName ? String(r[mapping.clientName] ?? "").trim() : "",
        saleDate: mapping.saleDate ? parseDateValue(r[mapping.saleDate]) : "",
        effectiveDate: mapping.effectiveDate ? parseDateValue(r[mapping.effectiveDate]) : "",
        status: mapping.status ? (String(r[mapping.status] ?? "").trim() || "Active") : "Active",
        commissionAmount: parseMoney(r[mapping.commissionAmount]),
        commissionType: mapping.commissionType ? String(r[mapping.commissionType] ?? "").trim() : "",
        paymentDate: mapping.paymentDate ? parseDateValue(r[mapping.paymentDate]) : "",
        uploadBatchId: batchId,
        sourceFile: fileName,
        importedAt: new Date().toISOString(),
      };
    });
    const batchEntry = { id: batchId, carrier, fileName, uploadedAt: new Date().toISOString(), rowCount: newRecords.length };
    const mappingPreset = { mapping, planTypeMode, planTypeColumn, planTypeFixed };

    if (cloudCfg) {
      try {
        const inserted = await sbFetch(cloudCfg, "policies", { method: "POST", headers: { Prefer: "return=representation" }, body: JSON.stringify(newRecords.map(toSnakeRow)) });
        setRecords((prev) => [...prev, ...(inserted || []).map(toCamelRow)]);
        await sbFetch(cloudCfg, "upload_batches", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify([toSnakeBatch(batchEntry)]) });
        setBatches((prev) => [...prev, batchEntry]);
        await sbFetch(cloudCfg, "carrier_mappings", { method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify([{ carrier, mapping: mappingPreset.mapping, plan_type_mode: mappingPreset.planTypeMode, plan_type_column: mappingPreset.planTypeColumn, plan_type_fixed: mappingPreset.planTypeFixed }]) });
        setCarrierMappings((prev) => ({ ...prev, [carrier]: mappingPreset }));
        showToast(`Imported ${newRecords.length} rows from ${carrier} to your database.`);
      } catch (e) {
        showToast("Import failed to save: " + e.message, "error");
        return;
      }
    } else {
      const nextRecords = [...records, ...newRecords];
      const nextBatches = [...batches, batchEntry];
      const nextMappings = { ...carrierMappings, [carrier]: mappingPreset };
      setRecords(nextRecords); setBatches(nextBatches); setCarrierMappings(nextMappings);
      try {
        await storage.set("records", JSON.stringify(nextRecords), false);
        await storage.set("upload-batches", JSON.stringify(nextBatches), false);
        await storage.set("carrier-mappings", JSON.stringify(nextMappings), false);
      } catch (e) { showToast("Could not save locally.", "error"); }
      showToast(`Imported ${newRecords.length} rows from ${carrier}.`);
    }
    resetImportStaging();
    setView("dashboard");
  }

  // ---------- MANAGE DATA ----------
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [confirmClearAll, setConfirmClearAll] = useState(false);

  async function deleteBatch(batchId) {
    if (cloudCfg) {
      try {
        await sbFetch(cloudCfg, `policies?upload_batch_id=eq.${encodeURIComponent(batchId)}`, { method: "DELETE", headers: { Prefer: "return=minimal" } });
        await sbFetch(cloudCfg, `upload_batches?id=eq.${encodeURIComponent(batchId)}`, { method: "DELETE", headers: { Prefer: "return=minimal" } });
        setRecords((prev) => prev.filter((r) => r.uploadBatchId !== batchId));
        setBatches((prev) => prev.filter((b) => b.id !== batchId));
        showToast("Import removed.");
      } catch (e) { showToast("Could not delete: " + e.message, "error"); }
    } else {
      const nextRecords = records.filter((r) => r.uploadBatchId !== batchId);
      const nextBatches = batches.filter((b) => b.id !== batchId);
      setRecords(nextRecords); setBatches(nextBatches);
      try { await storage.set("records", JSON.stringify(nextRecords), false); await storage.set("upload-batches", JSON.stringify(nextBatches), false); } catch (e) {}
      showToast("Import removed.");
    }
    setConfirmDeleteId(null);
  }
  async function clearAllData() {
    if (cloudCfg) {
      try {
        await sbFetch(cloudCfg, "policies?id=not.is.null", { method: "DELETE", headers: { Prefer: "return=minimal" } });
        await sbFetch(cloudCfg, "upload_batches?id=not.is.null", { method: "DELETE", headers: { Prefer: "return=minimal" } });
        await sbFetch(cloudCfg, "carrier_mappings?carrier=not.is.null", { method: "DELETE", headers: { Prefer: "return=minimal" } });
        setRecords([]); setBatches([]); setCarrierMappings({});
        showToast("All data cleared from your database.");
      } catch (e) { showToast("Could not clear: " + e.message, "error"); }
    } else {
      setRecords([]); setBatches([]); setCarrierMappings({});
      try {
        await storage.set("records", JSON.stringify([]), false);
        await storage.set("upload-batches", JSON.stringify([]), false);
        await storage.set("carrier-mappings", JSON.stringify({}), false);
      } catch (e) {}
      showToast("All data cleared.");
    }
    setConfirmClearAll(false);
  }

  // ---------- AGENT DIRECTORY ----------
  const [directoryTab, setDirectoryTab] = useState("directory");
  const [selectedDirAgent, setSelectedDirAgent] = useState(null);
  const [newAgentName, setNewAgentName] = useState("");
  const [newAgentNpn, setNewAgentNpn] = useState("");
  const [editAgentName, setEditAgentName] = useState("");
  const [editAgentNpn, setEditAgentNpn] = useState("");
  const [newAliasType, setNewAliasType] = useState("npn");
  const [newAliasCarrier, setNewAliasCarrier] = useState("");
  const [newAliasValue, setNewAliasValue] = useState("");
  const [linkChoice, setLinkChoice] = useState({});

  async function addAgentManual() {
    if (!newAgentName.trim() || !cloudCfg) return;
    try {
      const inserted = await sbFetch(cloudCfg, "agent_directory", { method: "POST", headers: { Prefer: "return=representation" }, body: JSON.stringify([{ canonical_name: newAgentName.trim(), npn: newAgentNpn.trim() || null }]) });
      const agent = inserted[0];
      if (newAgentNpn.trim()) {
        await sbFetch(cloudCfg, "agent_aliases", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify([{ agent_id: agent.id, alias_type: "npn", alias_value: newAgentNpn.trim() }]) });
      }
      await loadDirectory(cloudCfg);
      setNewAgentName(""); setNewAgentNpn("");
      showToast("Agent added to directory.");
    } catch (e) { showToast("Could not add agent: " + e.message, "error"); }
  }
  async function updateAgentDetails(agentId) {
    if (!cloudCfg) return;
    try {
      await sbFetch(cloudCfg, `agent_directory?id=eq.${agentId}`, { method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify({ canonical_name: editAgentName.trim(), npn: editAgentNpn.trim() || null }) });
      await sbFetch(cloudCfg, `agent_aliases?agent_id=eq.${agentId}&alias_type=eq.npn`, { method: "DELETE", headers: { Prefer: "return=minimal" } });
      if (editAgentNpn.trim()) {
        await sbFetch(cloudCfg, "agent_aliases", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify([{ agent_id: agentId, alias_type: "npn", alias_value: editAgentNpn.trim() }]) });
      }
      await loadDirectory(cloudCfg);
      showToast("Agent updated.");
    } catch (e) { showToast("Could not update: " + e.message, "error"); }
  }
  async function deleteDirAgent(agentId) {
    if (!cloudCfg) return;
    try {
      await sbFetch(cloudCfg, `agent_directory?id=eq.${agentId}`, { method: "DELETE", headers: { Prefer: "return=minimal" } });
      setSelectedDirAgent(null);
      await loadDirectory(cloudCfg);
      showToast("Agent removed from directory.");
    } catch (e) { showToast("Could not delete: " + e.message, "error"); }
  }
  async function addAliasToAgent(agentId) {
    if (!newAliasValue.trim() || !cloudCfg) return;
    try {
      await sbFetch(cloudCfg, "agent_aliases", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify([{ agent_id: agentId, alias_type: newAliasType, carrier: newAliasType === "carrier_id" ? newAliasCarrier.trim() : null, alias_value: newAliasType === "name_text" ? normalizeNameKey(newAliasValue) : newAliasValue.trim() }]) });
      await loadDirectory(cloudCfg);
      setNewAliasValue(""); setNewAliasCarrier("");
      showToast("Alias added.");
    } catch (e) { showToast("Could not add alias: " + e.message, "error"); }
  }
  async function deleteAlias(aliasId) {
    if (!cloudCfg) return;
    try {
      await sbFetch(cloudCfg, `agent_aliases?id=eq.${aliasId}`, { method: "DELETE", headers: { Prefer: "return=minimal" } });
      await loadDirectory(cloudCfg);
    } catch (e) { showToast("Could not remove alias: " + e.message, "error"); }
  }
  async function linkRawNameToAgent(rawName, agentId, canonicalName) {
    if (!cloudCfg) return;
    try {
      await sbFetch(cloudCfg, "agent_aliases", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify([{ agent_id: agentId, alias_type: "name_text", alias_value: normalizeNameKey(rawName) }]) });
      await sbFetch(cloudCfg, `policies?agent=eq.${encodeURIComponent(rawName)}`, { method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify({ agent: canonicalName }) });
      setRecords((prev) => prev.map((r) => (r.agent === rawName ? { ...r, agent: canonicalName } : r)));
      await loadDirectory(cloudCfg);
      showToast(`Linked "${rawName}" to ${canonicalName}.`);
    } catch (e) { showToast("Could not link: " + e.message, "error"); }
  }
  async function createAgentFromRawName(rawName) {
    if (!cloudCfg) return;
    try {
      const inserted = await sbFetch(cloudCfg, "agent_directory", { method: "POST", headers: { Prefer: "return=representation" }, body: JSON.stringify([{ canonical_name: rawName, npn: null }]) });
      const agent = inserted[0];
      await sbFetch(cloudCfg, "agent_aliases", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify([{ agent_id: agent.id, alias_type: "name_text", alias_value: normalizeNameKey(rawName) }]) });
      await loadDirectory(cloudCfg);
      showToast(`"${rawName}" added as a new agent.`);
    } catch (e) { showToast("Could not create agent: " + e.message, "error"); }
  }

  const unmatchedAgentNames = useMemo(() => {
    const { nameTextMap } = agentLookupMaps;
    const distinct = [...new Set(records.map((r) => r.agent))];
    return distinct
      .filter((n) => !nameTextMap[normalizeNameKey(n)])
      .map((n) => {
        const recs = records.filter((r) => r.agent === n);
        return { name: n, count: recs.length, revenue: recs.reduce((s, r) => s + r.commissionAmount, 0) };
      })
      .sort((a, b) => b.revenue - a.revenue);
  }, [records, agentLookupMaps]);

  // Roster import (bulk-create agents from a CRM export)
  const [rosterFileName, setRosterFileName] = useState("");
  const [rosterHeaders, setRosterHeaders] = useState([]);
  const [rosterRows, setRosterRows] = useState([]);
  const [rosterNameCol, setRosterNameCol] = useState("");
  const [rosterNpnCol, setRosterNpnCol] = useState("");

  function handleRosterFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!json.length) return;
        setRosterHeaders(Object.keys(json[0]));
        setRosterRows(json);
        setRosterFileName(file.name);
      } catch (err) { showToast("Couldn't read that file.", "error"); }
    };
    reader.readAsArrayBuffer(file);
  }
  async function commitRosterImport() {
    if (!rosterNameCol || !cloudCfg) return;
    try {
      const toInsert = rosterRows
        .map((r) => ({ canonical_name: String(r[rosterNameCol] ?? "").trim(), npn: rosterNpnCol ? String(r[rosterNpnCol] ?? "").trim() || null : null }))
        .filter((r) => r.canonical_name);
      const inserted = await sbFetch(cloudCfg, "agent_directory", { method: "POST", headers: { Prefer: "return=representation" }, body: JSON.stringify(toInsert) });
      const aliasRows = [];
      (inserted || []).forEach((a) => {
        aliasRows.push({ agent_id: a.id, alias_type: "name_text", alias_value: normalizeNameKey(a.canonical_name) });
        if (a.npn) aliasRows.push({ agent_id: a.id, alias_type: "npn", alias_value: a.npn });
      });
      if (aliasRows.length) await sbFetch(cloudCfg, "agent_aliases", { method: "POST", headers: { Prefer: "return=minimal" }, body: JSON.stringify(aliasRows) });
      await loadDirectory(cloudCfg);
      showToast(`Added ${inserted.length} agents to your directory.`);
      setRosterFileName(""); setRosterHeaders([]); setRosterRows([]); setRosterNameCol(""); setRosterNpnCol("");
    } catch (e) { showToast("Roster import failed: " + e.message, "error"); }
  }

  // ---------- DASHBOARD FILTERS ----------
  const [filterCarrier, setFilterCarrier] = useState("All");
  const [filterAgent, setFilterAgent] = useState("All");
  const [filterPlanType, setFilterPlanType] = useState("All");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const filteredRecords = useMemo(() => records.filter((r) =>
    (filterCarrier === "All" || r.carrier === filterCarrier) &&
    (filterAgent === "All" || r.agent === filterAgent) &&
    (filterPlanType === "All" || r.planType === filterPlanType) &&
    (!dateFrom || (r.saleDate && r.saleDate >= dateFrom)) &&
    (!dateTo || (r.saleDate && r.saleDate <= dateTo))
  ), [records, filterCarrier, filterAgent, filterPlanType, dateFrom, dateTo]);

  const netRevenue = useMemo(() => filteredRecords.reduce((s, r) => s + r.commissionAmount, 0), [filteredRecords]);
  const chargebacks = useMemo(() => filteredRecords.filter((r) => r.commissionAmount < 0).reduce((s, r) => s + r.commissionAmount, 0), [filteredRecords]);
  const byCarrier = useMemo(() => groupBy(filteredRecords, (r) => r.carrier).sort((a, b) => b.revenue - a.revenue), [filteredRecords]);
  const byAgent = useMemo(() => groupBy(filteredRecords, (r) => r.agent).sort((a, b) => b.revenue - a.revenue), [filteredRecords]);
  const byPlanType = useMemo(() => groupBy(filteredRecords, (r) => r.planType).sort((a, b) => b.revenue - a.revenue), [filteredRecords]);
  const byMonth = useMemo(() => {
    const grouped = groupBy(filteredRecords.filter((r) => r.saleDate), (r) => r.saleDate.slice(0, 7));
    return grouped.sort((a, b) => a.key.localeCompare(b.key));
  }, [filteredRecords]);
  const activeCarrierCount = new Set(filteredRecords.map((r) => r.carrier)).size;
  const activeAgentCount = new Set(filteredRecords.map((r) => r.agent)).size;

  // ---------- AGENTS / CARRIERS ----------
  const [agentSearch, setAgentSearch] = useState("");
  const [selectedAgent, setSelectedAgent] = useState(null);
  const agentSummary = useMemo(() => groupBy(records, (r) => r.agent).sort((a, b) => b.revenue - a.revenue)
    .filter((a) => a.key.toLowerCase().includes(agentSearch.toLowerCase())), [records, agentSearch]);
  const selectedAgentRecords = useMemo(() => records.filter((r) => r.agent === selectedAgent), [records, selectedAgent]);
  const selectedAgentByCarrier = useMemo(() => groupBy(selectedAgentRecords, (r) => r.carrier).sort((a, b) => b.revenue - a.revenue), [selectedAgentRecords]);
  const selectedAgentByPlanType = useMemo(() => groupBy(selectedAgentRecords, (r) => r.planType).sort((a, b) => b.revenue - a.revenue), [selectedAgentRecords]);

  const [carrierSearch, setCarrierSearch] = useState("");
  const [selectedCarrier, setSelectedCarrier] = useState(null);
  const carrierSummary = useMemo(() => groupBy(records, (r) => r.carrier).sort((a, b) => b.revenue - a.revenue)
    .filter((c) => c.key.toLowerCase().includes(carrierSearch.toLowerCase())), [records, carrierSearch]);
  const selectedCarrierRecords = useMemo(() => records.filter((r) => r.carrier === selectedCarrier), [records, selectedCarrier]);
  const selectedCarrierByAgent = useMemo(() => groupBy(selectedCarrierRecords, (r) => r.agent).sort((a, b) => b.revenue - a.revenue), [selectedCarrierRecords]);
  const selectedCarrierByPlanType = useMemo(() => groupBy(selectedCarrierRecords, (r) => r.planType).sort((a, b) => b.revenue - a.revenue), [selectedCarrierRecords]);

  // ---------- CLIENT LOOKUP ----------
  const [clientSearch, setClientSearch] = useState("");
  const clientMatches = useMemo(() => {
    if (clientSearch.trim().length < 2) return [];
    const q = clientSearch.toLowerCase();
    return records.filter((r) => r.clientName && r.clientName.toLowerCase().includes(q));
  }, [records, clientSearch]);

  if (loading) {
    return <div className="pt-app pt-loading"><style>{CSS}</style>Loading your data\u2026</div>;
  }

  const NAV = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "import", label: "Import statement", icon: UploadCloud },
    { key: "agents", label: "Agents", icon: Users },
    { key: "directory", label: "Agent directory", icon: Contact },
    { key: "carriers", label: "Carriers", icon: Building2 },
    { key: "clients", label: "Client lookup", icon: Search },
    { key: "manage", label: "Manage data", icon: Database },
    { key: "settings", label: "Database connection", icon: Settings },
  ];

  return (
    <div className="pt-app">
      <style>{CSS}</style>
      <aside className="pt-sidebar">
        <div className="pt-brand">
          <div className="pt-brand-mark">L</div>
          <div>
            <div className="pt-brand-name">Ledger</div>
            <div className="pt-brand-sub">Production Tracking</div>
          </div>
        </div>
        <nav className="pt-nav">
          {NAV.map((n) => (
            <button key={n.key} className={"pt-nav-item" + (view === n.key ? " active" : "")} onClick={() => setView(n.key)}>
              <n.icon size={17} strokeWidth={1.75} />
              <span>{n.label}</span>
            </button>
          ))}
        </nav>
        <div className="pt-sidebar-footer">
          <div className={"pt-conn-dot " + (cloudCfg ? "on" : "off")}>
            {cloudCfg ? <Cloud size={12} /> : <CloudOff size={12} />}
            {cloudCfg ? "Cloud connected" : "Local only"}
          </div>
          <div className="pt-footer-label">All-time production</div>
          <div className="pt-footer-value">{fmtMoneyShort(totalAllRevenue)}</div>
          <div className="pt-footer-sub">{records.length.toLocaleString()} policies on file</div>
        </div>
      </aside>

      <main className="pt-main">
        {toast && (
          <div className={"pt-toast " + toast.type}>
            {toast.type === "error" ? <AlertTriangle size={16} /> : <CheckCircle2 size={16} />}
            <span>{toast.msg}</span>
          </div>
        )}

        {view === "dashboard" && (
          <div>
            <div className="pt-page-head">
              <div>
                <h1>Dashboard</h1>
                <p>Production across every carrier statement you've imported.</p>
              </div>
              {filteredRecords.length > 0 && (
                <button className="pt-btn ghost" onClick={() => exportCSV(filteredRecords, "production-export.csv")}>
                  <Download size={15} /> Export view
                </button>
              )}
            </div>

            {records.length === 0 ? (
              <EmptyState onGo={() => setView("import")} />
            ) : (
              <>
                <div className="pt-filters">
                  <FilterSelect label="Carrier" value={filterCarrier} onChange={setFilterCarrier} options={carriersList} />
                  <FilterSelect label="Agent" value={filterAgent} onChange={setFilterAgent} options={agentsListAll} />
                  <FilterSelect label="Plan type" value={filterPlanType} onChange={setFilterPlanType} options={planTypesListAll} />
                  <div className="pt-filter">
                    <label>From</label>
                    <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} />
                  </div>
                  <div className="pt-filter">
                    <label>To</label>
                    <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
                  </div>
                  {(filterCarrier !== "All" || filterAgent !== "All" || filterPlanType !== "All" || dateFrom || dateTo) && (
                    <button className="pt-btn text" onClick={() => { setFilterCarrier("All"); setFilterAgent("All"); setFilterPlanType("All"); setDateFrom(""); setDateTo(""); }}>
                      Clear filters
                    </button>
                  )}
                </div>

                <div className="pt-cards">
                  <StatCard label="Net revenue" value={fmtMoneyShort(netRevenue)} tone={netRevenue < 0 ? "rose" : "ink"} />
                  <StatCard label="Policies" value={filteredRecords.length.toLocaleString()} tone="ink" />
                  <StatCard label="Avg per policy" value={fmtMoneyShort(filteredRecords.length ? netRevenue / filteredRecords.length : 0)} tone="ink" />
                  <StatCard label="Chargebacks" value={fmtMoneyShort(chargebacks)} tone="rose" />
                  <StatCard label="Carriers" value={activeCarrierCount} tone="ink" />
                  <StatCard label="Agents" value={activeAgentCount} tone="ink" />
                </div>

                <div className="pt-grid-2">
                  <div className="pt-card">
                    <h3>Revenue by carrier</h3>
                    <ResponsiveContainer width="100%" height={260}>
                      <BarChart data={byCarrier} margin={{ top: 8, right: 8, left: 0, bottom: 8 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#E2E4E9" vertical={false} />
                        <XAxis dataKey="key" tick={{ fontSize: 11, fill: "#64748B" }} axisLine={{ stroke: "#E2E4E9" }} tickLine={false} />
                        <YAxis tick={{ fontSize: 11, fill: "#64748B" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtMoneyShort(v)} />
                        <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #E2E4E9" }} />
                        <Bar dataKey="revenue" fill="#B8863B" radius={[3, 3, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <div className="pt-card">
                    <h3>Revenue by plan type</h3>
                    <ResponsiveContainer width="100%" height={260}>
                      <PieChart>
                        <Pie data={byPlanType} dataKey="revenue" nameKey="key" innerRadius={55} outerRadius={90} paddingAngle={2}>
                          {byPlanType.map((entry, i) => <Cell key={entry.key} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                        </Pie>
                        <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #E2E4E9" }} />
                        <Legend wrapperStyle={{ fontSize: 12 }} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="pt-card">
                  <h3>Monthly trend</h3>
                  <ResponsiveContainer width="100%" height={220}>
                    <LineChart data={byMonth} margin={{ top: 8, right: 8, left: 0, bottom: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#E2E4E9" vertical={false} />
                      <XAxis dataKey="key" tick={{ fontSize: 11, fill: "#64748B" }} axisLine={{ stroke: "#E2E4E9" }} tickLine={false} />
                      <YAxis tick={{ fontSize: 11, fill: "#64748B" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtMoneyShort(v)} />
                      <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #E2E4E9" }} />
                      <Line type="monotone" dataKey="revenue" stroke="#12203B" strokeWidth={2} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>

                <div className="pt-card">
                  <h3>Top agents in this view</h3>
                  <table className="pt-table">
                    <thead><tr><th>Agent</th><th className="num">Policies</th><th className="num">Revenue</th></tr></thead>
                    <tbody>
                      {byAgent.slice(0, 10).map((a) => (
                        <tr key={a.key}>
                          <td>{a.key}</td>
                          <td className="num">{a.count}</td>
                          <td className="num mono">{fmtMoney(a.revenue)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}

        {view === "import" && (
          <div>
            <div className="pt-page-head">
              <div>
                <h1>Import a carrier statement</h1>
                <p>Upload a spreadsheet, tell us what each column means once, and we'll remember it next time.</p>
              </div>
            </div>

            {!fileName ? (
              <div className="pt-card pt-upload-zone">
                <FileSpreadsheet size={32} strokeWidth={1.3} color="#B8863B" />
                <p className="pt-upload-title">Choose a .xlsx, .xls, or .csv file</p>
                <p className="pt-upload-sub">Any carrier's export format works \u2014 you'll map the columns next.</p>
                <label className="pt-btn primary" style={{ marginTop: 12 }}>
                  Choose file
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
                </label>
                {importError && <p className="pt-error">{importError}</p>}
              </div>
            ) : (
              <>
                <div className="pt-card">
                  <div className="pt-row-between">
                    <div>
                      <div className="pt-file-chip"><FileSpreadsheet size={14} /> {fileName} \u00b7 {rawRows.length} rows</div>
                    </div>
                    <button className="pt-btn text" onClick={resetImportStaging}><X size={14} /> Start over</button>
                  </div>

                  <div className="pt-field" style={{ marginTop: 16 }}>
                    <label>Carrier</label>
                    <input list="carrier-options" value={carrierInput} onChange={(e) => applyCarrierPreset(e.target.value)} placeholder="e.g. Humana, UHC, Aetna\u2026" />
                    <datalist id="carrier-options">
                      {carriersList.map((c) => <option key={c} value={c} />)}
                    </datalist>
                    {carrierMappings[carrierInput.trim()] && <p className="pt-hint">Loaded your saved column mapping for {carrierInput.trim()}. Adjust below if this file is different.</p>}
                  </div>
                </div>

                <div className="pt-card">
                  <h3>Map your columns</h3>
                  <p className="pt-hint" style={{ marginBottom: 12 }}>Match each field to a column from your file. Fields marked * are required.</p>
                  <div className="pt-mapping-grid">
                    {MAPPING_FIELDS.map((f) => (
                      <div className="pt-field" key={f.key}>
                        <label>{f.label}{f.required && " *"}</label>
                        <select value={mapping[f.key]} onChange={(e) => setMapping({ ...mapping, [f.key]: e.target.value })}>
                          <option value="">\u2014 not in file \u2014</option>
                          {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                        </select>
                      </div>
                    ))}
                  </div>

                  <div className="pt-plantype-block">
                    <label>Plan type (D-SNP, MA-HMO, MA-PPO, Med Supp, etc.)</label>
                    <div className="pt-radio-row">
                      <label className="pt-radio"><input type="radio" checked={planTypeMode === "column"} onChange={() => setPlanTypeMode("column")} /> Read from a column</label>
                      <label className="pt-radio"><input type="radio" checked={planTypeMode === "fixed"} onChange={() => setPlanTypeMode("fixed")} /> This whole file is one plan type</label>
                    </div>
                    {planTypeMode === "column" ? (
                      <select value={planTypeColumn} onChange={(e) => setPlanTypeColumn(e.target.value)}>
                        <option value="">\u2014 not in file \u2014</option>
                        {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                      </select>
                    ) : (
                      <input list="plan-type-options" value={planTypeFixed} onChange={(e) => setPlanTypeFixed(e.target.value)} />
                    )}
                    <datalist id="plan-type-options">
                      {PLAN_TYPES.map((p) => <option key={p} value={p} />)}
                    </datalist>
                  </div>
                </div>

                <div className="pt-card">
                  <h3>Preview \u2014 first 5 rows from your file</h3>
                  <div className="pt-preview-scroll">
                    <table className="pt-table">
                      <thead><tr>{headers.map((h) => <th key={h}>{h}</th>)}</tr></thead>
                      <tbody>
                        {rawRows.slice(0, 5).map((r, i) => (
                          <tr key={i}>{headers.map((h) => <td key={h}>{String(r[h])}</td>)}</tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="pt-row-between">
                  <div>{!mappingValid && <p className="pt-error">Set the carrier name and map agent, product, and commission amount to continue.</p>}</div>
                  <button className="pt-btn primary" disabled={!mappingValid} onClick={commitImport}>
                    Import {rawRows.length} rows
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {view === "agents" && (
          <div>
            <div className="pt-page-head"><div><h1>Agents</h1><p>Production by agent, across every carrier.</p></div></div>
            {records.length === 0 ? <EmptyState onGo={() => setView("import")} /> : (
              <div className="pt-grid-list">
                <div className="pt-card">
                  <input className="pt-search" placeholder="Search agents\u2026" value={agentSearch} onChange={(e) => setAgentSearch(e.target.value)} />
                  <table className="pt-table">
                    <thead><tr><th>Agent</th><th className="num">Policies</th><th className="num">Revenue</th></tr></thead>
                    <tbody>
                      {agentSummary.map((a) => (
                        <tr key={a.key} className={"pt-clickable" + (selectedAgent === a.key ? " selected" : "")} onClick={() => setSelectedAgent(a.key)}>
                          <td>{a.key}</td>
                          <td className="num">{a.count}</td>
                          <td className="num mono">{fmtMoney(a.revenue)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {selectedAgent && (
                  <div className="pt-card">
                    <div className="pt-row-between">
                      <h3>{selectedAgent}</h3>
                      <button className="pt-btn ghost" onClick={() => exportCSV(selectedAgentRecords, selectedAgent.replace(/\s+/g, "_") + ".csv")}><Download size={14} /> Export</button>
                    </div>
                    <div className="pt-mini-grid">
                      <div>
                        <div className="pt-mini-label">By carrier</div>
                        {selectedAgentByCarrier.map((c) => <div key={c.key} className="pt-mini-row"><span>{c.key}</span><span className="mono">{fmtMoney(c.revenue)}</span></div>)}
                      </div>
                      <div>
                        <div className="pt-mini-label">By plan type</div>
                        {selectedAgentByPlanType.map((p) => <div key={p.key} className="pt-mini-row"><span>{p.key}</span><span className="mono">{fmtMoney(p.revenue)}</span></div>)}
                      </div>
                    </div>
                    <div className="pt-mini-label" style={{ marginTop: 16 }}>All sales ({selectedAgentRecords.length})</div>
                    <div className="pt-preview-scroll">
                      <table className="pt-table">
                        <thead><tr><th>Carrier</th><th>Product</th><th>Plan type</th><th>Sale date</th><th>Status</th><th className="num">Amount</th></tr></thead>
                        <tbody>
                          {selectedAgentRecords.map((r) => (
                            <tr key={r.id}><td>{r.carrier}</td><td>{r.product}</td><td>{r.planType}</td><td>{fmtDate(r.saleDate)}</td><td>{r.status}</td><td className="num mono">{fmtMoney(r.commissionAmount)}</td></tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {view === "carriers" && (
          <div>
            <div className="pt-page-head"><div><h1>Carriers</h1><p>Production by carrier, across every agent.</p></div></div>
            {records.length === 0 ? <EmptyState onGo={() => setView("import")} /> : (
              <div className="pt-grid-list">
                <div className="pt-card">
                  <input className="pt-search" placeholder="Search carriers\u2026" value={carrierSearch} onChange={(e) => setCarrierSearch(e.target.value)} />
                  <table className="pt-table">
                    <thead><tr><th>Carrier</th><th className="num">Policies</th><th className="num">Revenue</th></tr></thead>
                    <tbody>
                      {carrierSummary.map((c) => (
                        <tr key={c.key} className={"pt-clickable" + (selectedCarrier === c.key ? " selected" : "")} onClick={() => setSelectedCarrier(c.key)}>
                          <td>{c.key}</td>
                          <td className="num">{c.count}</td>
                          <td className="num mono">{fmtMoney(c.revenue)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {selectedCarrier && (
                  <div className="pt-card">
                    <div className="pt-row-between">
                      <h3>{selectedCarrier}</h3>
                      <button className="pt-btn ghost" onClick={() => exportCSV(selectedCarrierRecords, selectedCarrier.replace(/\s+/g, "_") + ".csv")}><Download size={14} /> Export</button>
                    </div>
                    <div className="pt-mini-grid">
                      <div>
                        <div className="pt-mini-label">By agent</div>
                        {selectedCarrierByAgent.map((a) => <div key={a.key} className="pt-mini-row"><span>{a.key}</span><span className="mono">{fmtMoney(a.revenue)}</span></div>)}
                      </div>
                      <div>
                        <div className="pt-mini-label">By plan type</div>
                        {selectedCarrierByPlanType.map((p) => <div key={p.key} className="pt-mini-row"><span>{p.key}</span><span className="mono">{fmtMoney(p.revenue)}</span></div>)}
                      </div>
                    </div>
                    <div className="pt-mini-label" style={{ marginTop: 16 }}>All sales ({selectedCarrierRecords.length})</div>
                    <div className="pt-preview-scroll">
                      <table className="pt-table">
                        <thead><tr><th>Agent</th><th>Product</th><th>Plan type</th><th>Sale date</th><th>Status</th><th className="num">Amount</th></tr></thead>
                        <tbody>
                          {selectedCarrierRecords.map((r) => (
                            <tr key={r.id}><td>{r.agent}</td><td>{r.product}</td><td>{r.planType}</td><td>{fmtDate(r.saleDate)}</td><td>{r.status}</td><td className="num mono">{fmtMoney(r.commissionAmount)}</td></tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {view === "clients" && (
          <div>
            <div className="pt-page-head"><div><h1>Client lookup</h1><p>Search across every policy you've imported.</p></div></div>
            <div className="pt-card">
              <input className="pt-search" placeholder="Search by client name\u2026 (min 2 characters)" value={clientSearch} onChange={(e) => setClientSearch(e.target.value)} />
              {clientSearch.trim().length >= 2 && (
                clientMatches.length === 0 ? (
                  <p className="pt-hint" style={{ marginTop: 12 }}>No matches. Note: client name only shows if you mapped that column during import.</p>
                ) : (
                  <table className="pt-table" style={{ marginTop: 12 }}>
                    <thead><tr><th>Client</th><th>Agent</th><th>Carrier</th><th>Product</th><th>Plan type</th><th>Sale date</th><th>Status</th><th className="num">Amount</th></tr></thead>
                    <tbody>
                      {clientMatches.map((r) => (
                        <tr key={r.id}><td>{r.clientName}</td><td>{r.agent}</td><td>{r.carrier}</td><td>{r.product}</td><td>{r.planType}</td><td>{fmtDate(r.saleDate)}</td><td>{r.status}</td><td className="num mono">{fmtMoney(r.commissionAmount)}</td></tr>
                      ))}
                    </tbody>
                  </table>
                )
              )}
            </div>
          </div>
        )}

        {view === "directory" && (
          <div>
            <div className="pt-page-head"><div><h1>Agent directory</h1><p>The master list your imports match against, so one agent always shows up as one name.</p></div></div>

            {!cloudCfg ? (
              <div className="pt-card"><p className="pt-hint">Connect your database (Database connection tab) to use the Agent Directory \u2014 it needs somewhere to store your agent list.</p></div>
            ) : !directoryAvailable ? (
              <div className="pt-card">
                <p className="pt-error">Your database doesn't have the directory tables yet.</p>
                <p className="pt-hint" style={{ marginTop: 6, marginBottom: 10 }}>Run this once in your Supabase SQL Editor, then refresh this page:</p>
                <pre className="pt-sql">{MIGRATION_SQL}</pre>
              </div>
            ) : (
              <>
                <div className="pt-tabs">
                  <button className={"pt-tab" + (directoryTab === "directory" ? " active" : "")} onClick={() => setDirectoryTab("directory")}>Directory ({agentDirectory.length})</button>
                  <button className={"pt-tab" + (directoryTab === "unmatched" ? " active" : "")} onClick={() => setDirectoryTab("unmatched")}>Unmatched names ({unmatchedAgentNames.length})</button>
                </div>

                {directoryTab === "directory" && (
                  <div className="pt-grid-list">
                    <div className="pt-card">
                      <div className="pt-inline-form">
                        <input placeholder="Agent name" value={newAgentName} onChange={(e) => setNewAgentName(e.target.value)} />
                        <input placeholder="NPN (optional)" value={newAgentNpn} onChange={(e) => setNewAgentNpn(e.target.value)} style={{ maxWidth: 140 }} />
                        <button className="pt-btn primary small" disabled={!newAgentName.trim()} onClick={addAgentManual}><UserPlus size={13} /> Add</button>
                      </div>
                      <table className="pt-table" style={{ marginTop: 14 }}>
                        <thead><tr><th>Agent</th><th>NPN</th><th className="num">Aliases</th></tr></thead>
                        <tbody>
                          {agentDirectory.map((a) => (
                            <tr key={a.id} className={"pt-clickable" + (selectedDirAgent === a.id ? " selected" : "")} onClick={() => { setSelectedDirAgent(a.id); setEditAgentName(a.canonicalName); setEditAgentNpn(a.npn); }}>
                              <td>{a.canonicalName}</td>
                              <td className="mono">{a.npn || "\u2014"}</td>
                              <td className="num">{agentAliases.filter((al) => al.agentId === a.id).length}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {selectedDirAgent && (
                      <div className="pt-card">
                        <div className="pt-row-between">
                          <h3>Edit agent</h3>
                          <button className="pt-btn ghost small" onClick={() => deleteDirAgent(selectedDirAgent)}><Trash2 size={13} /> Delete</button>
                        </div>
                        <div className="pt-field" style={{ marginTop: 10 }}>
                          <label>Canonical name</label>
                          <input value={editAgentName} onChange={(e) => setEditAgentName(e.target.value)} />
                        </div>
                        <div className="pt-field" style={{ marginTop: 8 }}>
                          <label>NPN</label>
                          <input value={editAgentNpn} onChange={(e) => setEditAgentNpn(e.target.value)} />
                        </div>
                        <button className="pt-btn primary small" style={{ marginTop: 10 }} onClick={() => updateAgentDetails(selectedDirAgent)}>Save</button>

                        <div className="pt-mini-label" style={{ marginTop: 18 }}>Aliases (other IDs/names that mean this agent)</div>
                        {agentAliases.filter((al) => al.agentId === selectedDirAgent).map((al) => (
                          <div key={al.id} className="pt-alias-row">
                            <span className="pt-alias-type">{al.aliasType === "npn" ? "NPN" : al.aliasType === "carrier_id" ? `${al.carrier || "Carrier"} ID` : "Name"}</span>
                            <span className="mono">{al.aliasValue}</span>
                            <button className="pt-btn ghost small" onClick={() => deleteAlias(al.id)}><X size={12} /></button>
                          </div>
                        ))}
                        <div className="pt-inline-form" style={{ marginTop: 10 }}>
                          <select value={newAliasType} onChange={(e) => setNewAliasType(e.target.value)}>
                            <option value="npn">NPN</option>
                            <option value="carrier_id">Carrier-specific ID</option>
                            <option value="name_text">Name variant</option>
                          </select>
                          {newAliasType === "carrier_id" && <input placeholder="Carrier (e.g. Humana)" value={newAliasCarrier} onChange={(e) => setNewAliasCarrier(e.target.value)} style={{ maxWidth: 130 }} />}
                          <input placeholder="Value" value={newAliasValue} onChange={(e) => setNewAliasValue(e.target.value)} />
                          <button className="pt-btn ghost small" disabled={!newAliasValue.trim()} onClick={() => addAliasToAgent(selectedDirAgent)}>Add</button>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {directoryTab === "unmatched" && (
                  <div className="pt-card">
                    <p className="pt-hint" style={{ marginBottom: 12 }}>These are raw agent names from your imports that aren't linked to a directory entry yet. Link each one to the right agent, or add them as new.</p>
                    {unmatchedAgentNames.length === 0 ? (
                      <p className="pt-hint">Nothing unmatched \u2014 every name in your data is linked. \ud83c\udf89</p>
                    ) : (
                      <table className="pt-table">
                        <thead><tr><th>Name as imported</th><th className="num">Records</th><th className="num">Revenue</th><th>Link to\u2026</th><th></th></tr></thead>
                        <tbody>
                          {unmatchedAgentNames.map((u) => (
                            <tr key={u.name}>
                              <td>{u.name}</td>
                              <td className="num">{u.count}</td>
                              <td className="num mono">{fmtMoney(u.revenue)}</td>
                              <td>
                                <select value={linkChoice[u.name] || ""} onChange={(e) => setLinkChoice({ ...linkChoice, [u.name]: e.target.value })} style={{ minWidth: 160 }}>
                                  <option value="">Choose agent\u2026</option>
                                  {agentDirectory.map((a) => <option key={a.id} value={a.id}>{a.canonicalName}</option>)}
                                </select>
                              </td>
                              <td className="num">
                                <div className="pt-btn-row">
                                  <button className="pt-btn ghost small" disabled={!linkChoice[u.name]} onClick={() => linkRawNameToAgent(u.name, linkChoice[u.name], agentDirectory.find((a) => a.id === linkChoice[u.name])?.canonicalName)}><Link2 size={12} /> Link</button>
                                  <button className="pt-btn ghost small" onClick={() => createAgentFromRawName(u.name)}><UserPlus size={12} /> New</button>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                )}

                <div className="pt-card">
                  <h3>Import your agent roster</h3>
                  <p className="pt-hint" style={{ marginBottom: 10 }}>Bulk-add agents from a CRM export. Map the name column, and NPN if it has one.</p>
                  {!rosterFileName ? (
                    <label className="pt-btn ghost">
                      Choose roster file
                      <input type="file" accept=".xlsx,.xls,.csv" onChange={handleRosterFile} style={{ display: "none" }} />
                    </label>
                  ) : (
                    <>
                      <div className="pt-file-chip"><FileSpreadsheet size={14} /> {rosterFileName} \u00b7 {rosterRows.length} rows</div>
                      <div className="pt-mapping-grid" style={{ marginTop: 12 }}>
                        <div className="pt-field">
                          <label>Name column *</label>
                          <select value={rosterNameCol} onChange={(e) => setRosterNameCol(e.target.value)}>
                            <option value="">\u2014 choose \u2014</option>
                            {rosterHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                          </select>
                        </div>
                        <div className="pt-field">
                          <label>NPN column (optional)</label>
                          <select value={rosterNpnCol} onChange={(e) => setRosterNpnCol(e.target.value)}>
                            <option value="">\u2014 not in file \u2014</option>
                            {rosterHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                          </select>
                        </div>
                      </div>
                      <button className="pt-btn primary" style={{ marginTop: 12 }} disabled={!rosterNameCol} onClick={commitRosterImport}>Import {rosterRows.length} agents</button>
                    </>
                  )}
                </div>
              </>
            )}
          </div>
        )}

        {view === "manage" && (
          <div>
            <div className="pt-page-head"><div><h1>Manage data</h1><p>Every import you've run. Remove one if something loaded wrong.</p></div></div>
            <div className="pt-card">
              {batches.length === 0 ? <p className="pt-hint">No imports yet.</p> : (
                <table className="pt-table">
                  <thead><tr><th>Carrier</th><th>File</th><th>Imported</th><th className="num">Rows</th><th></th></tr></thead>
                  <tbody>
                    {[...batches].reverse().map((b) => (
                      <tr key={b.id}>
                        <td>{b.carrier}</td>
                        <td>{b.fileName}</td>
                        <td>{new Date(b.uploadedAt).toLocaleString()}</td>
                        <td className="num">{b.rowCount}</td>
                        <td className="num">
                          {confirmDeleteId === b.id ? (
                            <span className="pt-confirm-inline">
                              Remove {b.rowCount} rows?
                              <button className="pt-btn danger small" onClick={() => deleteBatch(b.id)}>Yes</button>
                              <button className="pt-btn ghost small" onClick={() => setConfirmDeleteId(null)}>Cancel</button>
                            </span>
                          ) : (
                            <button className="pt-btn ghost small" onClick={() => setConfirmDeleteId(b.id)}><Trash2 size={13} /></button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
            {records.length > 0 && (
              <div className="pt-card">
                <div className="pt-row-between">
                  <div>
                    <h3>Backup / reset</h3>
                    <p className="pt-hint">Export everything as a CSV backup, or wipe all data to start fresh.</p>
                  </div>
                  <div className="pt-btn-row">
                    <button className="pt-btn ghost" onClick={() => exportCSV(records, "all-production-data.csv")}><Download size={14} /> Export all</button>
                    {confirmClearAll ? (
                      <span className="pt-confirm-inline">
                        Clear everything?
                        <button className="pt-btn danger small" onClick={clearAllData}>Yes, clear</button>
                        <button className="pt-btn ghost small" onClick={() => setConfirmClearAll(false)}>Cancel</button>
                      </span>
                    ) : (
                      <button className="pt-btn danger" onClick={() => setConfirmClearAll(true)}><Trash2 size={14} /> Clear all data</button>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
        {view === "settings" && (
          <div>
            <div className="pt-page-head">
              <div>
                <h1>Database connection</h1>
                <p>Connect Ledger to your own Supabase database so your data is backed up and available anywhere.</p>
              </div>
            </div>

            <div className="pt-card">
              <div className="pt-row-between">
                <h3>Status</h3>
                <span className={"pt-status-badge " + (cloudCfg ? "on" : "off")}>
                  {cloudCfg ? <><Cloud size={14} /> Connected</> : <><CloudOff size={14} /> Local only (this browser)</>}
                </span>
              </div>
              {cloudError && <p className="pt-error">{cloudError}</p>}
              <div className="pt-field" style={{ marginTop: 12 }}>
                <label>Supabase project URL</label>
                <input value={cloudTestUrl} onChange={(e) => setCloudTestUrl(e.target.value)} placeholder="https://xxxxx.supabase.co" />
              </div>
              <div className="pt-field" style={{ marginTop: 8 }}>
                <label>Supabase anon public API key</label>
                <input value={cloudTestKey} onChange={(e) => setCloudTestKey(e.target.value)} placeholder="eyJhbGciOi..." />
              </div>
              <div className="pt-btn-row" style={{ marginTop: 14 }}>
                <button className="pt-btn primary" disabled={!cloudTestUrl.trim() || !cloudTestKey.trim() || cloudStatus === "connecting"} onClick={connectCloud}>
                  {cloudStatus === "connecting" ? "Connecting\u2026" : "Connect"}
                </button>
                {cloudCfg && <button className="pt-btn ghost" onClick={disconnectCloud}>Disconnect / use local only</button>}
              </div>
              <p className="pt-hint" style={{ marginTop: 10 }}>Your URL and key are stored only in this browser and sent only to your own Supabase project \u2014 never to Claude or anyone else.</p>
            </div>

            <div className="pt-card">
              <h3>Setup SQL</h3>
              <p className="pt-hint" style={{ marginBottom: 10 }}>Run this once in your Supabase project's SQL Editor before connecting.</p>
              <pre className="pt-sql">{SETUP_SQL}</pre>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function EmptyState({ onGo }) {
  return (
    <div className="pt-card pt-empty">
      <UploadCloud size={30} strokeWidth={1.3} color="#B8863B" />
      <p className="pt-upload-title">No production data yet</p>
      <p className="pt-upload-sub">Import your first carrier statement to start seeing trends.</p>
      <button className="pt-btn primary" style={{ marginTop: 12 }} onClick={onGo}>Import a statement</button>
    </div>
  );
}
function StatCard({ label, value, tone }) {
  return (
    <div className="pt-stat-card">
      <div className="pt-stat-label">{label}</div>
      <div className={"pt-stat-value " + tone}>{value}</div>
    </div>
  );
}
function FilterSelect({ label, value, onChange, options }) {
  return (
    <div className="pt-filter">
      <label>{label}</label>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        <option value="All">All</option>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

const CSS = `
.pt-app {
  --ink: #12203B;
  --ink-2: #1E2A3A;
  --paper: #F5F6F8;
  --card: #FFFFFF;
  --border: #E2E4E9;
  --muted: #64748B;
  --gold: #B8863B;
  --gold-soft: #F3E9D6;
  --green: #1F7A4D;
  --rose: #B4433D;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
  color: var(--ink-2);
  background: var(--paper);
  display: flex;
  min-height: 100vh;
  width: 100%;
}
.pt-loading { align-items: center; justify-content: center; color: var(--muted); }
.mono { font-family: "SF Mono", "Roboto Mono", Menlo, monospace; font-variant-numeric: tabular-nums; }

.pt-sidebar {
  width: 220px;
  flex-shrink: 0;
  background: var(--ink);
  color: #E7EAF2;
  display: flex;
  flex-direction: column;
  padding: 20px 14px;
  background-image: repeating-linear-gradient(180deg, rgba(255,255,255,0.025) 0px, rgba(255,255,255,0.025) 1px, transparent 1px, transparent 25px);
}
.pt-brand { display: flex; align-items: center; gap: 10px; padding: 6px 8px 22px; }
.pt-brand-mark {
  width: 30px; height: 30px; border-radius: 6px; background: var(--gold);
  color: var(--ink); font-family: Georgia, "Iowan Old Style", serif; font-weight: 700;
  display: flex; align-items: center; justify-content: center; font-size: 16px;
}
.pt-brand-name { font-family: Georgia, "Iowan Old Style", serif; font-size: 17px; letter-spacing: 0.02em; }
.pt-brand-sub { font-size: 10px; color: #8A96AE; text-transform: uppercase; letter-spacing: 0.08em; margin-top: 1px; }
.pt-nav { display: flex; flex-direction: column; gap: 2px; margin-top: 6px; }
.pt-nav-item {
  display: flex; align-items: center; gap: 10px; padding: 9px 10px; border-radius: 6px;
  background: transparent; border: none; border-left: 2px solid transparent; color: #B7C0D4;
  font-size: 13.5px; text-align: left; cursor: pointer; transition: background 0.12s;
}
.pt-nav-item:hover { background: rgba(255,255,255,0.05); }
.pt-nav-item.active { background: rgba(184,134,59,0.14); border-left-color: var(--gold); color: #fff; }
.pt-sidebar-footer { margin-top: auto; padding: 14px 10px 6px; border-top: 1px solid rgba(255,255,255,0.08); }
.pt-conn-dot { display: flex; align-items: center; gap: 5px; font-size: 11px; color: #8A96AE; margin-bottom: 10px; }
.pt-conn-dot.on { color: #8FD1A8; }
.pt-status-badge { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; padding: 4px 10px; border-radius: 20px; }
.pt-status-badge.on { background: #EAF5EE; color: var(--green); }
.pt-status-badge.off { background: #F1F2F4; color: var(--muted); }
.pt-sql { background: var(--ink); color: #E7EAF2; padding: 14px; border-radius: 6px; font-size: 11.5px; overflow: auto; line-height: 1.6; font-family: "SF Mono", Menlo, monospace; white-space: pre; }
.pt-tabs { display: flex; gap: 4px; margin-bottom: 14px; border-bottom: 1px solid var(--border); }
.pt-tab { background: none; border: none; padding: 8px 14px; font-size: 13px; color: var(--muted); cursor: pointer; border-bottom: 2px solid transparent; margin-bottom: -1px; }
.pt-tab.active { color: var(--ink); border-bottom-color: var(--gold); font-weight: 600; }
.pt-inline-form { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.pt-inline-form input, .pt-inline-form select { border: 1px solid var(--border); border-radius: 5px; padding: 6px 9px; font-size: 12.5px; }
.pt-alias-row { display: flex; align-items: center; gap: 10px; font-size: 12.5px; padding: 6px 0; border-bottom: 1px solid var(--border); }
.pt-alias-type { background: var(--gold-soft); color: var(--ink); border-radius: 4px; padding: 2px 7px; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.04em; }
.pt-footer-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.06em; color: #8A96AE; }
.pt-footer-value { font-family: "SF Mono", monospace; font-size: 19px; margin-top: 3px; color: #fff; }
.pt-footer-sub { font-size: 11px; color: #8A96AE; margin-top: 2px; }

.pt-main { flex: 1; padding: 30px 36px 60px; max-width: 1180px; overflow-x: hidden; }
.pt-page-head { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 20px; gap: 12px; }
.pt-page-head h1 { font-family: Georgia, "Iowan Old Style", serif; font-size: 26px; margin: 0 0 4px; color: var(--ink); }
.pt-page-head p { margin: 0; color: var(--muted); font-size: 13.5px; }

.pt-toast { display: flex; align-items: center; gap: 8px; padding: 10px 14px; border-radius: 6px; font-size: 13px; margin-bottom: 16px; border: 1px solid; }
.pt-toast.success { background: #EAF5EE; border-color: #BFE0CB; color: var(--green); }
.pt-toast.error { background: #FBEDEC; border-color: #EFC4C0; color: var(--rose); }

.pt-card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 20px; margin-bottom: 18px; }
.pt-card h3 { font-size: 14.5px; margin: 0 0 14px; font-weight: 600; color: var(--ink); }

.pt-cards { display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; margin-bottom: 18px; }
.pt-stat-card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 14px 16px; }
.pt-stat-label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); margin-bottom: 6px; }
.pt-stat-value { font-family: "SF Mono", monospace; font-size: 20px; letter-spacing: -0.01em; border-bottom: 2px solid var(--gold-soft); padding-bottom: 2px; display: inline-block; }
.pt-stat-value.rose { color: var(--rose); border-bottom-color: #F1D4D1; }
.pt-stat-value.ink { color: var(--ink); }

.pt-grid-2 { display: grid; grid-template-columns: 1.3fr 1fr; gap: 18px; margin-bottom: 0; }
.pt-grid-2 .pt-card { margin-bottom: 18px; }
.pt-grid-list { display: grid; grid-template-columns: 0.9fr 1.4fr; gap: 18px; align-items: start; }

.pt-filters { display: flex; flex-wrap: wrap; gap: 14px; align-items: flex-end; margin-bottom: 18px; background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 14px 16px; }
.pt-filter { display: flex; flex-direction: column; gap: 4px; }
.pt-filter label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); }
.pt-filter select, .pt-filter input { border: 1px solid var(--border); border-radius: 5px; padding: 6px 8px; font-size: 13px; background: #fff; color: var(--ink-2); min-width: 130px; }

.pt-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.pt-table th { text-align: left; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); font-weight: 600; padding: 8px 10px; border-bottom: 1px solid var(--border); }
.pt-table td { padding: 9px 10px; border-bottom: 1px solid var(--border); }
.pt-table tr:last-child td { border-bottom: none; }
.pt-table th.num, .pt-table td.num { text-align: right; }
.pt-clickable { cursor: pointer; }
.pt-clickable:hover { background: #F9F7F1; }
.pt-clickable.selected { background: var(--gold-soft); }

.pt-mini-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 12px; }
.pt-mini-label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); margin-bottom: 6px; }
.pt-mini-row { display: flex; justify-content: space-between; font-size: 13px; padding: 5px 0; border-bottom: 1px solid var(--border); }
.pt-preview-scroll { max-height: 320px; overflow: auto; margin-top: 4px; }

.pt-btn { display: inline-flex; align-items: center; gap: 6px; border-radius: 6px; padding: 8px 14px; font-size: 13px; font-weight: 500; cursor: pointer; border: 1px solid transparent; }
.pt-btn.primary { background: var(--gold); color: var(--ink); border-color: var(--gold); }
.pt-btn.primary:disabled { opacity: 0.45; cursor: not-allowed; }
.pt-btn.ghost { background: #fff; color: var(--ink-2); border-color: var(--border); }
.pt-btn.text { background: transparent; color: var(--muted); border: none; padding: 6px 4px; }
.pt-btn.danger { background: #fff; color: var(--rose); border-color: #EFC4C0; }
.pt-btn.small { padding: 4px 8px; font-size: 12px; }
.pt-btn-row { display: flex; gap: 10px; align-items: center; }
.pt-confirm-inline { display: inline-flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--muted); }

.pt-row-between { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
.pt-search { width: 100%; border: 1px solid var(--border); border-radius: 6px; padding: 8px 10px; font-size: 13px; margin-bottom: 12px; }

.pt-upload-zone, .pt-empty { display: flex; flex-direction: column; align-items: center; text-align: center; padding: 48px 20px; }
.pt-upload-title { font-size: 15px; font-weight: 600; color: var(--ink); margin: 12px 0 2px; }
.pt-upload-sub { font-size: 13px; color: var(--muted); margin: 0; }
.pt-file-chip { display: inline-flex; align-items: center; gap: 6px; background: var(--gold-soft); color: var(--ink); padding: 5px 10px; border-radius: 5px; font-size: 12.5px; }

.pt-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 4px; }
.pt-field label { font-size: 11.5px; color: var(--muted); font-weight: 500; }
.pt-field input, .pt-field select { border: 1px solid var(--border); border-radius: 5px; padding: 7px 9px; font-size: 13px; }
.pt-mapping-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px 18px; }
.pt-plantype-block { margin-top: 18px; padding-top: 16px; border-top: 1px solid var(--border); display: flex; flex-direction: column; gap: 8px; }
.pt-plantype-block label:first-child { font-size: 11.5px; color: var(--muted); font-weight: 500; }
.pt-plantype-block select, .pt-plantype-block input { border: 1px solid var(--border); border-radius: 5px; padding: 7px 9px; font-size: 13px; max-width: 260px; }
.pt-radio-row { display: flex; gap: 18px; }
.pt-radio { display: flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink-2); }
.pt-hint { font-size: 12px; color: var(--muted); margin: 4px 0 0; }
.pt-error { font-size: 12.5px; color: var(--rose); margin: 8px 0 0; }

@media (max-width: 900px) {
  .pt-cards { grid-template-columns: repeat(2, 1fr); }
  .pt-grid-2, .pt-grid-list, .pt-mapping-grid, .pt-mini-grid { grid-template-columns: 1fr; }
}
`;
