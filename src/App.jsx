import { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { Upload, Search, Trash2, ChevronUp, ChevronDown, Loader2, AlertTriangle, X, Gauge, FileSpreadsheet, Download } from "lucide-react";

// ---- design tokens ----
// bg: #1B1D22 (warm charcoal, not pure black)
// panel: #24272E
// paper row: #22252B / #26292F stripe
// ink: #ECE7DC
// muted ink: #9A9C9E
// amber (in-stock): #F2A93B
// teal (service/detail): #3FA796
// rust (vendor/transit): #C1502E
// slate-blue (line/borders): #3A3F49

const STATUS_COLORS = {
  "IN-STOCK": "#3FA796",
  "SERVICE": "#F2A93B",
  "DETAIL": "#7C8CD8",
  "VENDOR": "#C1502E",
  "IN-TRANSIT": "#C98BD9",
};

function statusColor(s) {
  return STATUS_COLORS[(s || "").toUpperCase()] || "#9A9C9E";
}

function parseMoney(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

function parseNum(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("Couldn't read this file from your device."));
    r.readAsText(file);
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("Couldn't read this file from your device."));
    r.readAsArrayBuffer(file);
  });
}

// Small CSV parser that handles quoted fields (so commas/quotes inside a
// field, like a model description, don't break the columns).
function parseCSV(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  const pushField = () => { row.push(field); field = ""; };
  const pushRow = () => { pushField(); rows.push(row); row = []; };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ",") pushField();
      else if (c === "\n") { if (field !== "" || row.length > 0) pushRow(); }
      else if (c === "\r") { /* skip */ }
      else field += c;
    }
  }
  if (field !== "" || row.length > 0) pushRow();
  if (rows.length === 0) return [];
  const header = rows[0].map((h) => h.trim().toLowerCase());
  const idx = (name) => header.indexOf(name);
  const cols = {
    stock: idx("stock/order") >= 0 ? idx("stock/order") : idx("stock"),
    year: idx("year"),
    make: idx("make"),
    model: idx("model"),
    desc: idx("model desc") >= 0 ? idx("model desc") : idx("desc"),
    status: idx("status"),
    color: idx("exterior color") >= 0 ? idx("exterior color") : idx("color"),
    odometer: idx("odometer"),
    vin: idx("vin"),
    days: idx("days"),
    price: idx("list price") >= 0 ? idx("list price") : idx("price"),
    certified: idx("certified"),
  };
  return rows.slice(1)
    .filter((r) => r.some((v) => v && v.trim() !== ""))
    .map((r) => ({
      s: cols.stock >= 0 ? r[cols.stock] : "",
      y: cols.year >= 0 ? parseNum(r[cols.year]) : null,
      mk: cols.make >= 0 ? r[cols.make] : "",
      md: cols.model >= 0 ? r[cols.model] : "",
      de: cols.desc >= 0 ? r[cols.desc] : "",
      st: cols.status >= 0 ? r[cols.status] : "",
      c: cols.color >= 0 ? r[cols.color] : "",
      o: cols.odometer >= 0 ? r[cols.odometer] : "",
      v: cols.vin >= 0 ? r[cols.vin] : "",
      d: cols.days >= 0 ? r[cols.days] : "",
      p: cols.price >= 0 ? r[cols.price] : "",
      ce: cols.certified >= 0 ? /^y/i.test((r[cols.certified] || "").trim()) : false,
    }));
}

// Classifies a vehicle's body type from its Model name. Rule-based rather
// than guessed per-row, so it's consistent — checked in order, first match wins.
const TYPE_RULES = [
  [/CIVIC HATCHBACK|PRIUS/i, "Hatchback"],
  [/CIVIC SEDAN|S-CLASS|SENTRA|ACCORD|^CAMRY|ELANTRA|LS 500|MALIBU|^COROLLA(?! CROS)/i, "Sedan"],
  [/CORVETTE|CONV/i, "Convertible"],
  [/2 SERIES/i, "Coupe"],
  [/GRAND CARAVAN|ODYSSEY/i, "Minivan"],
  [/SIERRA|SILVERADO|^TACOMA|^TITAN$|^TUNDRA|^COLORADO|F-150|GLADIATOR|MAVERICK/i, "Truck"],
  [/^RAV4|^ROGUE|^SANTA FE|SPORTAGE|TERRAIN|^TUCSON|WRANGLER|^XC|^4RUNNER|ATLAS|BLAZER|COROLLA CROS|CR-V|CROSSTREK|^CX-|EXPLORER|GRAND CHEROKEE|GRAND PTM|^GX |KONA|^MDX|OUTBACK|PALISADE|PILOT|^HIGHLANDER/i, "SUV"],
];

function classifyType(model) {
  const m = (model || "").toUpperCase();
  for (const [re, type] of TYPE_RULES) {
    if (re.test(m)) return type;
  }
  return "Other";
}

function normalizeLegacyRow(r) {
  const model = (r.md ?? "").toString().trim();
  return {
    stock: r.s ?? "",
    year: r.y ?? "",
    make: (r.mk ?? "").toString().trim(),
    model,
    type: classifyType(model),
    desc: r.de ?? "",
    status: (r.st ?? "").toString().toUpperCase().trim(),
    color: r.c ?? "",
    odometer: parseNum(r.o),
    vin: (r.v ?? "").toString().trim().toUpperCase(),
    days: parseNum(r.d),
    price: parseMoney(r.p),
    certified: !!r.ce,
    recall: "",
    drivetrain: "",
  };
}

// "2024 Toyota Corolla LE" -> { year: 2024, make: "Toyota", model: "Corolla LE" }
function parseVehicleString(vehicle) {
  const s = (vehicle || "").toString().trim();
  const m = s.match(/^(\d{4})\s+(\S+)\s+(.*)$/);
  if (!m) return { year: null, make: "", model: s };
  return { year: parseInt(m[1], 10), make: m[2], model: m[3].trim() };
}

function normalizeKey(k) {
  return k.replace(/\s+/g, " ").trim().toLowerCase();
}

function getField(row, name) {
  for (const k of Object.keys(row)) {
    if (normalizeKey(k) === name) return row[k];
  }
  return undefined;
}

// Native pricing-export schema: Photos, Autowriter Description, Vehicle,
// Stock #, VIN, Class, Certified, Deleted Date, Status, Recall Status, Body,
// Color, Disp, Price / % Mkt, Last $ Change, Odometer.
//
// Ignored on purpose: Photos, Autowriter Description, Deleted Date, Status,
// Disp — Status in particular is always blank in this export (verified
// against a real pull), so it can't drive the status filter/badge the way
// the legacy CSV format could. Certified and Days are kept: Certified is
// read straight from the export, and Days is left null since this export
// doesn't include a days-on-lot column (unlike the legacy CSV, which does).
function normalizePricingRow(row) {
  const { year, make, model } = parseVehicleString(getField(row, "vehicle"));
  const classVal = (getField(row, "class") || "").toString();
  const certifiedVal = (getField(row, "certified") || "").toString();
  return {
    stock: getField(row, "stock #") ?? "",
    year,
    make,
    model,
    type: classVal.split(",")[0].trim() || "Other",
    desc: (getField(row, "body") || "").toString(),
    status: "", // not populated in this export — left blank rather than guessed
    color: (getField(row, "color") || "").toString(),
    odometer: parseNum(getField(row, "odometer")),
    vin: (getField(row, "vin") || "").toString().trim().toUpperCase(),
    days: null, // this export doesn't include a days-on-lot column
    price: parseMoney(getField(row, "price / % mkt")),
    certified: /^y/i.test(certifiedVal.trim()),
    recall: (getField(row, "recall status") || "").toString().trim(),
    drivetrain: "",
  };
}

// Days on lot, from an Inventory Date that may come through as a JS Date
// (XLSX.read with cellDates:true) or as a "MM/DD/YYYY" string.
function daysSince(dateVal) {
  if (!dateVal) return null;
  const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return null;
  const start = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.round((today - start) / 86400000);
}

// "Pricing View" export schema (wider report, 46 columns) — includes real
// Inventory Date and Recall Status Icon Small, which the plain pricing
// export above doesn't have. Detected by the presence of "Inventory Date".
function normalizePricingViewRow(row) {
  const { year, make, model } = parseVehicleString(getField(row, "vehicle"));
  const classVal = (getField(row, "class") || "").toString();
  const certifiedVal = (getField(row, "certified") || "").toString();
  const exteriorColor = (getField(row, "color") || "").toString().trim();
  const interiorColor = (getField(row, "interior color") || "").toString().trim();
  const engine = (getField(row, "engine") || "").toString().trim();
  const drivetrainType = (getField(row, "drivetrain type") || "").toString().trim();
  return {
    stock: getField(row, "stock #") ?? "",
    year,
    make,
    model,
    type: classVal.split(",")[0].trim() || "Other",
    desc: (getField(row, "body") || "").toString(),
    status: "",
    color: interiorColor ? `${exteriorColor} / ${interiorColor}` : exteriorColor,
    odometer: parseNum(getField(row, "odometer")),
    vin: (getField(row, "vin") || "").toString().trim().toUpperCase(),
    days: daysSince(getField(row, "inventory date")),
    price: parseMoney(getField(row, "price")),
    certified: /^y/i.test(certifiedVal.trim()),
    recall: (getField(row, "recall status icon small") || "").toString().trim(),
    drivetrain: engine && drivetrainType ? `${engine}/${drivetrainType}` : (engine || drivetrainType),
  };
}

// Reads a file (legacy report CSV, or a native .xlsx/.xls pricing export)
// and returns a flat array of normalized vehicle records — no scanDate yet,
// that's attached by the caller.
async function extractRows(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv")) {
    const text = await readFileAsText(file);
    const raw = parseCSV(text).filter((r) => r.v);
    return raw.map(normalizeLegacyRow);
  }
  const buf = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (json.length && getField(json[0], "inventory date") !== undefined) {
    return json.map(normalizePricingViewRow).filter((r) => r.vin);
  }
  return json.map(normalizePricingRow).filter((r) => r.vin);
}

function csvEscape(v) {
  const s = v === null || v === undefined ? "" : String(v);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function recordsToCSV(records) {
  const header = ["Stock/Order", "Year", "Make", "Model", "Type", "Model Desc", "Status", "Recall Status", "Exterior Color", "Odometer", "VIN", "Days", "List Price", "Certified", "Scan Date"];
  const lines = [header.join(",")];
  for (const r of records) {
    lines.push([
      r.stock, r.year, r.make, r.model, r.type, r.desc, r.status, r.recall, r.color,
      r.odometer ?? "", r.vin, r.days ?? "", r.price ?? "", r.certified ? "Yes" : "", r.scanDate,
    ].map(csvEscape).join(","));
  }
  return lines.join("\n");
}

function MultiSelect({ label, options, selected, onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  function toggleValue(v) {
    if (selected.includes(v)) onChange(selected.filter((x) => x !== v));
    else onChange([...selected, v]);
  }
  return (
    <div
      ref={ref}
      tabIndex={-1}
      onBlur={(e) => {
        if (!ref.current || !ref.current.contains(e.relatedTarget)) setOpen(false);
      }}
      style={{ position: "relative" }}
    >
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="lg-input"
        style={{ textAlign: "left", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", color: selected.length ? "#ECE7DC" : "#B4B8BF" }}
      >
        <span>{label}{selected.length > 0 ? ` (${selected.length})` : ""}</span>
        <ChevronDown size={13} style={{ opacity: 0.6, flexShrink: 0 }} />
      </button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0,
          background: "#3D4354", border: "1px solid #6B7280", borderRadius: 6,
          maxHeight: 230, overflowY: "auto", zIndex: 30, padding: 6,
        }}>
          {selected.length > 0 && (
            <button type="button" onClick={() => onChange([])}
              style={{ background: "none", border: "none", color: "#F2A93B", fontSize: 11.5, cursor: "pointer", padding: "4px 6px", display: "block" }}>
              Clear {label.toLowerCase()}
            </button>
          )}
          {options.length === 0 && <div style={{ fontSize: 12, color: "#6B6D70", padding: "4px 6px" }}>No options yet</div>}
          {options.map((opt) => (
            <label key={opt} className="lg-row" style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5, padding: "5px 6px", cursor: "pointer", borderRadius: 4 }}>
              <input type="checkbox" checked={selected.includes(opt)} onChange={() => toggleValue(opt)} />
              {opt}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

export default function LotLedger() {
  const [records, setRecords] = useState(() => {
    try {
      const saved = localStorage.getItem("lot-ledger-records");
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

  function saveRecords(next) {
    setRecords(next);
    try {
      localStorage.setItem("lot-ledger-records", JSON.stringify(next));
    } catch (e) {
      console.error("Couldn't save to this browser's storage", e);
    }
  }

  const [queue, setQueue] = useState([]); // {name, status, error, count}
  const [dragOver, setDragOver] = useState(false);
  const [sortField, setSortField] = useState("price");
  const [sortDir, setSortDir] = useState("desc");
  const [showFilters, setShowFilters] = useState(true);
  const [confirmingClear, setConfirmingClear] = useState(false);
  const [exportHref, setExportHref] = useState(null);
  const [exportName, setExportName] = useState("");
  const fileInputRef = useRef(null);
  const scrollRef = useRef(null);
  const tableRef = useRef(null);
  const edgeTouch = useRef({ startY: 0, startScrollTop: 0 });

  const [filters, setFilters] = useState({
    search: "",
    make: [],
    model: [],
    type: [],
    status: [],
    recall: [],
    scanDate: [],
    yearMin: "",
    yearMax: "",
    priceMin: "",
    priceMax: "",
    odoMax: "",
    certifiedOnly: false,
  });

  async function processFile(file, scanDate) {
    const qid = file.name + "-" + Date.now();
    setQueue((q) => [...q, { id: qid, name: file.name, status: "reading" }]);
    try {
      const rawRows = await extractRows(file);
      if (rawRows.length === 0) {
        setQueue((q) =>
          q.map((it) => (it.id === qid ? { ...it, status: "error", error: "No vehicle rows found — check the file has the expected columns (VIN, Stock #/Stock-Order, etc.)" } : it))
        );
        return;
      }
      const newRows = rawRows.map((r) => ({ ...r, scanDate }));
      setRecords((prev) => {
        // The batch-level wipe already happened in handleFiles, so every file
        // selected together in one import simply merges in by VIN.
        const map = new Map(prev.map((r) => [r.vin, r]));
        for (const row of newRows) {
          map.set(row.vin, row);
        }
        const next = Array.from(map.values());
        try {
          localStorage.setItem("lot-ledger-records", JSON.stringify(next));
        } catch (e) {
          console.error("Couldn't save to this browser's storage", e);
        }
        return next;
      });
      setQueue((q) =>
        q.map((it) => (it.id === qid ? { ...it, status: "done", count: newRows.length } : it))
      );
    } catch (e) {
      setQueue((q) => q.map((it) => (it.id === qid ? { ...it, status: "error", error: e.message } : it)));
    }
  }

  function handleFiles(fileList) {
    const files = Array.from(fileList).filter((f) => /\.(csv|xlsx|xls)$/i.test(f.name));
    if (files.length === 0) return;
    const scanDate = new Date().toLocaleDateString();
    // Wipe once per import action — every CSV picked together in this batch
    // merges into one dataset; a later, separate import replaces it.
    saveRecords([]);
    files.forEach((f) => processFile(f, scanDate));
  }

  function clearAll() {
    saveRecords([]);
    setQueue([]);
    setConfirmingClear(false);
  }

  function exportCSV() {
    const csv = recordsToCSV(records);
    const dateLabel = (scanDates[0] || new Date().toLocaleDateString()).replace(/[^0-9A-Za-z-]/g, "-");
    setExportHref("data:text/csv;charset=utf-8," + encodeURIComponent(csv));
    setExportName(`lot-ledger-${dateLabel}.csv`);
  }

  const makes = useMemo(() => Array.from(new Set(records.map((r) => r.make).filter(Boolean))).sort(), [records]);
  const models = useMemo(() => Array.from(new Set(records.map((r) => r.model).filter(Boolean))).sort(), [records]);
  const types = useMemo(() => Array.from(new Set(records.map((r) => r.type).filter(Boolean))).sort(), [records]);
  const scanDates = useMemo(() => Array.from(new Set(records.map((r) => r.scanDate).filter(Boolean))).sort().reverse(), [records]);

  const filtered = useMemo(() => {
    let out = records.filter((r) => {
      if (filters.search) {
        const s = filters.search.toLowerCase();
        const haystack = [
          r.stock, r.year, r.make, r.model, r.type, r.desc, r.status, r.recall,
          r.color, r.drivetrain, r.odometer, r.vin, r.days, r.price, r.certified ? "certified" : "",
          r.scanDate,
        ]
          .filter((v) => v !== null && v !== undefined)
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(s)) return false;
      }
      if (filters.make.length && !filters.make.includes(r.make)) return false;
      if (filters.model.length && !filters.model.includes(r.model)) return false;
      if (filters.type.length && !filters.type.includes(r.type)) return false;
      if (filters.status.length && !filters.status.includes(r.status)) return false;
      if (filters.recall.length && !filters.recall.includes(r.recall)) return false;
      if (filters.scanDate.length && !filters.scanDate.includes(r.scanDate)) return false;
      if (filters.yearMin && (!r.year || r.year < parseInt(filters.yearMin))) return false;
      if (filters.yearMax && (!r.year || r.year > parseInt(filters.yearMax))) return false;
      if (filters.priceMin && (r.price === null || r.price < parseFloat(filters.priceMin))) return false;
      if (filters.priceMax && (r.price === null || r.price > parseFloat(filters.priceMax))) return false;
      if (filters.odoMax && (r.odometer === null || r.odometer > parseFloat(filters.odoMax))) return false;
      if (filters.certifiedOnly && !r.certified) return false;
      return true;
    });
    out.sort((a, b) => {
      let av = a[sortField];
      let bv = b[sortField];
      if (av === null || av === undefined) av = typeof bv === "number" ? -Infinity : "";
      if (bv === null || bv === undefined) bv = typeof av === "number" ? -Infinity : "";
      if (typeof av === "string") av = av.toLowerCase();
      if (typeof bv === "string") bv = bv.toLowerCase();
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return out;
  }, [records, filters, sortField, sortDir]);

  const totalCount = records.length;

  function toggleSort(field) {
    if (sortField === field) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortField(field);
      setSortDir("asc");
    }
  }

  function SortIcon({ field }) {
    if (sortField !== field) return null;
    return sortDir === "asc" ? <ChevronUp size={13} style={{ display: "inline" }} /> : <ChevronDown size={13} style={{ display: "inline" }} />;
  }

  return (
    <div ref={scrollRef} className="lg-scroll" style={{ height: "100vh", overflowY: "auto", overscrollBehavior: "contain", WebkitOverflowScrolling: "touch", background: "#1E2027", color: "#ECE7DC", fontFamily: "'IBM Plex Sans', system-ui, sans-serif" }}>
      <style>{`
        html, body, #root { height: 100%; margin: 0; overscroll-behavior: none; overflow: hidden; }
        .lg-scroll { scrollbar-width: none; -ms-overflow-style: none; }
        .lg-scroll::-webkit-scrollbar { display: none; width: 0; height: 0; }
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .lg-mono { font-family: 'IBM Plex Mono', monospace; }
        .lg-display { font-family: 'Space Grotesk', sans-serif; }
        .lg-input {
          background: #3D4354; border: 1px solid #6B7280; color: #ECE7DC;
          border-radius: 6px; padding: 7px 10px; font-size: 13px; font-family: 'IBM Plex Sans', sans-serif;
          outline: none; width: 100%; box-sizing: border-box;
        }
        .lg-input:focus { border-color: #FFC15E; background: #454C60; }
        .lg-input::placeholder { color: #B4B8BF; }
        .lg-th { cursor: pointer; user-select: none; white-space: nowrap; }
        .lg-th:hover { color: #F2A93B; }
        .lg-row:hover { background: #2C303A !important; }
        ::-webkit-scrollbar { height: 10px; width: 10px; }
        ::-webkit-scrollbar-track { background: #1B1D22; }
        ::-webkit-scrollbar-thumb { background: #3A3F49; border-radius: 5px; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }
      `}</style>

      {/* Header */}
      <div style={{ background: "#000000", padding: "18px 16px", textAlign: "center" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
          <Gauge size={22} color="#E8C468" />
          <h1 className="lg-display" style={{ fontSize: 21, fontWeight: 700, margin: 0, letterSpacing: 2, color: "#E8C468", textTransform: "uppercase" }}>
            The Lot Ledger
          </h1>
        </div>
        <div style={{ fontSize: 11, color: "#7A7565", marginTop: 6 }}>
          designed by <b style={{ color: "#FFE29A" }}>Jeff Patrick</b>
        </div>
      </div>

      <div style={{ padding: "10px 12px", display: "flex", flexDirection: "column", gap: 8 }}>
        {/* Upload zone */}
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `1.5px dashed ${dragOver ? "#F2A93B" : "#3A3F49"}`,
              borderRadius: 10,
              padding: "8px",
              textAlign: "center",
              cursor: "pointer",
              background: dragOver ? "#24272E" : "transparent",
              transition: "border-color 0.15s, background 0.15s",
            }}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".csv,.xlsx,.xls,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              multiple
              style={{ display: "none" }}
              onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }}
            />
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14 }}>
              <div style={{ fontSize: 13.5, color: "#ECE7DC", lineHeight: 1.5 }}>
                Drop today's inventory CSV or Excel (.xlsx/.xls) export here, or{" "}
                <span style={{ color: "#F2A93B", fontWeight: 700 }}>CLICK</span> to choose a file
              </div>
              <FileSpreadsheet size={20} color="#F2A93B" style={{ flexShrink: 0 }} />
            </div>
          </div>
        </div>

        {/* Processing queue */}
        {queue.some((it) => it.status !== "done") && (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {queue.filter((it) => it.status !== "done").map((it) => (
              <div key={it.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, background: "#24272E", borderRadius: 6, padding: "7px 12px" }}>
                {it.status === "reading" ? (
                  <Loader2 size={14} className="lg-mono" style={{ animation: "spin 1s linear infinite" }} />
                ) : it.status === "error" ? (
                  <AlertTriangle size={14} color="#C1502E" />
                ) : (
                  <span style={{ color: "#3FA796" }}>✓</span>
                )}
                <span style={{ flex: 1 }}>{it.name}</span>
                <span style={{ color: "#9A9C9E" }}>
                  {it.status === "reading" && "reading…"}
                  {it.status === "done" && `${it.count} rows added`}
                  {it.status === "error" && it.error}
                </span>
                <X size={13} style={{ cursor: "pointer", color: "#6B6D70" }} onClick={() => setQueue((q) => q.filter((x) => x.id !== it.id))} />
              </div>
            ))}
          </div>
        )}

        {totalCount === 0 && queue.length === 0 && (
          <div style={{ textAlign: "center", color: "#6B6D70", fontSize: 13, padding: "30px 0" }}>
            No inventory loaded — import a CSV above to get started.
          </div>
        )}

        {totalCount > 0 && (
          <>
            {/* Filters */}
            <div style={{ background: "#24272E", borderRadius: 10, padding: "10px 12px 8px" }}>
              {/* General search */}
              <input
                className="lg-input"
                style={{ marginBottom: 8, padding: "6px 10px", textAlign: "center" }}
                placeholder="Search anything (stock, VIN, model, color, price…)"
                value={filters.search}
                onChange={(e) => setFilters((f) => ({ ...f, search: e.target.value }))}
              />

              <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: showFilters ? 8 : 0 }}>
                <button onClick={() => setShowFilters((s) => !s)} style={{ background: "none", border: "none", color: "#9A9C9E", fontSize: 12, cursor: "pointer" }}>
                  {showFilters ? "Hide" : "Show"}
                </button>
              </div>

              {showFilters && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8, alignItems: "start" }}>
                  <MultiSelect label="Make" options={makes} selected={filters.make}
                    onChange={(vals) => setFilters((f) => ({ ...f, make: vals }))} />
                  <MultiSelect label="Model" options={models} selected={filters.model}
                    onChange={(vals) => setFilters((f) => ({ ...f, model: vals }))} />
                  <MultiSelect label="Type" options={types} selected={filters.type}
                    onChange={(vals) => setFilters((f) => ({ ...f, type: vals }))} />
                  <input className="lg-input" type="number" placeholder="Max odometer" value={filters.odoMax}
                    onChange={(e) => setFilters((f) => ({ ...f, odoMax: e.target.value }))} />
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    <input className="lg-input" type="number" placeholder="Year min" value={filters.yearMin}
                      onChange={(e) => setFilters((f) => ({ ...f, yearMin: e.target.value }))} />
                    <input className="lg-input" type="number" placeholder="Year max" value={filters.yearMax}
                      onChange={(e) => setFilters((f) => ({ ...f, yearMax: e.target.value }))} />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    <input className="lg-input" type="number" placeholder="Price min ($)" value={filters.priceMin}
                      onChange={(e) => setFilters((f) => ({ ...f, priceMin: e.target.value }))} />
                    <input className="lg-input" type="number" placeholder="Price max ($)" value={filters.priceMax}
                      onChange={(e) => setFilters((f) => ({ ...f, priceMax: e.target.value }))} />
                  </div>
                  <div style={{ gridColumn: "1 / -1", display: "flex", alignItems: "center", justifyContent: "center", gap: 26, marginTop: 2 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13 }}>
                      <input type="checkbox" checked={filters.certifiedOnly}
                        onChange={(e) => setFilters((f) => ({ ...f, certifiedOnly: e.target.checked }))} />
                      Certified only
                    </label>
                    <span style={{ fontSize: 12.5, color: "#9A9C9E" }}>
                      {filtered.length}/{totalCount} vehicles
                    </span>
                  </div>
                </div>
              )}
            </div>

            {/* Table */}
            <div ref={tableRef} className="lg-scroll" style={{ background: "#24272E", borderRadius: 10, overflow: "auto", maxHeight: "60vh", overscrollBehavior: "contain", WebkitOverflowScrolling: "touch" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr style={{ position: "sticky", top: 0, background: "#1F2228", zIndex: 1 }}>
                    {[
                      ["stock", "Stock"], ["year", "Year"], ["make", "Make"], ["model", "Model"],
                      ["price", "Price"], ["odometer", "Odo"], ["color", "Color"], ["drivetrain", "Engine/Drivetrain"], ["certified", "Cert"],
                      ["vin", "VIN"],
                      ["type", "Type"], ["days", "Days"], ["recall", "Recall"],
                    ].map(([field, label]) => (
                      <th key={field} className="lg-th" onClick={() => toggleSort(field)}
                        style={{ textAlign: "left", padding: "7px 8px", color: "#9A9C9E", fontWeight: 600, borderBottom: "1px solid #3A3F49" }}>
                        {label} <SortIcon field={field} />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => (
                    <tr key={r.vin + r.scanDate + i} className="lg-row" style={{ background: i % 2 ? "#22252B" : "#24272E" }}>
                      <td className="lg-mono" style={{ padding: "6px 8px" }}>{r.stock}</td>
                      <td style={{ padding: "6px 8px" }}>{r.year}</td>
                      <td style={{ padding: "6px 8px" }}>{r.make}</td>
                      <td style={{ padding: "6px 8px" }}>{r.model}<div style={{ color: "#6B6D70", fontSize: 11 }}>{r.desc}</div></td>
                      <td className="lg-mono" style={{ padding: "6px 8px" }}>{r.price !== null ? `$${r.price.toLocaleString()}` : ""}</td>
                      <td className="lg-mono" style={{ padding: "6px 8px" }}>{r.odometer?.toLocaleString?.() ?? ""}</td>
                      <td style={{ padding: "6px 8px" }}>{r.color}</td>
                      <td style={{ padding: "6px 8px", color: "#9A9C9E" }}>{r.drivetrain}</td>
                      <td style={{ padding: "6px 8px" }}>{r.certified ? "Yes" : ""}</td>
                      <td className="lg-mono" style={{ padding: "6px 8px", fontSize: 11 }}>{r.vin}</td>
                      <td style={{ padding: "6px 8px", color: "#9A9C9E" }}>{r.type}</td>
                      <td className="lg-mono" style={{ padding: "6px 8px", color: "#9A9C9E" }}>{r.days ?? ""}</td>
                      <td style={{ padding: "6px 8px" }}>
                        {r.recall && (
                          <span style={{ color: /open/i.test(r.recall) ? "#C1502E" : "#3FA796", fontWeight: 600 }}>{r.recall}</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filtered.length === 0 && (
                <div style={{ textAlign: "center", color: "#6B6D70", padding: "24px", fontSize: 13 }}>No vehicles match these filters.</div>
              )}
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 12, color: "#9A9C9E" }}>
                {totalCount} vehicles — imported {scanDates[0] || "recently"}
              </span>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                {confirmingClear ? (
                  <>
                    <span style={{ fontSize: 12, color: "#9A9C9E" }}>Clear everything?</span>
                    <button onClick={clearAll} style={{
                      background: "#C1502E", border: "1px solid #C1502E", color: "#ECE7DC", borderRadius: 6,
                      padding: "7px 12px", fontSize: 12, cursor: "pointer",
                    }}>
                      Yes, clear it
                    </button>
                    <button onClick={() => setConfirmingClear(false)} style={{
                      background: "none", border: "1px solid #3A3F49", color: "#9A9C9E", borderRadius: 6,
                      padding: "7px 12px", fontSize: 12, cursor: "pointer",
                    }}>
                      Cancel
                    </button>
                  </>
                ) : (
                  <button onClick={() => setConfirmingClear(true)} style={{
                    background: "none", border: "1px solid #3A3F49", color: "#C1502E", borderRadius: 6,
                    padding: "7px 12px", fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
                  }}>
                    <Trash2 size={13} /> Clear current import
                  </button>
                )}
              </div>
            </div>
          </>
        )}
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>

      {["left", "right"].map((side) => (
        <div
          key={side}
          onTouchStart={(e) => {
            edgeTouch.current.startY = e.touches[0].clientY;
            edgeTouch.current.startScrollTop = scrollRef.current.scrollTop;
          }}
          onTouchMove={(e) => {
            const dy = e.touches[0].clientY - edgeTouch.current.startY;
            scrollRef.current.scrollTop = edgeTouch.current.startScrollTop - dy;
          }}
          style={{
            position: "fixed",
            top: 0,
            bottom: 0,
            [side]: 0,
            width: 48,
            zIndex: 40,
            background: "transparent",
            touchAction: "none",
          }}
        />
      ))}
    </div>
  );
}
