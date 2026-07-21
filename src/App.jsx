import { useState, useRef, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { Upload, Search, Trash2, ChevronUp, ChevronDown, Loader2, AlertTriangle, X, Gauge, FileSpreadsheet, Download, FolderOpen } from "lucide-react";

// ---- Google Drive integration ----
// Lets "Cloud Drive" list files from one specific folder and import
// whichever one is tapped, using the same pipeline as a local upload.
const GOOGLE_API_KEY = "AIzaSyB83DBxB4RhCKSVO204UAJwttYn9c5O7sM";
const GOOGLE_CLIENT_ID = "188213564865-kccradb2a14ghtr4hkjsvmlmvan6f1h8.apps.googleusercontent.com";
const GOOGLE_DRIVE_FOLDER_ID = "15FRZCsq1RHTn0Ur5B93bqv3MfKCC42kp";
const GOOGLE_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly";

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

// Added to every vehicle's price on import (not retroactive to already-saved
// data). Change the number here, or set to 0 to turn it off.
const PRICE_MARKUP = 2000;
function markUpPrice(price) {
  return price === null ? null : price + PRICE_MARKUP;
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

// Shortens make names that would otherwise force the (intentionally narrow,
// no-wrap) Make column to stretch wide — add more entries here as needed.
const MAKE_SHORTENINGS = {
  "mercedes-benz": "Mercedes",
  "volkswagen": "V.W.",
};
function shortenMake(make) {
  const m = (make || "").toString().trim();
  return MAKE_SHORTENINGS[m.toLowerCase()] || m;
}

// Shortens long trim/package words inside Model text (whole-word, case
// insensitive) so trims like "Atlas Cross Sport SE Technology" wrap to
// fewer lines in the (intentionally narrow) Model column. Add more pairs
// as you run into other long words.
const MODEL_WORD_SHORTENINGS = {
  "technology": "Tech",
  "premium": "Prem",
  "package": "Pkg",
  "performance": "Perf",
  "convenience": "Conv",
  "advanced": "Adv",
  "appearance": "Appr",
};
function shortenModelWords(model) {
  const s = (model || "").toString();
  return s.replace(/[A-Za-z]+/g, (word) => MODEL_WORD_SHORTENINGS[word.toLowerCase()] || word);
}

// Words that should always start a new line in the Model cell — e.g.
// "CX-30 Preferred Pkg" becomes "CX-30" / "Preferred Pkg" instead of
// wrapping wherever it happens to run out of room. Add more as needed.
const MODEL_LINE_BREAK_BEFORE = ["preferred"];
function modelLineBreakParts(model) {
  const words = (model || "").toString().split(" ");
  const breakIdx = words.findIndex((w, i) => i > 0 && MODEL_LINE_BREAK_BEFORE.includes(w.toLowerCase()));
  if (breakIdx === -1) return [model];
  return [words.slice(0, breakIdx).join(" "), words.slice(breakIdx).join(" ")];
}

function normalizeLegacyRow(r) {
  const model = shortenModelWords((r.md ?? "").toString().trim());
  return {
    stock: r.s ?? "",
    year: r.y ?? "",
    make: shortenMake(r.mk),
    model,
    type: classifyType(model),
    desc: r.de ?? "",
    status: (r.st ?? "").toString().toUpperCase().trim(),
    color: r.c ?? "",
    odometer: parseNum(r.o),
    vin: (r.v ?? "").toString().trim().toUpperCase(),
    days: parseNum(r.d),
    price: markUpPrice(parseMoney(r.p)),
    priceMarked: true,
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
  const shortMake = shortenMake(make);
  const shortModel = shortenModelWords(model);
  const classVal = (getField(row, "class") || "").toString();
  const certifiedVal = (getField(row, "certified") || "").toString();
  return {
    stock: getField(row, "stock #") ?? "",
    year,
    make: shortMake,
    model: shortModel,
    type: classVal.split(",")[0].trim() || "Other",
    desc: (getField(row, "body") || "").toString(),
    status: "", // not populated in this export — left blank rather than guessed
    color: (getField(row, "color") || "").toString(),
    odometer: parseNum(getField(row, "odometer")),
    vin: (getField(row, "vin") || "").toString().trim().toUpperCase(),
    days: null, // this export doesn't include a days-on-lot column
    price: markUpPrice(parseMoney(getField(row, "price / % mkt"))),
    priceMarked: true,
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
    make: shortenMake(make),
    model: shortenModelWords(model),
    type: classVal.split(",")[0].trim() || "Other",
    desc: (getField(row, "body") || "").toString(),
    status: "",
    color: interiorColor ? `${exteriorColor} / ${interiorColor}` : exteriorColor,
    odometer: parseNum(getField(row, "odometer")),
    vin: (getField(row, "vin") || "").toString().trim().toUpperCase(),
    days: daysSince(getField(row, "inventory date")),
    price: markUpPrice(parseMoney(getField(row, "price"))),
    priceMarked: true,
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
          maxHeight: 230, overflowY: "auto", zIndex: 50, padding: 6,
        }}>
          {selected.length > 0 && (
            <button type="button" onClick={() => onChange([])}
              style={{ background: "none", border: "none", color: "#F2A93B", fontSize: 13.5, cursor: "pointer", padding: "4px 6px", display: "block" }}>
              Clear {label.toLowerCase()}
            </button>
          )}
          {options.length === 0 && <div style={{ fontSize: 14, color: "#6B6D70", padding: "4px 6px" }}>No options yet</div>}
          {options.map((opt) => (
            <label key={opt} className="lg-row" style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 14.5, padding: "5px 6px", cursor: "pointer", borderRadius: 4 }}>
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
      if (!saved) return [];
      const parsed = JSON.parse(saved);
      // Records saved before the make/model shortening rules existed are
      // stuck with the old (longer) text forever unless re-processed here —
      // this keeps previously-imported data in sync with the current rules.
      const healed = parsed.map((r) => ({
        ...r,
        make: shortenMake(r.make),
        model: shortenModelWords(r.model),
        price: r.priceMarked ? r.price : markUpPrice(r.price),
        priceMarked: true,
      }));
      try {
        localStorage.setItem("lot-ledger-records", JSON.stringify(healed));
      } catch (e) {
        console.error("Couldn't save to this browser's storage", e);
      }
      return healed;
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
  const driveTokenClient = useRef(null);
  const driveAccessToken = useRef(null);
  const [driveScriptReady, setDriveScriptReady] = useState(false);
  const [showDrivePicker, setShowDrivePicker] = useState(false);
  const [driveFiles, setDriveFiles] = useState([]);
  const [driveLoading, setDriveLoading] = useState(false);
  const [driveError, setDriveError] = useState("");
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

  // Loads Google's Identity Services library once, on demand (first time the
  // Cloud Drive button is used) rather than on every page load.
  function loadDriveScript() {
    if (window.google?.accounts?.oauth2) {
      setDriveScriptReady(true);
      return;
    }
    const existing = document.getElementById("gis-script");
    if (existing) {
      existing.addEventListener("load", () => setDriveScriptReady(true));
      return;
    }
    const script = document.createElement("script");
    script.id = "gis-script";
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.onload = () => setDriveScriptReady(true);
    script.onerror = () => setDriveError("Couldn't load Google's sign-in library — check your connection.");
    document.head.appendChild(script);
  }

  async function fetchDriveFiles(accessToken) {
    setDriveLoading(true);
    setDriveError("");
    try {
      const q = encodeURIComponent(`'${GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false`);
      const fields = encodeURIComponent("files(id,name,mimeType,modifiedTime)");
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${q}&fields=${fields}&orderBy=modifiedTime desc&key=${GOOGLE_API_KEY}`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (!res.ok) throw new Error(`Drive returned ${res.status}`);
      const data = await res.json();
      const usable = (data.files || []).filter((f) =>
        /\.(csv|xlsx|xls)$/i.test(f.name) || f.mimeType === "application/vnd.google-apps.spreadsheet"
      );
      setDriveFiles(usable);
      if (usable.length === 0) setDriveError("No CSV/Excel files found in that folder.");
    } catch (e) {
      setDriveError("Couldn't load the folder — " + e.message);
    } finally {
      setDriveLoading(false);
    }
  }

  function openDrivePicker() {
    setShowDrivePicker(true);
    setDriveError("");
    if (!window.google?.accounts?.oauth2) {
      loadDriveScript();
      setDriveError("Still loading — try again in a second.");
      return;
    }
    if (!driveTokenClient.current) {
      driveTokenClient.current = window.google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: GOOGLE_DRIVE_SCOPE,
        callback: (resp) => {
          if (resp.error) {
            setDriveError("Sign-in was cancelled or failed.");
            return;
          }
          driveAccessToken.current = resp.access_token;
          fetchDriveFiles(resp.access_token);
        },
      });
    }
    driveTokenClient.current.requestAccessToken();
  }

  async function importDriveFile(accessToken, file) {
    setDriveLoading(true);
    setDriveError("");
    try {
      const isGoogleSheet = file.mimeType === "application/vnd.google-apps.spreadsheet";
      const url = isGoogleSheet
        ? `https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=${encodeURIComponent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}&key=${GOOGLE_API_KEY}`
        : `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media&key=${GOOGLE_API_KEY}`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
      if (!res.ok) throw new Error(`Drive returned ${res.status}`);
      const blob = await res.blob();
      const name = isGoogleSheet && !/\.xlsx$/i.test(file.name) ? `${file.name}.xlsx` : file.name;
      const localFile = new File([blob], name, { type: blob.type });
      handleFiles([localFile]);
      setShowDrivePicker(false);
    } catch (e) {
      setDriveError("Couldn't import that file — " + e.message);
    } finally {
      setDriveLoading(false);
    }
  }

  function selectDriveFile(file) {
    if (driveAccessToken.current) {
      importDriveFile(driveAccessToken.current, file);
    } else {
      setDriveError("Signed-in session expired — tap Cloud Drive again.");
    }
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

  useEffect(() => {
    loadDriveScript();
  }, []);

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
          border-radius: 6px; padding: 7px 10px; font-size: 15px; font-family: 'IBM Plex Sans', sans-serif;
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

      <div style={{ padding: "10px 12px 0", display: "flex", flexDirection: "column", gap: 10 }}>
        {/* Upload zone */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
          style={{
            border: `1.5px dashed ${dragOver ? "#F2A93B" : "#3A3F49"}`,
            borderRadius: 10,
            padding: "8px",
            display: "flex",
            alignItems: "stretch",
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

          <div
            onClick={openDrivePicker}
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
              fontSize: 15.5, color: "#ECE7DC", cursor: "pointer", padding: "6px 4px",
            }}
          >
            <Upload size={17} color="#F2A93B" style={{ flexShrink: 0 }} />
            Cloud Drive
          </div>

          <div style={{ width: 1, background: "#3A3F49", margin: "0 4px" }} />

          <div
            onClick={() => fileInputRef.current?.click()}
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
              fontSize: 15.5, color: "#ECE7DC", cursor: "pointer", padding: "6px 4px",
            }}
          >
            <FileSpreadsheet size={17} color="#F2A93B" style={{ flexShrink: 0 }} />
            Upload XLS
          </div>
        </div>
      </div>

      {showDrivePicker && (
        <div
          onClick={() => setShowDrivePicker(false)}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ background: "#24272E", borderRadius: 12, padding: 16, width: "100%", maxWidth: 420, maxHeight: "70vh", display: "flex", flexDirection: "column" }}
          >
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 16, fontWeight: 600 }}>
                <FolderOpen size={18} color="#F2A93B" />
                Cloud Drive
              </div>
              <X size={18} style={{ cursor: "pointer", color: "#9A9C9E" }} onClick={() => setShowDrivePicker(false)} />
            </div>

            {driveLoading && (
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#9A9C9E", fontSize: 14, padding: "16px 0" }}>
                <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> Working…
              </div>
            )}

            {driveError && !driveLoading && (
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#C1502E", fontSize: 13.5, padding: "8px 0" }}>
                <AlertTriangle size={14} /> {driveError}
              </div>
            )}

            {!driveLoading && driveFiles.length > 0 && (
              <div style={{ overflowY: "auto", display: "flex", flexDirection: "column", gap: 4 }}>
                {driveFiles.map((f) => (
                  <button
                    key={f.id}
                    onClick={() => selectDriveFile(f)}
                    style={{
                      display: "flex", alignItems: "center", gap: 8, background: "#2C303A", border: "1px solid #3A3F49",
                      borderRadius: 8, padding: "10px 12px", color: "#ECE7DC", fontSize: 14.5, textAlign: "left", cursor: "pointer",
                    }}
                  >
                    <FileSpreadsheet size={15} color="#F2A93B" style={{ flexShrink: 0 }} />
                    {f.name}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Header */}
      <div style={{ background: "#000000", padding: "20px 16px 0", textAlign: "center" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
          <Gauge size={22} color="#E8C468" />
          <h1 className="lg-display" style={{ fontSize: 21, fontWeight: 700, margin: 0, letterSpacing: 2, color: "#E8C468", textTransform: "uppercase" }}>
            The Lot Ledger
          </h1>
        </div>
        <div style={{ fontSize: 11, color: "#7A7565", marginTop: 6 }}>
          designed by <b style={{ color: "#FFE29A" }}>Jeff Patrick</b>
        </div>
        <div style={{ height: 14 }} />
      </div>

      <div style={{ padding: "10px 12px 10px 8px", display: "flex", flexDirection: "column", gap: 8 }}>
        {/* Processing queue */}
        {queue.some((it) => it.status !== "done") && (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {queue.filter((it) => it.status !== "done").map((it) => (
              <div key={it.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14.5, background: "#24272E", borderRadius: 6, padding: "7px 12px" }}>
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
          <div style={{ textAlign: "center", color: "#6B6D70", fontSize: 15, padding: "30px 0" }}>
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
                style={{ marginBottom: 4, padding: "6px 10px", textAlign: "center", maxWidth: 640, margin: "0 auto 4px", display: "block" }}
                placeholder="Search anything (stock, VIN, model, color, price…)"
                value={filters.search}
                onChange={(e) => setFilters((f) => ({ ...f, search: e.target.value }))}
              />
              <div style={{ textAlign: "center", fontSize: 13, color: "#9A9C9E", marginBottom: 8 }}>
                {filtered.length}/{totalCount} vehicles
              </div>

              <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: showFilters ? 4 : -4, marginTop: -4, paddingRight: 16 }}>
                <button onClick={() => setShowFilters((s) => !s)} style={{ background: "none", border: "none", color: "#9A9C9E", fontSize: 14, cursor: "pointer", padding: 18, margin: -18 }}>
                  {showFilters ? "Hide" : "Show"}
                </button>
              </div>

              {showFilters && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8, alignItems: "start", maxWidth: 640, margin: "0 auto" }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    <MultiSelect label="Make" options={makes} selected={filters.make}
                      onChange={(vals) => setFilters((f) => ({ ...f, make: vals }))} />
                    <MultiSelect label="Model" options={models} selected={filters.model}
                      onChange={(vals) => setFilters((f) => ({ ...f, model: vals }))} />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    <MultiSelect label="Type" options={types} selected={filters.type}
                      onChange={(vals) => setFilters((f) => ({ ...f, type: vals }))} />
                    <input className="lg-input" type="number" placeholder="Max odometer" value={filters.odoMax}
                      onChange={(e) => setFilters((f) => ({ ...f, odoMax: e.target.value }))} />
                  </div>
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
                    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 15 }}>
                      <input type="checkbox" checked={filters.certifiedOnly}
                        onChange={(e) => setFilters((f) => ({ ...f, certifiedOnly: e.target.checked }))} />
                      Certified only
                    </label>
                    <button
                      onClick={() => setFilters({
                        search: "", make: [], model: [], type: [], status: [], recall: [], scanDate: [],
                        yearMin: "", yearMax: "", priceMin: "", priceMax: "", odoMax: "", certifiedOnly: false,
                      })}
                      style={{ background: "none", border: "1px solid #3A3F49", color: "#9A9C9E", borderRadius: 6, padding: "5px 10px", fontSize: 13.5, cursor: "pointer" }}
                    >
                      Clear search
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Table */}
            <div ref={tableRef} className="lg-scroll" style={{ background: "#24272E", borderRadius: 10, overflow: "auto", maxHeight: "60vh", overscrollBehavior: "contain", WebkitOverflowScrolling: "touch" }}>
              <table style={{ width: "max-content", borderCollapse: "collapse", fontSize: 14.5 }}>
                <colgroup>
                  <col style={{ width: "44px" }} />  {/* Stock */}
                  <col style={{ width: "34px" }} />  {/* Year */}
                  <col style={{ width: "40px" }} />  {/* Make */}
                  <col style={{ width: "100px" }} /> {/* Model */}
                  <col style={{ width: "62px" }} />  {/* Price */}
                  <col style={{ width: "50px" }} />  {/* Odo */}
                  <col style={{ width: "62px" }} />  {/* Color */}
                  <col style={{ width: "58px" }} />  {/* Engine/Drivetrain */}
                  <col style={{ width: "28px" }} />  {/* Cert */}
                  <col style={{ width: "90px" }} />  {/* VIN */}
                  <col style={{ width: "42px" }} />  {/* Type */}
                  <col style={{ width: "32px" }} />  {/* Days */}
                  <col style={{ width: "48px" }} />  {/* Recall */}
                </colgroup>
                <thead>
                  <tr style={{ position: "sticky", top: 0, background: "#1F2228", zIndex: 1 }}>
                    {[
                      ["stock", "Stock"], ["year", "Year"], ["make", "Make"], ["model", "Model"],
                      ["price", "Price"], ["odometer", "Odo"], ["color", "Color"], ["drivetrain", "Engine/Drivetrain"], ["certified", "Cert"],
                      ["vin", "VIN"],
                      ["type", "Type"], ["days", "Days"], ["recall", "Recall"],
                    ].map(([field, label]) => (
                      <th key={field} className="lg-th" onClick={() => toggleSort(field)}
                        style={{ textAlign: "left", padding: "7px 5px", color: "#9A9C9E", fontWeight: 600, borderBottom: "1px solid #3A3F49" }}>
                        {label} <SortIcon field={field} />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => (
                    <tr key={r.vin + r.scanDate + i} className="lg-row" style={{ background: i % 2 ? "#22252B" : "#24272E" }}>
                      <td className="lg-mono" style={{ padding: "4px 5px" }}>{r.stock}</td>
                      <td style={{ padding: "4px 5px" }}>{r.year}</td>
                      <td style={{ padding: "4px 5px", whiteSpace: "nowrap" }}>{r.make}</td>
                      <td style={{ padding: "4px 5px" }}>
                        {modelLineBreakParts(r.model).map((part, pi) => (
                          <span key={pi}>{pi > 0 && <br />}{part}</span>
                        ))}
                        <div style={{ color: "#6B6D70", fontSize: 13 }}>{r.desc}</div>
                      </td>
                      <td className="lg-mono" style={{ padding: "4px 5px" }}>{r.price !== null ? `$${r.price.toLocaleString()}` : ""}</td>
                      <td className="lg-mono" style={{ padding: "4px 5px" }}>{r.odometer?.toLocaleString?.() ?? ""}</td>
                      <td style={{ padding: "4px 5px" }}>
                        {r.color && r.color.includes(" / ") ? (
                          r.color.split(" / ").map((c, ci) => <div key={ci}>{c}</div>)
                        ) : (
                          r.color
                        )}
                      </td>
                      <td style={{ padding: "4px 5px", color: "#9A9C9E" }}>{r.drivetrain}</td>
                      <td style={{ padding: "4px 5px" }}>{r.certified ? "Yes" : ""}</td>
                      <td className="lg-mono" style={{ padding: "4px 5px", fontSize: 12, wordBreak: "break-all" }}>{r.vin}</td>
                      <td style={{ padding: "4px 5px", color: "#9A9C9E" }}>{r.type}</td>
                      <td className="lg-mono" style={{ padding: "4px 5px", color: "#9A9C9E" }}>{r.days ?? ""}</td>
                      <td style={{ padding: "4px 5px" }}>
                        {r.recall && (
                          <span style={{ color: /open/i.test(r.recall) ? "#C1502E" : "#3FA796", fontWeight: 600 }}>{r.recall}</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filtered.length === 0 && (
                <div style={{ textAlign: "center", color: "#6B6D70", padding: "24px", fontSize: 15 }}>No vehicles match these filters.</div>
              )}
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 14, color: "#9A9C9E" }}>
                {totalCount} vehicles — imported {scanDates[0] || "recently"}
              </span>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                {confirmingClear ? (
                  <>
                    <span style={{ fontSize: 14, color: "#9A9C9E" }}>Clear everything?</span>
                    <button onClick={clearAll} style={{
                      background: "#C1502E", border: "1px solid #C1502E", color: "#ECE7DC", borderRadius: 6,
                      padding: "7px 12px", fontSize: 14, cursor: "pointer",
                    }}>
                      Yes, clear it
                    </button>
                    <button onClick={() => setConfirmingClear(false)} style={{
                      background: "none", border: "1px solid #3A3F49", color: "#9A9C9E", borderRadius: 6,
                      padding: "7px 12px", fontSize: 14, cursor: "pointer",
                    }}>
                      Cancel
                    </button>
                  </>
                ) : (
                  <button onClick={() => setConfirmingClear(true)} style={{
                    background: "none", border: "1px solid #3A3F49", color: "#C1502E", borderRadius: 6,
                    padding: "7px 12px", fontSize: 14, cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
                  }}>
                    <Trash2 size={13} /> Clear inventory
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
