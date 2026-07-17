import { useState, useEffect, useRef, useMemo } from "react";
import { Upload, Search, Trash2, ChevronUp, ChevronDown, Loader2, AlertTriangle, X, Gauge } from "lucide-react";

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
  if (v === null || v === undefined) return null;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}
function parseNum(v) {
  if (v === null || v === undefined) return null;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = () => reject(new Error("Could not read file"));
    r.readAsDataURL(file);
  });
}

// Recover as many complete row-objects as possible even if the JSON got
// truncated (e.g. a very long sheet cut off by the token limit).
function salvageRows(text) {
  const rows = [];
  const matches = text.match(/\{[^{}]*\}/g) || [];
  for (const m of matches) {
    try {
      rows.push(JSON.parse(m));
    } catch (e) {
      // skip the one broken fragment (usually the last, truncated row)
    }
  }
  return rows;
}

function normalizeRow(r, scanDate) {
  return {
    scanDate: scanDate || "unknown",
    stock: r.s ?? "",
    year: r.y ?? "",
    make: (r.mk ?? "").toString().trim(),
    model: (r.md ?? "").toString().trim(),
    desc: r.de ?? "",
    status: (r.st ?? "").toString().toUpperCase().trim(),
    color: r.c ?? "",
    odometer: parseNum(r.o),
    vin: (r.v ?? "").toString().trim().toUpperCase(),
    days: parseNum(r.d),
    price: parseMoney(r.p),
    certified: !!r.ce,
  };
}

const STORAGE_KEY = "lot-ledger-records";

export default function LotLedger() {
  const [records, setRecords] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [queue, setQueue] = useState([]); // {name, status, error}
  const [dragOver, setDragOver] = useState(false);
  const [sortField, setSortField] = useState("price");
  const [sortDir, setSortDir] = useState("desc");
  const [showFilters, setShowFilters] = useState(true);
  const fileInputRef = useRef(null);

  const [filters, setFilters] = useState({
    search: "",
    make: "",
    status: "",
    scanDate: "",
    yearMin: "",
    yearMax: "",
    priceMin: "",
    priceMax: "",
    odoMax: "",
    certifiedOnly: false,
  });

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) setRecords(JSON.parse(res.value));
      } catch (e) {
        // no saved data yet
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  async function persist(next) {
    setRecords(next);
    try {
      await window.storage.set(STORAGE_KEY, JSON.stringify(next), false);
    } catch (e) {
      console.error("Storage save failed", e);
    }
  }

  async function processFile(file) {
    const qid = file.name + "-" + Date.now();
    setQueue((q) => [...q, { id: qid, name: file.name, status: "reading" }]);
    try {
      const base64 = await fileToBase64(file);
      const mediaType = file.type || "image/jpeg";
      setQueue((q) => q.map((it) => (it.id === qid ? { ...it, status: "extracting" } : it)));

      const prompt = `This image is a page from a used-vehicle inventory report. Extract every data row from the table.
Return ONLY minified JSON, no markdown fences, no explanation, in exactly this shape:
{"rd":"<the report date shown at top, e.g. 07/17/2026, or null if not visible>","rows":[{"s":"<stock/order>","y":<year as number>,"mk":"<make>","md":"<model>","de":"<model desc>","st":"<status>","c":"<exterior color, trimmed of trailing ...>","o":<odometer as number>,"v":"<vin>","d":<days as number>,"p":<list price as number, no commas or $>,"ce":<true if Certified column says Yes, else false>},...]}
Use short keys exactly as shown. Include EVERY row visible in the image, in order. Keep strings short. Output nothing but the JSON object.`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          messages: [
            {
              role: "user",
              content: [
                { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
                { type: "text", text: prompt },
              ],
            },
          ],
        }),
      });
      const data = await response.json();
      const text = (data.content || [])
        .map((b) => (b.type === "text" ? b.text : ""))
        .join("")
        .replace(/```json|```/g, "")
        .trim();

      let rd = null;
      let rawRows = [];
      let truncated = false;
      try {
        const parsed = JSON.parse(text);
        rd = parsed.rd || null;
        rawRows = parsed.rows || [];
      } catch (e) {
        // likely truncated by the token limit — salvage what we can
        const rdMatch = text.match(/"rd"\s*:\s*"([^"]*)"/);
        rd = rdMatch ? rdMatch[1] : null;
        rawRows = salvageRows(text);
        truncated = true;
      }

      if (rawRows.length === 0) {
        setQueue((q) =>
          q.map((it) => (it.id === qid ? { ...it, status: "error", error: "No rows recognized in image" } : it))
        );
        return;
      }

      const scanDate = rd || new Date().toLocaleDateString();
      const newRows = rawRows.map((r) => normalizeRow(r, scanDate));

      setRecords((prev) => {
        const map = new Map(prev.map((r) => [r.vin + "|" + r.scanDate, r]));
        for (const row of newRows) {
          map.set(row.vin + "|" + row.scanDate, row);
        }
        const next = Array.from(map.values());
        persist(next);
        return next;
      });

      setQueue((q) =>
        q.map((it) =>
          it.id === qid
            ? {
                ...it,
                status: "done",
                count: newRows.length,
                warning: truncated ? "Sheet may be long — some trailing rows could be missing. Consider photographing it in two halves next time." : null,
              }
            : it
        )
      );
    } catch (e) {
      setQueue((q) => q.map((it) => (it.id === qid ? { ...it, status: "error", error: e.message } : it)));
    }
  }

  function handleFiles(fileList) {
    const files = Array.from(fileList).filter((f) => f.type.startsWith("image/"));
    files.forEach(processFile);
  }

  async function clearAll() {
    if (!window.confirm("Clear all scanned inventory history? This can't be undone.")) return;
    await persist([]);
    setQueue([]);
  }

  const makes = useMemo(() => Array.from(new Set(records.map((r) => r.make).filter(Boolean))).sort(), [records]);
  const statuses = useMemo(() => Array.from(new Set(records.map((r) => r.status).filter(Boolean))).sort(), [records]);
  const scanDates = useMemo(() => Array.from(new Set(records.map((r) => r.scanDate).filter(Boolean))).sort().reverse(), [records]);

  const filtered = useMemo(() => {
    let out = records.filter((r) => {
      if (filters.search) {
        const s = filters.search.toLowerCase();
        const hit =
          r.stock.toString().toLowerCase().includes(s) ||
          r.vin.toLowerCase().includes(s) ||
          r.model.toLowerCase().includes(s) ||
          r.desc.toLowerCase().includes(s);
        if (!hit) return false;
      }
      if (filters.make && r.make !== filters.make) return false;
      if (filters.status && r.status !== filters.status) return false;
      if (filters.scanDate && r.scanDate !== filters.scanDate) return false;
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

  const statusCounts = useMemo(() => {
    const c = {};
    for (const r of records) c[r.status] = (c[r.status] || 0) + 1;
    return c;
  }, [records]);
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
    <div style={{ minHeight: "100vh", background: "#1B1D22", color: "#ECE7DC", fontFamily: "'IBM Plex Sans', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .lg-mono { font-family: 'IBM Plex Mono', monospace; }
        .lg-display { font-family: 'Space Grotesk', sans-serif; }
        .lg-input {
          background: #14161A; border: 1px solid #3A3F49; color: #ECE7DC;
          border-radius: 6px; padding: 7px 10px; font-size: 13px; font-family: 'IBM Plex Sans', sans-serif;
          outline: none; width: 100%; box-sizing: border-box;
        }
        .lg-input:focus { border-color: #F2A93B; }
        .lg-input::placeholder { color: #6B6D70; }
        .lg-th { cursor: pointer; user-select: none; white-space: nowrap; }
        .lg-th:hover { color: #F2A93B; }
        .lg-row:hover { background: #2C303A !important; }
        ::-webkit-scrollbar { height: 10px; width: 10px; }
        ::-webkit-scrollbar-track { background: #1B1D22; }
        ::-webkit-scrollbar-thumb { background: #3A3F49; border-radius: 5px; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }
      `}</style>

      {/* Header */}
      <div style={{ borderBottom: "1px solid #3A3F49", padding: "20px 28px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <Gauge size={22} color="#F2A93B" />
          <h1 className="lg-display" style={{ fontSize: 22, fontWeight: 700, margin: 0, letterSpacing: 0.2 }}>
            The Lot Ledger
          </h1>
        </div>
        <div style={{ fontSize: 12.5, color: "#9A9C9E", marginTop: 4 }}>
          Scan daily inventory sheets. Search everything you've ever scanned.
        </div>
      </div>

      <div style={{ padding: "22px 28px", display: "flex", flexDirection: "column", gap: 20 }}>
        {/* Upload zone */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: `1.5px dashed ${dragOver ? "#F2A93B" : "#3A3F49"}`,
            borderRadius: 10,
            padding: "22px",
            textAlign: "center",
            cursor: "pointer",
            background: dragOver ? "#24272E" : "transparent",
            transition: "border-color 0.15s, background 0.15s",
          }}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            multiple
            style={{ display: "none" }}
            onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }}
          />
          <Upload size={20} color="#F2A93B" style={{ marginBottom: 6 }} />
          <div style={{ fontSize: 13.5 }}>Drop photos of today's inventory sheet here, or click to choose files</div>
          <div style={{ fontSize: 11.5, color: "#6B6D70", marginTop: 3 }}>Each page is read automatically and added to your history</div>
        </div>

        {/* Processing queue */}
        {queue.length > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {queue.map((it) => (
              <div key={it.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, background: "#24272E", borderRadius: 6, padding: "7px 12px" }}>
                {it.status === "reading" || it.status === "extracting" ? (
                  <Loader2 size={14} className="lg-mono" style={{ animation: "spin 1s linear infinite" }} />
                ) : it.status === "error" ? (
                  <AlertTriangle size={14} color="#C1502E" />
                ) : (
                  <span style={{ color: "#3FA796" }}>✓</span>
                )}
                <span style={{ flex: 1 }}>{it.name}</span>
                <span style={{ color: "#9A9C9E" }}>
                  {it.status === "reading" && "reading…"}
                  {it.status === "extracting" && "extracting rows…"}
                  {it.status === "done" && `${it.count} rows added`}
                  {it.status === "error" && it.error}
                </span>
                {it.warning && <span style={{ color: "#F2A93B" }}>{it.warning}</span>}
                <X size={13} style={{ cursor: "pointer", color: "#6B6D70" }} onClick={() => setQueue((q) => q.filter((x) => x.id !== it.id))} />
              </div>
            ))}
          </div>
        )}

        {loaded && totalCount === 0 && queue.length === 0 && (
          <div style={{ textAlign: "center", color: "#6B6D70", fontSize: 13, padding: "30px 0" }}>
            No vehicles scanned yet — upload a sheet above to get started.
          </div>
        )}

        {totalCount > 0 && (
          <>
            {/* Status gauge */}
            <div style={{ background: "#24272E", borderRadius: 10, padding: "14px 18px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#9A9C9E", marginBottom: 8 }}>
                <span>{totalCount} vehicles across {scanDates.length} scan{scanDates.length !== 1 ? "s" : ""}</span>
                <span>{filtered.length} matching current filters</span>
              </div>
              <div style={{ display: "flex", height: 10, borderRadius: 5, overflow: "hidden" }}>
                {Object.entries(statusCounts).map(([st, ct]) => (
                  <div key={st} title={`${st}: ${ct}`} style={{ width: `${(ct / totalCount) * 100}%`, background: statusColor(st) }} />
                ))}
              </div>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginTop: 8, fontSize: 11.5 }}>
                {Object.entries(statusCounts).map(([st, ct]) => (
                  <div key={st} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                    <span style={{ width: 8, height: 8, borderRadius: 2, background: statusColor(st), display: "inline-block" }} />
                    <span style={{ color: "#9A9C9E" }}>{st} ({ct})</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Filters */}
            <div style={{ background: "#24272E", borderRadius: 10, padding: "16px 18px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: showFilters ? 12 : 0 }}>
                <div className="lg-display" style={{ fontSize: 13.5, fontWeight: 700, display: "flex", alignItems: "center", gap: 6 }}>
                  <Search size={14} color="#F2A93B" /> Search & filter
                </div>
                <button onClick={() => setShowFilters((s) => !s)} style={{ background: "none", border: "none", color: "#9A9C9E", fontSize: 12, cursor: "pointer" }}>
                  {showFilters ? "Hide" : "Show"}
                </button>
              </div>
              {showFilters && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 10 }}>
                  <input className="lg-input" placeholder="Stock #, VIN, or model" value={filters.search}
                    onChange={(e) => setFilters((f) => ({ ...f, search: e.target.value }))} />
                  <select className="lg-input" value={filters.make} onChange={(e) => setFilters((f) => ({ ...f, make: e.target.value }))}>
                    <option value="">All makes</option>
                    {makes.map((m) => <option key={m} value={m}>{m}</option>)}
                  </select>
                  <select className="lg-input" value={filters.status} onChange={(e) => setFilters((f) => ({ ...f, status: e.target.value }))}>
                    <option value="">All statuses</option>
                    {statuses.map((s) => <option key={s} value={s}>{s}</option>)}
                  </select>
                  <select className="lg-input" value={filters.scanDate} onChange={(e) => setFilters((f) => ({ ...f, scanDate: e.target.value }))}>
                    <option value="">All scan dates</option>
                    {scanDates.map((d) => <option key={d} value={d}>{d}</option>)}
                  </select>
                  <input className="lg-input" type="number" placeholder="Year min" value={filters.yearMin}
                    onChange={(e) => setFilters((f) => ({ ...f, yearMin: e.target.value }))} />
                  <input className="lg-input" type="number" placeholder="Year max" value={filters.yearMax}
                    onChange={(e) => setFilters((f) => ({ ...f, yearMax: e.target.value }))} />
                  <input className="lg-input" type="number" placeholder="Price min ($)" value={filters.priceMin}
                    onChange={(e) => setFilters((f) => ({ ...f, priceMin: e.target.value }))} />
                  <input className="lg-input" type="number" placeholder="Price max ($)" value={filters.priceMax}
                    onChange={(e) => setFilters((f) => ({ ...f, priceMax: e.target.value }))} />
                  <input className="lg-input" type="number" placeholder="Max odometer" value={filters.odoMax}
                    onChange={(e) => setFilters((f) => ({ ...f, odoMax: e.target.value }))} />
                  <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13 }}>
                    <input type="checkbox" checked={filters.certifiedOnly}
                      onChange={(e) => setFilters((f) => ({ ...f, certifiedOnly: e.target.checked }))} />
                    Certified only
                  </label>
                </div>
              )}
            </div>

            {/* Table */}
            <div style={{ background: "#24272E", borderRadius: 10, overflow: "auto", maxHeight: 560 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr style={{ position: "sticky", top: 0, background: "#1F2228", zIndex: 1 }}>
                    {[
                      ["stock", "Stock"], ["year", "Year"], ["make", "Make"], ["model", "Model"],
                      ["status", "Status"], ["color", "Color"], ["odometer", "Odo"], ["vin", "VIN"],
                      ["days", "Days"], ["price", "Price"], ["certified", "Cert"], ["scanDate", "Scanned"],
                    ].map(([field, label]) => (
                      <th key={field} className="lg-th" onClick={() => toggleSort(field)}
                        style={{ textAlign: "left", padding: "9px 10px", color: "#9A9C9E", fontWeight: 600, borderBottom: "1px solid #3A3F49" }}>
                        {label} <SortIcon field={field} />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => (
                    <tr key={r.vin + r.scanDate + i} className="lg-row" style={{ background: i % 2 ? "#22252B" : "#24272E" }}>
                      <td className="lg-mono" style={{ padding: "8px 10px" }}>{r.stock}</td>
                      <td style={{ padding: "8px 10px" }}>{r.year}</td>
                      <td style={{ padding: "8px 10px" }}>{r.make}</td>
                      <td style={{ padding: "8px 10px" }}>{r.model}<div style={{ color: "#6B6D70", fontSize: 11 }}>{r.desc}</div></td>
                      <td style={{ padding: "8px 10px" }}>
                        <span style={{ color: statusColor(r.status), fontWeight: 600 }}>{r.status}</span>
                      </td>
                      <td style={{ padding: "8px 10px" }}>{r.color}</td>
                      <td className="lg-mono" style={{ padding: "8px 10px" }}>{r.odometer?.toLocaleString?.() ?? ""}</td>
                      <td className="lg-mono" style={{ padding: "8px 10px", fontSize: 11 }}>{r.vin}</td>
                      <td className="lg-mono" style={{ padding: "8px 10px" }}>{r.days}</td>
                      <td className="lg-mono" style={{ padding: "8px 10px" }}>{r.price !== null ? `$${r.price.toLocaleString()}` : ""}</td>
                      <td style={{ padding: "8px 10px" }}>{r.certified ? "Yes" : ""}</td>
                      <td className="lg-mono" style={{ padding: "8px 10px", color: "#9A9C9E" }}>{r.scanDate}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filtered.length === 0 && (
                <div style={{ textAlign: "center", color: "#6B6D70", padding: "24px", fontSize: 13 }}>No vehicles match these filters.</div>
              )}
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <button onClick={clearAll} style={{
                background: "none", border: "1px solid #3A3F49", color: "#C1502E", borderRadius: 6,
                padding: "7px 12px", fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
              }}>
                <Trash2 size={13} /> Clear all scanned history
              </button>
            </div>
          </>
        )}
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
