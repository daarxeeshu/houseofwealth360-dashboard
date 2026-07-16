import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import {
  Search, LayoutDashboard, Users, Lightbulb, UserCog, FileText,
  Upload, PieChart as PieIcon, Settings, Download, ArrowUpDown,
  Building2, ChevronLeft, ChevronRight, X, Moon, Sun, Wallet,
  TrendingUp, Layers, Landmark, Sparkles, CheckCircle2, AlertTriangle, CheckSquare,
} from "lucide-react";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, AreaChart, Area, Treemap,
} from "recharts";
import * as XLSX from "xlsx";

/* ============================================================
   HOUSE OF WEALTH FINANCIAL SERVICES — 360 Portfolio Console
   PAN is the primary unique identifier. Rows never stand alone.
   ============================================================ */

const GOLD = "#D4AF37";
const PALETTE = ["#D4AF37", "#8C7A3F", "#5F5A4E", "#B9A46A", "#3E3A33", "#9C8B5A", "#6E6455", "#CFC09A"];

const COLUMN_MAP = {
  "account no": "folio", accountno: "folio", "account number": "folio", panno: "pan", pan: "pan",
  folio: "folio", "folio no": "folio", "folio number": "folio", folionumber: "folio", "folio no.": "folio",
  cpcode: "cpCode", amccode: "amcCode", fund: "fund", "scheme code": "schemeCode",
  "scheme name": "schemeName", "asset type": "assetType", clasification: "classification",
  classification: "classification", category: "category", "client code": "clientCode",
  "client name": "clientName", segment: "segment", "manager code": "managerCode",
  "manager name": "managerName", "employee id": "employeeId", department: "department",
  "emp segment": "empSegment", "emp branch": "branch", "emp region": "region",
  "emp zone": "zone", hold: "units", nav: "nav", aum: "aum", date: "date",
  rtacode: "rtaCode", "reporting name": "reportingName", "reporting code": "reportingCode",
  recommendedscheme: "recommendedScheme", days: "days",
  "sourcing emp sap code": "sourcingCode", "sourcing employee name": "sourcingEmployee",
  "new fpd rm sap": "rmSap", "new fpd rm name": "rmName",
  "new fpd rm client zone": "rmZone", "asset type group": "assetGroup",
};

const norm = (s) => String(s ?? "").trim().toLowerCase().replace(/\s+/g, " ");
const num = (v) => { const n = parseFloat(String(v ?? "").replace(/[₹,\s]/g, "")); return isFinite(n) ? n : 0; };
const inr = (n) => {
  if (!isFinite(n)) return "₹0";
  if (n >= 1e7) return `₹${(n / 1e7).toFixed(2)} Cr`;
  if (n >= 1e5) return `₹${(n / 1e5).toFixed(2)} L`;
  return `₹${n.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`;
};
const inrFull = (n) => `₹${(n || 0).toLocaleString("en-IN", { maximumFractionDigits: 2 })}`;

// Shape a client's holdings for export, folio included, in a readable column order.
const exportHoldingRows = (inv) => inv.holdings.map((h) => ({
  "Client Name": inv.clientName, PAN: inv.pan,
  AMC: h.fund, "Scheme Name": h.schemeName, "Folio No": h.folio,
  "Scheme Code": h.schemeCode, "Asset Type": h.assetType, Category: h.category,
  Units: num(h.units), NAV: num(h.nav), AUM: num(h.aum),
  "Investment Date": h.date, "Recommended": h.recommendedScheme, Days: num(h.days),
}));

/* --------------------------- Demo data --------------------------- */
const AMCS = ["HDFC", "SBI", "ICICI Prudential", "Axis", "Tata", "Nippon India", "Kotak", "Mirae Asset", "Aditya Birla", "UTI"];
const SCHEMES = [
  ["Flexicap Fund", "Equity", "Flexi Cap"], ["Bluechip Fund", "Equity", "Large Cap"],
  ["Small Cap Fund", "Equity", "Small Cap"], ["Midcap Opportunities", "Equity", "Mid Cap"],
  ["Liquid Fund", "Debt", "Liquid"], ["Corporate Bond Fund", "Debt", "Corporate Bond"],
  ["Short Duration Fund", "Debt", "Short Duration"], ["Balanced Advantage", "Hybrid", "Dynamic Asset"],
  ["Aggressive Hybrid", "Hybrid", "Aggressive Hybrid"], ["Gold ETF FoF", "Others", "Commodity"],
  ["ELSS Tax Saver", "Equity", "ELSS"], ["Gilt Fund", "Debt", "Gilt"],
];
const NAMES = ["Arjun Mehta", "Priya Raghavan", "Rohit Bansal", "Sneha Kulkarni", "Vikram Iyer", "Ananya Desai", "Karan Malhotra", "Meera Nair", "Siddharth Rao", "Nisha Chopra", "Aditya Verma", "Kavya Menon", "Rahul Saxena", "Divya Pillai", "Manish Gupta", "Ritu Aggarwal", "Sanjay Kapoor", "Pooja Reddy", "Tarun Joshi", "Ishita Sharma", "Nikhil Bhatt", "Shreya Ghosh", "Amit Trivedi", "Lakshmi Krishnan"];
const MANAGERS = ["Deepak Ranjan", "Farah Qureshi", "Gaurav Sethi", "Hema Vasudevan", "Imran Sheikh"];
const BRANCHES = ["Bandra Kurla", "Connaught Place", "Koramangala", "Anna Salai", "Salt Lake"];
const REGIONS = ["West", "North", "South", "East"];
const ZONES = ["Zone A", "Zone B", "Zone C"];
const DEPTS = ["Wealth", "Private Client", "Retail Advisory"];

function seedRows() {
  let s = 20240117;
  const rnd = () => ((s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const rows = [];
  for (let i = 0; i < 96; i++) {
    const name = NAMES[i % NAMES.length] + (i >= NAMES.length ? ` ${Math.floor(i / NAMES.length) + 1}` : "");
    const pan = `${["A", "B", "C", "E", "K", "M", "P", "R"][i % 8]}${["Y", "X", "Z", "T"][i % 4]}TPR${String(1000 + i * 7).slice(0, 4)}${"ABCDEFGH"[i % 8]}`;
    const mgr = pick(MANAGERS), br = pick(BRANCHES), rg = pick(REGIONS), zn = pick(ZONES), dp = pick(DEPTS);
    const n = 1 + Math.floor(rnd() * 12);
    const singleAmc = rnd() < 0.16;
    const debtHeavy = rnd() < 0.14;
    const amc0 = pick(AMCS);
    for (let j = 0; j < n; j++) {
      let sch = pick(SCHEMES);
      if (debtHeavy && rnd() < 0.85) sch = pick(SCHEMES.filter((x) => x[1] === "Debt"));
      const amc = singleAmc ? amc0 : pick(AMCS);
      const nav = 12 + rnd() * 620;
      const units = 20 + rnd() * 9000;
      const d = new Date(2024, Math.floor(rnd() * 12), 1 + Math.floor(rnd() * 27));
      rows.push({
        pan, clientName: name, cpCode: `CP${100 + i}`,
        folio: `${1000000 + i * 3137 + j * 91}`,
        amcCode: amc.slice(0, 3).toUpperCase(), fund: `${amc} Mutual Fund`,
        schemeCode: `${amc.slice(0, 3).toUpperCase()}${1000 + j * 13 + i}`,
        schemeName: `${amc} ${sch[0]}`, assetType: sch[1], assetGroup: sch[1],
        classification: sch[1] === "Equity" ? "Growth" : "Income", category: sch[2],
        clientCode: `C${9000 + i}`, segment: "Retail", managerCode: `M${MANAGERS.indexOf(mgr) + 1}`,
        managerName: mgr, employeeId: `EMP${200 + i}`, department: dp, empSegment: "Advisory",
        branch: br, region: rg, zone: zn, units, nav,
        aum: Math.round(units * nav), date: d.toISOString().slice(0, 10),
        rtaCode: rnd() < 0.5 ? "CAMS" : "KFin", reportingName: mgr, reportingCode: `R${i}`,
        recommendedScheme: rnd() < 0.45 ? "Yes" : "No",
        days: 30 + Math.floor(rnd() * 900), sourcingEmployee: mgr, rmName: mgr, rmZone: zn,
      });
    }
  }
  return rows;
}

/* ---------------------- PAN consolidation ---------------------- */
function buildInvestors(rows) {
  const map = new Map();
  for (const r of rows) {
    const pan = String(r.pan || "").trim().toUpperCase();
    if (!pan) continue;
    if (!map.has(pan)) map.set(pan, { pan, clientName: r.clientName, managerName: r.managerName, employeeId: r.employeeId, branch: r.branch, region: r.region, zone: r.zone, department: r.department, segment: r.segment, holdings: [] });
    map.get(pan).holdings.push(r);
  }
  return [...map.values()].map((inv) => {
    const h = inv.holdings;
    const totalAum = h.reduce((a, x) => a + num(x.aum), 0);
    const totalUnits = h.reduce((a, x) => a + num(x.units), 0);
    const byAsset = {};
    for (const x of h) { const k = x.assetType || "Others"; byAsset[k] = (byAsset[k] || 0) + num(x.aum); }
    const amcs = [...new Set(h.map((x) => x.fund))];
    const folios = [...new Set(h.map((x) => x.folio).filter((f) => f && f !== "—"))];
    const sorted = [...h].sort((a, b) => num(b.aum) - num(a.aum));
    const lastDate = h.map((x) => x.date).sort().slice(-1)[0] || "—";
    const minDays = Math.min(...h.map((x) => num(x.days) || 0));
    return {
      ...inv, totalAum, totalUnits, byAsset, amcs, folios,
      folioCount: folios.length,
      schemeCount: new Set(h.map((x) => x.schemeCode)).size,
      amcCount: amcs.length,
      avgNav: h.length ? h.reduce((a, x) => a + num(x.nav), 0) / h.length : 0,
      largest: sorted[0], smallest: sorted[sorted.length - 1], lastDate, minDays,
      equityPct: totalAum ? ((byAsset.Equity || 0) / totalAum) * 100 : 0,
      debtPct: totalAum ? ((byAsset.Debt || 0) / totalAum) * 100 : 0,
      _q: `${inv.pan} ${inv.clientName} ${folios.join(" ")}`.toLowerCase(),
    };
  });
}

function recommend(inv) {
  const out = [];
  if (inv.debtPct > 80) out.push({ t: "Add equity exposure", d: `Debt is ${inv.debtPct.toFixed(0)}% of the portfolio. Introduce a large-cap or flexi-cap fund to restore growth potential.`, tone: "warn" });
  if (inv.equityPct < 20 && inv.totalAum > 0) out.push({ t: "Equity under 20%", d: "Long-term returns are capped at this allocation. Review the risk profile with the client.", tone: "warn" });
  if (inv.amcCount === 1) out.push({ t: "Diversify fund houses", d: `All holdings sit with ${inv.amcs[0]}. Spread across two or three AMCs to reduce concentration risk.`, tone: "warn" });
  if (inv.schemeCount > 15) out.push({ t: "Consolidate the portfolio", d: `${inv.schemeCount} schemes create overlap and reporting drag. Trim to eight to ten core funds.`, tone: "warn" });
  if (inv.totalAum > 5000000) out.push({ t: "PMS / AIF eligible", d: `AUM of ${inr(inv.totalAum)} clears the discretionary mandate threshold.`, tone: "good" });
  if (!inv.holdings.some((h) => (h.category || "").toUpperCase().includes("ELSS"))) out.push({ t: "Tax-saving gap", d: "No ELSS holding on file. A Section 80C allocation is available.", tone: "info" });
  if (inv.totalAum < 100000) out.push({ t: "SIP candidate", d: "A monthly SIP builds the book steadily at this ticket size.", tone: "info" });
  if (!out.length) out.push({ t: "Portfolio is balanced", d: "Allocation, diversification and scheme count are all within advisory guardrails.", tone: "good" });
  return out;
}

/* ------------------------------ UI atoms ------------------------------ */
const Card = ({ className = "", children }) => (
  <div className={`rounded-2xl border border-white/10 bg-white/[0.03] backdrop-blur-xl shadow-[0_8px_40px_-12px_rgba(0,0,0,0.5)] dark:border-white/10 ${className}`}>{children}</div>
);
const Kpi = ({ label, value, sub, icon: Icon }) => (
  <Card className="p-5">
    <div className="flex items-start justify-between">
      <div>
        <div className="text-[11px] uppercase tracking-[0.18em] text-neutral-400">{label}</div>
        <div className="mt-2 text-2xl font-semibold tabular-nums text-neutral-900 dark:text-neutral-50">{value}</div>
        {sub && <div className="mt-1 text-xs text-neutral-500">{sub}</div>}
      </div>
      {Icon && <Icon size={18} style={{ color: GOLD }} className="opacity-80" />}
    </div>
  </Card>
);
const Chip = ({ children }) => (
  <span className="rounded-full border border-white/10 bg-white/5 px-2.5 py-0.5 text-[11px] text-neutral-300">{children}</span>
);
const TT = ({ active, payload }) =>
  active && payload?.length ? (
    <div className="rounded-lg border border-white/15 bg-neutral-900/95 px-3 py-2 text-xs text-neutral-100 shadow-xl">
      <div className="font-medium">{payload[0].name ?? payload[0].payload?.name}</div>
      <div style={{ color: GOLD }}>{inr(payload[0].value)}</div>
    </div>
  ) : null;

/* ================================ APP ================================ */
export default function App() {
  const [dark, setDark] = useState(true);
  const [rows, setRows] = useState(() => seedRows());
  const [page, setPage] = useState("dashboard");
  const [query, setQuery] = useState("");
  const [selected, setSelected] = useState(null);
  const [recent, setRecent] = useState([]);
  const [uploading, setUploading] = useState(0);
  const [uploadMsg, setUploadMsg] = useState("");
  const [filters, setFilters] = useState({ fund: "", assetType: "", managerName: "", region: "", zone: "" });

  /* --------- Shortlist: ticked clients + per-client remarks, saved locally --------- */
  const [shortlist, setShortlist] = useState(() => {
    try {
      const raw = localStorage.getItem("how360_shortlist");
      return raw ? JSON.parse(raw) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    try { localStorage.setItem("how360_shortlist", JSON.stringify(shortlist)); } catch {}
  }, [shortlist]);

  const isPicked = useCallback((pan) => Object.prototype.hasOwnProperty.call(shortlist, pan), [shortlist]);

  const togglePick = useCallback((pan) => {
    setShortlist((s) => {
      const next = { ...s };
      if (next[pan]) delete next[pan];
      else next[pan] = { remark: "", addedAt: new Date().toISOString().slice(0, 10) };
      return next;
    });
  }, []);

  const setRemark = useCallback((pan, remark) => {
    setShortlist((s) => (s[pan] ? { ...s, [pan]: { ...s[pan], remark } } : s));
  }, []);

  const clearShortlist = useCallback(() => setShortlist({}), []);
  const pickedCount = Object.keys(shortlist).length;
  const fileRef = useRef(null);

  const investors = useMemo(() => buildInvestors(rows), [rows]);
  const byPan = useMemo(() => new Map(investors.map((i) => [i.pan, i])), [investors]);
  const inv = selected ? byPan.get(selected) : null;

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    return investors.filter((i) => i._q.includes(q)).slice(0, 40);
  }, [query, investors]);

  const open = useCallback((pan) => {
    setSelected(pan); setPage("profile"); setQuery("");
    setRecent((r) => [pan, ...r.filter((x) => x !== pan)].slice(0, 6));
  }, []);

  useEffect(() => {
    const q = query.trim().toUpperCase();
    if (q.length === 10 && byPan.has(q)) open(q);
  }, [query, byPan, open]);

  /* --------- Import --------- */
  const [diag, setDiag] = useState(null);

  const onFile = async (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    setDiag(null); setUploading(8); setUploadMsg(`Reading ${f.name}…`);
    try {
      const buf = await f.arrayBuffer();
      setUploading(30);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      setUploading(55); setUploadMsg("Detecting the header row and mapping columns…");
      const { rows: clean, diag: d } = parseWorkbook(wb);
      setUploading(85); setUploadMsg("Grouping holdings by PAN…");
      setDiag(d);
      setUploading(100);

      if (!clean.length) {
        setUploadMsg(
          d.missing.includes("pan")
            ? `No PAN column found on sheet "${d.sheet}". The header row looked like row ${d.headerRow}. Rename your PAN column to "panno" and re-upload.`
            : `Found the headers on row ${d.headerRow} of "${d.sheet}", but no rows had a valid PAN value. ${d.dropped} rows were skipped.`
        );
        return;
      }
      setRows(clean); setSelected(null);
      const pans = new Set(clean.map((r) => r.pan)).size;
      setUploadMsg(
        `Loaded ${clean.length.toLocaleString("en-IN")} holdings across ${pans.toLocaleString("en-IN")} PANs ` +
        `from "${d.sheet}" (headers on row ${d.headerRow}).`
      );
    } catch (err) {
      setUploading(0);
      setUploadMsg("That file couldn't be read. Upload a .xlsx, .xls or .csv file.");
    }
  };

  const exportCsv = (data, name) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    a.download = `${name}.csv`; a.click();
  };
  const exportXlsx = (data, name) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Report");
    XLSX.writeFile(wb, `${name}.xlsx`);
  };

  const NAV = [
    ["dashboard", "Dashboard", LayoutDashboard], ["investors", "Investors", Users],
    ["search", "Portfolio search", Search], ["shortlist", "Shortlist", CheckSquare],
    ["insights", "Insights", Lightbulb],
    ["managers", "Managers", UserCog], ["reports", "Reports", FileText],
    ["import", "Import data", Upload], ["analytics", "Analytics", PieIcon],
    ["settings", "Settings", Settings],
  ];

  return (
    <div className={dark ? "dark" : ""}>
      <div className="min-h-screen bg-[#0B0B0C] text-neutral-200 dark:bg-[#0B0B0C] [&:not(.dark)_&]:bg-neutral-50">
        <div
          className="min-h-screen"
          style={{ background: dark ? "radial-gradient(1200px 600px at 80% -10%, rgba(212,175,55,0.10), transparent 60%), #0B0B0C" : "radial-gradient(1200px 600px at 80% -10%, rgba(212,175,55,0.14), transparent 60%), #F6F6F4" }}
        >
          <div className="mx-auto flex max-w-[1500px] gap-6 p-4 md:p-6">
            {/* Sidebar */}
            <aside className="sticky top-6 hidden h-fit w-60 shrink-0 lg:block">
              <Card className="p-4">
                <div className="flex items-center gap-2.5 px-1 pb-4">
                  <div className="grid h-9 w-9 place-items-center rounded-lg border" style={{ borderColor: `${GOLD}55`, background: `${GOLD}14` }}>
                    <Landmark size={16} style={{ color: GOLD }} />
                  </div>
                  <div className="leading-tight">
                    <div className="text-[13px] font-semibold tracking-wide text-neutral-100">HOUSE OF WEALTH</div>
                    <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500">Financial Services</div>
                  </div>
                </div>
                <nav className="space-y-0.5">
                  {NAV.map(([k, label, Icon]) => (
                    <button key={k} onClick={() => { setPage(k); if (k !== "profile") setSelected(k === "dashboard" ? null : selected); }}
                      className={`flex w-full items-center gap-2.5 rounded-lg px-3 py-2 text-left text-[13px] transition ${page === k ? "bg-white/10 text-neutral-50" : "text-neutral-400 hover:bg-white/5 hover:text-neutral-200"}`}
                      style={page === k ? { boxShadow: `inset 2px 0 0 ${GOLD}` } : undefined}>
                      <Icon size={15} /> {label}
                      {k === "shortlist" && pickedCount > 0 && (
                        <span className="ml-auto rounded-full px-1.5 py-0.5 text-[10px] font-semibold tabular-nums"
                          style={{ background: `${GOLD}22`, color: GOLD }}>{pickedCount}</span>
                      )}
                    </button>
                  ))}
                </nav>
              </Card>
            </aside>

            {/* Main */}
            <main className="min-w-0 flex-1 space-y-6">
              {/* Search bar */}
              <Card className="relative z-50 p-3">
                <div className="flex items-center gap-3">
                  <Search size={16} className="ml-2 text-neutral-500" />
                  <input
                    value={query} onChange={(e) => setQuery(e.target.value)}
                    placeholder="Search by PAN or client name — e.g. AYTPR1000A or Priya"
                    className="w-full bg-transparent py-2 text-sm text-neutral-100 outline-none placeholder:text-neutral-600"
                  />
                  {query && <button onClick={() => setQuery("")} className="text-neutral-500 hover:text-neutral-300"><X size={15} /></button>}
                  <button onClick={() => setDark(!dark)} className="rounded-lg border border-white/10 p-2 text-neutral-400 hover:text-neutral-200">
                    {dark ? <Sun size={14} /> : <Moon size={14} />}
                  </button>
                </div>
                {results.length > 0 && (
                  <div className="absolute inset-x-3 top-full z-50 mt-2 max-h-80 overflow-auto rounded-xl border border-white/10 bg-[#141416] p-1.5 shadow-[0_20px_60px_-10px_rgba(0,0,0,0.9)]">
                    {results.map((r) => (
                      <button key={r.pan} onClick={() => open(r.pan)} className="flex w-full items-center justify-between rounded-lg px-3 py-2 text-left hover:bg-white/5">
                        <div>
                          <div className="text-sm text-neutral-100">{r.clientName}</div>
                          <div className="font-mono text-[11px] text-neutral-500">{r.pan} · {r.folioCount} folio{r.folioCount !== 1 ? "s" : ""}</div>
                        </div>
                        <div className="text-right">
                          <div className="text-sm tabular-nums" style={{ color: GOLD }}>{inr(r.totalAum)}</div>
                          <div className="text-[11px] text-neutral-500">{r.schemeCount} schemes</div>
                        </div>
                      </button>
                    ))}
                  </div>
                )}
              </Card>

              {page === "profile" && inv ? <Profile inv={inv} onBack={() => { setSelected(null); setPage("dashboard"); }} exportCsv={exportCsv} exportXlsx={exportXlsx}
                    picked={isPicked(inv.pan)} onTogglePick={() => togglePick(inv.pan)}
                    remark={shortlist[inv.pan]?.remark || ""} onRemark={(v) => setRemark(inv.pan, v)} />
                : page === "shortlist" ? <Shortlist shortlist={shortlist} byPan={byPan} open={open}
                    onTogglePick={togglePick} onRemark={setRemark} onClear={clearShortlist}
                    exportCsv={exportCsv} exportXlsx={exportXlsx} />
                : page === "insights" ? <Insights investors={investors} open={open} />
                : page === "investors" ? <InvestorList investors={investors} open={open} filters={filters} setFilters={setFilters}
                    isPicked={isPicked} onTogglePick={togglePick} />
                : page === "import" ? <Import fileRef={fileRef} onFile={onFile} progress={uploading} msg={uploadMsg} diag={diag} />
                : page === "managers" ? <Managers investors={investors} />
                : page === "reports" ? <Reports investors={investors} rows={rows} exportCsv={exportCsv} exportXlsx={exportXlsx} />
                : <Global investors={investors} rows={rows} recent={recent} byPan={byPan} open={open} page={page} />}
            </main>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ------------------ Robust sheet parsing ------------------
   Real MIS exports put a title above the headers ("Holding MIS North I")
   and leave columns like Asset Type blank. Both are handled here. */

// Resolve a raw header cell to a canonical field. Exact map first, then fuzzy.
function matchField(header) {
  const n = norm(header);
  if (!n) return null;
  if (COLUMN_MAP[n]) return COLUMN_MAP[n];
  const squashed = n.replace(/[^a-z]/g, "");
  for (const k of Object.keys(COLUMN_MAP)) {
    if (k.replace(/[^a-z]/g, "") === squashed) return COLUMN_MAP[k];
  }
  // Fuzzy fallbacks for common header drift
  if (/^pan/.test(squashed) || squashed === "panno" || squashed === "pannumber") return "pan";
  if (squashed.includes("clientname") || squashed === "investorname") return "clientName";
  if (squashed.includes("accountno") || squashed === "account" || squashed === "accountnumber") return "folio";
  if (squashed === "folio" || squashed === "foliono" || squashed === "folionumber" || squashed.includes("folio")) return "folio";
  if (squashed === "hold" || squashed === "units" || squashed === "unit" || squashed === "balanceunits") return "units";
  if (squashed === "nav") return "nav";
  if (squashed === "aum" || squashed === "amount" || squashed === "marketvalue" || squashed === "currentvalue") return "aum";
  if (squashed.includes("schemename")) return "schemeName";
  if (squashed.includes("schemecode")) return "schemeCode";
  if (squashed.includes("assettype")) return "assetType";
  if (squashed.startsWith("clasification") || squashed.startsWith("classification")) return "classification";
  if (squashed === "fund" || squashed.includes("amcname") || squashed.includes("fundhouse")) return "fund";
  if (squashed === "category") return "category";
  if (squashed.includes("managername")) return "managerName";
  if (squashed.includes("branch")) return "branch";
  if (squashed.includes("region")) return "region";
  if (squashed.includes("zone")) return "zone";
  if (squashed.includes("department")) return "department";
  if (squashed === "date" || squashed.includes("date")) return "date";
  return null;
}

// Scan the first 25 rows and pick the one that looks most like a header row.
function findHeaderRow(grid) {
  let best = { idx: 0, score: -1 };
  for (let i = 0; i < Math.min(25, grid.length); i++) {
    const cells = (grid[i] || []).filter((c) => String(c ?? "").trim() !== "");
    if (cells.length < 3) continue;
    let score = 0, hasPan = false;
    for (const c of cells) {
      const f = matchField(c);
      if (f) score++;
      if (f === "pan") hasPan = true;
    }
    if (hasPan) score += 5; // a PAN column is the strongest signal
    if (score > best.score) best = { idx: i, score };
  }
  return best;
}

// Derive Asset Type when the column is blank, using Clasification / Category / Scheme Name.
function deriveAsset(o) {
  const hay = `${o.assetType || ""} ${o.classification || ""} ${o.category || ""} ${o.schemeName || ""}`.toLowerCase();
  if (/equity|elss|cap fund|flexi|bluechip|midcap|smallcap|multi cap|large cap|focused|value|contra|sectoral|thematic/.test(hay)) return "Equity";
  if (/debt|liquid|bond|gilt|duration|income|money market|overnight|credit risk|banking and psu/.test(hay)) return "Debt";
  if (/hybrid|balanced|advantage|arbitrage|multi asset|asset alloc|savings/.test(hay)) return "Hybrid";
  if (/gold|silver|commodity|index|etf|fof|international|global/.test(hay)) return "Others";
  return "Others";
}

// Full pipeline: workbook -> clean, PAN-keyed holding rows. Returns diagnostics too.
function parseWorkbook(wb) {
  const diag = { sheet: "", headerRow: 0, mapped: [], missing: [], rawRows: 0, kept: 0, dropped: 0, dupes: 0 };
  let bestSheet = null, bestScore = -1, bestGrid = null, bestHead = null;

  for (const name of wb.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", blankrows: false });
    if (!grid.length) continue;
    const h = findHeaderRow(grid);
    if (h.score > bestScore) { bestScore = h.score; bestSheet = name; bestGrid = grid; bestHead = h; }
  }
  if (!bestGrid) return { rows: [], diag };

  diag.sheet = bestSheet;
  diag.headerRow = bestHead.idx + 1; // 1-based, matches what Excel shows

  const headers = bestGrid[bestHead.idx] || [];
  const fields = headers.map(matchField);
  diag.mapped = headers
    .map((h, i) => (fields[i] ? `${String(h).trim()} → ${fields[i]}` : null))
    .filter(Boolean);

  const panCol = fields.indexOf("pan");
  if (panCol === -1) {
    diag.missing.push("pan");
    return { rows: [], diag };
  }

  const body = bestGrid.slice(bestHead.idx + 1);
  diag.rawRows = body.length;

  const seen = new Set();
  const out = [];
  for (const row of body) {
    const o = {};
    for (let i = 0; i < fields.length; i++) {
      const f = fields[i];
      if (!f) continue;
      const v = row[i];
      o[f] = typeof v === "string" ? v.trim() : v;
    }
    const pan = String(o.pan ?? "").toUpperCase().trim();
    // must look like a PAN, not a stray total/footer row
    if (!/^[A-Z0-9]{8,12}$/.test(pan)) { diag.dropped++; continue; }
    o.pan = pan;

    o.units = num(o.units);
    o.nav = num(o.nav);
    o.aum = num(o.aum) || o.units * o.nav;   // blank AUM -> units x NAV
    o.assetType = deriveAsset(o);            // blank Asset Type -> from Clasification
    o.assetGroup = o.assetType;
    o.fund = String(o.fund || o.amcCode || "Unknown").trim();
    o.schemeName = String(o.schemeName || "—").trim();
    o.schemeCode = String(o.schemeCode ?? "").trim();
    o.clientName = String(o.clientName || "—").trim();
    o.category = String(o.category || o.classification || "—").trim();
    o.managerName = String(o.managerName || "—").trim();
    // Account No IS the folio number, and it differs per holding.
    o.folio = String(o.folio ?? "").trim() || String(o.cpCode ?? "").trim() || "—";
    for (const k of ["branch", "region", "zone", "department"]) {
      o[k] = String(o[k] ?? "").trim() || "—";
    }
    if (o.date instanceof Date) o.date = o.date.toISOString().slice(0, 10);
    else o.date = String(o.date ?? "").trim() || "—";
    o.days = num(o.days);

    // Folio in the dedupe key so two genuinely different folios never collapse.
    const key = `${o.pan}|${o.folio}|${o.schemeCode}|${o.units}|${o.aum}`;
    if (seen.has(key)) { diag.dupes++; continue; }
    seen.add(key);
    out.push(o);
  }
  diag.kept = out.length;
  return { rows: out, diag };
}

/* ------------------------- Global dashboard ------------------------- */
function Global({ investors, rows, recent, byPan, open }) {
  const totalAum = investors.reduce((a, i) => a + i.totalAum, 0);
  const top = (key, n = 10) => {
    const m = {};
    for (const r of rows) { const k = r[key] || "—"; m[k] = (m[k] || 0) + num(r.aum); }
    return Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, n).map(([name, value]) => ({ name, value }));
  };
  const assetMix = useMemo(() => {
    const m = {};
    for (const r of rows) { const k = r.assetType || "Others"; m[k] = (m[k] || 0) + num(r.aum); }
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [rows]);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 gap-4 md:grid-cols-4">
        <Kpi label="Total investors" value={investors.length.toLocaleString("en-IN")} sub="Unique PANs" icon={Users} />
        <Kpi label="Total portfolio value" value={inr(totalAum)} sub={`${rows.length.toLocaleString("en-IN")} holdings`} icon={Wallet} />
        <Kpi label="Average AUM" value={inr(totalAum / (investors.length || 1))} sub="Per investor" icon={TrendingUp} />
        <Kpi label="Schemes / AMCs" value={`${new Set(rows.map((r) => r.schemeCode)).size} / ${new Set(rows.map((r) => r.fund)).size}`} sub="Across the book" icon={Layers} />
      </div>

      <div className="grid gap-4 lg:grid-cols-3">
        <Card className="p-5 lg:col-span-2">
          <h3 className="mb-4 text-sm font-medium text-neutral-200">Top 10 fund houses by AUM</h3>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={top("fund")} layout="vertical" margin={{ left: 10 }}>
              <CartesianGrid strokeDasharray="2 4" stroke="#ffffff10" horizontal={false} />
              <XAxis type="number" tick={{ fill: "#8a8a8a", fontSize: 11 }} tickFormatter={inr} />
              <YAxis type="category" dataKey="name" width={140} tick={{ fill: "#a5a5a5", fontSize: 11 }} />
              <Tooltip content={<TT />} cursor={{ fill: "#ffffff08" }} />
              <Bar dataKey="value" fill={GOLD} radius={[0, 4, 4, 0]} barSize={13} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card className="p-5">
          <h3 className="mb-2 text-sm font-medium text-neutral-200">Asset allocation</h3>
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={assetMix} dataKey="value" nameKey="name" innerRadius={52} outerRadius={82} paddingAngle={3} stroke="none">
                {assetMix.map((_, i) => <Cell key={i} fill={PALETTE[i % PALETTE.length]} />)}
              </Pie>
              <Tooltip content={<TT />} />
              <Legend iconType="circle" wrapperStyle={{ fontSize: 11, color: "#a5a5a5" }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
        {[["Top schemes", "schemeName"], ["Top managers", "managerName"], ["Top branches", "branch"], ["Top zones", "zone"]].map(([title, key]) => (
          <Card key={key} className="p-5">
            <h3 className="mb-3 text-sm font-medium text-neutral-200">{title}</h3>
            <div className="space-y-2">
              {top(key, 6).map((x, i) => (
                <div key={x.name} className="flex items-center justify-between gap-3 text-xs">
                  <span className="truncate text-neutral-400"><span className="mr-2 tabular-nums text-neutral-600">{i + 1}</span>{x.name}</span>
                  <span className="shrink-0 tabular-nums" style={{ color: GOLD }}>{inr(x.value)}</span>
                </div>
              ))}
            </div>
          </Card>
        ))}
      </div>

      {recent.length > 0 && (
        <Card className="p-5">
          <h3 className="mb-3 text-sm font-medium text-neutral-200">Recent searches</h3>
          <div className="flex flex-wrap gap-2">
            {recent.map((p) => byPan.get(p) && (
              <button key={p} onClick={() => open(p)} className="rounded-lg border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-neutral-300 hover:border-white/25">
                {byPan.get(p).clientName} <span className="font-mono text-neutral-500">· {p}</span>
              </button>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}

/* ---------------------------- Investor list ---------------------------- */
function InvestorList({ investors, open, filters, setFilters, isPicked, onTogglePick }) {
  const opts = (k) => [...new Set(investors.map((i) => i[k]).filter(Boolean))].sort();
  const list = investors.filter((i) =>
    (!filters.managerName || i.managerName === filters.managerName) &&
    (!filters.region || i.region === filters.region) &&
    (!filters.zone || i.zone === filters.zone)
  ).sort((a, b) => b.totalAum - a.totalAum);

  return (
    <div className="space-y-4">
      <Card className="flex flex-wrap gap-2 p-4">
        {[["managerName", "Manager"], ["region", "Region"], ["zone", "Zone"]].map(([k, label]) => (
          <select key={k} value={filters[k]} onChange={(e) => setFilters({ ...filters, [k]: e.target.value })}
            className="rounded-lg border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-neutral-300 outline-none">
            <option value="">All {label.toLowerCase()}s</option>
            {opts(k).map((o) => <option key={o} value={o} className="bg-neutral-900">{o}</option>)}
          </select>
        ))}
        <span className="ml-auto self-center text-xs text-neutral-500">{list.length} investors</span>
      </Card>
      <Card className="overflow-hidden">
        <div className="max-h-[70vh] overflow-auto">
          <table className="w-full text-left text-xs">
            <thead className="sticky top-0 z-10 bg-neutral-900/95 backdrop-blur">
              <tr className="text-[11px] uppercase tracking-wider text-neutral-500">
                <th className="px-4 py-3 font-medium"><CheckSquare size={13} /></th>
                {["Client", "PAN", "Schemes", "AMCs", "Equity %", "Manager", "AUM"].map((h) => <th key={h} className="px-4 py-3 font-medium">{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {list.map((i) => (
                <tr key={i.pan} className="border-t border-white/5 hover:bg-white/5">
                  <td className="px-4 py-2.5">
                    <input type="checkbox" checked={isPicked(i.pan)} onChange={() => onTogglePick(i.pan)}
                      className="h-3.5 w-3.5 cursor-pointer rounded" style={{ accentColor: GOLD }} />
                  </td>
                  <td onClick={() => open(i.pan)} className="cursor-pointer px-4 py-2.5 text-neutral-200">{i.clientName}</td>
                  <td onClick={() => open(i.pan)} className="cursor-pointer px-4 py-2.5 font-mono text-neutral-500">{i.pan}</td>
                  <td className="px-4 py-2.5 tabular-nums text-neutral-400">{i.schemeCount}</td>
                  <td className="px-4 py-2.5 tabular-nums text-neutral-400">{i.amcCount}</td>
                  <td className="px-4 py-2.5 tabular-nums text-neutral-400">{i.equityPct.toFixed(0)}%</td>
                  <td className="px-4 py-2.5 text-neutral-400">{i.managerName}</td>
                  <td className="px-4 py-2.5 text-right tabular-nums" style={{ color: GOLD }}>{inr(i.totalAum)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ------------------------------ Profile ------------------------------ */
function Profile({ inv, onBack, exportCsv, exportXlsx, picked, onTogglePick, remark, onRemark }) {
  const [sort, setSort] = useState({ k: "aum", dir: -1 });
  const [q, setQ] = useState("");

  const cols = [["fund", "AMC / Fund"], ["schemeName", "Scheme"], ["folio", "Folio No"], ["schemeCode", "Code"], ["assetType", "Asset"], ["category", "Category"], ["units", "Units"], ["nav", "NAV"], ["aum", "AUM"], ["date", "Date"], ["recommendedScheme", "Rec."], ["days", "Days"]];

  const filtered = useMemo(() => {
    const s = q.toLowerCase();
    const base = s ? inv.holdings.filter((h) => `${h.schemeName} ${h.fund} ${h.category} ${h.assetType} ${h.folio}`.toLowerCase().includes(s)) : inv.holdings;
    return [...base].sort((a, b) => {
      const x = a[sort.k], y = b[sort.k];
      const cmp = typeof x === "number" || !isNaN(num(x)) && typeof x !== "string" ? num(x) - num(y) : String(x).localeCompare(String(y));
      return (isNaN(cmp) ? String(x).localeCompare(String(y)) : cmp) * sort.dir;
    });
  }, [inv, q, sort]);
  const pageRows = filtered;

  const assetData = Object.entries(inv.byAsset).map(([name, value]) => ({ name, value }));
  const amcData = useMemo(() => {
    const m = {}; for (const h of inv.holdings) m[h.fund] = (m[h.fund] || 0) + num(h.aum);
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [inv]);
  const topHold = [...inv.holdings].sort((a, b) => num(b.aum) - num(a.aum)).slice(0, 6).map((h) => ({ name: h.schemeName.replace(/ Fund$/, ""), value: num(h.aum) }));
  const treeData = inv.holdings.map((h) => ({ name: h.schemeName, size: num(h.aum) }));
  const areaData = [...inv.holdings].sort((a, b) => String(a.date).localeCompare(String(b.date)))
    .reduce((acc, h) => { const prev = acc.length ? acc[acc.length - 1].value : 0; acc.push({ name: h.date, value: prev + num(h.aum) }); return acc; }, []);
  const recs = recommend(inv);

  const meta = [["PAN", inv.pan], ["Folios", `${inv.folioCount} account${inv.folioCount !== 1 ? "s" : ""}`], ["Relationship manager", inv.managerName], ["Employee", inv.employeeId], ["Branch", inv.branch], ["Region", inv.region], ["Zone", inv.zone], ["Department", inv.department], ["Last updated", inv.lastDate]];

  return (
    <div className="space-y-6">
      <button onClick={onBack} className="flex items-center gap-1.5 text-xs text-neutral-500 hover:text-neutral-300"><ChevronLeft size={14} /> Back to dashboard</button>

      <Card className="p-6">
        <div className="flex flex-wrap items-start justify-between gap-6">
          <div>
            <div className="text-[11px] uppercase tracking-[0.2em]" style={{ color: GOLD }}>Investor profile</div>
            <h1 className="mt-1 text-3xl font-semibold text-neutral-50">{inv.clientName}</h1>
            <div className="mt-3 flex flex-wrap gap-2">{meta.slice(0, 3).map(([k, v]) => <Chip key={k}>{k}: <span className="ml-1 font-mono text-neutral-200">{v}</span></Chip>)}</div>
          </div>
          <div className="text-right">
            <div className="text-[11px] uppercase tracking-[0.18em] text-neutral-500">Total portfolio value</div>
            <div className="text-4xl font-semibold tabular-nums" style={{ color: GOLD }}>{inr(inv.totalAum)}</div>
            <div className="mt-1 text-xs text-neutral-500">{inrFull(inv.totalAum)}</div>
          </div>
        </div>
        <div className="mt-6 grid grid-cols-2 gap-x-6 gap-y-3 border-t border-white/10 pt-5 text-xs md:grid-cols-3 lg:grid-cols-5">
          {meta.slice(3).map(([k, v]) => (
            <div key={k}><div className="text-neutral-500">{k}</div><div className="mt-0.5 text-neutral-200">{v}</div></div>
          ))}
        </div>

        {/* Shortlist: tick the client, then leave a remark */}
        <div className="mt-5 rounded-xl border p-4 transition"
          style={{ borderColor: picked ? `${GOLD}55` : "#ffffff14", background: picked ? `${GOLD}0D` : "transparent" }}>
          <label className="flex cursor-pointer items-center gap-3">
            <input type="checkbox" checked={picked} onChange={onTogglePick}
              className="h-4 w-4 cursor-pointer rounded border-white/25 bg-transparent"
              style={{ accentColor: GOLD }} />
            <span className="text-sm text-neutral-200">
              {picked ? "On your shortlist" : "Add to shortlist"}
            </span>
            {picked && <CheckSquare size={14} style={{ color: GOLD }} className="ml-auto" />}
          </label>

          {picked ? (
            <div className="mt-3">
              <label className="text-[11px] uppercase tracking-[0.15em] text-neutral-500">Remark</label>
              <textarea
                value={remark}
                onChange={(e) => onRemark(e.target.value)}
                rows={2}
                placeholder="e.g. Call after 10 days"
                className="mt-1.5 w-full resize-none rounded-lg border border-white/10 bg-white/5 px-3 py-2 text-sm text-neutral-100 outline-none placeholder:text-neutral-600 focus:border-white/25"
              />
              <p className="mt-1 text-[11px] text-neutral-600">
                Saved automatically. Tick more clients, then export them together from Shortlist.
              </p>
            </div>
          ) : (
            <p className="mt-2 pl-7 text-xs text-neutral-600">
              Tick to collect this client, add a remark, and export the batch later.
            </p>
          )}
        </div>
      </Card>

      <div className="grid grid-cols-2 gap-4 md:grid-cols-4 lg:grid-cols-6">
        <Kpi label="Total units" value={inv.totalUnits.toLocaleString("en-IN", { maximumFractionDigits: 0 })} />
        <Kpi label="Schemes" value={inv.schemeCount} />
        <Kpi label="Folios" value={inv.folioCount} sub="Account numbers" />
        <Kpi label="Fund houses" value={inv.amcCount} />
        <Kpi label="Average NAV" value={`₹${inv.avgNav.toFixed(2)}`} />
        <Kpi label="Largest holding" value={inr(num(inv.largest?.aum))} sub={inv.largest?.schemeName} />
      </div>

      {/* Charts */}
      <div className="grid gap-4 lg:grid-cols-2">
        <Card className="p-5">
          <h3 className="mb-2 text-sm font-medium text-neutral-200">Asset allocation</h3>
          <ResponsiveContainer width="100%" height={230}>
            <PieChart>
              <Pie data={assetData} dataKey="value" nameKey="name" outerRadius={82} stroke="none">
                {assetData.map((_, i) => <Cell key={i} fill={PALETTE[i % PALETTE.length]} />)}
              </Pie>
              <Tooltip content={<TT />} /><Legend iconType="circle" wrapperStyle={{ fontSize: 11, color: "#a5a5a5" }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card className="p-5">
          <h3 className="mb-2 text-sm font-medium text-neutral-200">Fund house distribution</h3>
          <ResponsiveContainer width="100%" height={230}>
            <PieChart>
              <Pie data={amcData} dataKey="value" nameKey="name" innerRadius={52} outerRadius={82} paddingAngle={2} stroke="none">
                {amcData.map((_, i) => <Cell key={i} fill={PALETTE[i % PALETTE.length]} />)}
              </Pie>
              <Tooltip content={<TT />} /><Legend iconType="circle" wrapperStyle={{ fontSize: 10, color: "#a5a5a5" }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card className="p-5">
          <h3 className="mb-2 text-sm font-medium text-neutral-200">Top holdings</h3>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={topHold}>
              <CartesianGrid strokeDasharray="2 4" stroke="#ffffff10" vertical={false} />
              <XAxis dataKey="name" tick={{ fill: "#7a7a7a", fontSize: 9 }} interval={0} angle={-18} textAnchor="end" height={56} />
              <YAxis tick={{ fill: "#7a7a7a", fontSize: 10 }} tickFormatter={inr} />
              <Tooltip content={<TT />} cursor={{ fill: "#ffffff08" }} />
              <Bar dataKey="value" fill={GOLD} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card className="p-5">
          <h3 className="mb-2 text-sm font-medium text-neutral-200">Cumulative AUM by investment date</h3>
          <ResponsiveContainer width="100%" height={230}>
            <AreaChart data={areaData}>
              <defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={GOLD} stopOpacity={0.5} /><stop offset="100%" stopColor={GOLD} stopOpacity={0} />
              </linearGradient></defs>
              <CartesianGrid strokeDasharray="2 4" stroke="#ffffff10" vertical={false} />
              <XAxis dataKey="name" tick={{ fill: "#7a7a7a", fontSize: 9 }} />
              <YAxis tick={{ fill: "#7a7a7a", fontSize: 10 }} tickFormatter={inr} />
              <Tooltip content={<TT />} />
              <Area type="monotone" dataKey="value" stroke={GOLD} strokeWidth={2} fill="url(#g)" />
            </AreaChart>
          </ResponsiveContainer>
        </Card>
        <Card className="p-5 lg:col-span-2">
          <h3 className="mb-2 text-sm font-medium text-neutral-200">Scheme-wise allocation</h3>
          <ResponsiveContainer width="100%" height={230}>
            <Treemap data={treeData} dataKey="size" stroke="#0B0B0C" content={<TreeCell />} />
          </ResponsiveContainer>
        </Card>
      </div>

      {/* Table */}
      <Card className="overflow-hidden">
        <div className="flex flex-wrap items-center gap-2 border-b border-white/10 p-4">
          <h3 className="mr-auto text-sm font-medium text-neutral-200">Holdings <span className="text-neutral-500">({filtered.length})</span></h3>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Filter schemes"
            className="rounded-lg border border-white/10 bg-white/5 px-3 py-1.5 text-xs text-neutral-200 outline-none placeholder:text-neutral-600" />
          <button onClick={() => exportCsv(exportHoldingRows(inv), `${inv.pan}_portfolio`)} className="flex items-center gap-1.5 rounded-lg border border-white/10 px-3 py-1.5 text-xs text-neutral-300 hover:border-white/25"><Download size={13} /> CSV</button>
          <button onClick={() => exportXlsx(exportHoldingRows(inv), `${inv.pan}_portfolio`)} className="flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs" style={{ borderColor: `${GOLD}55`, color: GOLD }}><Download size={13} /> Excel</button>
        </div>
        <div className="max-h-[75vh] overflow-auto">
          <table className="w-full text-left text-xs">
            <thead className="sticky top-0 z-10 bg-neutral-900/95 backdrop-blur">
              <tr className="text-[11px] uppercase tracking-wider text-neutral-500">
                {cols.map(([k, label]) => (
                  <th key={k} onClick={() => setSort((s) => ({ k, dir: s.k === k ? -s.dir : -1 }))} className="cursor-pointer whitespace-nowrap px-4 py-3 font-medium hover:text-neutral-300">
                    <span className="inline-flex items-center gap-1">{label}<ArrowUpDown size={10} className={sort.k === k ? "opacity-100" : "opacity-25"} /></span>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {pageRows.map((h, i) => (
                <tr key={i} className="border-t border-white/5 hover:bg-white/5">
                  <td className="whitespace-nowrap px-4 py-2.5 text-neutral-300">{h.fund}</td>
                  <td className="whitespace-nowrap px-4 py-2.5 text-neutral-200">{h.schemeName}</td>
                  <td className="whitespace-nowrap px-4 py-2.5 font-mono text-[11px] text-neutral-400">{h.folio}</td>
                  <td className="px-4 py-2.5 font-mono text-neutral-600">{h.schemeCode}</td>
                  <td className="px-4 py-2.5"><Chip>{h.assetType}</Chip></td>
                  <td className="whitespace-nowrap px-4 py-2.5 text-neutral-500">{h.category}</td>
                  <td className="px-4 py-2.5 tabular-nums text-neutral-400">{num(h.units).toLocaleString("en-IN", { maximumFractionDigits: 2 })}</td>
                  <td className="px-4 py-2.5 tabular-nums text-neutral-400">₹{num(h.nav).toFixed(2)}</td>
                  <td className="px-4 py-2.5 text-right tabular-nums font-medium" style={{ color: GOLD }}>{inr(num(h.aum))}</td>
                  <td className="whitespace-nowrap px-4 py-2.5 text-neutral-500">{h.date}</td>
                  <td className="px-4 py-2.5 text-neutral-500">{h.recommendedScheme}</td>
                  <td className="px-4 py-2.5 tabular-nums text-neutral-500">{h.days}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="flex items-center justify-between border-t border-white/10 px-4 py-3 text-xs text-neutral-500">
          <span>Showing all {filtered.length} holding{filtered.length !== 1 ? "s" : ""}</span>
          <span className="tabular-nums">Total <span style={{ color: GOLD }}>{inr(filtered.reduce((a, h) => a + num(h.aum), 0))}</span></span>
        </div>
      </Card>

      <Card className="p-5">
        <div className="mb-4 flex items-center gap-2">
          <Sparkles size={15} style={{ color: GOLD }} />
          <h3 className="text-sm font-medium text-neutral-200">Advisor recommendations</h3>
        </div>
        <div className="grid gap-3 md:grid-cols-2">
          {recs.map((r) => (
            <div key={r.t} className="rounded-xl border border-white/10 bg-white/[0.02] p-4">
              <div className="flex items-center gap-2 text-sm text-neutral-100">
                {r.tone === "good" ? <CheckCircle2 size={14} className="text-emerald-500/80" /> : r.tone === "warn" ? <AlertTriangle size={14} style={{ color: GOLD }} /> : <Lightbulb size={14} className="text-neutral-500" />}
                {r.t}
              </div>
              <p className="mt-1.5 text-xs leading-relaxed text-neutral-500">{r.d}</p>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

function TreeCell({ x, y, width, height, index, name, value }) {
  if (width < 2 || height < 2) return null;
  return (
    <g>
      <rect x={x} y={y} width={width} height={height} fill={PALETTE[index % PALETTE.length]} fillOpacity={0.85} rx={3} />
      {width > 74 && height > 26 && (
        <text x={x + 7} y={y + 17} fill="#0B0B0C" fontSize={10} fontWeight={600}>
          {String(name).slice(0, Math.floor(width / 6.4))}
        </text>
      )}
    </g>
  );
}

/* ------------------------------ Shortlist ------------------------------ */
function Shortlist({ shortlist, byPan, open, onTogglePick, onRemark, onClear, exportCsv, exportXlsx }) {
  const picks = Object.entries(shortlist)
    .map(([pan, meta]) => ({ pan, ...meta, inv: byPan.get(pan) }))
    .filter((p) => p.inv)
    .sort((a, b) => b.inv.totalAum - a.inv.totalAum);

  const totalAum = picks.reduce((a, p) => a + p.inv.totalAum, 0);

  // One row per client, remark included — this is what gets exported.
  const exportRows = picks.map((p) => ({
    "Client Name": p.inv.clientName,
    PAN: p.pan,
    "Folios": p.inv.folioCount,
    "Folio Numbers": p.inv.folios.join(", "),
    Remark: p.remark || "",
    "Total AUM": p.inv.totalAum,
    "Total Units": Math.round(p.inv.totalUnits),
    Schemes: p.inv.schemeCount,
    AMCs: p.inv.amcCount,
    "Equity %": +p.inv.equityPct.toFixed(1),
    "Debt %": +p.inv.debtPct.toFixed(1),
    "Relationship Manager": p.inv.managerName,
    Branch: p.inv.branch,
    Region: p.inv.region,
    Zone: p.inv.zone,
    "Added On": p.addedAt || "",
  }));

  // Optional deeper export: every holding of every ticked client, remark carried on each row.
  const exportHoldings = picks.flatMap((p) =>
    p.inv.holdings.map((h) => ({
      "Client Name": p.inv.clientName, PAN: p.pan, Remark: p.remark || "",
      AMC: h.fund, "Scheme Name": h.schemeName, "Folio No": h.folio, "Scheme Code": h.schemeCode,
      "Asset Type": h.assetType, Category: h.category,
      Units: num(h.units), NAV: num(h.nav), AUM: num(h.aum), Date: h.date,
    }))
  );

  if (!picks.length) {
    return (
      <Card className="p-12 text-center">
        <div className="mx-auto grid h-12 w-12 place-items-center rounded-xl border" style={{ borderColor: `${GOLD}44`, background: `${GOLD}10` }}>
          <CheckSquare size={20} style={{ color: GOLD }} />
        </div>
        <h2 className="mt-4 text-lg font-semibold text-neutral-100">No clients shortlisted yet</h2>
        <p className="mx-auto mt-2 max-w-md text-sm text-neutral-500">
          Search a PAN or name, open the profile, and tick <span className="text-neutral-300">Add to shortlist</span>.
          Leave a remark like “Call after 10 days”, then come back here to export the batch.
        </p>
      </Card>
    );
  }

  return (
    <div className="space-y-4">
      <Card className="p-6">
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div>
            <div className="text-[11px] uppercase tracking-[0.2em]" style={{ color: GOLD }}>Shortlist</div>
            <h1 className="mt-1 text-2xl font-semibold text-neutral-50">
              {picks.length} client{picks.length > 1 ? "s" : ""} selected
            </h1>
            <p className="mt-1 text-sm text-neutral-500">
              Combined AUM {inr(totalAum)}. Remarks save as you type.
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button onClick={() => exportCsv(exportRows, "shortlist")}
              className="flex items-center gap-1.5 rounded-lg border border-white/10 px-3 py-1.5 text-xs text-neutral-300 hover:border-white/25">
              <Download size={13} /> CSV
            </button>
            <button onClick={() => exportXlsx(exportRows, "shortlist")}
              className="flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs"
              style={{ borderColor: `${GOLD}55`, color: GOLD }}>
              <Download size={13} /> Excel
            </button>
            <button onClick={() => exportXlsx(exportHoldings, "shortlist_holdings")}
              className="flex items-center gap-1.5 rounded-lg border border-white/10 px-3 py-1.5 text-xs text-neutral-300 hover:border-white/25">
              <Download size={13} /> Excel + holdings
            </button>
            <button onClick={onClear}
              className="flex items-center gap-1.5 rounded-lg border border-white/10 px-3 py-1.5 text-xs text-neutral-500 hover:border-red-500/40 hover:text-red-400">
              <X size={13} /> Clear all
            </button>
          </div>
        </div>
      </Card>

      <div className="space-y-3">
        {picks.map((p) => (
          <Card key={p.pan} className="p-4">
            <div className="flex flex-wrap items-start gap-4">
              <input type="checkbox" checked readOnly onClick={() => onTogglePick(p.pan)}
                className="mt-1 h-4 w-4 cursor-pointer rounded" style={{ accentColor: GOLD }} title="Remove from shortlist" />

              <div className="min-w-[190px] flex-1">
                <button onClick={() => open(p.pan)} className="text-left text-sm font-medium text-neutral-100 hover:underline">
                  {p.inv.clientName}
                </button>
                <div className="mt-0.5 font-mono text-[11px] text-neutral-500">{p.pan} · {p.inv.folioCount} folio{p.inv.folioCount !== 1 ? "s" : ""}</div>
                <div className="mt-2 flex flex-wrap gap-1.5">
                  <Chip>{p.inv.schemeCount} schemes</Chip>
                  <Chip>{p.inv.amcCount} AMCs</Chip>
                  <Chip>Equity {p.inv.equityPct.toFixed(0)}%</Chip>
                  <Chip>{p.inv.managerName}</Chip>
                </div>
              </div>

              <div className="flex-[2] basis-72">
                <label className="text-[11px] uppercase tracking-[0.15em] text-neutral-500">Remark</label>
                <input
                  value={p.remark}
                  onChange={(e) => onRemark(p.pan, e.target.value)}
                  placeholder="e.g. Call after 10 days"
                  className="mt-1 w-full rounded-lg border border-white/10 bg-white/5 px-3 py-2 text-sm text-neutral-100 outline-none placeholder:text-neutral-600 focus:border-white/25"
                />
              </div>

              <div className="text-right">
                <div className="text-[10px] uppercase tracking-[0.15em] text-neutral-600">AUM</div>
                <div className="text-lg font-semibold tabular-nums" style={{ color: GOLD }}>{inr(p.inv.totalAum)}</div>
                <div className="mt-0.5 text-[10px] text-neutral-600">added {p.addedAt}</div>
              </div>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------ Insights ------------------------------ */
function Insights({ investors, open }) {
  const groups = [
    ["Single fund house", investors.filter((i) => i.amcCount === 1), "Concentration risk. Pitch diversification."],
    ["Debt only", investors.filter((i) => i.debtPct === 100 && i.totalAum > 0), "No growth engine in the portfolio."],
    ["Equity under 20%", investors.filter((i) => i.equityPct < 20 && i.totalAum > 0), "Long-horizon returns are constrained."],
    ["More than 15 schemes", investors.filter((i) => i.schemeCount > 15), "Consolidation candidates."],
    ["AUM above ₹50 lakh", investors.filter((i) => i.totalAum > 5000000), "PMS and AIF eligible."],
    ["AUM below ₹1 lakh", investors.filter((i) => i.totalAum < 100000), "SIP candidates."],
    ["Needs diversification", investors.filter((i) => i.amcCount <= 2 && i.schemeCount >= 4), "Few AMCs, many schemes."],
    ["No ELSS on file", investors.filter((i) => !i.holdings.some((h) => (h.category || "").toUpperCase().includes("ELSS"))), "Section 80C opportunity."],
  ];
  return (
    <div className="space-y-4">
      <Card className="p-6">
        <div className="text-[11px] uppercase tracking-[0.2em]" style={{ color: GOLD }}>Business insights</div>
        <h1 className="mt-1 text-2xl font-semibold text-neutral-50">Where the book needs attention</h1>
        <p className="mt-2 max-w-2xl text-sm text-neutral-500">Every cohort is computed live from consolidated PAN portfolios. Click any client to open the full 360 profile.</p>
      </Card>
      <div className="grid gap-4 md:grid-cols-2">
        {groups.map(([title, list, why]) => (
          <Card key={title} className="p-5">
            <div className="flex items-baseline justify-between">
              <h3 className="text-sm font-medium text-neutral-200">{title}</h3>
              <span className="text-lg font-semibold tabular-nums" style={{ color: GOLD }}>{list.length}</span>
            </div>
            <p className="mt-1 text-xs text-neutral-600">{why}</p>
            <div className="mt-3 max-h-40 space-y-1 overflow-auto">
              {list.slice(0, 30).map((i) => (
                <button key={i.pan} onClick={() => open(i.pan)} className="flex w-full items-center justify-between rounded px-2 py-1 text-xs hover:bg-white/5">
                  <span className="text-neutral-400">{i.clientName}</span>
                  <span className="font-mono text-[10px] text-neutral-600">{i.pan} · {inr(i.totalAum)}</span>
                </button>
              ))}
              {!list.length && <p className="px-2 py-3 text-xs text-neutral-600">No clients in this cohort.</p>}
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------ Managers ------------------------------ */
function Managers({ investors }) {
  const m = {};
  for (const i of investors) {
    const k = i.managerName || "—";
    m[k] = m[k] || { name: k, clients: 0, aum: 0, branch: i.branch, region: i.region, zone: i.zone, schemes: 0 };
    m[k].clients++; m[k].aum += i.totalAum; m[k].schemes += i.schemeCount;
  }
  const list = Object.values(m).sort((a, b) => b.aum - a.aum);
  return (
    <Card className="overflow-hidden">
      <div className="border-b border-white/10 p-5"><h3 className="text-sm font-medium text-neutral-200">Relationship managers</h3></div>
      <table className="w-full text-left text-xs">
        <thead className="bg-neutral-900/60"><tr className="text-[11px] uppercase tracking-wider text-neutral-500">
          {["Manager", "Branch", "Region", "Zone", "Clients", "Schemes", "Book AUM"].map((h) => <th key={h} className="px-4 py-3 font-medium">{h}</th>)}
        </tr></thead>
        <tbody>
          {list.map((x) => (
            <tr key={x.name} className="border-t border-white/5 hover:bg-white/5">
              <td className="px-4 py-3 text-neutral-200">{x.name}</td>
              <td className="px-4 py-3 text-neutral-500">{x.branch}</td>
              <td className="px-4 py-3 text-neutral-500">{x.region}</td>
              <td className="px-4 py-3 text-neutral-500">{x.zone}</td>
              <td className="px-4 py-3 tabular-nums text-neutral-400">{x.clients}</td>
              <td className="px-4 py-3 tabular-nums text-neutral-400">{x.schemes}</td>
              <td className="px-4 py-3 text-right tabular-nums" style={{ color: GOLD }}>{inr(x.aum)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Card>
  );
}

/* ------------------------------- Reports ------------------------------- */
function Reports({ investors, rows, exportCsv, exportXlsx }) {
  const agg = (key) => {
    const m = {};
    for (const r of rows) { const k = r[key] || "—"; m[k] = m[k] || { [key]: k, holdings: 0, aum: 0 }; m[k].holdings++; m[k].aum += num(r.aum); }
    return Object.values(m).sort((a, b) => b.aum - a.aum);
  };
  const reports = [
    ["Client portfolio report", investors.map((i) => ({ PAN: i.pan, Client: i.clientName, Folios: i.folioCount, Schemes: i.schemeCount, AMCs: i.amcCount, Units: Math.round(i.totalUnits), AUM: i.totalAum, Manager: i.managerName, Branch: i.branch }))],
    ["Manager report", agg("managerName")], ["Branch report", agg("branch")],
    ["AMC report", agg("fund")], ["Scheme report", agg("schemeName")],
  ];
  return (
    <div className="grid gap-4 md:grid-cols-2">
      {reports.map(([title, data]) => (
        <Card key={title} className="p-5">
          <h3 className="text-sm font-medium text-neutral-200">{title}</h3>
          <p className="mt-1 text-xs text-neutral-600">{data.length.toLocaleString("en-IN")} rows, generated from the current dataset.</p>
          <div className="mt-4 flex gap-2">
            <button onClick={() => exportCsv(data, title.replace(/\s+/g, "_"))} className="flex items-center gap-1.5 rounded-lg border border-white/10 px-3 py-1.5 text-xs text-neutral-300 hover:border-white/25"><Download size={13} /> Download CSV</button>
            <button onClick={() => exportXlsx(data, title.replace(/\s+/g, "_"))} className="flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs" style={{ borderColor: `${GOLD}55`, color: GOLD }}><Download size={13} /> Download Excel</button>
          </div>
        </Card>
      ))}
    </div>
  );
}

/* -------------------------------- Import -------------------------------- */
function Import({ fileRef, onFile, progress, msg, diag }) {
  return (
    <div className="space-y-4">
      <Card className="p-8">
        <div className="text-[11px] uppercase tracking-[0.2em]" style={{ color: GOLD }}>Import data</div>
        <h1 className="mt-1 text-2xl font-semibold text-neutral-50">Upload the holdings file</h1>
        <p className="mt-2 max-w-2xl text-sm text-neutral-500">
          Accepts .xlsx, .xls and .csv. The header row is detected automatically, so a title above the
          headers is fine. Column names are matched loosely, blank AUM is recomputed from units × NAV,
          blank Asset Type is derived from Clasification, duplicates are dropped, and every row is
          grouped under its PAN.
        </p>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFile} className="hidden" />
        <button onClick={() => fileRef.current?.click()}
          className="mt-6 flex w-full items-center justify-center gap-3 rounded-2xl border-2 border-dashed py-16 text-sm transition hover:bg-white/[0.03]"
          style={{ borderColor: `${GOLD}44`, color: GOLD }}>
          <Upload size={18} /> Choose a file
        </button>
        {progress > 0 && (
          <div className="mt-5">
            <div className="h-1 overflow-hidden rounded-full bg-white/10">
              <div className="h-full transition-all duration-300" style={{ width: `${progress}%`, background: GOLD }} />
            </div>
            <p className="mt-2 text-xs text-neutral-400">{msg}</p>
          </div>
        )}
      </Card>

      {diag && (
        <Card className="p-6">
          <h3 className="text-sm font-medium text-neutral-200">What the importer saw</h3>
          <div className="mt-4 grid gap-4 text-xs md:grid-cols-4">
            {[["Sheet", diag.sheet || "—"], ["Header row", diag.headerRow || "—"],
              ["Rows read", (diag.rawRows || 0).toLocaleString("en-IN")],
              ["Rows kept", (diag.kept || 0).toLocaleString("en-IN")]].map(([k, v]) => (
              <div key={k}>
                <div className="text-neutral-500">{k}</div>
                <div className="mt-1 text-base text-neutral-100">{v}</div>
              </div>
            ))}
          </div>
          {(diag.dropped > 0 || diag.dupes > 0) && (
            <p className="mt-3 text-xs text-neutral-500">
              Skipped {diag.dropped.toLocaleString("en-IN")} rows without a valid PAN
              {diag.dupes > 0 && <> and {diag.dupes.toLocaleString("en-IN")} duplicate rows</>}.
            </p>
          )}
          {diag.mapped?.length > 0 && (
            <div className="mt-5 border-t border-white/10 pt-4">
              <div className="text-xs text-neutral-500">Column mapping</div>
              <div className="mt-2 flex flex-wrap gap-1.5">
                {diag.mapped.map((m) => (
                  <span key={m} className="rounded border border-white/10 bg-white/5 px-2 py-1 font-mono text-[10px] text-neutral-400">{m}</span>
                ))}
              </div>
            </div>
          )}
        </Card>
      )}
    </div>
  );
}
