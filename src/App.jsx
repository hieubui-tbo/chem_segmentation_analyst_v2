import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════
   CONSTANTS
   ═══════════════════════════════════════════════ */
const STEPS = [
  { id: "upload",   label: "Upload Data",       icon: "📁" },
  { id: "schema",   label: "Schema Mapping",    icon: "🗂️" },
  { id: "quality",  label: "Data Quality",      icon: "🔍" },
  { id: "backfill", label: "Numeric Backfill",  icon: "🔧" },
  { id: "validate", label: "Value Validation",  icon: "✅" },
  { id: "entity",   label: "Entity Resolution", icon: "🏢" },
  { id: "ranking",  label: "Top 80% Ranking",   icon: "📊" },
  { id: "industry", label: "Industry Master",   icon: "🏭" },
  { id: "keywords", label: "Keyword Matrix",    icon: "🔑" },
  { id: "classify", label: "Classification",    icon: "🏷️" },
  { id: "convert",  label: "Unit Conversion",   icon: "⚖️" },
  { id: "iqr",      label: "IQR Outlier",       icon: "📈" },
  { id: "metrics",  label: "Segment Metrics",   icon: "📋" },
  { id: "export",   label: "Export",            icon: "💾" },
];

const UNIT_FACTORS = {
  TNE: 1, TONNE: 1, TONNES: 1, "METRIC TON": 1, "METRIC TONS": 1, MT: 1,
  KGM: 0.001, KG: 0.001, KGS: 0.001, KILOGRAM: 0.001, KILOGRAMS: 0.001,
  GRM: 0.000001, G: 0.000001, GRAM: 0.000001, GRAMS: 0.000001,
  LBS: 0.00045359237, LB: 0.00045359237, POUND: 0.00045359237, POUNDS: 0.00045359237,
};
const AMBIGUOUS_UNITS = ["TON", "TONS", "T"];

// V16: Default placeholder purchaser names to exclude from Top 80% ranking
const DEFAULT_PLACEHOLDERS = [
  "NONE", "OTHER", "OTHERS", "N/A", "NA", "UNKNOWN", "NOT FOUND",
  "TO THE ORDER OF", "TO ORDER", "NO DATA", "NULL", "NOT APPLICABLE",
  "NOT AVAILABLE", "UNIDENTIFIED", "UNSPECIFIED", "TBD", "TBC",
];

// ─── ENTITY RESOLUTION ENGINE (V17 Enhanced) ───

// Legal suffixes: for REMOVAL from matching key (longest-first)
const LEGAL_SUFFIXES_KEY = [
  "GMBH & CO KG","GMBH AND CO KG","S.A. DE C.V.","SA DE CV","S A DE C V","SADECV",
  "S DE RL DE CV","S DE R L DE CV","SOCIEDAD ANONIMA",
  "TRACH NHIEM HUU HAN","TRÁCH NHIỆM HỮU HẠN",
  "INCORPORATED","CORPORATION","COMPANY LIMITED","PRIVATE LIMITED",
  "SP Z OO","SP ZOO","SPZOO","SDN BHD",
  "LIMITED","COMPANY","CÔNG TY","CONG TY","COMMA V",
  "S.R.L.","S.P.A.","S.N.C.","AND CO","& CO",
  "PVT LTD","PVT. LTD.","PVT. LTD",
  "CO LTD","CO.,LTD.","CO., LTD.","CO..LTD.","CO,.LTD",
  "CORP","GMBH","SARL","OJSC","CJSC","EIRL",
  "LLC","LTD","INC","PLC","SAS","SRL","PTY","SDN","BHD",
  "JSC","OOO","ZAO","OAO","KFT","DOO","TBK","VOF",
  "TNHH","CTCP","CTY","CP",
  "CO","AG","SA","BV","NV","AB","AS","OY","OYJ","SE",
  "KG","CV","PT","S.A.","S.L.",
  "PRIVATE","PVT",
].sort((a, b) => b.length - a.length);

// Leftover fragments that should never stand alone after suffix stripping
const SUFFIX_FRAGMENTS = new Set(["P","L","D","LT","CO","LD","PV","PVT","DE","CV","SA","PRIVATE","LIMITED","LTD"]);

// Malformed suffix patterns → cleaned display form
const SUFFIX_REPAIRS = [
  [/\bINTERNACIONAL\s*SA\s*DE\s*C\.?\s*V\.?/gi, "Internacional, S.A. de C.V."],
  [/\bCO\s*[\.,]{0,3}\s*LTD[\.,]*/gi, "Co., Ltd."],
  [/\bCOMPANY\s+LIMITED(\s+LTD\.?)?/gi, "Co., Ltd."],
  [/\bPRIVATE\s+LIMITED(\s+(PVT|LTD)\.?)+/gi, "Pvt. Ltd."],
  [/\bPVT[\.\s]*LTD\.?/gi, "Pvt. Ltd."],
  [/\bS\.?\s*A\.?\s*DE\s*C\.?\s*V\.?/gi, "S.A. de C.V."],
  [/\bS\s*DE\s*R\.?\s*L\.?\s*DE\s*C\.?\s*V\.?/gi, "S. de R.L. de C.V."],
  [/\bGMBH\s*&?\s*CO\.?\s*KG\.?/gi, "GmbH & Co. KG"],
  [/\bGMBH/gi, "GmbH"],
  [/\bS\.?\s*R\.?\s*L\.?/gi, "S.R.L."],
  [/\bS\.?\s*P\.?\s*A\.?/gi, "S.p.A."],
  [/\bSDN\.?\s*BHD\.?/gi, "Sdn. Bhd."],
  [/\bSP\.?\s*Z\.?\s*O\.?\s*O\.?/gi, "Sp. z o.o."],
  [/\bLIMITED/gi, "Ltd."],
  [/\bINCORPORATED/gi, "Inc."],
  [/\bCORPORATION/gi, "Corp."],
  [/\bCOMPANY/gi, "Co."],
  [/\bLLC/gi, "LLC"],
  [/\bPLC/gi, "PLC"],
  [/\bJSC/gi, "JSC"],
  [/\bSAS/gi, "SAS"],
  [/\bINC\.?$/gi, "Inc."],
  [/\bLTD\.?$/gi, "Ltd."],
  [/\bCORP\.?$/gi, "Corp."],
];

// Vietnamese ↔ English (for matching key: replace with English or remove)
const VIET_MAP_KEY = [
  // Longest first
  ["XUAT NHAP KHAU", ""], ["XUẤT NHẬP KHẨU", ""],
  ["TRACH NHIEM HUU HAN", ""], ["TRÁCH NHIỆM HỮU HẠN", ""],
  ["THUONG MAI", ""], ["THƯƠNG MẠI", ""],
  ["SAN XUAT", ""], ["SẢN XUẤT", ""],
  ["DICH VU", ""], ["DỊCH VỤ", ""],
  ["CO PHAN", ""], ["CỔ PHẦN", ""],
  ["HOA CHAT", ""], ["HÓA CHẤT", ""],
  ["CONG TY", ""], ["CÔNG TY", ""], ["CTY", ""],
  ["TNHH", ""], ["CTCP", ""],
];

// Vietnamese → English for DISPLAY (meaningful business words preserved)
const VIET_MAP_DISPLAY = {
  "THUONG MAI": "Trading", "THƯƠNG MẠI": "Trading",
  "SAN XUAT": "Manufacturing", "SẢN XUẤT": "Manufacturing",
  "DICH VU": "Service", "DỊCH VỤ": "Service",
  "HOA CHAT": "Chemical", "HÓA CHẤT": "Chemical",
  "XUAT NHAP KHAU": "Import Export", "XUẤT NHẬP KHẨU": "Import Export",
  "CO PHAN": "", "CỔ PHẦN": "",
  "CONG TY": "", "CÔNG TY": "", "CTY": "",
  "TNHH": "", "CTCP": "",
  "TRACH NHIEM HUU HAN": "", "TRÁCH NHIỆM HỮU HẠN": "",
};

const SCHEMA_FIELDS = [
  { key: "productDesc",  label: "Product Description" },
  { key: "supplier",     label: "Supplier" },
  { key: "purchaser",    label: "Purchaser" },
  { key: "countryOrigin",label: "Country of Origin" },
  { key: "purchCountry", label: "Purchasing Country" },
  { key: "unitPrice",    label: "Unit Price" },
  { key: "totalValue",   label: "Total Value" },
  { key: "quantity",     label: "Quantity / Volume" },
  { key: "unit",         label: "Unit" },
];

/* ═══════════════════════════════════════════════
   HELPERS
   ═══════════════════════════════════════════════ */
const toNum = (v) => (v != null && v !== "" && !isNaN(Number(v)) ? Number(v) : null);
const fmt = (n) => n == null ? "—" : typeof n === "number" ? n.toLocaleString("en-US", { maximumFractionDigits: 2 }) : String(n);
const pct = (n) => n == null ? "—" : (n * 100).toFixed(2) + "%";
const safeSheetName = (s) => String(s).replace(/[:\\/?*[\]]/g, "_").slice(0, 31);

function triggerDownload(arrayBuffer, filename) {
  const blob = new Blob([arrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
}

/* ── V17: Normalize for MATCHING KEY (strips everything non-core) ── */
function normalizeForKey(name) {
  if (!name) return "";
  let s = String(name).trim().toUpperCase();
  // Remove accents
  s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  // Apply Vietnamese → remove for matching
  VIET_MAP_KEY.forEach(([from, to]) => {
    s = s.replace(new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), to);
  });
  // Remove all legal suffixes
  LEGAL_SUFFIXES_KEY.forEach((sf) => {
    const esc = sf.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    s = s.replace(new RegExp("\\b" + esc + "\\b[.,]*", "gi"), " ");
  });
  // Remove punctuation
  s = s.replace(/[.,;:'"()\-\/\\#&!@$%^*+={}[\]|<>?~`_]+/g, " ");
  s = s.replace(/\s+/g, " ").trim();
  // Remove leftover suffix fragments
  let tokens = s.split(" ").filter(t => !SUFFIX_FRAGMENTS.has(t));
  // Deduplicate consecutive tokens
  const deduped = [];
  tokens.forEach(t => { if (t && (!deduped.length || deduped[deduped.length - 1] !== t)) deduped.push(t); });
  return deduped.join(" ");
}

/* ── V17: Build clean DISPLAY name from raw name ── */
function buildDisplayName(rawName) {
  if (!rawName) return "";
  let s = String(rawName).trim();

  // Step 1: Translate Vietnamese business terms → English for display
  const sUp = s.toUpperCase();
  // Sort by longest first to avoid partial replacements
  const vietEntries = Object.entries(VIET_MAP_DISPLAY).sort((a, b) => b[0].length - a[0].length);
  vietEntries.forEach(([vn, en]) => {
    const esc = vn.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(esc, "gi");
    if (re.test(s)) s = s.replace(re, en ? " " + en + " " : " ");
  });

  // Step 2: Fix malformed suffixes
  SUFFIX_REPAIRS.forEach(([re, fix]) => {
    if (re.test(s)) s = s.replace(re, " " + fix);
  });

  // Step 3: Clean up spacing and punctuation
  s = s.replace(/\s+/g, " ").trim();
  s = s.replace(/^[,.\s]+|[,.\s]+$/g, "").trim();

  // Step 4: Remove leftover fragments at end
  let tokens = s.split(" ");
  while (tokens.length > 1 && SUFFIX_FRAGMENTS.has(tokens[tokens.length - 1].toUpperCase().replace(/[.,]/g, ""))) {
    tokens.pop();
  }
  s = tokens.join(" ");

  // Step 5: Deduplicate suffix repetitions (e.g., "Ltd. Ltd." → "Ltd.")
  s = s.replace(/\b(Ltd\.?|Co\.?|Limited|Inc\.?|Corp\.?)\s+\1/gi, "$1");
  s = s.replace(/\s+/g, " ").trim();

  // Step 6: Title Case the core name part, preserve suffix casing from repairs
  // Split on known suffix boundary
  const suffixPatterns = SUFFIX_REPAIRS.map(([_, fix]) => fix).filter(f => f.length > 1);
  let corePart = s;
  let suffixPart = "";
  for (const sfx of suffixPatterns.sort((a, b) => b.length - a.length)) {
    const idx = s.lastIndexOf(sfx);
    if (idx > 0) {
      corePart = s.slice(0, idx).trim();
      suffixPart = s.slice(idx).trim();
      break;
    }
  }
  // Title-case the core
  corePart = corePart.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  // Combine
  s = suffixPart ? corePart + " " + suffixPart : corePart;
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

/* ── V17: Detect risk flags on a raw name ── */
function detectRiskFlags(rawName, normKey) {
  const flags = [];
  const up = String(rawName || "").toUpperCase();
  // Vietnamese alias
  if (/CONG TY|CÔNG TY|CTY|TNHH|CTCP|CỔ PHẦN|THƯƠNG MẠI|THUONG MAI/i.test(rawName)) flags.push("Vietnamese-English Alias Candidate");
  // Malformed suffix
  if (/CO\s*[.,]{2,}\s*LTD|INTERNACIONALSA|LTD\s+LTD|LIMITED\s+LTD/i.test(rawName)) flags.push("Malformed Legal Suffix");
  // Short name / acronym (core key ≤ 2 chars)
  if (normKey && normKey.length <= 2 && normKey.length > 0) flags.push("Short Name / Acronym");
  // Multiple entities hint
  if (/\b(AND|&|\/)\b/i.test(rawName) && rawName.length > 30) flags.push("Multiple Entities in One Cell");
  // Leftover fragments
  if (/\b[A-Z]\b\s*$/.test(up.trim())) flags.push("Leftover Suffix Fragment");
  return flags;
}

/* ── V17: Improved similarity (Jaccard + containment) ── */
function similarity(a, b) {
  if (!a || !b) return 0;
  const A = a.toUpperCase(), B = b.toUpperCase();
  if (A === B) return 1;
  const tokA = A.split(" ").filter(Boolean), tokB = B.split(" ").filter(Boolean);
  if (!tokA.length || !tokB.length) return 0;
  const setA = new Set(tokA), setB = new Set(tokB);
  const inter = [...setA].filter(x => setB.has(x)).length;
  // Jaccard
  const jaccard = inter / (setA.size + setB.size - inter);
  // Containment: if smaller set fully contained in larger
  const smaller = setA.size <= setB.size ? setA : setB;
  const containment = inter / smaller.size;
  // Weight: use whichever is higher
  return Math.max(jaccard, containment * 0.95);
}

/* ── V17: Full entity resolution with group IDs, risk flags, review ── */
let _entityGroupCounter = 0;
function groupEntities(rows, nameCol, countryCol, prefix, extraPlaceholders) {
  _entityGroupCounter = 0;
  const phSet = new Set([...DEFAULT_PLACEHOLDERS, ...(extraPlaceholders || [])]);

  // Phase 1: Normalize every row
  const entries = rows.map((r, i) => {
    const raw = String(r[nameCol] || "").trim();
    const country = String(r[countryCol] || "").trim();
    const rawUp = raw.toUpperCase();
    // Check placeholder
    const isPlh = !raw || phSet.has(rawUp);
    const normKey = isPlh ? "" : normalizeForKey(raw);
    return { idx: i, raw, country, normKey, isPlh };
  });

  // Phase 2: Group by exact normKey + country
  const buckets = {};
  entries.forEach((e) => {
    if (e.isPlh) return; // skip placeholders
    const bk = e.normKey + "||" + e.country.toUpperCase();
    if (!buckets[bk]) buckets[bk] = { normKey: e.normKey, country: e.country, entries: [] };
    buckets[bk].entries.push(e);
  });

  // Phase 3: Fuzzy merge within same country
  const bKeys = Object.keys(buckets);
  const merged = {};
  const visited = new Set();
  bKeys.forEach((k) => {
    if (visited.has(k)) return;
    const g = buckets[k];
    const group = { normKey: g.normKey, country: g.country, entries: [...g.entries], matchMethods: new Set(["Exact_Key"]) };
    bKeys.forEach((k2) => {
      if (k2 === k || visited.has(k2)) return;
      const g2 = buckets[k2];
      if (g.country.toUpperCase() !== g2.country.toUpperCase()) return;
      const sim = similarity(g.normKey, g2.normKey);
      if (sim >= 0.7) {
        group.entries.push(...g2.entries);
        if (sim < 1) group.matchMethods.add("Fuzzy_Token_" + (sim * 100).toFixed(0) + "pct");
        else group.matchMethods.add("Vietnamese_English_Alias");
        visited.add(k2);
      }
    });
    visited.add(k);
    merged[k] = group;
  });

  // Phase 4: Assign output columns
  const stdCol = prefix + "_Standardize";
  const normCol = prefix + "_NormalizedKey";
  const groupCol = prefix + "_Entity_Group_ID";
  const methodCol = prefix + "_MatchMethod";
  const confCol = prefix + "_Confidence";
  const riskCol = prefix + "_Risk_Flag";
  const reviewCol = prefix + "_Review_Status";
  const noteCol = prefix + "_Standardization_Note";

  // Handle placeholder rows
  entries.filter(e => e.isPlh).forEach((e) => {
    const r = rows[e.idx];
    r[stdCol] = e.raw || "(Placeholder)";
    r[normCol] = "";
    r[groupCol] = "";
    r[methodCol] = "Placeholder";
    r[confCol] = "N/A";
    r[riskCol] = "Placeholder Name";
    r[reviewCol] = "Excluded";
    r[noteCol] = "Excluded from entity resolution";
  });

  // Handle real entity groups
  Object.values(merged).forEach((g) => {
    _entityGroupCounter++;
    const gid = prefix.charAt(0) + String(_entityGroupCounter).padStart(4, "0");
    // Pick best raw name: longest non-Vietnamese-prefix name, or longest overall
    const rawNames = g.entries.map(e => e.raw);
    const uniqueRaws = [...new Set(rawNames)];
    // Prefer names without CONG TY/CTY prefix
    const nonViet = uniqueRaws.filter(n => !/^(CONG TY|CÔNG TY|CTY)\s/i.test(n));
    const candidates = nonViet.length ? nonViet : uniqueRaws;
    const bestRaw = candidates.sort((a, b) => b.length - a.length)[0];
    const displayName = buildDisplayName(bestRaw);
    const methods = [...g.matchMethods].join("; ");
    const isMultiSource = uniqueRaws.length > 1;
    const confidence = isMultiSource ? (methods.includes("Fuzzy") ? "Medium" : "High") : "Exact";
    const risks = detectRiskFlags(bestRaw, g.normKey);
    if (isMultiSource && methods.includes("Fuzzy")) risks.push("Possible Duplicate Entity");
    const needsReview = confidence === "Medium" || risks.length > 1;

    g.entries.forEach((e) => {
      const r = rows[e.idx];
      r[stdCol] = displayName;
      r[normCol] = g.normKey;
      r[groupCol] = gid;
      r[methodCol] = methods;
      r[confCol] = confidence;
      r[riskCol] = risks.join("; ") || "";
      r[reviewCol] = needsReview ? "Needs_Review" : "Auto_Accepted";
      r[noteCol] = uniqueRaws.length > 1 ? "Merged " + uniqueRaws.length + " variants" : "";
    });
  });

  return rows;
}

function deepCopyRows(rows) {
  return rows.map((r) => ({ ...r }));
}

function isPlaceholder(rawName, stdName, extraList) {
  const checkSet = new Set([...DEFAULT_PLACEHOLDERS, ...(extraList || [])]);
  const rawNorm = String(rawName || "").trim().toUpperCase();
  const stdNorm = String(stdName || "").trim().toUpperCase();
  if (!rawNorm && !stdNorm) return "Blank/Missing";
  if (checkSet.has(rawNorm)) return rawNorm;
  if (checkSet.has(stdNorm)) return stdNorm;
  return null; // not a placeholder
}

/* ═══════════════════════════════════════════════
   UI COMPONENTS
   ═══════════════════════════════════════════════ */
const Card = ({ title, children, accent }) => (
  <div style={{
    background: "#fff", borderRadius: 14, padding: "28px 36px",
    boxShadow: "0 2px 20px rgba(0,0,0,0.06)", borderLeft: "4px solid " + (accent || "#2563eb"),
    marginBottom: 20, width: "100%",
  }}>
    {title && <h3 style={{ margin: "0 0 16px", fontSize: 17, fontWeight: 700, color: "#0f172a", letterSpacing: "-0.02em" }}>{title}</h3>}
    {children}
  </div>
);

const StatBox = ({ label, value, sub, color }) => (
  <div style={{
    background: (color || "#e2e8f0") + "12", borderRadius: 12, padding: "18px 24px",
    minWidth: 160, flex: "1 1 160px", border: "1px solid " + (color || "#e2e8f0") + "33",
  }}>
    <div style={{ fontSize: 26, fontWeight: 800, color: color || "#2563eb", letterSpacing: "-0.03em" }}>{value}</div>
    <div style={{ fontSize: 11, color: "#475569", marginTop: 5, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>{label}</div>
    {sub && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 3 }}>{sub}</div>}
  </div>
);

const Btn = ({ children, onClick, disabled, primary, big }) => (
  <button onClick={onClick} disabled={disabled} style={{
    padding: big ? "14px 40px" : "12px 32px", borderRadius: big ? 12 : 10, border: "none",
    fontWeight: big ? 800 : 700, fontSize: big ? 15 : 14, fontFamily: "inherit",
    background: disabled ? "#94a3b8" : primary !== false ? (big ? "linear-gradient(135deg,#059669,#10b981)" : "#2563eb") : "transparent",
    color: primary !== false ? "#fff" : "#2563eb",
    cursor: disabled ? "not-allowed" : "pointer",
    border: primary === false ? "2px solid #2563eb" : "none",
    boxShadow: big ? "0 4px 20px rgba(16,185,129,0.3)" : "none",
    transition: "all 0.2s",
  }}>
    {children}
  </button>
);

const MiniTable = ({ headers, rows, maxH }) => (
  <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: maxH || 400, borderRadius: 10, border: "1px solid #e2e8f0", width: "100%" }}>
    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
      <thead>
        <tr>{headers.map((h, i) => (
          <th key={i} style={{
            padding: "11px 14px", textAlign: "left", background: "#f8fafc", color: "#0f172a",
            fontWeight: 700, fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.05em",
            position: "sticky", top: 0, borderBottom: "2px solid #e2e8f0", whiteSpace: "nowrap", zIndex: 1,
          }}>{h}</th>
        ))}</tr>
      </thead>
      <tbody>{rows.map((row, ri) => (
        <tr key={ri} style={{ background: ri % 2 === 0 ? "transparent" : "#fafbfc" }}>
          {row.map((cell, ci) => (
            <td key={ci} style={{
              padding: "9px 14px", borderBottom: "1px solid #f1f5f9",
              color: "#0f172a", whiteSpace: "nowrap", maxWidth: 340, overflow: "hidden", textOverflow: "ellipsis",
            }}>{cell != null ? String(cell) : ""}</td>
          ))}
        </tr>
      ))}</tbody>
    </table>
  </div>
);

const DropZone = ({ icon, text, inputRef, accept, onChange }) => (
  <div style={{width:"100%"}}>
    <div onClick={() => inputRef.current && inputRef.current.click()}
      style={{
        display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
        padding: "52px 24px", border: "2px dashed #cbd5e1", borderRadius: 14,
        cursor: "pointer", transition: "all 0.2s", background: "#f8fafc", width: "100%",
      }}
      onDragOver={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = "#2563eb"; }}
      onDragLeave={(e) => { e.currentTarget.style.borderColor = "#cbd5e1"; }}
      onDrop={(e) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = "#cbd5e1";
        const f = e.dataTransfer.files[0];
        if (f) onChange({ target: { files: [f] } });
      }}
    >
      <div style={{ fontSize: 52, marginBottom: 14 }}>{icon}</div>
      <div style={{ fontSize: 15, fontWeight: 600, color: "#0f172a", marginBottom: 6 }}>{text}</div>
      <div style={{ fontSize: 13, color: "#94a3b8" }}>Click or drag & drop</div>
    </div>
    <input ref={inputRef} type="file" accept={accept} onChange={onChange}
      style={{ position: "absolute", width: 0, height: 0, opacity: 0, pointerEvents: "none" }} />
  </div>
);

const InfoBar = ({ type, children }) => {
  const colors = { info: ["#eff6ff", "#1d4ed8"], warn: ["#fffbeb", "#b45309"], success: ["#f0fdf4", "#166534"], error: ["#fef2f2", "#991b1b"] };
  const [bg, fg] = colors[type] || colors.info;
  return (
    <div style={{ padding: "12px 20px", background: bg, borderRadius: 10, marginBottom: 14, fontSize: 13, color: fg, fontWeight: 600 }}>
      {children}
    </div>
  );
};


/* ═══════════════════════════════════════════════
   MAIN COMPONENT
   ═══════════════════════════════════════════════ */
export default function ChemicalSegmentationTool() {
  const [step, setStep] = useState(0);
  const [yearSheets, setYearSheets] = useState([]);
  const [sheetData, setSheetData] = useState({});
  const [schema, setSchema] = useState({});

  // Processing results — empty {} means "not yet run"
  const [qualityReport, setQualityReport] = useState(null);
  const [backfillLog, setBackfillLog] = useState(null);
  const [validationLog, setValidationLog] = useState(null);
  const [entityLog, setEntityLog] = useState(null);
  const [rankingData, setRankingData] = useState(null);
  const [industryLog, setIndustryLog] = useState(null);
  const [keywordMatrix, setKeywordMatrix] = useState(null);
  const [keywordConflicts, setKeywordConflicts] = useState([]);
  const [classificationLog, setClassificationLog] = useState(null);
  const [conversionLog, setConversionLog] = useState(null);
  const [iqrResults, setIqrResults] = useState(null);
  const [segmentMetrics, setSegmentMetrics] = useState(null);

  const [processing, setProcessing] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [customPlaceholders, setCustomPlaceholders] = useState([]); // V16: user-configurable extra placeholders
  const [placeholderInput, setPlaceholderInput] = useState(""); // V16: text input for adding custom placeholders
  const fileRef = useRef(null);
  const indRef = useRef(null);
  const kwRef = useRef(null);

  /* ─── columns from first sheet ─── */
  const allCols = useMemo(() => {
    if (!yearSheets.length || !sheetData[yearSheets[0]]) return [];
    return Object.keys(sheetData[yearSheets[0]][0] || {});
  }, [yearSheets, sheetData]);

  const schemaComplete = useMemo(() => SCHEMA_FIELDS.every((f) => schema[f.key]), [schema]);

  /* ════════════════════════════════════
     STEP 0 — UPLOAD
     ════════════════════════════════════ */
  const handleUpload = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setProcessing(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const yrs = wb.SheetNames.filter((n) => /\b(19|20)\d{2}\b/.test(n));
        if (!yrs.length) { alert("No year-based sheets found. Sheet names must contain a year (e.g. 2022, 2023)."); setProcessing(false); return; }
        const data = {};
        yrs.forEach((yr) => { data[yr] = XLSX.utils.sheet_to_json(wb.Sheets[yr], { defval: "" }); });
        setYearSheets(yrs);
        setSheetData(data);
        // Auto-detect schema
        const cols = Object.keys(data[yrs[0]][0] || {});
        const maps = {
          productDesc: ["product description","product desc","description","product","product name","hàng hóa","hang hoa"],
          supplier: ["supplier","exporter","seller","shipper","nhà cung cấp"],
          purchaser: ["purchaser","buyer","importer","consignee","người mua"],
          countryOrigin: ["country of origin","origin country","origin","exporting country","nước xuất xứ"],
          purchCountry: ["purchasing country","importing country","destination","buyer country","nước nhập khẩu"],
          unitPrice: ["unit price","price per unit","price","đơn giá"],
          totalValue: ["total value","total amount","value","amount","trị giá"],
          quantity: ["quantity","volume","qty","net weight","weight","khối lượng","số lượng"],
          unit: ["unit","uom","unit of measure","đơn vị"],
        };
        const auto = {};
        Object.entries(maps).forEach(([key, aliases]) => {
          const m = cols.find((c) => aliases.some((a) => c.toLowerCase().includes(a)));
          if (m) auto[key] = m;
        });
        setSchema(auto);
        // Reset downstream
        setQualityReport(null); setBackfillLog(null); setValidationLog(null);
        setEntityLog(null); setRankingData(null); setIndustryLog(null);
        setKeywordMatrix(null); setKeywordConflicts([]); setClassificationLog(null);
        setConversionLog(null); setIqrResults(null); setSegmentMetrics(null);
        setExportDone(false);
        setStep(1);
      } catch (err) { alert("Error reading file: " + err.message); }
      setProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  /* ════════════════════════════════════
     STEP 2 — DATA QUALITY
     ════════════════════════════════════ */
  const runQuality = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const report = {};
      const textKeys = ["productDesc","supplier","purchaser","countryOrigin","purchCountry","unit"];
      const numKeys = ["unitPrice","totalValue","quantity"];
      yearSheets.forEach((yr) => {
        const rows = sheetData[yr];
        const yr_r = { text: {}, numeric: {}, totalRows: rows.length };
        textKeys.forEach((f) => {
          const col = schema[f]; if (!col) return;
          let blank=0,none=0,other=0,filled=0;
          rows.forEach((r) => {
            const v = String(r[col]||"").trim();
            if (!v) blank++;
            else if (["NONE","N/A","NA","-"].includes(v.toUpperCase())) none++;
            else if (["OTHER","OTHERS"].includes(v.toUpperCase())) other++;
            else filled++;
          });
          yr_r.text[f] = {blank,none,other,filled};
        });
        numKeys.forEach((f) => {
          const col = schema[f]; if (!col) return;
          let missing=0,zero=0,negative=0,invalid=0,textPh=0,valid=0;
          rows.forEach((r) => {
            const v = r[col]; const s = String(v??"").trim();
            if (s==="" || v==null) missing++;
            else if (isNaN(Number(s))) { /[a-zA-Z]/.test(s) ? textPh++ : invalid++; }
            else { const n=Number(s); n===0?zero++:n<0?negative++:valid++; }
          });
          yr_r.numeric[f] = {missing,zero,negative,invalid,textPh,valid};
        });
        report[yr] = yr_r;
      });
      setQualityReport(report);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData, schema]);

  /* ════════════════════════════════════
     STEP 3 — NUMERIC BACKFILL
     ════════════════════════════════════ */
  const runBackfill = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const log = {}; const newData = {};
      yearSheets.forEach((yr) => {
        const rows = deepCopyRows(sheetData[yr]);
        let bf = 0;
        rows.forEach((r) => {
          const up=toNum(r[schema.unitPrice]), tv=toNum(r[schema.totalValue]), q=toNum(r[schema.quantity]);
          const h = [up>0, tv>0, q>0];
          const miss = h.filter(x=>!x).length;
          if (miss===1) {
            if (!h[0] && tv>0 && q>0)     { r[schema.unitPrice]=tv/q; r._backfill="UnitPrice"; bf++; }
            else if (!h[1] && up>0 && q>0) { r[schema.totalValue]=up*q; r._backfill="TotalValue"; bf++; }
            else if (!h[2] && up>0 && tv>0){ r[schema.quantity]=tv/up; r._backfill="Quantity"; bf++; }
          }
        });
        newData[yr] = rows;
        log[yr] = { backfilled: bf, total: rows.length };
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setBackfillLog(log);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData, schema]);

  /* ════════════════════════════════════
     STEP 4 — VALUE VALIDATION
     ════════════════════════════════════ */
  const runValidation = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const log = {}; const newData = {};
      yearSheets.forEach((yr) => {
        const rows = deepCopyRows(sheetData[yr]);
        let flagged = 0;
        rows.forEach((r) => {
          const up=toNum(r[schema.unitPrice]), tv=toNum(r[schema.totalValue]), q=toNum(r[schema.quantity]);
          if (up>0 && q>0 && tv>0) {
            const diff = Math.abs(up*q - tv);
            r.Pre_Conversion_Value_Check = diff > 500 ? "Outlier" : "OK";
            if (diff>500) flagged++;
          } else { r.Pre_Conversion_Value_Check = "Insufficient_Data"; }
        });
        newData[yr] = rows;
        log[yr] = { flagged, total: rows.length };
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setValidationLog(log);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData, schema]);

  /* ════════════════════════════════════
     STEP 5 — ENTITY RESOLUTION (V17 Enhanced)
     ════════════════════════════════════ */
  const runEntity = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const log = {}; const newData = {};
      yearSheets.forEach((yr) => {
        let rows = deepCopyRows(sheetData[yr]);
        rows = groupEntities(rows, schema.supplier, schema.countryOrigin, "Supplier", customPlaceholders);
        rows = groupEntities(rows, schema.purchaser, schema.purchCountry, "Purchaser", customPlaceholders);

        // Collect stats
        const suppStd = new Set(), purchStd = new Set();
        let suppReview = 0, purchReview = 0, suppPlh = 0, purchPlh = 0;
        const riskSamples = [];
        rows.forEach(r => {
          if (r.Supplier_Standardize) suppStd.add(r.Supplier_Standardize);
          if (r.Purchaser_Standardize) purchStd.add(r.Purchaser_Standardize);
          if (r.Supplier_Review_Status === "Needs_Review") suppReview++;
          if (r.Purchaser_Review_Status === "Needs_Review") purchReview++;
          if (r.Supplier_Review_Status === "Excluded") suppPlh++;
          if (r.Purchaser_Review_Status === "Excluded") purchPlh++;
          // Collect a few risk-flagged samples for display
          if (r.Purchaser_Risk_Flag && riskSamples.length < 15) {
            const key = r.Purchaser_Standardize + "||" + (r.Purchaser_Risk_Flag || "");
            if (!riskSamples.find(s => s.key === key)) {
              riskSamples.push({
                key, raw: r[schema.purchaser], std: r.Purchaser_Standardize,
                normKey: r.Purchaser_NormalizedKey, risk: r.Purchaser_Risk_Flag,
                method: r.Purchaser_MatchMethod, review: r.Purchaser_Review_Status,
              });
            }
          }
        });
        newData[yr] = rows;
        log[yr] = {
          suppGroups: suppStd.size, purchGroups: purchStd.size, total: rows.length,
          suppReview, purchReview, suppPlh, purchPlh, riskSamples,
        };
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setEntityLog(log);
      setProcessing(false);
    }, 80);
  }, [yearSheets, sheetData, schema, customPlaceholders]);

  /* ════════════════════════════════════
     STEP 6 — TOP 80% RANKING (V16: Placeholder Exclusion)
     ════════════════════════════════════ */
  const runRanking = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const ranking = {}; const newData = {};
      yearSheets.forEach((yr) => {
        const rows = deepCopyRows(sheetData[yr]);

        // --- Phase 1: Identify placeholders & compute row ranking values ---
        let rawAnnualTotal = 0;
        let placeholderValue = 0, placeholderRows = 0;
        let placeholderNoneValue = 0, placeholderOtherValue = 0;
        let validRealTotal = 0, notCalc = 0, notCalcRealRows = 0;
        const purchMap = {};

        rows.forEach((r) => {
          const rawTV = toNum(r[schema.totalValue]), rawUP = toNum(r[schema.unitPrice]), rawQ = toNum(r[schema.quantity]);
          // Compute Row_Ranking_Value
          let rrv = null, src = "";
          if (rawTV > 0) { rrv = rawTV; src = "Raw_TotalValue"; }
          else if (rawUP > 0 && rawQ > 0) { rrv = rawUP * rawQ; src = "Calculated_UP_x_Q"; }
          r.Row_Ranking_Value = rrv;
          r.Row_Ranking_Value_Source = src || "Not_Calculable";
          if (rrv > 0) rawAnnualTotal += rrv;

          // V16: Check placeholder
          const phReason = isPlaceholder(r[schema.purchaser], r.Purchaser_Standardize, customPlaceholders);
          if (phReason) {
            r.Purchaser_Ranking_Eligibility = "Excluded_Placeholder";
            r.Purchaser_Ranking_Exclusion_Reason = phReason;
            if (rrv > 0) {
              placeholderValue += rrv;
              placeholderRows++;
              const upper = phReason.toUpperCase();
              if (upper === "NONE") placeholderNoneValue += rrv;
              else if (upper === "OTHER" || upper === "OTHERS") placeholderOtherValue += rrv;
            }
            return; // skip from ranking map
          }

          r.Purchaser_Ranking_Eligibility = "Eligible";
          r.Purchaser_Ranking_Exclusion_Reason = "";

          if (rrv === null) { notCalc++; notCalcRealRows++; r.Ranking_Value_Issue_Flag = "Not_Calculable"; return; }
          r.Ranking_Value_Issue_Flag = "";

          // Aggregate by purchaser
          const key = (r.Purchaser_Standardize || "") + "||" + (r[schema.purchCountry] || "").trim().toUpperCase();
          if (!purchMap[key]) purchMap[key] = { std: r.Purchaser_Standardize || "", country: (r[schema.purchCountry] || "").trim(), raws: new Set(), value: 0 };
          purchMap[key].raws.add(r[schema.purchaser]);
          purchMap[key].value += rrv;
          validRealTotal += rrv;
        });

        // --- Phase 2: Rank eligible real purchasers ---
        const purchList = Object.values(purchMap).sort((a, b) => b.value - a.value);
        let cum = 0; const threshold = validRealTotal * 0.8; const selNames = [];
        purchList.forEach((p, idx) => {
          p.rank = idx + 1;
          p.share = validRealTotal > 0 ? p.value / validRealTotal : 0;
          cum += p.value;
          p.cumValue = cum;
          p.cumShare = validRealTotal > 0 ? cum / validRealTotal : 0;
          p.scope = (cum - p.value) < threshold ? "Top_80pct_Value" : "Below_80pct";
          p.eligibility = "Eligible";
          if (p.scope === "Top_80pct_Value") selNames.push(p.std);
        });

        // --- Phase 3: Assign flags to rows ---
        const selSet = new Set(selNames);
        rows.forEach((r) => {
          if (r.Purchaser_Ranking_Eligibility === "Excluded_Placeholder") {
            r.Research_Scope_Flag = "Placeholder_Excluded";
            r.Top80_Research_Flag = "No";
          } else if (r.Ranking_Value_Issue_Flag === "Not_Calculable") {
            r.Research_Scope_Flag = "Value_Not_Calculable";
            r.Top80_Research_Flag = "No";
          } else if (selSet.has(r.Purchaser_Standardize)) {
            r.Research_Scope_Flag = "Top_80pct_Value";
            r.Top80_Research_Flag = "Yes";
          } else {
            r.Research_Scope_Flag = "Below_80pct";
            r.Top80_Research_Flag = "No";
          }
          r.Valid_Year_Ranking_Value_Total = validRealTotal;
        });

        newData[yr] = rows;
        ranking[yr] = {
          purchList, validRankVal: validRealTotal, rawAnnualTotal,
          notCalc: notCalcRealRows, selectedCount: selNames.length,
          totalPurchasers: purchList.length,
          // V16 summary fields
          placeholderValue, placeholderRows,
          placeholderNoneValue, placeholderOtherValue,
          excludedPlaceholderCount: placeholderRows,
        };
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setRankingData(ranking);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData, schema, customPlaceholders]);

  /* ─── Export Research List (V16) ─── */
  const exportResearchList = useCallback(() => {
    if (!rankingData) return;
    const wb = XLSX.utils.book_new();
    yearSheets.forEach((yr) => {
      const rd = rankingData[yr]; if (!rd) return;
      const rows = rd.purchList.filter(p => p.scope === "Top_80pct_Value").map(p => ({
        Year: yr, Raw_Purchaser_Names: [...p.raws].join("; "), Purchaser_Standardize: p.std,
        Purchasing_Country: p.country,
        Purchaser_Ranking_Eligibility: p.eligibility,
        Purchaser_Ranking_Value_Year: p.value,
        Valid_Year_Ranking_Value_Total: rd.validRankVal,
        Purchaser_Value_Share: p.share,
        Cumulative_Value_Share: p.cumShare,
        Purchaser_Value_Rank: p.rank,
        Top80_Research_Flag: "Yes",
        Industry: "", Industry_Segment: "",
      }));
      if (rows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), safeSheetName(yr));
    });
    triggerDownload(XLSX.write(wb, { bookType: "xlsx", type: "array" }), "Purchaser_Research_List.xlsx");
  }, [yearSheets, rankingData]);

  /* ════════════════════════════════════
     STEP 7 — INDUSTRY MASTER UPLOAD
     ════════════════════════════════════ */
  const handleIndustryUpload = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setProcessing(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        const cols = Object.keys(data[0] || {});
        const reqd = ["Purchaser_Standardize","Purchasing Country","Industry","Industry Segment"];
        const miss = reqd.filter(r => !cols.some(c => c.trim().toLowerCase() === r.toLowerCase()));
        if (miss.length) { alert("Industry Master missing columns:\n" + miss.join(", ")); setProcessing(false); return; }
        const log = {}; const newData = {};
        yearSheets.forEach((yr) => {
          const rows = deepCopyRows(sheetData[yr]);
          let matched=0, unmatched=0;
          rows.forEach((r) => {
            if (r.Research_Scope_Flag !== "Top_80pct_Value") { r.Industry=""; r.Industry_Segment=""; unmatched++; return; }
            const m = data.find(d =>
              (d["Purchaser_Standardize"]||"").trim().toUpperCase() === (r.Purchaser_Standardize||"").toUpperCase() &&
              (d["Purchasing Country"]||"").trim().toUpperCase() === (r[schema.purchCountry]||"").trim().toUpperCase()
            );
            if (m) { r.Industry = m.Industry||""; r.Industry_Segment = m["Industry Segment"]||""; matched++; }
            else { r.Industry=""; r.Industry_Segment=""; unmatched++; }
          });
          newData[yr] = rows;
          log[yr] = { matched, unmatched };
        });
        setSheetData((prev) => ({ ...prev, ...newData }));
        setIndustryLog(log);
      } catch (err) { alert("Error: " + err.message); }
      setProcessing(false);
      // Reset ref so same file can be re-selected
      if (indRef.current) indRef.current.value = "";
    };
    reader.readAsArrayBuffer(file);
  }, [sheetData, yearSheets, schema]);

  /* ════════════════════════════════════
     STEP 8 — KEYWORD MATRIX UPLOAD
     ════════════════════════════════════ */
  const handleKeywordUpload = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setProcessing(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        const segCols = Object.keys(data[0]||{}).filter(c => !["stt","no","#",""].includes(c.toLowerCase().trim()));
        if (!segCols.length) { alert("No segment columns found."); setProcessing(false); return; }
        const matrix = {}; const allKws = {}; const conflicts = [];
        segCols.forEach((seg) => {
          matrix[seg] = [];
          data.forEach((row) => {
            const kw = String(row[seg]||"").trim();
            if (kw) {
              const kwUp = kw.toUpperCase();
              matrix[seg].push(kwUp);
              if (allKws[kwUp] && allKws[kwUp] !== seg) conflicts.push({ keyword: kw, segments: [allKws[kwUp], seg] });
              allKws[kwUp] = seg;
            }
          });
        });
        setKeywordMatrix(matrix);
        setKeywordConflicts(conflicts);
      } catch (err) { alert("Error: " + err.message); }
      setProcessing(false);
      if (kwRef.current) kwRef.current.value = "";
    };
    reader.readAsArrayBuffer(file);
  }, []);

  /* ════════════════════════════════════
     STEP 9 — CLASSIFICATION
     ════════════════════════════════════ */
  const runClassification = useCallback(() => {
    if (!keywordMatrix) return;
    setProcessing(true);
    setTimeout(() => {
      const log = {}; const newData = {};
      yearSheets.forEach((yr) => {
        const rows = deepCopyRows(sheetData[yr]);
        let byKw=0, byInd=0, unclass=0;
        rows.forEach((r) => {
          const desc = (r[schema.productDesc]||"").toUpperCase();
          const kwMatches = [];
          Object.entries(keywordMatrix).forEach(([seg, kws]) => {
            const found = kws.filter(k => desc.includes(k));
            if (found.length) kwMatches.push({ seg, count: found.length, kws: found });
          });
          kwMatches.sort((a,b) => b.count - a.count);
          const kwSeg = kwMatches.length ? kwMatches[0].seg : null;
          r.Keyword_Segment = kwSeg || "";
          r.Keyword_Match_Detail = kwMatches.map(m => m.seg+"("+m.kws.join("+")+")").join("; ");
          if (r.Research_Scope_Flag === "Top_80pct_Value" && r.Industry_Segment) {
            if (kwSeg) { r.Final_Segment=kwSeg; r.Segment_Method="Keyword"; byKw++; }
            else { r.Final_Segment=r.Industry_Segment; r.Segment_Method="Industry_Fallback"; byInd++; }
          } else {
            if (kwSeg) { r.Final_Segment=kwSeg; r.Segment_Method="Keyword_Only"; byKw++; }
            else { r.Final_Segment="Unclassified"; r.Segment_Method="None"; unclass++; }
          }
        });
        newData[yr] = rows;
        log[yr] = { byKeyword: byKw, byIndustry: byInd, unclassified: unclass, total: rows.length };
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setClassificationLog(log);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData, schema, keywordMatrix]);

  /* ════════════════════════════════════
     STEP 10 — UNIT CONVERSION
     ════════════════════════════════════ */
  const runConversion = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const log = {}; const newData = {};
      yearSheets.forEach((yr) => {
        const rows = deepCopyRows(sheetData[yr]);
        let conv=0, ambig=0, unsup=0;
        rows.forEach((r) => {
          const rawUnit = String(r[schema.unit]||"").trim().toUpperCase();
          const cQ=toNum(r[schema.quantity]), cTV=toNum(r[schema.totalValue]), cUP=toNum(r[schema.unitPrice]);
          r.TotalValue_Clean=cTV; r.Quantity_Clean=cQ; r.UnitPrice_Clean=cUP;
          if (AMBIGUOUS_UNITS.includes(rawUnit)) {
            r.Unit_Status="Ambiguous"; r.Conversion_Factor=null; r.Quantity_MT=null; r.UnitPrice_per_MT=null; ambig++; return;
          }
          const factor = UNIT_FACTORS[rawUnit];
          if (factor != null) {
            r.Unit_Status="Converted"; r.Conversion_Factor=factor;
            if (cQ>0) {
              r.Quantity_MT = cQ * factor;
              if (cTV>0) r.UnitPrice_per_MT = cTV / r.Quantity_MT;
              else if (cUP>0) r.UnitPrice_per_MT = cUP / factor;
              else r.UnitPrice_per_MT = null;
            } else { r.Quantity_MT=null; r.UnitPrice_per_MT=null; }
            conv++;
          } else { r.Unit_Status="Unsupported"; r.Conversion_Factor=null; r.Quantity_MT=null; r.UnitPrice_per_MT=null; unsup++; }
        });
        newData[yr] = rows;
        log[yr] = { converted: conv, ambiguous: ambig, unsupported: unsup, total: rows.length };
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setConversionLog(log);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData, schema]);

  /* ════════════════════════════════════
     STEP 11 — IQR
     ════════════════════════════════════ */
  const runIQR = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const results = {}; const newData = {};
      yearSheets.forEach((yr) => {
        const rows = deepCopyRows(sheetData[yr]);
        const segGroups = {};
        rows.forEach((r,i) => {
          const seg = r.Final_Segment || "Unclassified";
          if (!segGroups[seg]) segGroups[seg] = [];
          if (r.UnitPrice_per_MT > 0) segGroups[seg].push({ idx:i, price: r.UnitPrice_per_MT });
        });
        const yrRes = {};
        Object.entries(segGroups).forEach(([seg, items]) => {
          if (items.length < 4) {
            items.forEach(({idx}) => { rows[idx].IQR_Segment_Check="Insufficient_Data"; rows[idx].IQR_Method="N/A"; rows[idx].IQR_Note="Only "+items.length+" values"; });
            yrRes[seg] = { count:items.length, q1:null, q3:null, iqr:null, lower:null, upper:null, outliers:0, note:"Insufficient" };
            return;
          }
          const sorted = items.map(x=>x.price).sort((a,b)=>a-b);
          const n=sorted.length;
          const q1=sorted[Math.floor(n*0.25)], q3=sorted[Math.floor(n*0.75)];
          const iqr=q3-q1, lower=q1-1.5*iqr, upper=q3+1.5*iqr;
          const wide = upper > 0 && lower > 0 && upper/lower > 100;
          let outliers = 0;
          items.forEach(({idx, price}) => {
            const isOut = price < lower || price > upper;
            rows[idx].IQR_Q1=q1; rows[idx].IQR_Q3=q3; rows[idx].IQR_IQR=iqr;
            rows[idx].IQR_Lower=lower; rows[idx].IQR_Upper=upper;
            rows[idx].IQR_Segment_Check = isOut ? "IQR Outlier" : "Within_Range";
            rows[idx].IQR_Method = "1.5xIQR";
            rows[idx].IQR_Note = wide ? "IQR_Range_Too_Wide" : "";
            if (isOut) outliers++;
          });
          yrRes[seg] = { count:n, q1, q3, iqr, lower, upper, outliers, note: wide?"Range_Too_Wide":"" };
        });
        rows.forEach(r => { if (!r.IQR_Segment_Check) r.IQR_Segment_Check = "No_Price_Data"; });
        newData[yr] = rows;
        results[yr] = yrRes;
      });
      setSheetData((prev) => ({ ...prev, ...newData }));
      setIqrResults(results);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData]);

  /* ════════════════════════════════════
     STEP 12 — SEGMENT METRICS
     ════════════════════════════════════ */
  const runMetrics = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      const metrics = {};
      yearSheets.forEach((yr) => {
        const rows = sheetData[yr]; const segMap = {};
        rows.forEach((r) => {
          const seg = r.Final_Segment || "Unclassified";
          if (!segMap[seg]) segMap[seg] = { before:{sumTV:0,sumQty:0,count:0}, after:{sumTV:0,sumQty:0,count:0} };
          const tv=toNum(r.TotalValue_Clean), qty=toNum(r.Quantity_MT);
          if (tv>0 && qty>0) { segMap[seg].before.sumTV+=tv; segMap[seg].before.sumQty+=qty; segMap[seg].before.count++; }
          const excluded = r.Pre_Conversion_Value_Check==="Outlier" && r.IQR_Segment_Check==="IQR Outlier";
          if (tv>0 && qty>0 && !excluded) { segMap[seg].after.sumTV+=tv; segMap[seg].after.sumQty+=qty; segMap[seg].after.count++; }
        });
        const sm = {};
        Object.entries(segMap).forEach(([seg,d]) => {
          sm[seg] = {
            before: { ...d.before, avgPrice: d.before.sumQty>0 ? d.before.sumTV/d.before.sumQty : 0 },
            after:  { ...d.after,  avgPrice: d.after.sumQty>0  ? d.after.sumTV/d.after.sumQty   : 0 },
          };
        });
        metrics[yr] = sm;
      });
      setSegmentMetrics(metrics);
      setProcessing(false);
    }, 50);
  }, [yearSheets, sheetData]);

  /* ════════════════════════════════════
     STEP 13 — EXPORT
     ════════════════════════════════════ */
  const handleExport = useCallback(() => {
    setProcessing(true);
    setTimeout(() => {
      try {
        const wb = XLSX.utils.book_new();
        // 00_Summary (V16: includes placeholder exclusion audit)
        const sum = yearSheets.map(yr => ({
          Year: yr, Total_Rows: sheetData[yr].length,
          Backfilled: backfillLog?.[yr]?.backfilled||0, Value_Flagged: validationLog?.[yr]?.flagged||0,
          Supplier_Groups: entityLog?.[yr]?.suppGroups||0, Purchaser_Groups: entityLog?.[yr]?.purchGroups||0,
          Raw_Annual_Total_Value: rankingData?.[yr]?.rawAnnualTotal||0,
          Valid_Ranking_Denominator: rankingData?.[yr]?.validRankVal||0,
          Excluded_NONE_Value: rankingData?.[yr]?.placeholderNoneValue||0,
          Excluded_OTHER_Value: rankingData?.[yr]?.placeholderOtherValue||0,
          Excluded_All_Placeholder_Value: rankingData?.[yr]?.placeholderValue||0,
          Excluded_Placeholder_Rows: rankingData?.[yr]?.placeholderRows||0,
          Top80_Purchasers: rankingData?.[yr]?.selectedCount||0,
          Industry_Matched: industryLog?.[yr]?.matched||0,
          Classified_KW: classificationLog?.[yr]?.byKeyword||0, Classified_Ind: classificationLog?.[yr]?.byIndustry||0,
          Unclassified: classificationLog?.[yr]?.unclassified||0, Units_Converted: conversionLog?.[yr]?.converted||0,
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sum), "00_Summary");

        // Data Quality
        if (qualityReport) {
          const dq = [];
          yearSheets.forEach(yr => { const q=qualityReport[yr]; if(!q) return;
            Object.entries(q.numeric).forEach(([f,v]) => dq.push({Year:yr,Field:f,...v}));
          });
          if (dq.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dq), "DataQuality_Summary");
        }
        // Backfill Log
        const bf = [];
        yearSheets.forEach(yr => sheetData[yr].forEach((r,i) => { if (r._backfill) bf.push({Year:yr,Row:i+2,Backfilled:r._backfill}); }));
        if (bf.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bf), "Backfill_Log");

        // Entity Resolution Log (V17: detailed with review items)
        if (entityLog) {
          const el = yearSheets.map(yr => ({
            Year: yr, Supplier_Groups: entityLog[yr]?.suppGroups, Purchaser_Groups: entityLog[yr]?.purchGroups,
            Total_Rows: entityLog[yr]?.total,
            Supplier_Needs_Review: entityLog[yr]?.suppReview, Purchaser_Needs_Review: entityLog[yr]?.purchReview,
            Supplier_Placeholders: entityLog[yr]?.suppPlh, Purchaser_Placeholders: entityLog[yr]?.purchPlh,
          }));
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(el), "Entity_Log");
          // Entity review items
          const reviewRows = [];
          yearSheets.forEach(yr => {
            const samples = entityLog[yr]?.riskSamples || [];
            samples.forEach(s => reviewRows.push({ Year: yr, ...s }));
          });
          if (reviewRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(reviewRows), "Entity_Review");
        }
        // Purchaser Research Scope (V16: with eligibility + exclusion audit)
        if (rankingData) {
          yearSheets.forEach(yr => {
            const rd=rankingData[yr]; if(!rd) return;
            const sc = rd.purchList.filter(p=>p.scope==="Top_80pct_Value").map(p => ({
              Year:yr, Raw_Names:[...p.raws].join("; "), Purchaser_Std:p.std, Country:p.country,
              Purchaser_Ranking_Eligibility: p.eligibility,
              Purchaser_Ranking_Value_Year: p.value,
              Valid_Year_Ranking_Value_Total: rd.validRankVal,
              Excluded_Placeholder_Value: rd.placeholderValue,
              Ranking_Value_Not_Calculable_Count: rd.notCalc,
              Purchaser_Value_Share: p.share, Cumulative_Value_Share: p.cumShare,
              Purchaser_Value_Rank: p.rank, Top80_Research_Flag: "Yes",
            }));
            if (sc.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sc), safeSheetName("Scope_"+yr));
          });
        }
        // Industry Matching Log
        if (industryLog) {
          const il = yearSheets.map(yr => ({Year:yr, ...industryLog[yr]}));
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(il), "Industry_Match_Log");
        }
        // IQR Summary
        if (iqrResults) {
          const iq = [];
          yearSheets.forEach(yr => { const ir=iqrResults[yr]; if(!ir) return;
            Object.entries(ir).forEach(([seg,v]) => iq.push({Year:yr,Segment:seg,...v})); });
          if (iq.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(iq), "IQR_Summary");
        }
        // Segment Metrics
        if (segmentMetrics) {
          const sm = [];
          yearSheets.forEach(yr => { const s=segmentMetrics[yr]; if(!s) return;
            Object.entries(s).forEach(([seg,v]) => sm.push({
              Year:yr, Segment:seg, B_TV:v.before.sumTV, B_QtyMT:v.before.sumQty, B_AvgP:v.before.avgPrice, B_N:v.before.count,
              A_TV:v.after.sumTV, A_QtyMT:v.after.sumQty, A_AvgP:v.after.avgPrice, A_N:v.after.count,
            })); });
          if (sm.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sm), "Segment_Metrics");
        }
        // Keyword Conflicts
        if (keywordConflicts.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(keywordConflicts), "Keyword_Conflicts");

        // Per-year Before/After
        yearSheets.forEach(yr => {
          const rows = sheetData[yr];
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), safeSheetName(yr+"_Before"));
          const after = rows.filter(r => !(r.Pre_Conversion_Value_Check==="Outlier" && r.IQR_Segment_Check==="IQR Outlier"));
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(after), safeSheetName(yr+"_After"));
        });

        triggerDownload(XLSX.write(wb, { bookType:"xlsx", type:"array" }), "ChemSeg_Output_"+new Date().toISOString().slice(0,10)+".xlsx");
        setExportDone(true);
      } catch (err) { alert("Export error: " + err.message); }
      setProcessing(false);
    }, 100);
  }, [yearSheets, sheetData, qualityReport, backfillLog, validationLog, entityLog, rankingData, industryLog, classificationLog, conversionLog, iqrResults, segmentMetrics, keywordConflicts]);


  /* ═══════════════════════════════════════════════
     RENDER — Each case matches STEPS[i] exactly
     ═══════════════════════════════════════════════ */
  const renderStep = () => {
    switch (step) {

      /* ── 0: Upload ── */
      case 0:
        return <Card title="Upload Multi-Year Workbook" accent="#3b82f6">
          <p style={{color:"#475569",marginBottom:20,lineHeight:1.7}}>
            Upload an Excel file (.xlsx) with separate worksheets per year. Sheet names must contain a year (e.g. "2022", "2023_Data").
          </p>
          <DropZone icon="📊" text="Drop your Excel workbook here" inputRef={fileRef} accept=".xlsx,.xls" onChange={handleUpload} />
        </Card>;

      /* ── 1: Schema ── */
      case 1:
        return <Card title="Schema Mapping" accent="#8b5cf6">
          <p style={{color:"#475569",marginBottom:20,lineHeight:1.6,fontSize:14}}>
            Map your column headers to required fields. {yearSheets.length} sheet(s) detected: {yearSheets.join(", ")}
          </p>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"20px 28px"}}>
            {SCHEMA_FIELDS.map(f => (
              <div key={f.key}>
                <label style={{display:"block",fontSize:11.5,fontWeight:700,color:"#475569",marginBottom:8,textTransform:"uppercase",letterSpacing:"0.04em"}}>
                  {f.label} {schema[f.key] ? " ✅" : " ⚠️"}
                </label>
                <select value={schema[f.key]||""} onChange={e => setSchema(p=>({...p,[f.key]:e.target.value}))}
                  style={{width:"100%",padding:"11px 14px",borderRadius:8,border:"1px solid #e2e8f0",background:"#fff",color:"#0f172a",fontSize:14,fontFamily:"inherit"}}>
                  <option value="">— Select column —</option>
                  {allCols.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{marginTop:24}}>
            <Btn disabled={!schemaComplete} onClick={() => setStep(2)}>Confirm Mapping →</Btn>
            {!schemaComplete && <span style={{marginLeft:14,fontSize:13,color:"#ef4444"}}>All fields must be mapped</span>}
          </div>
        </Card>;

      /* ── 2: Data Quality ── */
      case 2:
        return <>
          <Card title="Data Quality Check" accent="#f59e0b">
            {!qualityReport ? <>
              <p style={{color:"#475569",marginBottom:16}}>Analyze text and numeric columns across all years.</p>
              <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
                {yearSheets.map(yr => <StatBox key={yr} label={yr} value={fmt(sheetData[yr]?.length)} sub="rows" color="#3b82f6"/>)}
              </div>
              <Btn onClick={runQuality} disabled={processing}>{processing ? "Analyzing..." : "Run Quality Check"}</Btn>
            </> : <>
              {yearSheets.map(yr => { const q=qualityReport[yr]; if(!q) return null; return (
                <div key={yr} style={{marginBottom:20}}>
                  <h4 style={{fontSize:14,fontWeight:700,margin:"0 0 10px",color:"#f59e0b"}}>{yr} — {q.totalRows} rows</h4>
                  <div style={{fontSize:12,fontWeight:700,color:"#475569",marginBottom:6,textTransform:"uppercase"}}>Numeric Fields</div>
                  <MiniTable headers={["Field","Valid","Missing","Zero","Negative","Invalid","Text"]}
                    rows={Object.entries(q.numeric).map(([f,v])=>[f,v.valid,v.missing,v.zero,v.negative,v.invalid,v.textPh])} />
                  <div style={{fontSize:12,fontWeight:700,color:"#475569",marginBottom:6,marginTop:14,textTransform:"uppercase"}}>Text Fields</div>
                  <MiniTable headers={["Field","Filled","Blank","None/N/A","Other"]}
                    rows={Object.entries(q.text).map(([f,v])=>[f,v.filled,v.blank,v.none,v.other])} />
                </div>
              );})}
              <Btn onClick={() => setStep(3)}>Next: Numeric Backfill →</Btn>
            </>}
          </Card>
        </>;

      /* ── 3: Backfill ── */
      case 3:
        return <Card title="Numeric Backfill" accent="#6366f1">
          {!backfillLog ? <>
            <p style={{color:"#475569",marginBottom:16}}>
              Backfill exactly one missing field (Unit Price, Total Value, or Quantity) when the other two are valid.
            </p>
            <Btn onClick={runBackfill} disabled={processing}>{processing ? "Processing..." : "Run Backfill"}</Btn>
          </> : <>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
              {yearSheets.map(yr => <StatBox key={yr} label={yr+" Backfilled"} value={fmt(backfillLog[yr]?.backfilled)} sub={"of "+backfillLog[yr]?.total} color="#6366f1"/>)}
            </div>
            <Btn onClick={() => setStep(4)}>Next: Value Validation →</Btn>
          </>}
        </Card>;

      /* ── 4: Validate ── */
      case 4:
        return <Card title="Pre-Conversion Value Validation" accent="#ef4444">
          {!validationLog ? <>
            <p style={{color:"#475569",marginBottom:16}}>
              Flag rows where |UnitPrice × Quantity − TotalValue| &gt; $500 as outlier.
            </p>
            <Btn onClick={runValidation} disabled={processing}>{processing ? "Processing..." : "Run Validation"}</Btn>
          </> : <>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
              {yearSheets.map(yr => <StatBox key={yr} label={yr+" Flagged"} value={fmt(validationLog[yr]?.flagged)} sub={"of "+validationLog[yr]?.total} color="#ef4444"/>)}
            </div>
            <Btn onClick={() => setStep(5)}>Next: Entity Resolution →</Btn>
          </>}
        </Card>;

      /* ── 5: Entity (V17 Enhanced) ── */
      case 5:
        return <>
          <Card title="Company Entity Resolution" accent="#0ea5e9">
            {!entityLog ? <>
              <p style={{color:"#475569",marginBottom:16,lineHeight:1.7}}>
                <b>V17 Enhanced:</b> Country-aware entity resolution with Vietnamese-English alias detection, malformed suffix repair,
                normalized matching keys (separate from display names), risk flagging, and review status.
              </p>
              <Btn onClick={runEntity} disabled={processing}>{processing ? "Processing..." : "Run Entity Resolution"}</Btn>
            </> : <>
              {yearSheets.map(yr => { const el=entityLog[yr]; if(!el) return null;
                return <div key={yr} style={{marginBottom:24}}>
                  <h4 style={{fontSize:14,fontWeight:700,margin:"0 0 12px",color:"#0ea5e9"}}>{yr} — {el.total} rows</h4>
                  <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
                    <StatBox label="Supplier Groups" value={el.suppGroups} color="#0ea5e9"/>
                    <StatBox label="Purchaser Groups" value={el.purchGroups} color="#8b5cf6"/>
                    <StatBox label="Needs Review" value={el.suppReview + el.purchReview} sub="supplier + purchaser" color="#f59e0b"/>
                    <StatBox label="Placeholders" value={el.suppPlh + el.purchPlh} sub="excluded" color="#94a3b8"/>
                  </div>
                  {el.riskSamples && el.riskSamples.length > 0 && <>
                    <div style={{fontSize:12,fontWeight:700,color:"#475569",marginBottom:6,marginTop:14,textTransform:"uppercase"}}>
                      Sample Standardizations & Risk Flags
                    </div>
                    <MiniTable maxH={240}
                      headers={["Raw Name","Standardized","Norm Key","Method","Risk","Review"]}
                      rows={el.riskSamples.map(s => [s.raw, s.std, s.normKey, s.method, s.risk, s.review])} />
                  </>}
                </div>;
              })}
              <div style={{display:"flex",gap:12}}>
                <Btn primary={false} onClick={() => { setEntityLog(null); }}>↻ Re-run</Btn>
                <Btn onClick={() => setStep(6)}>Next: Top 80% Ranking →</Btn>
              </div>
            </>}
          </Card>
        </>;

      /* ── 6: Ranking (V16: Placeholder Exclusion) ── */
      case 6:
        return <>
          <Card title="Top 80% Purchaser Ranking" accent="#f97316">
            {!rankingData ? <>
              <p style={{color:"#475569",marginBottom:16,lineHeight:1.7}}>
                Rank purchasers by raw-basis transaction value per year. <b>V16:</b> Placeholder purchasers
                (NONE, OTHER, blank, N/A, UNKNOWN, etc.) are excluded from both numerator and denominator.
              </p>

              {/* Placeholder exclusion list editor */}
              <div style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:10,padding:"16px 20px",marginBottom:16}}>
                <div style={{fontSize:12,fontWeight:700,color:"#9a3412",marginBottom:8,textTransform:"uppercase",letterSpacing:"0.04em"}}>
                  Excluded Placeholder Names
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>
                  {DEFAULT_PLACEHOLDERS.map(p => (
                    <span key={p} style={{padding:"3px 10px",borderRadius:6,background:"#ffedd5",color:"#9a3412",fontSize:11,fontWeight:600}}>{p||"(blank)"}</span>
                  ))}
                  {customPlaceholders.map((p,i) => (
                    <span key={"c"+i} style={{padding:"3px 10px",borderRadius:6,background:"#fde68a",color:"#92400e",fontSize:11,fontWeight:600,cursor:"pointer"}}
                      onClick={() => setCustomPlaceholders(prev => prev.filter((_,j) => j!==i))} title="Click to remove">
                      {p} ✕
                    </span>
                  ))}
                </div>
                <div style={{display:"flex",gap:8}}>
                  <input value={placeholderInput} onChange={e => setPlaceholderInput(e.target.value)}
                    placeholder="Add custom placeholder name..."
                    onKeyDown={e => { if (e.key==="Enter" && placeholderInput.trim()) { setCustomPlaceholders(prev => [...prev, placeholderInput.trim().toUpperCase()]); setPlaceholderInput(""); } }}
                    style={{flex:1,padding:"8px 12px",borderRadius:8,border:"1px solid #e2e8f0",fontSize:13,fontFamily:"inherit"}} />
                  <Btn onClick={() => { if (placeholderInput.trim()) { setCustomPlaceholders(prev => [...prev, placeholderInput.trim().toUpperCase()]); setPlaceholderInput(""); } }}>Add</Btn>
                </div>
              </div>

              <Btn onClick={runRanking} disabled={processing}>{processing ? "Processing..." : "Calculate Ranking"}</Btn>
            </> : <>
              {yearSheets.map(yr => { const rd=rankingData[yr]; if(!rd) return null;
                const top = rd.purchList.filter(p=>p.scope==="Top_80pct_Value");
                return <div key={yr} style={{marginBottom:24}}>
                  <h4 style={{fontSize:14,fontWeight:700,margin:"0 0 12px",color:"#f97316"}}>{yr}</h4>
                  <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
                    <StatBox label="Top 80% Selected" value={rd.selectedCount} color="#f97316"/>
                    <StatBox label="Real Purchasers" value={rd.totalPurchasers} color="#64748b"/>
                    <StatBox label="Valid Denom." value={"$"+(rd.validRankVal/1e6).toFixed(2)+"M"} color="#10b981"/>
                  </div>
                  {/* V16: Placeholder exclusion stats */}
                  <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
                    <StatBox label="Excl. Placeholder Rows" value={rd.placeholderRows} color="#dc2626"/>
                    <StatBox label="Excl. NONE Value" value={"$"+fmt(rd.placeholderNoneValue)} color="#b91c1c"/>
                    <StatBox label="Excl. OTHER Value" value={"$"+fmt(rd.placeholderOtherValue)} color="#b91c1c"/>
                    <StatBox label="Total Excl. Value" value={"$"+fmt(rd.placeholderValue)} color="#991b1b"/>
                  </div>
                  {rd.rawAnnualTotal > 0 && (
                    <InfoBar type="info">
                      Raw annual total: ${fmt(rd.rawAnnualTotal)} → Valid ranking denominator: ${fmt(rd.validRankVal)}
                      (excluded {pct(rd.placeholderValue / rd.rawAnnualTotal)} placeholder value)
                    </InfoBar>
                  )}
                  <MiniTable maxH={280}
                    headers={["Rank","Purchaser","Country","Value","Share","Cum%"]}
                    rows={top.slice(0,30).map(p=>[p.rank,p.std,p.country,"$"+fmt(p.value),pct(p.share),pct(p.cumShare)])} />
                  {top.length>30 && <div style={{fontSize:12,color:"#94a3b8",marginTop:6}}>...and {top.length-30} more</div>}
                </div>;
              })}
              <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
                <Btn primary={false} onClick={exportResearchList}>⬇ Export Research List</Btn>
                <Btn primary={false} onClick={() => { setRankingData(null); }}>↻ Re-run with changes</Btn>
                <Btn onClick={() => setStep(7)}>Next: Industry Master →</Btn>
              </div>
            </>}
          </Card>
        </>;


      /* ── 7: Industry Master ── */
      case 7:
        return <Card title="Upload Industry Master" accent="#8b5cf6">
          <p style={{color:"#475569",marginBottom:16,lineHeight:1.7}}>
            Required columns: <b>Purchaser_Standardize</b>, <b>Purchasing Country</b>, <b>Industry</b>, <b>Industry Segment</b>
          </p>
          <DropZone icon="🏭" text="Drop Industry Master file here (.xlsx)" inputRef={indRef} accept=".xlsx,.xls" onChange={handleIndustryUpload} />
          {industryLog && <>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginTop:16}}>
              {yearSheets.map(yr => <StatBox key={yr} label={yr+" Matched"} value={fmt(industryLog[yr]?.matched)} sub={industryLog[yr]?.unmatched+" unmatched"} color="#8b5cf6"/>)}
            </div>
            <div style={{marginTop:16}}>
              <Btn onClick={() => setStep(8)}>Next: Keyword Matrix →</Btn>
            </div>
          </>}
          {!industryLog && <div style={{marginTop:16}}>
            <Btn primary={false} onClick={() => setStep(8)}>Skip (no industry data) →</Btn>
          </div>}
        </Card>;

      /* ── 8: Keywords ── */
      case 8:
        return <Card title="Upload Keyword Matrix" accent="#ec4899">
          <p style={{color:"#475569",marginBottom:16,lineHeight:1.7}}>
            Excel format: <b>STT | Segment 1 | Segment 2 | ...</b> with keywords listed under each segment column.
          </p>
          <DropZone icon="🔑" text="Drop Keyword Matrix file here (.xlsx)" inputRef={kwRef} accept=".xlsx,.xls" onChange={handleKeywordUpload} />
          {keywordMatrix && <>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginTop:16}}>
              {Object.entries(keywordMatrix).map(([seg,kws]) => <StatBox key={seg} label={seg} value={kws.length} sub="keywords" color="#ec4899"/>)}
            </div>
            {keywordConflicts.length>0 && <InfoBar type="warn">⚠ {keywordConflicts.length} keyword conflict(s) found across segments</InfoBar>}
            <div style={{marginTop:16}}>
              <Btn onClick={() => setStep(9)}>Next: Run Classification →</Btn>
            </div>
          </>}
        </Card>;

      /* ── 9: Classification ── */
      case 9:
        return <Card title="Row Classification" accent="#14b8a6">
          {!classificationLog ? <>
            <p style={{color:"#475569",marginBottom:16}}>
              Classify rows by Product Description keywords. For Top 80% purchasers, validate/fallback using Industry Segment.
            </p>
            <Btn onClick={runClassification} disabled={processing}>{processing ? "Classifying..." : "Run Classification"}</Btn>
          </> : <>
            {yearSheets.map(yr => { const cl=classificationLog[yr]; if(!cl) return null;
              return <div key={yr} style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:14}}>
                <StatBox label={yr+" Keyword"} value={fmt(cl.byKeyword)} color="#14b8a6"/>
                <StatBox label={yr+" Industry"} value={fmt(cl.byIndustry)} color="#8b5cf6"/>
                <StatBox label={yr+" Unclassified"} value={fmt(cl.unclassified)} color="#ef4444"/>
              </div>;
            })}
            <Btn onClick={() => setStep(10)}>Next: Unit Conversion →</Btn>
          </>}
        </Card>;

      /* ── 10: Conversion ── */
      case 10:
        return <Card title="Unit Conversion" accent="#0d9488">
          {!conversionLog ? <>
            <p style={{color:"#475569",marginBottom:16}}>Convert quantities to Metric Tons and recalculate UnitPrice_per_MT.</p>
            <Btn onClick={runConversion} disabled={processing}>{processing ? "Converting..." : "Run Conversion"}</Btn>
          </> : <>
            {yearSheets.map(yr => { const c=conversionLog[yr]; if(!c) return null;
              return <div key={yr} style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:14}}>
                <StatBox label={yr+" Converted"} value={fmt(c.converted)} color="#0d9488"/>
                <StatBox label={yr+" Ambiguous"} value={fmt(c.ambiguous)} color="#f59e0b"/>
                <StatBox label={yr+" Unsupported"} value={fmt(c.unsupported)} color="#ef4444"/>
              </div>;
            })}
            <Btn onClick={() => setStep(11)}>Next: IQR Analysis →</Btn>
          </>}
        </Card>;

      /* ── 11: IQR ── */
      case 11:
        return <Card title="IQR Outlier Detection" accent="#7c3aed">
          {!iqrResults ? <>
            <p style={{color:"#475569",marginBottom:16}}>Run 1.5×IQR on UnitPrice_per_MT by Year + Final Segment.</p>
            <Btn onClick={runIQR} disabled={processing}>{processing ? "Analyzing..." : "Run IQR"}</Btn>
          </> : <>
            {yearSheets.map(yr => { const ir=iqrResults[yr]; if(!ir) return null;
              return <div key={yr} style={{marginBottom:20}}>
                <h4 style={{fontSize:14,fontWeight:700,margin:"0 0 10px",color:"#7c3aed"}}>{yr}</h4>
                <MiniTable headers={["Segment","N","Q1","Q3","IQR","Lower","Upper","Outliers","Note"]}
                  rows={Object.entries(ir).map(([seg,v])=>[seg,v.count,fmt(v.q1),fmt(v.q3),fmt(v.iqr),fmt(v.lower),fmt(v.upper),v.outliers,v.note])} />
              </div>;
            })}
            <Btn onClick={() => setStep(12)}>Next: Segment Metrics →</Btn>
          </>}
        </Card>;

      /* ── 12: Metrics ── */
      case 12:
        return <Card title="Segment Metrics" accent="#059669">
          {!segmentMetrics ? <>
            <p style={{color:"#475569",marginBottom:16}}>Calculate per-segment summary: Average Price = Sum(TotalValue) / Sum(Quantity_MT), before and after outlier removal.</p>
            <Btn onClick={runMetrics} disabled={processing}>{processing ? "Calculating..." : "Calculate Metrics"}</Btn>
          </> : <>
            {yearSheets.map(yr => { const sm=segmentMetrics[yr]; if(!sm) return null;
              return <div key={yr} style={{marginBottom:20}}>
                <h4 style={{fontSize:14,fontWeight:700,margin:"0 0 10px",color:"#059669"}}>{yr}</h4>
                <MiniTable headers={["Segment","B.Count","B.Value","B.Qty(MT)","B.AvgPrice","A.Count","A.Value","A.Qty(MT)","A.AvgPrice"]}
                  rows={Object.entries(sm).map(([seg,v])=>[seg,v.before.count,"$"+fmt(v.before.sumTV),fmt(v.before.sumQty),"$"+fmt(v.before.avgPrice),
                    v.after.count,"$"+fmt(v.after.sumTV),fmt(v.after.sumQty),"$"+fmt(v.after.avgPrice)])} />
              </div>;
            })}
            <Btn onClick={() => setStep(13)}>Next: Export →</Btn>
          </>}
        </Card>;

      /* ── 13: Export ── */
      case 13:
        return <Card title="Export Final Audit Workbook" accent="#10b981">
          <p style={{color:"#475569",marginBottom:16,lineHeight:1.7}}>
            Output includes: 00_Summary, Data Quality, Backfill Log, Entity Log, Research Scope per year, Industry Match Log,
            IQR Summary, Segment Metrics, Keyword Conflicts, and per-year Before/After Outlier sheets with all raw + audit columns.
          </p>
          {exportDone && <InfoBar type="success">✅ Export complete! Check your downloads folder.</InfoBar>}
          <Btn big onClick={handleExport} disabled={processing}>
            {processing ? "Generating..." : "💾 Export Full Audit Workbook"}
          </Btn>
        </Card>;

      default: return null;
    }
  };

  /* ═══════════════════════════════════════════════
     LAYOUT
     ═══════════════════════════════════════════════ */
  return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif", display:"flex", minHeight:"100vh", background:"#f1f5f9", color:"#0f172a" }}>

      {/* ── Sidebar ── */}
      <aside style={{ width:220, minWidth:220, background:"#0f172a", padding:"20px 0", display:"flex", flexDirection:"column", overflowY:"auto", flexShrink:0 }}>
        <div style={{ padding:"0 18px 18px", borderBottom:"1px solid rgba(255,255,255,0.08)", marginBottom:8 }}>
          <div style={{ fontSize:16, fontWeight:800, color:"#fff", letterSpacing:"-0.02em" }}>⚗️ ChemSeg V17</div>
          <div style={{ fontSize:11, color:"#94a3b8", marginTop:4 }}>Chemical Segmentation Tool</div>
        </div>
        {STEPS.map((s, i) => {
          const active = i === step;
          const done = i < step;
          return (
            <button key={s.id} onClick={() => { if (i<=step) setStep(i); }}
              style={{
                display:"flex", alignItems:"center", gap:10, width:"100%",
                padding:"9px 18px", border:"none", textAlign:"left", fontFamily:"inherit",
                background: active ? "#1e40af" : "transparent",
                color: active ? "#fff" : done ? "#93c5fd" : "#64748b",
                cursor: i<=step ? "pointer" : "default", fontSize:12.5, fontWeight: active ? 700 : 500,
                opacity: i>step ? 0.35 : 1, transition:"all 0.15s",
                borderLeft: active ? "3px solid #60a5fa" : "3px solid transparent",
              }}>
              <span style={{fontSize:13,width:20,textAlign:"center",flexShrink:0}}>{done && !active ? "✓" : s.icon}</span>
              <span style={{lineHeight:1.3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{s.label}</span>
            </button>
          );
        })}
        <div style={{flex:1}}/>
        <div style={{padding:"14px 18px",borderTop:"1px solid rgba(255,255,255,0.08)",fontSize:11,color:"#475569"}}>
          {yearSheets.length>0 && <div>📅 {yearSheets.join(", ")}</div>}
          <div style={{marginTop:4}}>Step {step+1} / {STEPS.length}</div>
        </div>
      </aside>

      {/* ── Main Content ── */}
      <main style={{ flex:1, padding:"32px 48px", overflowY:"auto" }}>
        <div style={{ maxWidth:1280, margin:"0 auto" }}>
          <div style={{ marginBottom:28 }}>
            <h1 style={{ fontSize:24, fontWeight:800, margin:0, letterSpacing:"-0.03em", display:"flex", alignItems:"center", gap:10 }}>
              <span>{STEPS[step]?.icon}</span> {STEPS[step]?.label}
            </h1>
            <div style={{ fontSize:13, color:"#94a3b8", marginTop:5 }}>
              Step {step+1} of {STEPS.length}
              {yearSheets.length>0 && (" — "+yearSheets.join(", "))}
            </div>
          </div>

          {processing && (
            <InfoBar type="info">
              <span style={{display:"inline-block",animation:"spin 1s linear infinite",marginRight:8}}>⏳</span>
              Processing...
              <style>{`@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}`}</style>
            </InfoBar>
          )}

          {renderStep()}
        </div>
      </main>
    </div>
  );
}
