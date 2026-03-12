import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Palette ──────────────────────────────────────────────────────────────────
const C = {
  dark: "#1a3c2e",
  mid: "#2e7d32",
  bright: "#3dcd58",
  light: "#e8f5e9",
  teal: "#00897b",
  white: "#fff",
  off: "#f4f6f4",
  grey: "#78909c",
  text: "#1a2e28",
  rLow: "#2e7d32",
  rMed: "#e65100",
  rHigh: "#b71c1c",
  rDel: "#1565c0",
  upBand: "#3dcd58",
  dnBand: "#00897b",
  scBand: "#1a3c2e",
};

// ── Domain constants ─────────────────────────────────────────────────────────
const HIGH_RISK = "under investigation - high risk";
const RISK_LONG = {
  low: "Comitted or very low risk",
  medium: "under investigation but medium risk",
  high: "under investigation - high risk",
  delivered: "Delivered",
  na: "n/a",
};
const LEVER_META = {
  design_with_less: { driverType: "Upstream design", driverName: "Design with Less", subDriver: "Less materials", updown: "Upstream", applyNew: "Yes" },
  mat_eff_legacy: { driverType: "Upstream Improvements", driverName: "Material efficiency legacy", subDriver: "Efficient material", updown: "Upstream", applyNew: "No" },
  mat_eff_new_offer: { driverType: "Upstream Improvements", driverName: "Material efficiency new offer", subDriver: "Efficient material", updown: "Upstream", applyNew: "Yes" },
  lifetime_extension: { driverType: "Upstream Improvements", driverName: "Lifetime extention new offer", subDriver: "Extension", updown: "Upstream", applyNew: "Yes" },
  circular_offers: { driverType: "Upstream Improvements", driverName: "Circular offers", subDriver: "Circularity", updown: "Upstream", applyNew: "Yes" },
  other: { driverType: "Upstream Improvements", driverName: "Additional improvements", subDriver: "Additional improvements", updown: "Upstream", applyNew: "Yes" },
  energy_efficiency_1: { driverType: "Dowstream improvement", driverName: "Energy efficiency - 1", subDriver: "Energy efficiency - 1", updown: "Downstream", applyNew: "Yes" },
  energy_efficiency_2: { driverType: "Dowstream improvement", driverName: "Energy efficiency - 2", subDriver: "Energy efficiency - 2", updown: "Downstream", applyNew: "Yes" },
  energy_efficiency_3: { driverType: "Dowstream improvement", driverName: "Energy efficiency - 3", subDriver: "Energy efficiency - 3", updown: "Downstream", applyNew: "Yes" },
};
const LEVERS_TAB_COLS = [
  "Driver Type",
  "Driver Name",
  "Sub-driver name",
  "Lever Type Tag",
  "Attribute impacted (EF or Volume)",
  "LoB",
  "product line code",
  "product family",
  "hub",
  "country (dest)",
  "Upstream/Downstream",
  "Life cycle stage",
  "year",
  "Growth/degrowth",
  "Apply to new references?",
  "Migration rate",
  "Annual growth/degrowth",
  "Annual growth/degrowth to be applied",
  "Type of growth/degrowth",
  "Starting year",
  "Completion year",
  "Risk/Confidence status",
  "Risk/Confidence status ",
  "Hypothesis",
];

// ── Trajectory Summary columns ───────────────────────────────────────────────
const TRAJ_COLS = [
  "Ref product family",
  "Ref product line code",
  "Ref commercial reference",
  "Ref countries dest",
  "Ref hub",
  "year",
  "scalar distribution emissions (kgco2e)",
  "scalar end-of-life emissions (kgco2e)",
  "scalar installation emissions (kgco2e)",
  "scalar manufacturing emissions (kgco2e)",
  "scalar use phase emissions (kgco2e)",
  "Overall emissions (kgco2e)",
  "scalar quantity",
];
const TRAJ_NUM_COLS = new Set([
  "scalar distribution emissions (kgco2e)",
  "scalar end-of-life emissions (kgco2e)",
  "scalar installation emissions (kgco2e)",
  "scalar manufacturing emissions (kgco2e)",
  "scalar use phase emissions (kgco2e)",
  "Overall emissions (kgco2e)",
  "scalar quantity",
  "year",
]);

// ── Emission column groups (used by lever application) ────────────────────────
const UPSTREAM_COLS = [
  "scalar distribution emissions (kgco2e)",
  "scalar manufacturing emissions (kgco2e)",
];
const DOWNSTREAM_COLS = [
  "scalar end-of-life emissions (kgco2e)",
  "scalar use phase emissions (kgco2e)",
  "scalar installation emissions (kgco2e)",
];
const EMISSION_COLS = [...UPSTREAM_COLS, ...DOWNSTREAM_COLS];

// ── Read Excel sheet handling duplicate column names (like pandas .1 suffix) ──
function xlsxToJsonSafe(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!raw.length) return [];
  const counts = {};
  const headers = raw[0].map((h) => {
    const key = String(h || "").trim();
    if (!key) return "__empty__";
    if (counts[key] === undefined) { counts[key] = 0; return key; }
    counts[key]++;
    return `${key}.${counts[key]}`;
  });
  return raw.slice(1)
    .filter((cells) => cells.some((c) => String(c ?? "").trim() !== ""))
    .map((cells) => {
      const row = {};
      headers.forEach((h, i) => { row[h] = cells[i] ?? ""; });
      return row;
    });
}

// ── Parse numbers that may use European comma decimal separator ───────────────
function parseNum(v) {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  if (v == null || v === "") return 0;
  const s = String(v).trim().replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// ── Pre-aggregate 340K rows → unique (family × line × country × hub) groups ──
// Levers match on product_family + product_line_code.
// Display groups on product_line_code + countries_dest.
// Including all four in the key preserves full granularity for both.
const AGG_KEY_COLS = [
  "Ref product family",
  "Ref product line code",
  "Ref countries dest",
  "Ref hub",
];
function aggregateBaseline(rows) {
  const map = new Map();
  for (const row of rows) {
    const key = AGG_KEY_COLS.map((k) => String(row[k] ?? "")).join("|||");
    if (!map.has(key)) {
      const agg = { _rowCount: 0 };
      for (const k of AGG_KEY_COLS) agg[k] = row[k] ?? "";
      for (const c of EMISSION_COLS) agg[c] = 0;
      agg["scalar quantity"] = 0;
      map.set(key, agg);
    }
    const agg = map.get(key);
    for (const c of EMISSION_COLS) agg[c] += parseNum(row[c]);
    agg["scalar quantity"] += parseNum(row["scalar quantity"]);
    agg._rowCount++;
  }
  return [...map.values()];
}

// ── Apply levers to 2025 baseline (JS port of apply_levers_v3) ───────────────
function applyLeversV3(baseline, levers) {
  const FUTURE_YEARS = [2026, 2027, 2028, 2029, 2030];
  const n = baseline.length;
  if (!n || !levers.length) return null;

  // Drop levers whose Type of growth/degrowth contains "growth"
  const TYPE_COL = "Type of growth/degrowth";
  // Separate EF vs Volume levers (mirrors Python logic — no growth-type pre-filter;
  // "Percentual direct (de)growth" and "Compound growth" both contain "growth" so
  // the old filter incorrectly removed every lever)
  const ATTR_COL = "Attribute impacted (EF or Volume)";
  const leversEF  = levers.filter((l) => String(l[ATTR_COL] || "").trim() !== "Volume");
  const leversVol = levers.filter((l) => String(l[ATTR_COL] || "").trim() === "Volume");

  // Risk/Confidence status.1 = second occurrence (pandas naming convention)
  const RISK_COL = "Risk/Confidence status.1";
  const baselineLevers = leversEF.filter(
    (l) =>
      (!l["product line code"] || String(l["product line code"]).trim() === "") &&
      (!l["product family"]    || String(l["product family"]).trim()    === "")
  );
  const lowLevers    = leversEF.filter((l) => String(l[RISK_COL] || "").trim().toLowerCase() === "low");
  const mediumLevers = leversEF.filter((l) => String(l[RISK_COL] || "").trim().toLowerCase() === "medium");

  function dedupConcat(...arrays) {
    const seen = new Set();
    const out = [];
    for (const arr of arrays)
      for (const item of arr)
        if (!seen.has(item)) { seen.add(item); out.push(item); }
    return out;
  }

  const SCENARIOS = {
    baseline:         baselineLevers,
    baseline_low:     dedupConcat(baselineLevers, lowLevers),
    baseline_low_med: dedupConcat(baselineLevers, lowLevers, mediumLevers),
    all:              leversEF,
  };

  function buildMultipliers(scenarioLevers, useDriverType = true) {
    const mult = {};
    for (const yr of FUTURE_YEARS) {
      mult[yr] = {};
      for (const col of EMISSION_COLS) mult[yr][col] = new Float64Array(n).fill(1);
    }
    for (const lever of scenarioLevers) {
      const growthValue = parseFloat(lever["Annual growth/degrowth to be applied"]);
      if (isNaN(growthValue)) continue;

      const productLine   = String(lever["product line code"] || "").trim();
      const productFamily = String(lever["product family"]    || "").trim();
      const driverType    = String(lever["Driver Type"] || "").toLowerCase();
      const growthType    = String(lever[TYPE_COL] || "").toLowerCase();
      const startYear     = parseInt(lever["Starting year"])   || 2026;
      const endYear       = parseInt(lever["Completion year"]) || 2030;

      let targetCols;
      if (useDriverType) {
        if (driverType.includes("upstream")) targetCols = UPSTREAM_COLS;
        else if (driverType.includes("downstream") || driverType.includes("dowstream")) targetCols = DOWNSTREAM_COLS;
        else targetCols = EMISSION_COLS;
      } else {
        targetCols = EMISSION_COLS;
      }

      const matchedIdx = [];
      for (let i = 0; i < n; i++) {
        const row = baseline[i];
        if (productLine   && String(row["Ref product line code"] || "").trim() !== productLine)   continue;
        if (productFamily && String(row["Ref product family"]    || "").trim() !== productFamily) continue;
        matchedIdx.push(i);
      }
      if (!matchedIdx.length) continue;

      const duration = endYear !== startYear ? endYear - startYear : 1;
      for (const yr of FUTURE_YEARS) {
        if (yr < startYear || yr > endYear) continue;
        const t = yr - startYear + 1;
        const leverMult = growthType.includes("compound")
          ? Math.pow(1 + growthValue, t)
          : 1 + (growthValue / duration) * t;
        for (const col of targetCols) {
          const arr = mult[yr][col];
          for (const idx of matchedIdx) arr[idx] *= leverMult;
        }
      }
    }
    return mult;
  }

  // Pre-build all scenario multipliers
  const scenarioMults = {};
  for (const [name, leversSet] of Object.entries(SCENARIOS))
    scenarioMults[name] = buildMultipliers(leversSet, true);

  // 2025 base totals
  let base2025Up = 0, base2025Down = 0;
  for (const row of baseline) {
    for (const col of UPSTREAM_COLS)   base2025Up   += parseNum(row[col]);
    for (const col of DOWNSTREAM_COLS) base2025Down += parseNum(row[col]);
  }

  // EF scenarios – grand totals
  const scenarioResults = {};
  for (const [name, mult] of Object.entries(scenarioMults)) {
    scenarioResults[name] = {};
    for (const yr of FUTURE_YEARS) {
      let up = 0, dn = 0;
      for (let i = 0; i < n; i++) {
        const row = baseline[i];
        for (const col of UPSTREAM_COLS)   up += parseNum(row[col]) * mult[yr][col][i];
        for (const col of DOWNSTREAM_COLS) dn += parseNum(row[col]) * mult[yr][col][i];
      }
      scenarioResults[name][yr] = { upstream: up, downstream: dn, total: up + dn };
    }
  }

  // Volume growth – net incremental vs 2025
  const volMult = buildMultipliers(leversVol, false);
  const volumeGrowth = {};
  for (const yr of FUTURE_YEARS) {
    let delta = 0;
    for (let i = 0; i < n; i++) {
      const row = baseline[i];
      for (const col of EMISSION_COLS)
        delta += parseNum(row[col]) * (volMult[yr][col][i] - 1);
    }
    volumeGrowth[yr] = delta;
  }

  // Grouped breakdown by (Ref product line code, Ref countries dest) — like df_grouped2
  const groupMap = {};
  for (let i = 0; i < n; i++) {
    const row = baseline[i];
    const lc = String(row["Ref product line code"] || "").trim();
    const cd = String(row["Ref countries dest"]    || "").trim();
    const key = `${lc}|||${cd}`;
    if (!groupMap[key]) groupMap[key] = { lineCode: lc, countriesDest: cd, indices: [] };
    groupMap[key].indices.push(i);
  }
  const byGroup = Object.values(groupMap).map(({ lineCode, countriesDest, indices }) => {
    let overall2025 = 0;
    for (const i of indices)
      for (const col of EMISSION_COLS)
        overall2025 += parseNum(baseline[i][col]);
    const years = {};
    for (const yr of FUTURE_YEARS) {
      years[yr] = {};
      for (const [sName, mult] of Object.entries(scenarioMults)) {
        let total = 0;
        for (const i of indices)
          for (const col of EMISSION_COLS)
            total += parseNum(baseline[i][col]) * mult[yr][col][i];
        years[yr][sName] = total;
      }
    }
    return { lineCode, countriesDest, overall2025, years };
  });

  return {
    base2025: { upstream: base2025Up, downstream: base2025Down, total: base2025Up + base2025Down },
    scenarios: scenarioResults,
    volumeGrowth,
    byGroup,
    debugFirstRow: baseline[0] ?? {},
    rowCount: n,
    leverCount: levers.length,
  };
}

// ── Scenario calculation ─────────────────────────────────────────────────────
function calcScenarios(li) {
  const f = li.migration_rate || 0;
  const u = li.upstream,
    d = li.downstream;
  const sc1_up =
    (u.design_with_less.status === HIGH_RISK ? 0 : f * u.design_with_less.value) +
    (u.mat_eff_legacy.status === HIGH_RISK ? 0 : u.mat_eff_legacy.value * (1 - f)) +
    (u.mat_eff_new_offer.status === HIGH_RISK ? 0 : u.mat_eff_new_offer.value * f) +
    (u.lifetime_extension.status === HIGH_RISK ? 0 : u.lifetime_extension.value * f) +
    (u.circular_offers.status === HIGH_RISK ? 0 : u.circular_offers.value * f) +
    (u.other.status === HIGH_RISK ? 0 : u.other.value * f);
  const sc2_up =
    f * u.design_with_less.value +
    u.mat_eff_legacy.value * (1 - f) +
    u.mat_eff_new_offer.value * f +
    u.lifetime_extension.value * f +
    u.circular_offers.value * f +
    u.other.value * f;
  const sc1_dn =
    (d.energy_efficiency_1.status === HIGH_RISK ? 0 : f * d.energy_efficiency_1.value) +
    (d.energy_efficiency_2.status === HIGH_RISK ? 0 : f * d.energy_efficiency_2.value) +
    (d.energy_efficiency_3.status === HIGH_RISK ? 0 : f * d.energy_efficiency_3.value);
  const sc2_dn = f * d.energy_efficiency_1.value + f * d.energy_efficiency_2.value + f * d.energy_efficiency_3.value;
  return { sc1_up, sc2_up, sc1_dn, sc2_dn };
}

// ── Build outputs from spec ──────────────────────────────────────────────────
function buildLeversInputRow(spec) {
  const li = spec.levers_input;
  const { sc1_up, sc2_up, sc1_dn, sc2_dn } = calcScenarios(li);
  const u = li.upstream,
    d = li.downstream;
  const pct = (v) => `${((v || 0) * 100).toFixed(0)}%`;
  return {
    lob: spec.lob || "",
    product_line: spec.product_line_code || "",
    product_family: spec.product_family || "",
    comment: spec.comment || "",
    family_lifecycle: spec.family_lifecycle || "",
    migration_rate: pct(li.migration_rate),
    growth_low: pct(li.growth_low),
    growth_high: pct(li.growth_high),
    dwl_value: pct(u.design_with_less.value),
    dwl_hypothesis: u.design_with_less.hypothesis,
    dwl_status: u.design_with_less.status,
    mel_value: pct(u.mat_eff_legacy.value),
    mel_hypothesis: u.mat_eff_legacy.hypothesis,
    mel_status: u.mat_eff_legacy.status,
    mno_value: pct(u.mat_eff_new_offer.value),
    mno_hypothesis: u.mat_eff_new_offer.hypothesis,
    mno_status: u.mat_eff_new_offer.status,
    lte_value: pct(u.lifetime_extension.value),
    lte_hypothesis: u.lifetime_extension.hypothesis,
    lte_status: u.lifetime_extension.status,
    co_value: pct(u.circular_offers.value),
    co_hypothesis: u.circular_offers.hypothesis,
    co_status: u.circular_offers.status,
    oth_value: pct(u.other.value),
    oth_hypothesis: u.other.hypothesis,
    oth_status: u.other.status,
    scenario1_up: pct(sc1_up),
    scenario2_up: pct(sc2_up),
    ee1_value: pct(d.energy_efficiency_1.value),
    ee1_hypothesis: d.energy_efficiency_1.hypothesis,
    ee1_status: d.energy_efficiency_1.status,
    ee2_value: pct(d.energy_efficiency_2.value),
    ee2_hypothesis: d.energy_efficiency_2.hypothesis,
    ee2_status: d.energy_efficiency_2.status,
    ee3_value: pct(d.energy_efficiency_3.value),
    ee3_hypothesis: d.energy_efficiency_3.hypothesis,
    ee3_status: d.energy_efficiency_3.status,
    scenario1_dn: pct(sc1_dn),
    scenario2_dn: pct(sc2_dn),
  };
}
function buildLeversTabRows(spec) {
  const rows = [];
  const pline = spec.product_line_code || "",
    family = spec.product_family || "",
    lob = spec.lob || "";
  const defaultMig = spec.levers_input?.migration_rate || 0;
  for (const lev of spec.levers_tab || []) {
    const meta = LEVER_META[lev.lever_type];
    if (!meta) continue;
    for (const entry of lev.entries || []) {
      const hub = entry.hub || "Hub1";
      const mig = typeof entry.migration_rate === "number" ? entry.migration_rate : defaultMig;
      const q = entry.annual_rate;
      const r = meta.applyNew === "No" ? +(q * (1 - mig)).toFixed(4) : +(q * mig).toFixed(4);
      rows.push({
        "Driver Type": meta.driverType,
        "Driver Name": meta.driverName,
        "Sub-driver name": meta.subDriver,
        "Lever Type Tag": `${hub} - ${meta.subDriver}-${pline}`,
        "Attribute impacted (EF or Volume)": "EF",
        "LoB": lob,
        "product line code": pline,
        "product family": family,
        hub: hub,
        "country (dest)": entry.country || "",
        "Upstream/Downstream": meta.updown,
        "Life cycle stage": "",
        year: spec.year || 2025,
        "Growth/degrowth": q <= 0 ? "Degrowth" : "Growth",
        "Apply to new references?": meta.applyNew,
        "Migration rate": mig,
        "Annual growth/degrowth": q,
        "Annual growth/degrowth to be applied": r,
        "Type of growth/degrowth": "Percentual direct (de)growth",
        "Starting year": spec.starting_year || 2026,
        "Completion year": spec.completion_year || 2030,
        "Risk/Confidence status": RISK_LONG[entry.risk] || RISK_LONG.low,
        "Risk/Confidence status ": entry.risk === "delivered" ? "low" : entry.risk || "low",
        Hypothesis: entry.hypothesis || "",
      });
    }
  }
  return rows;
}

// ── CSV helpers ──────────────────────────────────────────────────────────────
const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
const LI_COLS = [
  "lob", "product_line", "product_family", "comment", "family_lifecycle",
  "migration_rate", "growth_low", "growth_high",
  "dwl_value", "dwl_hypothesis", "dwl_status",
  "mel_value", "mel_hypothesis", "mel_status",
  "mno_value", "mno_hypothesis", "mno_status",
  "lte_value", "lte_hypothesis", "lte_status",
  "co_value", "co_hypothesis", "co_status",
  "oth_value", "oth_hypothesis", "oth_status",
  "scenario1_up", "scenario2_up",
  "ee1_value", "ee1_hypothesis", "ee1_status",
  "ee2_value", "ee2_hypothesis", "ee2_status",
  "ee3_value", "ee3_hypothesis", "ee3_status",
  "scenario1_dn", "scenario2_dn",
];
const LI_HEADERS = [
  "LoB", "Product Line", "Product Familly", "comment", "Familly lifecycle",
  "migration rate (old to new offer) by 2030",
  "Annual growth rates Low", "Annual growth rates high",
  "improvement value", "hypothesis", "status",
  "improvement value", "hypothesis", "status",
  "improvement value", "hypothesis", "status",
  "improvement value", "hypothesis", "status",
  "improvement value", "hypothesis", "status",
  "improvement value", "hypothesis", "status",
  "SCENARIO 1", "SCENARIO 2",
  "Improvement Value", "Hypothesis", "status",
  "Improvement Value", "Hypothesis", "status",
  "Improvement Value", "Hypothesis", "status",
  "Improvement %", "Improvement %",
];
function dlCSV(filename, csv) {
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  Object.assign(document.createElement("a"), { href: url, download: filename }).click();
  URL.revokeObjectURL(url);
}
function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') { cell += '"'; i++; }
      else if (ch === '"') { inQuotes = false; }
      else { cell += ch; }
    } else {
      if (ch === '"') { inQuotes = true; }
      else if (ch === ",") { row.push(cell); cell = ""; }
      else if (ch === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; }
      else if (ch === "\r") { /* ignore */ }
      else { cell += ch; }
    }
  }
  row.push(cell);
  rows.push(row);
  while (rows.length && rows[rows.length - 1].every((c) => String(c ?? "").trim() === "")) rows.pop();
  return rows;
}
function normalizeHeader(h) {
  return String(h ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}
function isLikelyLeversInputHeaders(headers) {
  const H = headers.map(normalizeHeader);
  const must = ["lob", "product line", "product familly", "comment", "familly lifecycle"];
  const okMust = must.every((m, idx) => H[idx] === m);
  return okMust && headers.length >= LI_HEADERS.length;
}

// ── AI ───────────────────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are a Schneider Electric sustainability analyst specialising in industrial automation carbon footprint modelling.
Extract lever assumptions from the user's description. Return ONLY valid JSON — no markdown, no preamble.
Schema:
{
  "lob": string, "product_line_code": string, "product_family": string,
  "comment": string, "family_lifecycle": string,
  "year": number, "starting_year": number, "completion_year": number,
  "levers_input": {
    "migration_rate": number, "growth_low": number, "growth_high": number,
    "upstream": {
      "design_with_less":   {"value":number,"hypothesis":string,"status":string},
      "mat_eff_legacy":     {"value":number,"hypothesis":string,"status":string},
      "mat_eff_new_offer":  {"value":number,"hypothesis":string,"status":string},
      "lifetime_extension": {"value":number,"hypothesis":string,"status":string},
      "circular_offers":    {"value":number,"hypothesis":string,"status":string},
      "other":              {"value":number,"hypothesis":string,"status":string}
    },
    "downstream": {
      "energy_efficiency_1": {"value":number,"hypothesis":string,"status":string},
      "energy_efficiency_2": {"value":number,"hypothesis":string,"status":string},
      "energy_efficiency_3": {"value":number,"hypothesis":string,"status":string}
    }
  },
  "levers_tab": [
    {"lever_type":string,"entries":[{"hub":string,"country":string|null,"migration_rate":number,"annual_rate":number,"risk":string,"hypothesis":string}]}
  ]
}
Rules:
- levers_input values: POSITIVE fractions (10% improvement = 0.10)
- levers_tab annual_rate: NEGATIVE for reductions (-0.10 for 10% improvement)
- status exactly one of: "Comitted or very low risk" | "under investigation but medium risk" | "under investigation - high risk" | "Delivered"
- risk: "committed"/"delivered"→"delivered", "low"→"low", "medium"→"medium", "high"/"uncertain"→"high"
- Omitted levers: value=0, hypothesis="", status="Comitted or very low risk"; omit from levers_tab
- lever_type values: "design_with_less"|"mat_eff_legacy"|"mat_eff_new_offer"|"lifetime_extension"|"circular_offers"|"other"|"energy_efficiency_1"|"energy_efficiency_2"|"energy_efficiency_3"
- Default hub="Hub1"; if Hub1≠Hub2 create two entries; migration_rate per entry defaults to levers_input.migration_rate
- year=2025, starting_year=2026, completion_year=2030`;

async function callClaude(msg, history) {
  const messages = [...history, { role: "user", content: msg }];
  const res = await fetch("/hf-api/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "MiniMaxAI/MiniMax-M2.5",
      max_tokens: 2500,
      messages: [{ role: "system", content: SYSTEM_PROMPT }, ...messages],
    }),
  });
  const data = await res.json();
  const text = data.choices?.[0]?.message?.content || "";
  return { text, messages };
}
function tryJSON(t) {
  return JSON.parse(t.replace(/```json|```/g, "").trim());
}

// ── Shared UI atoms ──────────────────────────────────────────────────────────
function RiskChip({ status }) {
  const s = status || "";
  const cfg = s.includes("high")
    ? { bg: "#fce4e4", color: C.rHigh, dot: "#f44336" }
    : s.includes("medium")
    ? { bg: "#fff3e0", color: C.rMed, dot: "#ff9800" }
    : s === "Delivered"
    ? { bg: "#e3f2fd", color: C.rDel, dot: "#2196f3" }
    : { bg: "#e8f5e9", color: C.rLow, dot: "#4caf50" };
  const label = s.includes("high") ? "High risk" : s.includes("medium") ? "Medium" : s === "Delivered" ? "Delivered" : "Low risk";
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: cfg.bg, color: cfg.color, borderRadius: 20, padding: "2px 8px", fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: cfg.dot, flexShrink: 0 }} />
      {label}
    </span>
  );
}
function pctFmt(v) {
  if (typeof v === "number") return `${(v * 100).toFixed(0)}%`;
  if (typeof v === "string" && v.endsWith("%")) return v;
  return v || "0%";
}

// ── Trajectory Summary component ─────────────────────────────────────────────
// ── Trajectory Watch helpers ──────────────────────────────────────────────────
const WATCH_SCENARIOS = {
  baseline:         { label: "Baseline",    color: "#78909c" },
  baseline_low:     { label: "+ Low risk",  color: "#2e7d32" },
  baseline_low_med: { label: "+ Med risk",  color: "#00897b" },
  all:              { label: "All levers",  color: "#1a3c2e" },
};
const WATCH_YEARS = [2026, 2027, 2028, 2029, 2030];
const ALL_YEARS   = [2025, ...WATCH_YEARS];

// ── SVG Line Chart for scenario trajectories ──────────────────────────────────
const TARGET_KT = 5222 * 1e6; // 5 222 kt CO₂e target (in kg, matching data units)

function TrajectoryLineChart({ watchResult, mult = 1 }) {
  const W = 760, H = 260, PAD = { top: 20, right: 160, bottom: 20, left: 60 };
  const innerW = W - PAD.left - PAD.right;
  const innerH = H - PAD.top  - PAD.bottom;

  const lines = Object.entries(WATCH_SCENARIOS).map(([key, { label, color }]) => {
    const pts = ALL_YEARS.map((yr) => {
      const raw = yr === 2025
        ? watchResult.base2025.total
        : watchResult.scenarios[key]?.[yr]?.total ?? null;
      const total = raw != null ? raw * mult : null;
      return { yr, total };
    });
    return { key, label, color, pts };
  });

  const allVals = lines.flatMap((l) => l.pts.map((p) => p.total)).filter((v) => v != null);
  if (!allVals.length || allVals.every((v) => v === 0)) return (
    <div style={{ padding: 20, textAlign: "center", color: C.grey, fontSize: 12 }}>
      No data yet — check debug panel below.
    </div>
  );

  // Include target in domain so it's always visible
  const domainVals = [...allVals, TARGET_KT];
  const minV = Math.min(...domainVals) * 0.97;
  const maxV = Math.max(...domainVals) * 1.03;
  const xScale = (yr) => PAD.left + ((yr - 2025) / 5) * innerW;
  const yScale = (v)  => PAD.top  + innerH - ((v - minV) / (maxV - minV)) * innerH;

  const fmtKt = (v) => v == null ? "" : (v / 1e6).toLocaleString(undefined, { maximumFractionDigits: 1 }) + " kt";
  const yTicks = 5;
  const gridRight = W - PAD.right;
  const legendX = gridRight + 12;

  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", display: "block" }}>
      {/* Y grid + labels */}
      {Array.from({ length: yTicks + 1 }, (_, i) => {
        const v = minV + (i / yTicks) * (maxV - minV);
        const y = yScale(v);
        return (
          <g key={i}>
            <line x1={PAD.left} x2={gridRight} y1={y} y2={y} stroke="#e0ece0" strokeWidth={1} />
            <text x={PAD.left - 6} y={y + 4} fontSize={9} fill={C.grey} textAnchor="end">{fmtKt(v)}</text>
          </g>
        );
      })}
      {/* X axis labels */}
      {ALL_YEARS.map((yr) => (
        <text key={yr} x={xScale(yr)} y={H - 4} fontSize={10} fill={C.grey} textAnchor="middle">{yr}</text>
      ))}
      {/* Target line at 5 222 kt */}
      {(() => {
        const yT = yScale(TARGET_KT);
        return (
          <g>
            <line x1={PAD.left} x2={gridRight} y1={yT} y2={yT}
              stroke="#e53935" strokeWidth={1.5} strokeDasharray="6,4" />
            <text x={PAD.left + 4} y={yT - 3} fontSize={8} fill="#e53935">
              Target 5 222 kt
            </text>
          </g>
        );
      })()}
      {/* Lines */}
      {lines.map(({ key, color, pts }) => {
        const valid = pts.filter((p) => p.total != null);
        if (valid.length < 2) return null;
        const d = valid.map((p, i) => `${i === 0 ? "M" : "L"}${xScale(p.yr)},${yScale(p.total)}`).join(" ");
        return (
          <g key={key}>
            <path d={d} fill="none" stroke={color} strokeWidth={2.2} strokeLinejoin="round" />
            {valid.map((p) => (
              <circle key={p.yr} cx={xScale(p.yr)} cy={yScale(p.total)} r={3.5} fill={color} />
            ))}
          </g>
        );
      })}
      {/* Legend — right of grid */}
      {lines.map(({ key, label, color }, i) => (
        <g key={key} transform={`translate(${legendX}, ${PAD.top + i * 22})`}>
          <line x1={0} x2={16} y1={5} y2={5} stroke={color} strokeWidth={2.5} />
          <circle cx={8} cy={5} r={3} fill={color} />
          <text x={22} y={9} fontSize={10} fill={C.text}>{label}</text>
        </g>
      ))}
      {/* Target legend entry */}
      <g transform={`translate(${legendX}, ${PAD.top + lines.length * 22 + 6})`}>
        <line x1={0} x2={16} y1={5} y2={5} stroke="#e53935" strokeWidth={1.5} strokeDasharray="5,3" />
        <text x={22} y={9} fontSize={10} fill="#e53935">Target (5 222 kt)</text>
      </g>
    </svg>
  );
}

function WatchPill({ icon, label, ok, loading, error, count, unit }) {
  const bg = loading ? "#e8f0e8" : error ? "#ffebee" : ok ? "#e8f5e9" : "#f4f6f4";
  const color = loading ? C.grey : error ? C.rHigh : ok ? C.mid : C.grey;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, background: bg, border: `1px solid ${color}55`, borderRadius: 20, padding: "4px 10px", fontSize: 11 }}>
      <span>{icon}</span>
      <span style={{ fontWeight: 700, color }}>{label}</span>
      {loading && <span style={{ color: C.grey }}>loading…</span>}
      {!loading && error && <span style={{ color: C.rHigh }}>error</span>}
      {!loading && !error && ok && <span style={{ color: C.grey }}>{count?.toLocaleString()} {unit}</span>}
      {!loading && !error && !ok && <span style={{ color: C.grey }}>not loaded</span>}
    </div>
  );
}

function WatchCard({ label, value, sub, color }) {
  return (
    <div style={{ background: C.white, border: `1px solid ${color}55`, borderRadius: 8, padding: "10px 14px", minWidth: 140 }}>
      <div style={{ fontSize: 10, color: C.grey, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 18, fontWeight: 800, color }}>{value}</div>
      <div style={{ fontSize: 10, color: C.grey, marginTop: 2 }}>{sub}</div>
    </div>
  );
}

// ── Trajectory Watch tab ──────────────────────────────────────────────────────
function TrajectoryWatch({ df2025, df2025Loading, df2025Error, levers, leversLoading, leversError, onUploadLevers, watchResult, watchComputing }) {
  const [growthRate, setGrowthRate] = React.useState(4);
  const mult = 1 + growthRate / 100;

  const fmtKt = (v) =>
    v != null ? ((v * mult) / 1e6).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : "—";
  const fmtPct = (v, base) => {
    if (base == null || v == null) return "—";
    const p = ((v - base) / base) * 100;
    return (p > 0 ? "+" : "") + p.toFixed(1) + "%";
  };

  const thS = (align = "left") => ({
    padding: "6px 8px", color: C.white, fontWeight: 700, whiteSpace: "nowrap",
    textAlign: align, borderRight: "1px solid #2d5a3d", fontSize: 10,
  });
  const tdS = (align = "left", bold, color) => ({
    padding: "4px 8px", borderBottom: "1px solid #e0ece0", borderRight: "1px solid #f0f4f0",
    whiteSpace: "nowrap", textAlign: align, fontWeight: bold ? 700 : 400, color: color || C.text,
  });

  return (
    <div style={{ padding: "14px 16px", overflowY: "auto", height: "100%" }}>
      {/* Status + upload */}
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <WatchPill icon="📊" label="2025 baseline" ok={!!df2025?.length} loading={df2025Loading} error={df2025Error}
          count={df2025?.length} unit="aggregated groups" />
        <WatchPill icon="⚙️" label="Levers" ok={!!levers?.length} loading={leversLoading} error={leversError} count={levers?.length} unit="rows" />
        {watchComputing && <span style={{ fontSize: 11, color: C.grey, fontStyle: "italic" }}>Computing scenarios…</span>}
        <label style={{ marginLeft: "auto", background: C.mid, color: C.white, border: "none", borderRadius: 6, padding: "6px 12px", fontWeight: 700, fontSize: 11, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6 }}>
          ⬆ {levers?.length ? "Replace" : "Upload"} levers.xlsx
          <input type="file" accept=".xlsx,.xls" onChange={onUploadLevers} style={{ display: "none" }} />
        </label>
      </div>

      {/* ── Growth rate control ── */}
      <div style={{
        display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap",
        background: C.white, border: "1px solid #cdd8d0", borderRadius: 8,
        padding: "10px 14px", marginBottom: 12,
      }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.dark, whiteSpace: "nowrap" }}>
          Volume Growth Rate
        </div>
        <input
          type="range" min={0} max={20} step={0.5} value={growthRate}
          onChange={(e) => setGrowthRate(parseFloat(e.target.value))}
          style={{ flex: "1 1 180px", accentColor: C.mid, cursor: "pointer" }}
        />
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          <input
            type="number" min={0} max={20} step={0.5} value={growthRate}
            onChange={(e) => setGrowthRate(Math.min(20, Math.max(0, parseFloat(e.target.value) || 0)))}
            style={{
              width: 54, padding: "3px 6px", fontSize: 13, fontWeight: 700,
              border: "1px solid #cdd8d0", borderRadius: 5, textAlign: "right", color: C.mid,
            }}
          />
          <span style={{ fontSize: 13, fontWeight: 700, color: C.mid }}>%</span>
        </div>
        <div style={{ fontSize: 11, color: C.grey, whiteSpace: "nowrap" }}>
          multiplier&nbsp;<strong style={{ color: C.dark }}>×{mult.toFixed(4)}</strong>
        </div>
        <div style={{ fontSize: 10, color: C.grey, marginLeft: "auto" }}>
          All emissions figures include this growth factor
        </div>
      </div>

      {leversError && (
        <div style={{ background: "#ffebee", border: "1px solid #ef9a9a", borderRadius: 7, padding: "8px 12px", fontSize: 12, color: C.rHigh, marginBottom: 12 }}>
          {leversError}
        </div>
      )}

      {!watchResult && !watchComputing && (
        <div style={{ color: C.grey, fontSize: 13, padding: "40px 0", textAlign: "center" }}>
          {!df2025?.length ? "Waiting for df_2025.xlsx…" : "Upload a levers.xlsx file to compute scenarios."}
        </div>
      )}

      {watchResult && (
        <>
          {/* ── Charts (top) ─────────────────────────────────────────────── */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 340px", gap: 12, marginBottom: 16 }}>
            {/* Line chart */}
            <div style={{ background: C.white, border: "1px solid #cdd8d0", borderRadius: 8, padding: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.text, marginBottom: 10 }}>
                Emissions Trajectory 2025–2030 (kt CO₂e)
              </div>
              <TrajectoryLineChart watchResult={watchResult} mult={mult} />
            </div>

            {/* Bar chart */}
            <div style={{ background: C.white, border: "1px solid #cdd8d0", borderRadius: 8, padding: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.text, marginBottom: 10 }}>
                2030 Scenario Comparison
              </div>
              {(() => {
                const maxVal = watchResult.base2025.total * mult * 1.05;
                const targetW = Math.min(100, (TARGET_KT / maxVal) * 100);
                return Object.entries(WATCH_SCENARIOS).map(([key, { label, color }]) => {
                  const rawTotal = watchResult.scenarios[key]?.[2030]?.total ?? 0;
                  const barW  = Math.min(100, Math.max(0, (rawTotal * mult / maxVal) * 100));
                  return (
                    <div key={key} style={{ marginBottom: 10 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ width: 90, fontSize: 11, fontWeight: 700, color, flexShrink: 0 }}>{label}</div>
                        <div style={{ flex: 1, background: "#e8f0e8", borderRadius: 4, height: 18, position: "relative" }}>
                          <div style={{ width: `${barW}%`, height: "100%", background: color, borderRadius: 4, opacity: 0.85 }} />
                          <div style={{ position: "absolute", left: `${targetW}%`, top: 0, bottom: 0, width: 2, background: "#e53935" }} title="5 222 kt target" />
                        </div>
                        <div style={{ width: 70, textAlign: "right", fontSize: 11, fontWeight: 600 }}>
                          {fmtKt(rawTotal)} kt
                        </div>
                      </div>
                    </div>
                  );
                });
              })()}
              <div style={{ fontSize: 10, color: "#e53935", marginTop: 6 }}>
                Red line = 5 222 kt CO₂e target
              </div>
            </div>
          </div>

          {/* ── Summary cards ────────────────────────────────────────────── */}
          <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
            <WatchCard
              label="2025 baseline (kt CO₂e)"
              value={fmtKt(watchResult.base2025.total)}
              sub={`${fmtKt(watchResult.base2025.upstream)} up · ${fmtKt(watchResult.base2025.downstream)} dn`}
              color={C.dark}
            />
            {Object.entries(WATCH_SCENARIOS).map(([key, { label, color }]) => {
              const y2030 = watchResult.scenarios[key]?.[2030];
              if (!y2030) return null;
              return (
                <WatchCard
                  key={key}
                  label={`2030 ${label} (kt CO₂e)`}
                  value={fmtKt(y2030.total)}
                  sub={`${fmtPct(y2030.total, watchResult.base2025.total)} vs 2025`}
                  color={color}
                />
              );
            })}
          </div>

          {/* ── Scenario table ───────────────────────────────────────────── */}
          <div style={{ overflowX: "auto", borderRadius: 7, border: "1px solid #8fbe8f", boxShadow: "0 1px 4px #0001", marginBottom: 16 }}>
            <table style={{ borderCollapse: "collapse", fontSize: 11, width: "100%" }}>
              <thead>
                <tr style={{ background: C.dark }}>
                  <th style={thS("left")}>Scenario</th>
                  <th style={thS("right")}>2025 (kt)</th>
                  {WATCH_YEARS.map((yr) => (
                    <th key={yr} colSpan={2} style={thS("center")}>{yr}</th>
                  ))}
                </tr>
                <tr style={{ background: "#2d5a3d" }}>
                  <th style={thS("left")} />
                  <th style={thS("right")} />
                  {WATCH_YEARS.map((yr) => (
                    <React.Fragment key={yr}>
                      <th style={thS("right")}>kt CO₂e</th>
                      <th style={{ ...thS("right"), color: "#a5d6a7" }}>vs 2025</th>
                    </React.Fragment>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.entries(WATCH_SCENARIOS).map(([key, { label, color }], si) => (
                  <tr key={key} style={{ background: si % 2 === 0 ? C.white : C.off }}>
                    <td style={{ ...tdS("left", true, color) }}>{label}</td>
                    <td style={tdS("right", true)}>{fmtKt(watchResult.base2025.total)}</td>
                    {WATCH_YEARS.map((yr) => {
                      const row = watchResult.scenarios[key]?.[yr];
                      const pctStr = fmtPct(row?.total, watchResult.base2025.total);
                      const isPos = row && row.total > watchResult.base2025.total;
                      return (
                        <React.Fragment key={yr}>
                          <td style={tdS("right", false)}>{row ? fmtKt(row.total) : "—"}</td>
                          <td style={{ ...tdS("right", true, isPos ? C.rHigh : C.mid) }}>{pctStr}</td>
                        </React.Fragment>
                      );
                    })}
                  </tr>
                ))}
                <tr style={{ background: "#e8f0e8", borderTop: "2px solid #8fbe8f" }}>
                  <td style={{ ...tdS("left", false, C.grey), fontStyle: "italic" }}>Volume growth (net)</td>
                  <td style={tdS("right")}>—</td>
                  {WATCH_YEARS.map((yr) => (
                    <React.Fragment key={yr}>
                      <td style={{ ...tdS("right", false, C.grey), fontStyle: "italic" }}>{fmtKt(watchResult.volumeGrowth[yr])}</td>
                      <td style={tdS("right")}>—</td>
                    </React.Fragment>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>

          {/* ── Breakdown table (df_grouped2) — first 10 rows + download ── */}
          {watchResult.byGroup?.length > 0 && (() => {
            const allRows = watchResult.byGroup;
            const preview = allRows.slice(0, 10);

            function downloadCsv() {
              const scenarioKeys = Object.keys(WATCH_SCENARIOS);
              const header = [
                "Product Line Code", "Countries Dest", "Overall 2025 (kgco2e)",
                ...WATCH_YEARS.flatMap((yr) =>
                  scenarioKeys.map((s) => `Overall ${yr} ${s} (kgco2e)`)
                ),
              ];
              const csvRows = allRows.map(({ lineCode, countriesDest, overall2025, years }) => [
                lineCode, countriesDest, overall2025.toFixed(2),
                ...WATCH_YEARS.flatMap((yr) =>
                  scenarioKeys.map((s) => (years[yr]?.[s] ?? 0).toFixed(2))
                ),
              ]);
              const csv = [header, ...csvRows].map((r) => r.join(",")).join("\n");
              const a = document.createElement("a");
              a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
              a.download = "trajectory_by_line_country.csv";
              a.click();
            }

            return (
              <div style={{ marginBottom: 16 }}>
                {/* Header row with download button */}
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between",
                  background: C.dark, padding: "6px 10px", borderRadius: "7px 7px 0 0" }}>
                  <span style={{ fontSize: 11, fontWeight: 700, color: C.white }}>
                    2030 Trajectory — by Product Line &amp; Country (kt CO₂e)
                    <span style={{ fontWeight: 400, opacity: 0.7, marginLeft: 8 }}>
                      showing 10 of {allRows.length} rows
                    </span>
                  </span>
                  <button onClick={downloadCsv} style={{
                    background: C.bright, color: C.dark, border: "none", borderRadius: 5,
                    padding: "4px 10px", fontWeight: 700, fontSize: 11, cursor: "pointer",
                  }}>
                    ⬇ Download full table (.csv)
                  </button>
                </div>
                <div style={{ overflowX: "auto", border: "1px solid #8fbe8f", borderTop: "none", borderRadius: "0 0 7px 7px", boxShadow: "0 1px 4px #0001" }}>
                  <table style={{ borderCollapse: "collapse", fontSize: 10, width: "100%" }}>
                    <thead>
                      <tr style={{ background: "#2d5a3d" }}>
                        <th style={{ ...thS("left"), position: "sticky", left: 0, background: "#2d5a3d", zIndex: 1 }}>Line</th>
                        <th style={thS("left")}>Country</th>
                        <th style={thS("right")}>2025</th>
                        {WATCH_YEARS.map((yr) => (
                          <th key={yr} colSpan={Object.keys(WATCH_SCENARIOS).length}
                            style={{ ...thS("center"), borderLeft: "2px solid #4a8a5a" }}>{yr}</th>
                        ))}
                      </tr>
                      <tr style={{ background: "#3a6a4a" }}>
                        <th style={{ ...thS("left"), position: "sticky", left: 0, background: "#3a6a4a", zIndex: 1 }} />
                        <th style={thS("left")} />
                        <th style={thS("right")} />
                        {WATCH_YEARS.map((yr) =>
                          Object.entries(WATCH_SCENARIOS).map(([key, { label, color }]) => (
                            <th key={`${yr}-${key}`} style={{ ...thS("right"), color, borderLeft: key === "baseline" ? "2px solid #4a8a5a" : undefined }}>
                              {label}
                            </th>
                          ))
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {preview.map(({ lineCode, countriesDest, overall2025, years }, ri) => (
                        <tr key={ri} style={{ background: ri % 2 === 0 ? C.white : C.off }}>
                          <td style={{ ...tdS("left", true), position: "sticky", left: 0, background: ri % 2 === 0 ? C.white : C.off, zIndex: 1 }}>{lineCode || "—"}</td>
                          <td style={tdS("left")}>{countriesDest || "—"}</td>
                          <td style={tdS("right", true)}>{fmtKt(overall2025)}</td>
                          {WATCH_YEARS.map((yr) =>
                            Object.keys(WATCH_SCENARIOS).map((sKey, si) => (
                              <td key={`${yr}-${sKey}`} style={{ ...tdS("right", false), borderLeft: si === 0 ? "2px solid #e0ece0" : undefined }}>
                                {fmtKt(years[yr]?.[sKey])}
                              </td>
                            ))
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}

          {/* ── Debug panel ──────────────────────────────────────────────── */}
          <details style={{ background: "#fffde7", border: "1px solid #f9a825", borderRadius: 7, padding: "8px 12px", fontSize: 11 }}>
            <summary style={{ fontWeight: 700, color: "#e65100", cursor: "pointer", marginBottom: 6 }}>
              🐛 Debug — raw df_2025 first row &amp; byGroup
            </summary>
            <div style={{ marginTop: 8 }}>
              <div style={{ fontWeight: 700, marginBottom: 4 }}>df_2025 first row (all keys + values):</div>
              <pre style={{ background: "#fff8e1", padding: 8, borderRadius: 4, overflowX: "auto", maxHeight: 200, fontSize: 10 }}>
                {JSON.stringify(watchResult.debugFirstRow || {}, null, 2)}
              </pre>
              <div style={{ fontWeight: 700, margin: "8px 0 4px" }}>byGroup (df_grouped2) — first 3 rows:</div>
              <pre style={{ background: "#fff8e1", padding: 8, borderRadius: 4, overflowX: "auto", maxHeight: 300, fontSize: 10 }}>
                {JSON.stringify(watchResult.byGroup?.slice(0, 3) || [], null, 2)}
              </pre>
            </div>
          </details>
        </>
      )}
    </div>
  );
}

function TrajectorySummary({ data, loading, error }) {
  const [growthRate, setGrowthRate] = React.useState(4);

  if (loading) return <div style={{ padding: 40, textAlign: "center", color: C.grey, fontSize: 13 }}>Loading 2025 data…</div>;
  if (error) return <div style={{ padding: 40, color: C.rHigh, fontSize: 13 }}>Error loading df_2025.xlsx: {error}</div>;
  if (!data?.length) return <div style={{ padding: 40, color: C.grey, fontSize: 13 }}>No data found in df_2025.xlsx</div>;

  const mult = 1 + growthRate / 100;
  const totalBase = data.reduce((s, r) => s + (parseNum(r["Overall emissions (kgco2e)"]) || 0), 0);
  const totalWithGrowth = totalBase * mult;

  const GROWTH_COL = "With growth (kgco2e)";
  const displayCols = [...TRAJ_COLS, GROWTH_COL];

  return (
    <div style={{ padding: "14px 16px", overflowY: "auto", height: "100%" }}>

      {/* ── Growth rate control ── */}
      <div style={{
        display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap",
        background: C.white, border: "1px solid #cdd8d0", borderRadius: 8,
        padding: "10px 14px", marginBottom: 12,
      }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.dark, whiteSpace: "nowrap" }}>
          Volume Growth Rate
        </div>
        <input
          type="range" min={0} max={20} step={0.5} value={growthRate}
          onChange={(e) => setGrowthRate(parseFloat(e.target.value))}
          style={{ flex: "1 1 180px", accentColor: C.mid, cursor: "pointer" }}
        />
        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
          <input
            type="number" min={0} max={20} step={0.5} value={growthRate}
            onChange={(e) => setGrowthRate(Math.min(20, Math.max(0, parseFloat(e.target.value) || 0)))}
            style={{
              width: 54, padding: "3px 6px", fontSize: 13, fontWeight: 700,
              border: "1px solid #cdd8d0", borderRadius: 5, textAlign: "right",
              color: C.mid,
            }}
          />
          <span style={{ fontSize: 13, fontWeight: 700, color: C.mid }}>%</span>
        </div>
        <div style={{ fontSize: 11, color: C.grey, whiteSpace: "nowrap" }}>
          multiplier&nbsp;<strong style={{ color: C.dark }}>×{mult.toFixed(4)}</strong>
        </div>

        {/* Summary chips */}
        <div style={{ display: "flex", gap: 8, marginLeft: "auto", flexWrap: "wrap" }}>
          <div style={{
            background: "#e8f0e8", borderRadius: 6, padding: "4px 10px", fontSize: 11,
          }}>
            Baseline total&nbsp;
            <strong style={{ color: C.dark }}>{fmtKt(totalBase)} kt</strong>
          </div>
          <div style={{
            background: "#fff3e0", borderRadius: 6, padding: "4px 10px", fontSize: 11,
          }}>
            With {growthRate}% growth&nbsp;
            <strong style={{ color: "#e65100" }}>{fmtKt(totalWithGrowth)} kt</strong>
          </div>
          <div style={{
            background: "#fce4e4", borderRadius: 6, padding: "4px 10px", fontSize: 11,
          }}>
            Delta&nbsp;
            <strong style={{ color: C.rHigh }}>+{fmtKt(totalWithGrowth - totalBase)} kt</strong>
          </div>
        </div>
      </div>

      <div style={{ marginBottom: 8, fontSize: 12, color: C.grey }}>{data.length} rows · 2025 baseline emissions data</div>
      <div style={{ overflowX: "auto", borderRadius: 7, border: "1px solid #8fbe8f", boxShadow: "0 1px 4px #0001" }}>
        <table style={{ borderCollapse: "collapse", fontSize: 11, width: "100%" }}>
          <thead>
            <tr style={{ background: C.dark }}>
              {displayCols.map((h) => (
                <th
                  key={h}
                  style={{
                    padding: "6px 8px",
                    color: h === GROWTH_COL ? "#ffe082" : C.white,
                    fontWeight: 700,
                    whiteSpace: "nowrap",
                    textAlign: TRAJ_NUM_COLS.has(h) || h === GROWTH_COL ? "right" : "left",
                    borderRight: "1px solid #2d5a3d",
                    fontSize: 10,
                    background: h === GROWTH_COL ? "#b45309" : undefined,
                  }}
                >
                  {h === GROWTH_COL ? `With ${growthRate}% growth (kgco2e)` : h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? C.white : C.off }}>
                {displayCols.map((h) => {
                  let v, display;
                  if (h === GROWTH_COL) {
                    const base = parseNum(row["Overall emissions (kgco2e)"]);
                    v = base * mult;
                    display = v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                  } else {
                    v = row[h] ?? "";
                    const isNum = TRAJ_NUM_COLS.has(h);
                    display = isNum && typeof v === "number"
                      ? h === "year" || h === "scalar quantity"
                        ? v.toLocaleString()
                        : v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
                      : String(v);
                  }
                  const isGrowthCol = h === GROWTH_COL;
                  const isOverall   = h === "Overall emissions (kgco2e)";
                  const isNum = TRAJ_NUM_COLS.has(h) || isGrowthCol;
                  return (
                    <td
                      key={h}
                      style={{
                        padding: "4px 8px",
                        borderBottom: "1px solid #e0ece0",
                        borderRight: "1px solid #f0f4f0",
                        whiteSpace: "nowrap",
                        textAlign: isNum ? "right" : "left",
                        fontWeight: isOverall || isGrowthCol ? 700 : 400,
                        color: isGrowthCol ? "#b45309" : isOverall ? C.mid : C.text,
                        background: isGrowthCol ? "#fff8ee" : undefined,
                      }}
                    >
                      {display}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── Levers Input table ───────────────────────────────────────────────────────
function LeversInputTable({ rows }) {
  if (!rows || !rows.length) return null;
  const TH = ({ children, colspan = 1, bg = C.dark, color = C.white, align = "center", minW }) => (
    <th
      colSpan={colspan}
      style={{
        background: bg, color, padding: "5px 7px", fontSize: 10, fontWeight: 700,
        textAlign: align, border: "1px solid #5a8a6a", whiteSpace: "nowrap",
        minWidth: minW, verticalAlign: "middle",
      }}
    >
      {children}
    </th>
  );
  const TD = ({ v, bold, align = "center", italic, color, bg = C.white }) => (
    <td
      style={{
        background: bg, color: color || C.text, padding: "5px 7px", fontSize: 11,
        fontWeight: bold ? 700 : 400, textAlign: align, border: "1px solid #d0e0d0",
        whiteSpace: "normal", maxWidth: 160, lineHeight: 1.3,
        fontStyle: italic ? "italic" : "normal", verticalAlign: "top",
      }}
    >
      {v ?? ""}
    </td>
  );
  const RiskTD = ({ s }) => {
    const cfg =
      (s || "").includes("high") ? { bg: "#fce4e4", color: C.rHigh } :
      (s || "").includes("medium") ? { bg: "#fff3e0", color: C.rMed } :
      s === "Delivered" ? { bg: "#e3f2fd", color: C.rDel } :
      { bg: "#f0fff4", color: C.rLow };
    return (
      <td
        style={{
          background: cfg.bg, color: cfg.color, padding: "5px 7px", fontSize: 10,
          fontWeight: 700, textAlign: "center", border: "1px solid #d0e0d0",
          whiteSpace: "normal", maxWidth: 120, lineHeight: 1.3, verticalAlign: "top",
        }}
      >
        {s || "—"}
      </td>
    );
  };
  return (
    <div style={{ overflowX: "auto", borderRadius: 7, border: "1px solid #8fbe8f", boxShadow: "0 1px 4px #0001" }}>
      <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
        <thead>
          <tr>
            <TH colspan={8} bg="#2a3d2a" color="#aaa"> </TH>
            <TH colspan={20} bg={C.upBand} color={C.dark}>↑ Upstream</TH>
            <TH colspan={11} bg={C.dnBand} color={C.white}>↓ Downstream</TH>
          </tr>
          <tr>
            <TH colspan={8} bg="#2a3d2a" color="#aaa"> </TH>
            {[
              ["Design with Less", 3], ["Material efficiency legacy", 3],
              ["Material efficiency new offer", 3], ["Lifetime extention new offer", 3],
              ["Circular Offers", 3], ["Other", 3],
            ].map(([l, n]) => (
              <TH key={l} colspan={n} bg={C.upBand} color={C.dark}>{l}</TH>
            ))}
            <TH colspan={2} bg={C.scBand} color={C.white}>Scenarios</TH>
            {[["Energy efficiency", 3], ["Energy efficiency", 3], ["Energy efficiency", 3]].map(([l, n], idx) => (
              <TH key={`${l}_${idx}`} colspan={n} bg={C.dnBand} color={C.white}>{l}</TH>
            ))}
            <TH colspan={2} bg={C.scBand} color={C.white}>Scenarios</TH>
          </tr>
          <tr>
            {[
              "LoB", "Product Line", "Product Familly", "comment", "Familly lifecycle",
              "migration rate (old to new offer) by 2030",
              "Annual growth rates Low", "Annual growth rates high",
            ].map((h) => (
              <TH key={h} bg="#2a3d2a" color={C.white} align="left" minW={h === "comment" || h === "Familly lifecycle" ? "120px" : "60px"}>
                {h}
              </TH>
            ))}
            {Array(6).fill(0).flatMap((_, i) => [
              <TH key={`v${i}`} bg="#1a5c28" color={C.white}>improvement value</TH>,
              <TH key={`h${i}`} bg="#1a5c28" color={C.white} minW="120px">hypothesis</TH>,
              <TH key={`s${i}`} bg="#1a5c28" color={C.white} minW="90px">status</TH>,
            ])}
            <TH bg={C.scBand} color={C.white}>SCENARIO 1</TH>
            <TH bg={C.scBand} color={C.white}>SCENARIO 2</TH>
            {Array(3).fill(0).flatMap((_, i) => [
              <TH key={`ev${i}`} bg="#004d40" color={C.white}>Improvement Value</TH>,
              <TH key={`eh${i}`} bg="#004d40" color={C.white} minW="120px">Hypothesis</TH>,
              <TH key={`es${i}`} bg="#004d40" color={C.white} minW="90px">status</TH>,
            ])}
            <TH bg={C.scBand} color={C.white}>Improvement %</TH>
            <TH bg={C.scBand} color={C.white}>Improvement %</TH>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? C.white : C.off }}>
              <TD v={r.lob} bold />
              <TD v={r.product_line} bold />
              <TD v={r.product_family} bold />
              <TD v={r.comment} align="left" color="#555" />
              <TD v={r.family_lifecycle} align="left" italic color="#555" />
              <TD v={pctFmt(r.migration_rate)} bold />
              <TD v={pctFmt(r.growth_low)} />
              <TD v={pctFmt(r.growth_high)} />
              <TD v={pctFmt(r.dwl_value)} bold color={C.mid} />
              <TD v={r.dwl_hypothesis} align="left" />
              <RiskTD s={r.dwl_status} />
              <TD v={pctFmt(r.mel_value)} bold color={C.mid} />
              <TD v={r.mel_hypothesis} align="left" />
              <RiskTD s={r.mel_status} />
              <TD v={pctFmt(r.mno_value)} bold color={C.mid} />
              <TD v={r.mno_hypothesis} align="left" />
              <RiskTD s={r.mno_status} />
              <TD v={pctFmt(r.lte_value)} bold color={C.mid} />
              <TD v={r.lte_hypothesis} align="left" />
              <RiskTD s={r.lte_status} />
              <TD v={pctFmt(r.co_value)} bold color={C.mid} />
              <TD v={r.co_hypothesis} align="left" />
              <RiskTD s={r.co_status} />
              <TD v={pctFmt(r.oth_value)} bold color={C.mid} />
              <TD v={r.oth_hypothesis} align="left" />
              <RiskTD s={r.oth_status} />
              <TD v={pctFmt(r.scenario1_up)} bold bg="#eaf4ea" color={C.dark} />
              <TD v={pctFmt(r.scenario2_up)} bold bg="#eaf4ea" color={C.dark} />
              <TD v={pctFmt(r.ee1_value)} bold color={C.teal} />
              <TD v={r.ee1_hypothesis} align="left" />
              <RiskTD s={r.ee1_status} />
              <TD v={pctFmt(r.ee2_value)} bold color={C.teal} />
              <TD v={r.ee2_hypothesis} align="left" />
              <RiskTD s={r.ee2_status} />
              <TD v={pctFmt(r.ee3_value)} bold color={C.teal} />
              <TD v={r.ee3_hypothesis} align="left" />
              <RiskTD s={r.ee3_status} />
              <TD v={pctFmt(r.scenario1_dn)} bold bg="#e0f2f1" color={C.teal} />
              <TD v={pctFmt(r.scenario2_dn)} bold bg="#e0f2f1" color={C.teal} />
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── Levers Tab table ─────────────────────────────────────────────────────────
function LeversTabTable({ rows }) {
  if (!rows || !rows.length) return null;
  const fmtN = (v, d = 1) => (typeof v === "number" ? `${(v * 100).toFixed(d)}%` : v || "");
  return (
    <div style={{ overflowX: "auto", borderRadius: 7, border: "1px solid #8fbe8f", boxShadow: "0 1px 4px #0001" }}>
      <table style={{ borderCollapse: "collapse", fontSize: 10, width: "100%" }}>
        <thead>
          <tr style={{ background: C.dark }}>
            {["Driver Type", "Driver Name", "Sub-driver", "Prod. Line", "Family", "Hub", "Up/Down", "Apply", "Migration", "Annual rate", "Applied", "Risk", "Hypothesis"].map((h) => (
              <th key={h} style={{ padding: "5px 7px", color: C.white, fontWeight: 700, whiteSpace: "nowrap", textAlign: "left", borderRight: "1px solid #2d5a3d" }}>
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const updown = r["Upstream/Downstream"];
            const accent = updown === "Upstream" ? "#e8f5e9" : "#e0f2f1";
            return (
              <tr key={i} style={{ background: i % 2 === 0 ? C.white : C.off }}>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", fontWeight: 600, color: updown === "Upstream" ? C.mid : C.teal }}>{r["Driver Type"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap" }}>{r["Driver Name"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", color: C.grey }}>{r["Sub-driver name"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap" }}>{r["product line code"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", fontWeight: 600 }}>{r["product family"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap" }}>{r["hub"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", background: accent, fontWeight: 600, color: updown === "Upstream" ? C.mid : C.teal }}>{updown}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap" }}>{r["Apply to new references?"]}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", textAlign: "right" }}>{fmtN(r["Migration rate"], 0)}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", textAlign: "right", color: Number(r["Annual growth/degrowth"]) < 0 ? C.mid : C.rHigh, fontWeight: 600 }}>{fmtN(r["Annual growth/degrowth"])}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", whiteSpace: "nowrap", textAlign: "right", fontWeight: 600 }}>{fmtN(r["Annual growth/degrowth to be applied"])}</td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  <RiskChip status={r["Risk/Confidence status"]} />
                </td>
                <td style={{ padding: "4px 6px", borderBottom: "1px solid #e0ece0", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", color: C.grey, fontStyle: "italic" }}>{r["Hypothesis"]}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("levers"); // "levers" | "watch"
  const [trajData, setTrajData] = useState(null);
  const [trajLoading, setTrajLoading] = useState(true);
  const [trajError, setTrajError] = useState(null);

  const [leversData, setLeversData]     = useState(null);
  const [leversLoading, setLeversLoading] = useState(true);
  const [leversError, setLeversError]   = useState(null);

  const [watchResult, setWatchResult]     = useState(null);
  const [watchComputing, setWatchComputing] = useState(false);

  const [history, setHistory] = useState([]);
  const [chatLog, setChatLog] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [library, setLibrary] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [viewTab, setViewTab] = useState("levers_input");
  const endRef = useRef(null);

  // ── Load df_2025.xlsx from public/ on mount ──────────────────────────────
  // raw:true skips SheetJS type coercion (faster); defval:"" avoids undefined.
  // aggregateBaseline collapses 340K rows → unique key groups before any
  // computation, cutting applyLeversV3 work from ~500M ops to ~7M ops.
  useEffect(() => {
    fetch("/df_2025.xlsx")
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status} — make sure df_2025.xlsx is in the public/ folder`);
        return r.arrayBuffer();
      })
      .then((buf) => {
        const wb = XLSX.read(buf, { type: "array", cellDates: false });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { raw: true, defval: "" });
        const aggregated = aggregateBaseline(rows);
        console.log(`df_2025: ${rows.length} rows → ${aggregated.length} aggregated groups`);
        setTrajData(aggregated);
      })
      .catch((e) => setTrajError(e.message))
      .finally(() => setTrajLoading(false));
  }, []);

  // ── Auto-load levers.xlsx from public/ on mount ───────────────────────────
  useEffect(() => {
    fetch("/levers.xlsx")
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status} — place levers.xlsx in the public/ folder or upload below`);
        return r.arrayBuffer();
      })
      .then((buf) => {
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        setLeversData(xlsxToJsonSafe(ws));
      })
      .catch((e) => setLeversError(e.message))
      .finally(() => setLeversLoading(false));
  }, []);

  // ── Compute scenarios when df_2025 + levers are both ready ───────────────
  useEffect(() => {
    if (!trajData?.length || !leversData?.length) { setWatchResult(null); return; }
    setWatchComputing(true);
    const id = setTimeout(() => {
      try {
        setWatchResult(applyLeversV3(trajData, leversData));
      } catch (e) {
        console.error("applyLeversV3 error:", e);
      }
      setWatchComputing(false);
    }, 0);
    return () => clearTimeout(id);
  }, [trajData, leversData]);

  // ── AI send ──────────────────────────────────────────────────────────────
  async function send(text) {
    const msg = (text || input).trim();
    if (!msg) return;
    setInput("");
    setError(null);
    setLoading(true);
    const log = [...chatLog, { role: "user", text: msg }];
    setChatLog(log);
    try {
      const { text: aiText, messages: hist } = await callClaude(msg, history);
      try {
        const parsed = tryJSON(aiText);
        const liRow = buildLeversInputRow(parsed);
        const ltRows = buildLeversTabRows(parsed);
        const entry = {
          id: `ai_${Date.now()}`,
          product_family: parsed.product_family || "Unknown",
          product_line: parsed.product_line_code || "",
          liRow, ltRows, _spec: parsed,
        };
        setLibrary((prev) => [...prev, entry]);
        setActiveId(entry.id);
        setHistory(hist.concat([{ role: "assistant", content: aiText }]));
        setChatLog([
          ...log,
          { role: "assistant", ok: true, text: `✨ Added "${parsed.product_family}" — ${ltRows.length} Levers tab row${ltRows.length !== 1 ? "s" : ""}. Now in library.` },
        ]);
      } catch {
        setHistory(hist.concat([{ role: "assistant", content: aiText }]));
        setChatLog([...log, { role: "assistant", text: aiText }]);
      }
    } catch (e) {
      setError("API error: " + e.message);
    } finally {
      setLoading(false);
      setTimeout(() => endRef.current?.scrollIntoView({ behavior: "smooth" }), 80);
    }
  }

  // ── Derived data ─────────────────────────────────────────────────────────
  const allLiRows = library.filter((e) => e.liRow).map((e) => e.liRow);
  const allLtRows = library.flatMap((e) => e.ltRows || []);

  // ── Downloads ────────────────────────────────────────────────────────────
  const downloadLeversInput = () => {
    const liCSV = [LI_HEADERS.map(esc).join(","), ...allLiRows.map((r) => LI_COLS.map((c) => esc(r[c])).join(","))].join("\n");
    dlCSV("levers_input.csv", liCSV);
  };
  const downloadLeversTab = () => {
    const ltCSV = [LEVERS_TAB_COLS.map(esc).join(","), ...allLtRows.map((r) => LEVERS_TAB_COLS.map((c) => esc(r[c])).join(","))].join("\n");
    dlCSV("levers_tab.csv", ltCSV);
  };
  const downloadAll = () => { downloadLeversInput(); setTimeout(downloadLeversTab, 300); };

  // ── Upload Levers Input CSV ──────────────────────────────────────────────
  async function onUploadLeversInput(ev) {
    const file = ev.target.files?.[0];
    ev.target.value = "";
    if (!file) return;
    try {
      const text = await file.text();
      const rows = parseCSV(text);
      if (!rows.length) throw new Error("Empty file");
      const headers = rows[0] || [];
      if (!isLikelyLeversInputHeaders(headers)) {
        throw new Error("This file does not look like a Levers Input CSV (unexpected headers).");
      }
      const dataRows = rows.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
      if (!dataRows.length) throw new Error("No data rows found.");
      const newEntries = dataRows.map((cells, idx) => {
        const slice = cells.slice(0, LI_COLS.length);
        const liRow = {};
        LI_COLS.forEach((k, i) => { liRow[k] = slice[i] ?? ""; });
        const fam = liRow.product_family || `Imported ${idx + 1}`;
        const pl = liRow.product_line || "";
        return { id: `csv_${Date.now()}_${idx}`, product_family: fam, product_line: pl, liRow, ltRows: [], _spec: null, _source: file.name };
      });
      setLibrary((prev) => [...prev, ...newEntries]);
      setActiveId(newEntries[0]?.id || null);
      setChatLog((prev) => [...prev, { role: "assistant", ok: true, text: `📄 Imported ${newEntries.length} Levers Input row(s) from "${file.name}".` }]);
      setError(null);
    } catch (e) {
      setError(`Import error: ${e.message}`);
    }
  }

  // ── Upload levers.xlsx ────────────────────────────────────────────────────
  async function onUploadLevers(ev) {
    const file = ev.target.files?.[0];
    ev.target.value = "";
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = xlsxToJsonSafe(ws);
      if (!rows.length) throw new Error("No data rows found in the file.");
      setLeversData(rows);
      setLeversError(null);
    } catch (e) {
      setLeversError(`Levers import error: ${e.message}`);
    }
  }

  const EXAMPLES = [
    `Product: ATV750, product line IDVSD, LoB IC&D.\nComment: Variable Speed Drive (FAKE data). Lifecycle: new range under development, commercialised in 2028.\nMigration rate 60%, growth 6–7%.\nDesign with Less: 10%, "New range will be 10% lighter", committed.\nMat efficiency legacy: 10%, "mainly plastic, aluminum & Steel", committed.\nMat efficiency new offer: 20%, "New range will include 100% of responsible plastic & Steel", committed.\nLifetime extension: 5%, high risk.\nEE1: 10%, "10% more efficient with new electronic design", committed.\nEE2: 20%, "extra 20% efficiency with new SiC components", high risk.`,
    `Family EcoVSD, line L5, LoB IC&D. Migration 70%, growth 5–7%.\nDesign with Less: 12%, committed.\nMat eff legacy: 8%, low risk.\nMat eff new offer: 18%, medium risk.\nEE1 Hub1+Hub2: 25%, committed.\nEE2 Hub1 only: 10%, high risk.`,
  ];

  return (
    <div style={{ fontFamily: "'DM Sans','Segoe UI',sans-serif", background: C.off, height: "100vh", display: "flex", flexDirection: "column", overflow: "hidden" }}>

      {/* ── Header ── */}
      <div
        style={{
          flexShrink: 0,
          background: `linear-gradient(90deg,${C.dark} 0%,#2d6a35 60%,${C.mid} 100%)`,
          padding: "10px 18px",
          display: "flex",
          alignItems: "center",
          gap: 12,
          boxShadow: "0 2px 10px #0005",
        }}
      >
        <div style={{ width: 30, height: 30, borderRadius: 7, background: C.bright, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 900, color: C.dark }}>⚡</div>

        {/* Title + nav tabs */}
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 800, color: C.white, letterSpacing: 0.3 }}>Schneider Electric · Lever Generator</div>
            <div style={{ fontSize: 10, color: C.white + "99" }}>Prompt new entries · View & download everything together</div>
          </div>
          <div style={{ display: "flex", gap: 4 }}>
            {[
              ["levers", "⚡ Lever Generator"],
              ["watch", "🎯 2030 Trajectory Watch"],
            ].map(([id, label]) => (
              <button
                key={id}
                onClick={() => setPage(id)}
                style={{
                  padding: "5px 12px",
                  fontSize: 11,
                  fontWeight: 700,
                  borderRadius: 6,
                  border: "none",
                  cursor: "pointer",
                  background: page === id ? C.bright : "rgba(255,255,255,0.15)",
                  color: page === id ? C.dark : C.white,
                }}
              >
                {label}
              </button>
            ))}
          </div>
        </div>

        {/* Upload + Download (only shown on levers page) */}
        {page === "levers" && (
          <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
            <label
              style={{
                background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.25)",
                borderRadius: 6, padding: "6px 10px", color: C.white, fontWeight: 700,
                fontSize: 11, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6, userSelect: "none",
              }}
              title="Import Levers Input CSV"
            >
              ⬆ Import Levers Input
              <input type="file" accept=".csv,text/csv" onChange={onUploadLeversInput} style={{ display: "none" }} />
            </label>
            {library.length > 0 && (
              <button
                onClick={downloadAll}
                style={{ background: C.bright, border: "none", borderRadius: 6, padding: "6px 12px", color: C.dark, fontWeight: 700, fontSize: 11, cursor: "pointer" }}
                title="Download both files"
              >
                ↓ Download All ({library.length})
              </button>
            )}
          </div>
        )}
      </div>

      {/* ── Page body ── */}
      {page === "watch" ? (
        <div style={{ flex: 1, minHeight: 0, overflowY: "auto" }}>
          <TrajectoryWatch
            df2025={trajData} df2025Loading={trajLoading} df2025Error={trajError}
            levers={leversData} leversLoading={leversLoading} leversError={leversError}
            onUploadLevers={onUploadLevers}
            watchResult={watchResult} watchComputing={watchComputing}
          />
        </div>
      ) : (
        <div style={{ flex: 1, minHeight: 0, display: "grid", gridTemplateColumns: library.length > 0 ? "340px 1fr" : "1fr" }}>

          {/* ── Left: Chat + Library ── */}
          <div style={{ display: "flex", flexDirection: "column", borderRight: library.length > 0 ? "1px solid #cdd8d0" : "none", minHeight: 0, background: C.white }}>

            {/* Library panel */}
            {library.length > 0 && (
              <div style={{ flexShrink: 0, borderBottom: "1px solid #e0ece0", background: "#f8fdf8" }}>
                <div style={{ padding: "8px 12px 6px" }}>
                  <span style={{ fontSize: 11, fontWeight: 700, color: C.text }}>Library ({library.length})</span>
                </div>
                <div style={{ maxHeight: 160, overflowY: "auto", padding: "0 8px 8px" }}>
                  {library.map((e) => (
                    <div
                      key={e.id}
                      onClick={() => setActiveId(e.id)}
                      style={{
                        display: "flex", alignItems: "center", gap: 8, padding: "6px 8px",
                        borderRadius: 6, cursor: "pointer", marginBottom: 3,
                        background: activeId === e.id ? C.light : "transparent",
                        border: activeId === e.id ? `1px solid ${C.bright}55` : "1px solid transparent",
                      }}
                    >
                      <span style={{ fontSize: 10, width: 16, textAlign: "center" }}>{e._source ? "📄" : "✨"}</span>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: C.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.product_family}</div>
                        <div style={{ fontSize: 10, color: C.grey }}>
                          {e.product_line} · {e.ltRows?.length || 0} lever rows{e._source ? ` · ${e._source}` : ""}
                        </div>
                      </div>
                      <button
                        onClick={(ev) => {
                          ev.stopPropagation();
                          setLibrary((prev) => prev.filter((x) => x.id !== e.id));
                          if (activeId === e.id) setActiveId(null);
                        }}
                        style={{ background: "none", border: "none", cursor: "pointer", color: C.grey, fontSize: 13, padding: "0 2px", lineHeight: 1 }}
                        title="Remove entry"
                      >
                        ×
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Chat messages */}
            <div style={{ flex: 1, overflowY: "auto", padding: "10px 12px", display: "flex", flexDirection: "column", gap: 8 }}>
              {chatLog.length === 0 && (
                <div>
                  <div style={{ padding: "10px 0 6px", fontSize: 11, fontWeight: 700, color: C.grey, textTransform: "uppercase", letterSpacing: 1 }}>
                    Examples — click to try
                  </div>
                  {EXAMPLES.map((p, i) => (
                    <button
                      key={i}
                      onClick={() => send(p)}
                      style={{
                        display: "block", width: "100%", background: C.off,
                        border: `1px solid #c8d8c8`, borderLeft: `3px solid ${C.bright}`,
                        borderRadius: 6, padding: "8px 10px", textAlign: "left",
                        cursor: "pointer", fontSize: 11, color: C.text,
                        lineHeight: 1.5, marginBottom: 6, whiteSpace: "pre-line",
                      }}
                    >
                      {p.slice(0, 140)}…
                    </button>
                  ))}
                </div>
              )}
              {chatLog.map((m, i) => (
                <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start" }}>
                  <div
                    style={{
                      maxWidth: "90%",
                      background: m.role === "user" ? `linear-gradient(135deg,${C.mid},${C.dark})` : m.ok ? C.light : C.white,
                      color: m.role === "user" ? C.white : C.text,
                      border: m.role === "user" ? "none" : `1px solid ${m.ok ? C.bright + "55" : "#dde8dd"}`,
                      borderRadius: m.role === "user" ? "14px 14px 4px 14px" : "14px 14px 14px 4px",
                      padding: "8px 12px", fontSize: 12, lineHeight: 1.6,
                    }}
                  >
                    {m.text}
                  </div>
                </div>
              ))}
              {loading && (
                <div style={{ display: "flex" }}>
                  <div style={{ background: C.white, border: "1px solid #dde8dd", borderRadius: "14px 14px 14px 4px", padding: "9px 13px", display: "flex", gap: 5 }}>
                    {[0, 1, 2].map((j) => (
                      <div key={j} style={{ width: 6, height: 6, borderRadius: "50%", background: C.bright, animation: `pulse 1.2s ${j * 0.2}s ease-in-out infinite` }} />
                    ))}
                  </div>
                </div>
              )}
              {error && <div style={{ background: "#ffebee", border: "1px solid #ef9a9a", borderRadius: 7, padding: "8px 12px", fontSize: 11, color: C.rHigh }}>{error}</div>}
              <div ref={endRef} />
            </div>

            {/* Input */}
            <div style={{ padding: "9px 12px", borderTop: "1px solid #e0ece0", background: C.white, flexShrink: 0 }}>
              <div style={{ display: "flex", gap: 6 }}>
                <textarea
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
                  placeholder="Describe a product family — lever types, improvement %, migration rate, hub, risk…"
                  disabled={loading}
                  rows={2}
                  style={{
                    flex: 1, border: `1.5px solid ${input ? C.bright : "#cdd8d0"}`,
                    borderRadius: 8, padding: "8px 10px", fontSize: 12,
                    fontFamily: "inherit", resize: "none", outline: "none",
                    lineHeight: 1.5, color: C.text,
                  }}
                />
                <button
                  onClick={() => send()}
                  disabled={loading || !input.trim()}
                  style={{
                    background: loading || !input.trim() ? "#ccc" : `linear-gradient(135deg,${C.mid},${C.dark})`,
                    border: "none", borderRadius: 8, padding: "0 12px", color: C.white,
                    fontWeight: 700, fontSize: 12,
                    cursor: loading || !input.trim() ? "not-allowed" : "pointer",
                    minWidth: 56,
                  }}
                >
                  {loading ? "…" : "Send"}
                </button>
              </div>
              <div style={{ fontSize: 10, color: C.grey, marginTop: 3 }}>Enter to send · Shift+Enter for newline</div>
            </div>
          </div>

          {/* ── Right: Data view ── */}
          {library.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", minHeight: 0 }}>
              <div style={{ flexShrink: 0, borderBottom: "1px solid #cdd8d0", padding: "8px 16px", display: "flex", alignItems: "center", gap: 10, background: C.white }}>
                {[
                  ["levers_input", "📋 Levers Input", allLiRows.length],
                  ["levers_tab", "📊 Levers tab", allLtRows.length],
                ].map(([id, label, count]) => (
                  <button
                    key={id}
                    onClick={() => setViewTab(id)}
                    style={{
                      padding: "5px 12px", fontSize: 12, fontWeight: 700, borderRadius: 6,
                      border: "none", cursor: "pointer",
                      background: viewTab === id ? C.mid : "#e8f0e8",
                      color: viewTab === id ? C.white : C.grey,
                    }}
                  >
                    {label} <span style={{ marginLeft: 4, fontSize: 10, opacity: 0.8 }}>({count})</span>
                  </button>
                ))}
                <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
                  {viewTab === "levers_input" && (
                    <button onClick={downloadLeversInput} style={{ background: "#e8f0e8", border: "none", borderRadius: 6, padding: "6px 10px", fontWeight: 700, fontSize: 11, cursor: "pointer", color: C.text }}>
                      ↓ Download Levers Input
                    </button>
                  )}
                  {viewTab === "levers_tab" && (
                    <button onClick={downloadLeversTab} style={{ background: "#e8f0e8", border: "none", borderRadius: 6, padding: "6px 10px", fontWeight: 700, fontSize: 11, cursor: "pointer", color: C.text }}>
                      ↓ Download Levers Tab
                    </button>
                  )}
                  <div style={{ fontSize: 11, color: C.grey }}>{library.length} product famil{library.length === 1 ? "y" : "ies"}</div>
                </div>
              </div>
              <div style={{ flex: 1, overflowY: "auto", padding: "14px 16px" }}>
                {viewTab === "levers_input" && (allLiRows.length === 0
                  ? <div style={{ color: C.grey, fontSize: 13, padding: "40px 0", textAlign: "center" }}>No data yet. Add products via chat or import CSV.</div>
                  : <LeversInputTable rows={allLiRows} />)}
                {viewTab === "levers_tab" && (allLtRows.length === 0
                  ? <div style={{ color: C.grey, fontSize: 13, padding: "40px 0", textAlign: "center" }}>No Levers tab data yet. Add products via chat.</div>
                  : <LeversTabTable rows={allLtRows} />)}
              </div>
            </div>
          )}
        </div>
      )}

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap');
        @keyframes pulse{0%,80%,100%{opacity:.2;transform:scale(.8)}40%{opacity:1;transform:scale(1)}}
        *{box-sizing:border-box}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-thumb{background:#b0c4b0;border-radius:3px}
      `}</style>
    </div>
  );
}
