/*
  EPC Dashboard — tests
  =====================
  Usage:
    node tools/tests/tests.js                 -> logic tests on dist data
    node tools/tests/tests.js path/to/export  -> also test ERP parsing (xls/xlsx)
*/
const fs = require("fs");
const path = require("path");
const DIST = path.join(__dirname, "..", "..", "dist");

let XLSX = null;
try { XLSX = require(path.join(DIST, "js", "vendor", "xlsx.full.min.js")); } catch (e) { /* optional */ }

globalThis.window = globalThis;
eval(fs.readFileSync(path.join(DIST, "js", "data.js"), "utf8"));
eval(fs.readFileSync(path.join(DIST, "js", "app.js"), "utf8"));
const A = globalThis.__APP__;

let fails = 0;
const near = (n, a, b, t) => { t = t === undefined ? 0.02 : t; if (Math.abs(a - b) >= t) { fails++; console.log("FAIL " + n + ": " + a + " != " + b); } else console.log("ok   " + n + " = " + (a && a.toFixed ? a.toFixed(3) : a)); };
const eq = (n, a, b) => { if (a !== b) { fails++; console.log("FAIL " + n + ": " + a + " != " + b); } else console.log("ok   " + n + " = " + a); };

// --- KPI logic on the built data ---
const k = A.computeKPIs(A.DATA);
eq("packs", k.packs, A.DATA.length);
near("total = AL+CU+Scr+Arm", k.totalT, k.al + k.cu + k.st + k.at);
near("fam sum = total", Object.values(k.famT).reduce((s, v) => s + v, 0), k.totalT);
near("cum last day = total", k.cum[A.MDAYS], k.totalT);
if (A.HAS_PLAN) {
  near("remaining = plan - done", k.remaining, A.PLAN - k.totalT);
  eq("daysLeft", k.daysLeft, A.MDAYS - k.maxDay);
  near("paceToFinish", k.paceToFinish, k.remaining / Math.max(1, A.MDAYS - k.maxDay));
}
// filters
A.state.mkt = "exp";
const ke = A.computeKPIs(A.filteredRows());
near("export filter", ke.totalT, k.expT);
A.state.mkt = "all";
A.state.fams.add("MV");
const kmv = A.computeKPIs(A.filteredRows());
near("MV filter", kmv.totalT, k.famT.MV);
A.state.fams.clear();

// --- order (WO) search tests ---
A.state.wo = "803/26/01";
const r803 = A.filteredRows();
eq("search exact wo returns rows", r803.length > 0, true);
eq("all rows match exact wo", r803.every(r => r.wo === "803/26/01"), true);

A.state.wo = "803";
const rPartial = A.filteredRows();
eq("partial search wo returns rows", rPartial.length >= r803.length, true);
eq("partial rows contain 803", rPartial.every(r => r.wo.includes("803")), true);

// Arabic numerals search
A.state.wo = "٨٠٣";
const rArabic = A.filteredRows();
eq("arabic numerals wo search matches ascii", rArabic.length, rPartial.length);

// whitespace & slash tolerance
A.state.wo = " 803 / 26 / 01 ";
const rSpaces = A.filteredRows();
eq("wo search with slashes and spaces", rSpaces.length, r803.length);

// order totals calculation
A.state.wo = "803/26/01";
const rows803 = A.filteredRows();
const tTons = rows803.reduce((s, r) => s + r.tt, 0);
const tMeters = rows803.reduce((s, r) => s + r.q, 0);
near("order 803/26/01 total tons", tTons, 33.392, 0.01);
eq("order 803/26/01 total meters", tMeters, 10891);
eq("order 803/26/01 total packs", rows803.length, 27);

// renderWoSummary mock test
const mockEl = { style: {}, innerHTML: "" };
globalThis.document = { getElementById: id => (id === "woSummary" ? mockEl : null) };
A.renderWoSummary(rows803, A.computeKPIs(rows803));
eq("wo summary rendered visible", mockEl.style.display, "block");
eq("wo summary has tonnage", mockEl.innerHTML.includes("33.392"), true);
eq("wo summary has meters", mockEl.innerHTML.includes("10,891"), true);
delete globalThis.document;

A.state.wo = "";

// --- month helpers ---
eq("month key", A.cfgMonthKey({ year: 2026, month_num: 10 }), "2026-10");
eq("month label", A.monthLabel("2026-10"), "أكتوبر 2026");
A.storeMonth("2026-10", A.DATA, A.CONFIG);
eq("archive stores packs", (A.getMONTHS()["2026-10"] || {}).data ? A.getMONTHS()["2026-10"].data.length : -1, A.DATA.length);

// --- optional: parse a real ERP export ---
const target = process.argv[2];
if (target && XLSX) {
  const buf = fs.readFileSync(target);
  const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  const sn = wb.SheetNames.includes("0040_qa") ? "0040_qa" : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
  const res = A.parseErpRows(rows, path.basename(target));
  console.log("parsed " + path.basename(target) + ": " + res.data.length + " packs, " + res.year + "-" + res.month);
  eq("parse packs>0", res.data.length > 0, true);
  const t = res.data.reduce((s, r) => s + r.tt, 0);
  console.log("parsed total: " + t.toFixed(3) + " t");
} else if (target) {
  console.log("SKIP parse test (vendor lib not found)");
}

console.log(fails === 0 ? "TESTS: ALL PASS" : "TESTS: " + fails + " FAILURES");
process.exit(fails === 0 ? 0 : 1);
