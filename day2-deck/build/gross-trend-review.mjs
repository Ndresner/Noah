import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* REVIEW (Noah, Oct 20 2026): trailing-12 gross line replaced by Jan-Sep 2026 vs Jan-Sep 2025.
   Option A = two-line chart, table slide kept. Option B = chart and YOY card merged, table slide dropped.
   Option B + Q4 = the 2025 line runs through December. Plus the headcount trend on the same basis.
   Data: gross/monthly.json - Qlik P&L app bc30d66f, pulled 2026-10-20, GM lines Total Gross Profit - Service /
   - Parts, INEOS excluded, signs flipped. Jan-Sep reconciles to the verified fixed gross $181,331,157 / $188,799,724. */
const SEG = "Fixed Ops Gross";
const SRC = "Qlik - Profit & Loss. Same-store, INEOS excluded (opened 2026). Fixed ops = service + parts; GMC Prescott body shop gross excluded.";
const G = JSON.parse(fs.readFileSync("./gross/monthly.json", "utf8"));
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const FULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const sum = a => a.reduce((t, x) => t + x, 0);
const usd = v => (v >= 0 ? "+" : "−") + "$" + Math.abs(Math.round(v)).toLocaleString("en-US");
const M = v => "$" + (v / 1e6).toFixed(2) + "M";
const pc = (d, b) => (d >= 0 ? "+" : "−") + Math.abs(d / b * 100).toFixed(1) + "%";
const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);

const fx = y => G[y].sv.map((v, i) => v + G[y].pt[i]);
const LINES = [
  { title: "Total Gross Profit - Fixed Ops", metric: "fixed ops gross", short: "Fixed ops", a: fx("2025"), b: fx("2026").slice(0, 9) },
  { title: "Total Gross Profit - Service",   metric: "service gross",   short: "Service",   a: G["2025"].sv, b: G["2026"].sv.slice(0, 9) },
  { title: "Total Gross Profit - Parts",     metric: "parts gross",     short: "Parts",     a: G["2025"].pt, b: G["2026"].pt.slice(0, 9) },
];
const EXPECT = { "fixed ops gross": 188799724 - 181331157 };

function facts(L) {
  const a = L.a.slice(0, 9), b = L.b, d = b.map((v, i) => v - a[i]), v = sum(b) - sum(a);
  const ahead = d.filter(x => x > 0).length, behind = MON.slice(0, 9).filter((_, i) => d[i] <= 0);
  let run = 0; for (let i = 8; i >= 0 && d[i] > 0; i--) run++;
  const bi = d.indexOf(Math.max(...d));
  if (EXPECT[L.metric] != null && Math.abs(v - EXPECT[L.metric]) > 1) throw new Error(`${L.title}: ${v} vs ${EXPECT[L.metric]}`);
  const since = run >= 2 ? `, and in every month since ${FULL[9 - run]}` : "";
  return { a, b, d, v, ahead, behind, bi, run,
    takeaway: `${L.short} gross ${v >= 0 ? "up" : "down"} ${usd(v).slice(1)} (${pc(v, sum(a)).replace("+", "")}) year to date; ${ahead === 9 ? "ahead of 2025 in every month" : `ahead of 2025 in ${ahead} of 9 months${since}`}.`,
    bullets: [
      ahead === 9 ? "Ahead of 2025 in every month." : `Ahead of 2025 in ${ahead} of 9 months.`,
      `Best month: ${FULL[bi]}, ${d[bi] >= 0 ? "up" : "down"} ${usd(d[bi]).slice(1)} (${pc(d[bi], a[bi]).slice(1)}).`,
      behind.length ? `Behind in ${list(behind)}.` : `No month behind 2025.`,
    ] };
}
const card = "Group · January–September · same-store, 33 stores";

// ---- Option A: two-line chart, table slide kept
for (const L of LINES) {
  const f = facts(L);
  H.trendYoy(pptx, pptx, { eyebrow: `${SEG} · Option A`, title: L.title, metric: L.metric,
    panelLabel: "Group total by month, January–September ($M)", months: MON.slice(0, 9), a: f.a, b: f.b,
    bullets: f.bullets, takeaway: f.takeaway, tone: f.v >= 0 ? "g" : "b",
    source: `${SRC} Jan-Sep 2026 vs Jan-Sep 2025 by month.` });
  H.yoy(pptx, pptx, { eyebrow: `${SEG} · Option A`, title: L.title, metric: L.metric, cardLabel: card,
    months: MON.slice(0, 9).map((m, i) => [m.toUpperCase(), f.a[i], f.b[i]]), expect: f.v,
    takeaway: `${usd(f.v)} (${pc(f.v, sum(f.a)).replace("+", "")}) year to date; ${f.ahead} of nine months ahead of 2025, best ${FULL[f.bi]} at ${usd(f.d[f.bi])}.`,
    source: `${SRC} Jan-Sep 2026 vs 2025.` });
}

// ---- Option B: chart and card merged on one slide
for (const L of LINES) {
  const f = facts(L);
  H.trendYoy(pptx, pptx, { eyebrow: `${SEG} · Option B`, title: L.title, metric: L.metric, card: true, cardLabel: card,
    panelLabel: "Group total by month ($M)", months: MON.slice(0, 9), a: f.a, b: f.b,
    takeaway: f.takeaway, tone: f.v >= 0 ? "g" : "b",
    source: `${SRC} Jan-Sep 2026 vs Jan-Sep 2025 by month. Replaces the trend and table slides.` });
}

// ---- Option B + Q4: 2025 runs through December
{
  const L = LINES[0], f = facts(L), q4 = L.a.slice(9);
  H.trendYoy(pptx, pptx, { eyebrow: `${SEG} · Option B + Q4`, title: L.title, metric: L.metric, card: true, cardLabel: card,
    panelLabel: "Group total by month ($M); 2025 shown through December", months: MON, a: L.a, b: f.b,
    takeaway: f.takeaway, tone: f.v >= 0 ? "g" : "b",
    source: `${SRC} Card and takeaway are Jan-Sep. 2025 Oct-Dec shown for context: ${q4.map(M).join(", ")}; Q4 2025 total ${M(sum(q4))}.` });
}

// ---- Headcount on the same basis (prod/trend.json: Jan 25 - Sep 26)
{
  const T = JSON.parse(fs.readFileSync("./prod/trend.json", "utf8"));
  const a = T.fe.slice(0, 9), b = T.fe.slice(12, 21), d = b[8] - a[8];
  const sg = v => (v > 0 ? "+" : v < 0 ? "−" : "") + Math.abs(v).toLocaleString("en-US");
  const mv = k => sg(T[k][20] - T[k][8]);
  H.trendYoy(pptx, pptx, { eyebrow: "Fixed Ops Productivity · Option A", title: "Fixed Ops Headcount", metric: "fixed headcount",
    panelLabel: "Fixed employees at month end, January–September", months: MON.slice(0, 9), a, b,
    scale: 1, axisFormat: "#,##0", majorUnit: 20,
    min: Math.floor((Math.min(...a, ...b) - 10) / 20) * 20, max: Math.ceil((Math.max(...a, ...b) + 10) / 20) * 20,
    bullets: [
      `${sg(d)} fixed staff, Sep to Sep, to ${b[8].toLocaleString("en-US")}.`,
      `Technicians ${mv("tech")}, advisors ${mv("sa")}, managers ${mv("mgr")}.`,
      (() => { const up = b.filter((v, i) => v > a[i]).length, dn = b.filter((v, i) => v < a[i]).length;
        return up === 9 ? "2026 above 2025 in every month." : dn === 9 ? "2026 below 2025 in every month."
          : `2026 above 2025 in ${up} of 9 months.`; })(),
    ],
    takeaway: `Fixed ops headcount ${d >= 0 ? "rose" : "fell"} ${Math.abs(d)} to ${b[8].toLocaleString("en-US")}, Sep to Sep.`, tone: "n",
    source: "Headcount: ADP Employee Master extract, 10/19/2026, month-end snapshots, Qlik's fixed-employee title list (managers included). Each person counted once: someone open at two stores is counted at the old store until that record closes. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop staff excluded." });
}

await pptx.writeFile({ fileName: "Gross Trend - Review.pptx" }); console.log("gross trend review built");
