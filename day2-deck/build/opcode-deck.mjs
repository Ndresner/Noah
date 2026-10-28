import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
import { moversSlide } from "./blocks.mjs";

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };

/* ---- Qlik Closed Repair Orders app c3efc739 (Fixed space), pulled 2026-10-14.
   Pay type C, closed ROs Jan-Sep by close month. Same-store: INEOS excluded.
   GMC Prescott body shop labor types excluded. 33 store rows reconcile to the
   group on numerator and denominator, both years. ---- */
const A = JSON.parse(fs.readFileSync("./all.json", "utf8"));
const MON = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP"];
const SRC = "Qlik - Closed Repair Orders. Customer pay, Jan-Sep by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor excluded.";
const SEG = "Opcode Benchmarks";
const sum = a => a.reduce((t, x) => t + x, 0);
const p1  = v => (v * 100).toFixed(1) + "%";
const fmt = {
  v: v => (Math.floor(v * 1000 + 1e-9) / 10).toFixed(1) + "%",   // truncated, same as the ranker
  d: v => (v >= 0 ? "+" : "\u2212") + Math.abs(v * 100).toFixed(1) + " pts",
  p: (d, a) => (d >= 0 ? "+" : "\u2212") + Math.abs(d / a * 100).toFixed(1) + "%",
};
const pt = v => (Math.floor(v * 1000 + 1e-9) / 10).toFixed(1) + "%";   // truncated, matches the benchmark rankers
const tr = v => Math.floor(v * 1000 + 1e-9) / 10;   // the truncated rate in points
const NUM = ["zero","one","two","three","four","five","six","seven","eight","nine"];

H.section(pptx, pptx, {
  kicker: "Fixed Ops", title: SEG,
  chips: ["Alignment 15%", "Balance 15%", "Brake fluid 10%", "Fuel service 7.5%"],
});

function pair(key, title, metric, measure, bench, noun, nouns) {
  const M = A[key], d25 = A.den25, d26 = A.den26;
  const y25 = sum(M.m25) / sum(d25), y26 = sum(M.m26) / sum(d26);
  const up = MON.filter((m, i) => M.m26[i] / d26[i] > M.m25[i] / d25[i]).length;
  const gap = Math.round(bench * sum(d26) - sum(M.m26));
  const B = p1(bench).replace(".0%", "%");

  // Day 1 group picture (Oct 21 rebuild): a rate, so a line, in percent
  H.groupYoy(pptx, pptx, {
    eyebrow: SEG, title, metric, measure,
    prior: [...MON.map((m, i) => M.m25[i] / d25[i] * 100), ...M.q4.map((v, i) => v / A.den_q4[i] * 100)], current: MON.map((m, i) => M.m26[i] / d26[i] * 100),
    priorName: "2025", currentName: "2026", ytdPrior: y25 * 100, ytdCurrent: y26 * 100, unit: "pct", chart: "line",
    fmt: v => fmt.v(v / 100), dfmt: v => fmt.d(v / 100),
    chartLabel: `${title} by month, 2025 against 2026`,
    // dashed benchmark line (Noah, Oct 24); axis fitted to the data, not stretched by the line (Oct 26)
    benchLines: [bench * 100], benchLabel: `Benchmark ${B}`, tightBench: true,
    // Oct 28 audit: points from the shown (truncated) rates, so 11.9% -> 13.2% reads +1.3 and the gap is off the shown 13.2%
    takeaway: `Up ${(tr(y26) - tr(y25)).toFixed(1)} pts to ${fmt.v(y26)}, ahead of 2025 in ${up === 9 ? "all nine" : NUM[up] + " of nine"} months, `
            + (y26 >= bench ? `and above the ${B} benchmark.` : `but ${(bench * 100 - tr(y26)).toFixed(1)} pts under the ${B} benchmark \u2014 ${gap.toLocaleString()} ${nouns} short.`),
    source: `${SRC} ${measure} = ${noun} ROs \u00f7 customer-pay ROs on the op-code report basis. YTD from YTD totals, not the mean of months.`,
  });

  const rows = M.S.map(([n, a, b, c, d]) => [n, c / d, a / b]);
  const at = rows.filter(r => r[1] >= bench).length;
  const lead = [...rows].sort((p, q) => q[1] - p[1])[0];
  const zero = rows.filter(r => r[1] < 0.005).map(r => r[0]);
  const jump = rows.filter(r => r[2] < 0.005 && r[1] >= 0.02).map(r => r[0]);
  let dq = "";
  // Oct 28 (Noah): Lexus balance near zero is real, not a mapping problem - no CHECK line
  // Oct 28 audit: near-zero-in-2025 QA line removed from the notes (Mazda's jump confirmed real, Oct 19)

  H.levelRanker(pptx, pptx, {
    eyebrow: SEG, title, metric,
    rows: rows.map(r => [r[0], r[1]]),
    group: y26, bench, fmt: pt, measure, groupLabel: "SEPTEMBER 2026 YTD",   // truncated so just-under-benchmark never reads as the benchmark
        takeaway: `${at} of 33 stores at or above ${B}, led by ${lead[0]} at ${pt(lead[1])}. The group sits at ${pt(y26)}.`,
    source: `${SRC} Jan-Sep 2026 YTD; green at or above the ${B} benchmark, red below.${dq}\n`
          + rows.sort((p, q) => q[1] - p[1]).map(r => `${r[0]}: ${pt(r[1])} (2025 ${pt(r[2])})`).join("\n"),   // notes match the slide (Oct 28 audit)
  });
  // top and bottom five by change in points (Noah, Oct 22)
  moversSlide(pptx, { seg: SEG, title, metric, fmtD: fmt.d,
    rows: rows.map(r => ({ n: r[0], v0: pt(r[2]), v1: pt(r[1]), d: r[1] - r[2] })),
    // moving now (Oct 23): Jul-Sep vs Jan-Mar 2026 closing %, from all.json mn = [store, num_q1, den_q1, num_q3, den_q3]
    now: M.mn.map(([n, a, b, c, d]) => ({ n, v0: pt(a / b), v1: pt(c / d), d: c / d - a / b })),
    nowGroup: fmt.d(M.mn.reduce((t, r) => t + r[3], 0) / M.mn.reduce((t, r) => t + r[4], 0) - M.mn.reduce((t, r) => t + r[1], 0) / M.mn.reduce((t, r) => t + r[2], 0)),
    src: `${SRC} Top and bottom five stores by change in ${title}, Jan-Sep 2026 vs 2025, in points.` });
}

pair("align", "Alignment %",     "alignment %",    "ALIGNMENT %",    0.15,  "alignment", "alignments");
pair("bal",   "Balance %",       "balance %",      "BALANCE %",      0.15,  "balance", "balances");
pair("bf",    "Brake Fluid %",   "brake fluid %",  "BRAKE FLUID %",  0.10,  "brake fluid service", "brake fluid services");   // title per Noah, Oct 15 2026
pair("fuel",  "Fuel Service %",  "fuel service %", "FUEL SERVICE %", 0.075, "fuel service", "fuel services");

await pptx.writeFile({ fileName: "Opcode Benchmarks - House Style.pptx" });
console.log("built");
