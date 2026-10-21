import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* Qlik P&L app bc30d66f, pulled 2026-10-15. Service = GROSS PROFIT / INCOME - Mechanical Dept, Parts = GROSS PROFIT /
   INCOME - P&A Dept (Detail statement; these equal the GM lines Total Gross Profit - Service / - Parts). Fixed Ops here =
   Service + Parts, i.e. the P&L's Total Gross Profit - Fixed Ops less GMC Prescott's body shop gross (verified to the
   dollar). Same-store, INEOS excluded. RangeSum per store; store rows reconcile to the group, both years. */
const SEG = "Fixed Ops Gross";
const SRC = "Qlik - Profit & Loss. Same-store, INEOS excluded (opened 2026). Fixed ops = service + parts; GMC Prescott body shop gross excluded.";
const N = JSON.parse(fs.readFileSync("./ntgs.json", "utf8")), P = JSON.parse(fs.readFileSync("./ntgp.json", "utf8"));
const MON = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP"];
const FULL = ["January","February","March","April","May","June","July","August","September"];
const sv25 = N.m25.map(r => r[1]), sv26 = N.m26.map(r => r[1]), pt25 = P.m25.map(r => r[1]), pt26 = P.m26.map(r => r[1]);
const Q4 = { sv: [13742904.77, 11835094.76, 12919268.17], pt: [7528553.86, 5908964.47, 6398186.93] };   // Oct-Dec 2025
const svS = N.S.map(s => [s[0], s[2], s[4]]), ptS = P.S.map(s => [s[0], s[2], s[4]]);
const fxS = svS.map((r, i) => [r[0], r[1] + ptS[i][1], r[2] + ptS[i][2]]);
const sum = a => a.reduce((t, x) => t + x, 0);
const M = v => "$" + (v / 1e6).toFixed(2) + "M";
const usd = v => (v >= 0 ? "+" : "\u2212") + "$" + Math.abs(Math.round(v)).toLocaleString("en-US");
const MLAB = ["Oct '25","Nov '25","Dec '25","Jan '26","Feb '26","Mar '26","Apr '26","May '26","Jun '26","Jul '26","Aug '26","Sep '26"];

H.section(pptx, pptx, { kicker: "Fixed Ops", title: "Gross", chips: ["Fixed ops", "Service", "Parts"] });

function grossBlock(title, metric, m25, m26, q4, S, chain) {
  const t12 = [...q4, ...m26];
  const v25 = sum(m25), v26 = sum(m26), d = v26 - v25;
  const yo = m26.map((v, i) => v / m25[i] - 1);
  const lo = Math.min(...t12), hi = Math.max(...t12), hiI = t12.indexOf(hi), loI = t12.indexOf(lo);
  const up = yo.filter(x => x > 0).length;
  const last4 = yo.slice(-4).map(x => (x >= 0 ? "+" : "\u2212") + Math.abs(x * 100).toFixed(1) + "%");
  // Day 1 group picture (Oct 21 rebuild): 2025 January-December against 2026 to date, bars
  const G = JSON.parse(fs.readFileSync("./gross/monthly.json", "utf8"));
  const full25 = title.endsWith("Fixed Ops") ? G["2025"].sv.map((v, i) => v + G["2025"].pt[i]) : title.endsWith("Service") ? G["2025"].sv : G["2025"].pt;
  H.groupYoy(pptx, pptx, {
    eyebrow: SEG, title, metric, measure: title.replace("Total Gross Profit - ", "") + " Gross",
    prior: full25, current: m26, priorName: "2025", currentName: "2026", expectPrior: v25, expectCurrent: v26,
    chartLabel: `${title.replace("Total Gross Profit - ", "")} gross by month, 2025 against 2026 ($M)`,
    takeaway: `${usd(d)} (${(d / v25 * 100).toFixed(1)}%) year to date; ${up} of nine months ahead of 2025, best ${FULL[m26.map((v, i) => v - m25[i]).indexOf(Math.max(...m26.map((v, i) => v - m25[i])))]} at ${usd(Math.max(...m26.map((v, i) => v - m25[i])))}.`,
    source: `${SRC} Jan-Sep 2026 vs Jan-Sep 2025; 2025 shown January-December.`,
  });
  H.totalRanker(pptx, pptx, {
    eyebrow: SEG, title, metric, measure: title.replace("Total Gross Profit - ", "") + " Gross", rows: S.map(r => [r[0], r[2]]), expect: v26,
    groupLabel: "SEPTEMBER 2026 YTD",
    takeaway: (() => { const t = [...S].sort((p, q) => q[2] - p[2]); return `${H.storeName(t[0][0])} is ${H.Mmixed(t[0][2])} of the group's ${H.Mmixed(v26)}; ${H.storeName(t.at(-1)[0])} is smallest at ${H.Mmixed(t.at(-1)[2])}.`; })(),
    tone: "n", source: `${SRC} Jan-Sep 2026.`,
  });
  const ds = S.map(r => [r[0], r[2] - r[1]]).sort((a, b) => b[1] - a[1]);
  const grew = ds.filter(r => r[1] > 0).length;
  H.ranker(pptx, pptx, {
    eyebrow: SEG, title, metric, rows: S, expect: d, groupLabel: "2026 VS 2025 YTD · **YEAR-OVER-YEAR CHANGE**", measure: title.replace("Total Gross Profit - ", "") + " Gross",
    takeaway: `${grew} of 33 stores grew, led by ${ds[0][0]} at ${usd(ds[0][1])}; ${33 - grew} declined, ${ds[32][0]} most at ${usd(ds[32][1])}.`,
    source: `${SRC} Jan-Sep 2026 vs 2025, best first.`,
  });
  if (chain) {
    const top = ds.slice(0, 5), bot = ds.slice(-4);
    H.chain(pptx, pptx, {
      eyebrow: SEG, title, panelLabel: "How the group got there, Jan\u2013Sep",
      steps: [{ v: v25, l: "2025 group" }, { v: d, l: "same-store growth" }], total: v26, totalLabel: "2026 group",
      moversLabel: "Who moved it", movers: [...top, ...bot],
      takeaway: `Five stores added ${usd(sum(top.map(r => r[1])))}; the bottom four gave back ${usd(sum(bot.map(r => r[1]))).replace("\u2212", "")}.`,
      source: `${SRC} Top five and bottom four stores by variance.`,
    });
  }
}
grossBlock("Total Gross Profit - Fixed Ops", "fixed ops gross", sv25.map((v, i) => v + pt25[i]), sv26.map((v, i) => v + pt26[i]),
  Q4.sv.map((v, i) => v + Q4.pt[i]), fxS, true);
grossBlock("Total Gross Profit - Service", "service gross", sv25, sv26, Q4.sv, svS, false);
grossBlock("Total Gross Profit - Parts", "parts gross", pt25, pt26, Q4.pt, ptS, false);
await pptx.writeFile({ fileName: "Fixed Ops Gross - House Style.pptx" }); console.log("gross built");
