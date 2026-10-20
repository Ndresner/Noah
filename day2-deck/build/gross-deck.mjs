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
  H.trend(pptx, pptx, {
    eyebrow: SEG, title, panelLabel: "Group total, trailing 12 months ($M)",
    series: MLAB.map((m, i) => [m, t12[i]]), min: Math.floor(lo / 1e6 - 0.5), max: Math.ceil(hi / 1e6 + 0.5),
    bullets: [
      `Peak ${MLAB[hiI]} at ${M(hi)}. Low ${MLAB[loI]} at ${M(lo)}.`,
      `Jun\u2013Sep year over year: ${last4.join(", ")}.`,
      `${up} of 9 months in 2026 beat the same month of 2025.`,
    ],
    takeaway: `${title.replace("Total Gross Profit - ", "")} gross ran ${M(sum(t12) / 12)} a month over the last twelve, ${d >= 0 ? "up" : "down"} ${usd(d).replace(/^[+\u2212]/, "")} year to date.`,
    tone: d >= 0 ? "g" : "b",
    source: `${SRC} Trailing 12 months Oct 2025 - Sep 2026.`,
  });
  H.yoy(pptx, pptx, {
    eyebrow: SEG, title, metric,
    cardLabel: "Group \u00b7 January\u2013September \u00b7 same-store, 33 stores",
    months: MON.map((m, i) => [m, m25[i], m26[i]]), expect: d,
    takeaway: `${usd(d)} (${(d / v25 * 100).toFixed(1)}%) year to date; ${up} of nine months ahead of 2025, best ${FULL[m26.map((v, i) => v - m25[i]).indexOf(Math.max(...m26.map((v, i) => v - m25[i])))]} at ${usd(Math.max(...m26.map((v, i) => v - m25[i])))}.`,
    source: `${SRC} Jan-Sep 2026 vs 2025.`,
  });
  const ds = S.map(r => [r[0], r[2] - r[1]]).sort((a, b) => b[1] - a[1]);
  const grew = ds.filter(r => r[1] > 0).length;
  H.ranker(pptx, pptx, {
    eyebrow: SEG, title, metric, rows: S, expect: d, groupLabel: "SAME-STORE, 33 STORES",
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
