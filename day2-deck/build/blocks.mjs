/* Shared builder for one metric block: yoy card -> level (group, benchmark or
   segment) -> change ranker. All takeaways and counts derived from the data. */
import fs from "fs";
import * as H from "./house.mjs";

export const MON = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP"];
export const LUX = new Set(["LEXUS","LINCOLN","CADILLAC","VOLVO","AUDI RENO","LR RENO","AUDI HENDERSON","LR LV","LR HENDERSON"]);
const NUM = ["zero","one","two","three","four","five","six","seven","eight","nine"];
const sg = v => v >= 0 ? "+" : "\u2212";
export const F = {
  usd:  { v: v => "$" + Math.round(v).toLocaleString("en-US"), d: v => sg(v) + "$" + Math.abs(Math.round(v)).toLocaleString("en-US") },
  usd2: { v: v => "$" + v.toFixed(2), d: v => sg(v) + "$" + Math.abs(v).toFixed(2) },
  int:  { v: v => Math.round(v).toLocaleString("en-US"), d: v => sg(v) + Math.abs(Math.round(v)).toLocaleString("en-US") },
  dec1: { v: v => v.toFixed(1), d: v => sg(v) + Math.abs(v).toFixed(1) },
  dec2: { v: v => v.toFixed(2), d: v => sg(v) + Math.abs(v).toFixed(2) },
  pct:  { v: v => (v * 100).toFixed(1) + "%", d: v => sg(v) + Math.abs(v * 100).toFixed(1) + " pts" },
  pctc: { v: v => (v * 100).toFixed(1) + "%", d: v => sg(v) + Math.abs(v * 100).toFixed(1) + "%" },
};
export const rel = (d, a) => sg(d) + Math.abs(d / a * 100).toFixed(1) + "%";

export function load(file) {
  const D = JSON.parse(fs.readFileSync(file, "utf8"));
  const w = D.m25[0].length;
  const tot = rows => rows.reduce((t, r) => t.map((x, i) => x + r[i]), Array(w).fill(0));
  return { D, tot, G25: tot(D.m25), G26: tot(D.m26),
    st: D.S.map(s => ({ n: s[0], a25: s.slice(1, 1 + w), a26: s.slice(1 + w, 1 + 2 * w), lux: LUX.has(s[0]) })) };
}

export function block(pptx, X, M) {
  const { D, tot, G25, G26, st } = X, f = M.fn, FM = M.f;
  const g25 = f(G25), g26 = f(G26);
  const up = D.m26.filter((r, i) => f(r) > f(D.m25[i])).length;
  const rose = g26 >= g25;
  const months = rose ? (up === 9 ? "all nine" : `${NUM[up]} of nine`) : (up === 0 ? "all nine" : `${NUM[9 - up]} of nine`);

  // 1. group card
  H.yoy(pptx, pptx, {
    eyebrow: M.seg, title: M.title, metric: M.metric,
    cardLabel: `Group \u00b7 January\u2013September \u00b7 ${M.basis ?? "same-store, 33 stores"}`,
    months: MON.map((m, i) => [m, f(D.m25[i]), f(D.m26[i])]),
    ytd: [g25, g26], fmt: { v: FM.v, d: FM.d, p: rel },
    takeaway: `${rose ? "Up" : "Down"} ${FM.d(g26 - g25).replace(/^[+\u2212]/, "")} to ${FM.v(g26)}, ${rose ? "ahead of" : "behind"} 2025 in ${months} months.${M.extra ? " " + M.extra(X) : ""}`,
    source: `${M.src} ${M.def} YTD from YTD totals, not the mean of months. Store rows reconcile to the group on every input, both years.`,
  });

  // 2. level
  const lv = st.map(x => [x.n, f(x.a26)]);
  if (M.segment) {
    const L = st.filter(x => x.lux), N = st.filter(x => !x.lux);
    const la = f(tot(L.map(x => x.a26))), na = f(tot(N.map(x => x.a26)));
    const aL = L.filter(x => f(x.a26) >= la).length, aN = N.filter(x => f(x.a26) >= na).length;
    const topN = [...N].sort((p, q) => f(q.a26) - f(p.a26))[0];
    H.segmentRanker(pptx, pptx, {
      eyebrow: M.seg, title: M.title, metric: M.metric, fmt: FM.v, fmtD: FM.d,
      segments: [
        { label: "LUXURY", rows: L.map(x => [x.n, f(x.a26)]), avg: la, prev: f(tot(L.map(x => x.a25))) },
        { label: "NON-LUXURY", rows: N.map(x => [x.n, f(x.a26)]), avg: na, prev: f(tot(N.map(x => x.a25))) },
      ],
      takeaway: `Against their own segment, ${aL} of ${L.length} luxury stores beat ${FM.v(la)} and ${aN} of ${N.length} others beat ${FM.v(na)}. ${topN.n} leads non-luxury at ${FM.v(f(topN.a26))}.`,
      source: `${M.src} ${M.def} Jan-Sep 2026 YTD. Green at or above own segment average, red below. Segment averages from segment totals. Luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson; Acura counted as non-luxury.`,
    });
  } else if (M.level !== false) {
    const sorted = [...lv].sort((p, q) => q[1] - p[1]);
    const above = lv.filter(r => r[1] >= g26).length;
    H.levelRanker(pptx, pptx, {
      eyebrow: M.seg, title: M.title, metric: M.metric, rows: lv, group: g26, fmt: FM.v, floor: M.floor,   // red below group - Tray's rule, no exceptions (Noah, Oct 15 2026)
      takeaway: `${above} of 33 stores at or above the ${FM.v(g26)} group, led by ${sorted[0][0]} at ${FM.v(sorted[0][1])}. ${sorted[sorted.length - 1][0]} is lowest at ${FM.v(sorted[sorted.length - 1][1])}.`,
      source: `${M.src} ${M.def} Jan-Sep 2026 YTD. Green at or above the group, red below.${M.floor ? ` Bars start at ${FM.v(M.floor)} so differences are visible.` : ""}`,
    });
  }

  // 3. change
  const ch = M.chg ?? ((a, b) => f(b) - f(a));
  const CF = M.cf ?? FM.d;
  const deltas = st.map(x => [x.n, ch(x.a25, x.a26)]);
  const grew = deltas.filter(r => r[1] > 0).sort((p, q) => q[1] - p[1]);
  const fell = deltas.filter(r => r[1] < 0).sort((p, q) => p[1] - q[1]);
  const lead = st.find(x => x.n === grew[0][0]);
  H.ranker(pptx, pptx, {
    eyebrow: M.seg, title: M.title, metric: M.metric, deltas, groupD: ch(G25, G26), fmtD: CF,
    groupLabel: "SAME-STORE, 33 STORES",
    takeaway: `${grew.length} of 33 stores ${M.verb}, led by ${grew[0][0]} at ${CF(grew[0][1])} to ${FM.v(f(lead.a26))}. `
            + (fell.length ? `${fell.length} declined; ${fell[0][0]} fell most at ${CF(fell[0][1])}.` : "None declined."),
    source: `${M.src} Change in ${M.chgDesc ?? M.title}, Jan-Sep 2026 vs 2025, best first.`,
  });
}
