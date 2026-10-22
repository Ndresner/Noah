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

/* Day 1 sequence (Tommy, Oct 18 2026; Day 2 rebuilt to match, Noah Oct 21 2026):
   1. group picture - YTD card over the month-by-month chart, 2025 against 2026
   2. store level   - totalRanker for a count or dollar total, levelRanker for a
                      rate (segmentRanker where luxury runs on its own scale)
   3. store change  - ranker
   Labels name the period, never a store count. Store names Title Case in prose. */
export const LABEL = {
  level:  "SEPTEMBER 2026 YTD",
  change: "2026 VS 2025 YTD · **YEAR-OVER-YEAR CHANGE**",
};
const nm = H.storeName;
// how a block's format reads on the group chart (groupYoy units)
function chartSpec(FM) {
  if (FM === F.pct)  return { unit: "pct", scale: 100, chart: "line", fmt: v => (v).toFixed(1) + "%", dfmt: v => (v >= 0 ? "+" : "−") + Math.abs(v).toFixed(1) + " pts" };
  if (FM === F.usd2) return { unit: "usd", scale: 1, chart: "line", labelFormat: '"$"#,##0.00', axisFormat: '"$"#,##0', fmt: FM.v, dfmt: FM.d };
  if (FM === F.usd)  return { unit: "usd", scale: 1, chart: "line", fmt: FM.v, dfmt: FM.d };
  if (FM === F.dec2) return { unit: "count", scale: 1, chart: "line", labelFormat: "0.00", axisFormat: "0.0", fmt: FM.v, dfmt: FM.d };
  if (FM === F.dec1) return { unit: "count", scale: 1, chart: "line", labelFormat: "0.0", axisFormat: "0", fmt: FM.v, dfmt: FM.d };
  return { unit: "count", scale: 1, chart: "line", labelFormat: "#,##0", axisFormat: "#,##0", fmt: FM.v, dfmt: FM.d };
}

export function block(pptx, X, M) {
  const { D, tot, G25, G26, st } = X, f = M.fn, FM = M.f;
  const g25 = f(G25), g26 = f(G26);
  const up = D.m26.filter((r, i) => f(r) > f(D.m25[i])).length;
  const rose = g26 >= g25;
  const months = rose ? (up === 9 ? "all nine" : `${NUM[up]} of nine`) : (up === 0 ? "all nine" : `${NUM[9 - up]} of nine`);
  const measure = M.measure ?? M.title;
  // a TOTAL's year is the sum of its months (bars); anything else is a rate (a line)
  const total = M.level === false && FM === F.int;   // a count (ROs): bars, totalRanker
  const CS = chartSpec(FM);

  // 1. group picture
  H.groupYoy(pptx, pptx, {
    eyebrow: M.seg, title: M.title, metric: M.metric, measure,
    // 2025 runs Jan-Dec on the chart (D.q4 = Oct-Dec group months; Noah, Oct 22); YTD stays Jan-Sep
    prior: [...D.m25, ...(D.q4 ?? [])].map(r => f(r) * CS.scale), current: D.m26.map(r => f(r) * CS.scale),
    priorName: "2025", currentName: "2026",
    ytdPrior: total ? undefined : g25 * CS.scale, ytdCurrent: total ? undefined : g26 * CS.scale,
    unit: total ? "count" : CS.unit, chart: "line",   // totals are lines too (Noah, Oct 22: no bar charts)
    labelFormat: total ? undefined : CS.labelFormat, axisFormat: total ? undefined : CS.axisFormat,
    fmt: total ? FM.v : CS.fmt, dfmt: total ? FM.d : CS.dfmt,
    chartLabel: `${M.title} by month, 2025 against 2026`,
    takeaway: `${rose ? "Up" : "Down"} ${FM.d(g26 - g25).replace(/^[+−]/, "")} to ${FM.v(g26)}, ${rose ? "ahead of" : "behind"} 2025 in ${months} months.${M.extra ? " " + M.extra(X) : ""}`,
    source: `${M.src} ${M.def} YTD from YTD totals, not the mean of months. Store rows reconcile to the group on every input, both years.${D.q4 ? " 2025 shown January-December; October-December pulled Oct 22 on the same basis (chart only)." : ""}`,
  });

  // 2. level
  const lv = st.map(x => [x.n, f(x.a26)]);
  if (M.segment) {
    const L = st.filter(x => x.lux), N = st.filter(x => !x.lux);
    const la = f(tot(L.map(x => x.a26))), na = f(tot(N.map(x => x.a26)));
    const aL = L.filter(x => f(x.a26) >= la).length, aN = N.filter(x => f(x.a26) >= na).length;
    const topN = [...N].sort((p, q) => f(q.a26) - f(p.a26))[0];
    // one ranking of all 33 (Noah, Oct 21: the side-by-side luxury / non-luxury lists read badly);
    // each store is still scored against its own segment average, shown as a tick on its row
    const lux = new Set(L.map(x => x.n));
    H.levelRanker(pptx, pptx, {
      eyebrow: M.seg, title: M.title, metric: M.metric, measure, rows: lv, group: g26, fmt: FM.v, groupLabel: LABEL.level,
      refOf: n => lux.has(n) ? la : na, mark: lux,
      legend: `Shaded = luxury. Navy tick = the store's segment average: luxury ${FM.v(la)}, non-luxury ${FM.v(na)}. Green at or above it.`,
      takeaway: `Against their own segment, ${aL} of ${L.length} luxury stores beat ${FM.v(la)} and ${aN} of ${N.length} others beat ${FM.v(na)}. ${nm(topN.n)} leads non-luxury at ${FM.v(f(topN.a26))}.`,
      source: `${M.src} ${M.def} Jan-Sep 2026 YTD. Green at or above own segment average, red below. Segment averages from segment totals. Luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson; Acura counted as non-luxury.`,
    });
  } else if (total) {
    const sorted = [...lv].sort((p, q) => q[1] - p[1]);
    H.totalRanker(pptx, pptx, {
      eyebrow: M.seg, title: M.title, metric: M.metric, measure, rows: lv, expect: g26, fmt: FM.v,
      groupLabel: LABEL.level,
      takeaway: `${nm(sorted[0][0])} leads at ${FM.v(sorted[0][1])}; ${nm(sorted.at(-1)[0])} is smallest at ${FM.v(sorted.at(-1)[1])}.`,
      tone: "n",
      source: `${M.src} ${M.def} Jan-Sep 2026 YTD.`,
    });
  } else if (M.level !== false) {
    const sorted = [...lv].sort((p, q) => q[1] - p[1]);
    const above = lv.filter(r => r[1] >= g26).length;
    H.levelRanker(pptx, pptx, {
      eyebrow: M.seg, title: M.title, metric: M.metric, measure, rows: lv, group: g26, fmt: FM.v, floor: M.floor,   // red below group - Tommy's rule, no exceptions (Noah, Oct 15 2026)
      groupLabel: LABEL.level,
      takeaway: `${above} of 33 stores at or above the ${FM.v(g26)} group, led by ${nm(sorted[0][0])} at ${FM.v(sorted[0][1])}. ${nm(sorted[sorted.length - 1][0])} is lowest at ${FM.v(sorted[sorted.length - 1][1])}.`,
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
    eyebrow: M.seg, title: M.title, metric: M.metric, measure, deltas, groupD: ch(G25, G26), fmtD: CF,
    groupLabel: LABEL.change,
    takeaway: `${grew.length} of 33 stores ${M.verb}, led by ${nm(grew[0][0])} at ${CF(grew[0][1])} to ${FM.v(f(lead.a26))}. `
            + (fell.length ? `${fell.length} declined; ${nm(fell[0][0])} fell most at ${CF(fell[0][1])}.` : "None declined."),
    source: `${M.src} Change in ${M.chgDesc ?? M.title}, Jan-Sep 2026 vs 2025, best first.`,
  });

  // 4. top and bottom five, a conversation starter (Noah, Oct 22)
  if (M.movers !== false) moversSlide(pptx, {
    seg: M.seg, title: M.title, metric: M.metric, fmtD: CF,
    rows: st.map(x => ({ n: x.n, v0: FM.v(f(x.a25)), v1: FM.v(f(x.a26)), d: ch(x.a25, x.a26) })),
    sub: `Jan–Sep 2026 against Jan–Sep 2025. Group ${FM.v(g25)} → ${FM.v(g26)} (${CF(ch(G25, G26))}). Ranked by ${M.chgDesc ? "percent change" : "change"}.`,
    src: `${M.src} ${M.def ?? ""} Top and bottom five stores by change in ${M.chgDesc ?? M.title}, Jan-Sep 2026 vs 2025.`,
  });
}

/* the top-and-bottom-five slide for any metric (also used by the gross and opcode decks) */
export function moversSlide(pptx, o) {
  const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
  const up = H.dirOf(o.metric) !== "down";
  const r = [...o.rows].sort((a, b) => up ? b.d - a.d : a.d - b.d).map(x => ({ ...x, n: nm(x.n) }));
  const top = r.slice(0, 3).map(x => x.n), bot = r.slice(-3).reverse().map(x => x.n);
  return H.movers(pptx, pptx, {
    eyebrow: o.seg, title: o.title, metric: o.metric, rows: r, fmtD: o.fmtD, sub: o.sub, valHead: o.valHead,
    takeaway: `Ask ${list(top)} what changed. Ask ${list(bot)} what got in the way.`, tone: "n",
    source: o.src,
    notes: "Conversation starter: each top store says one thing they started doing; each bottom store says what got in the way and what they would try. The room trades ideas.",
  });
}
