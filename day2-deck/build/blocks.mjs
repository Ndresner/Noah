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
  // mn: Jan-Mar and Jul-Sep 2026 inputs per store (pulls/mn/build_mn.py) for the moving-now column
  const mn = Object.fromEntries((D.mn ?? []).map(r => [r[0], [r.slice(1, 1 + w), r.slice(1 + w, 1 + 2 * w)]]));
  return { D, tot, G25: tot(D.m25), G26: tot(D.m26),
    st: D.S.map(s => ({ n: s[0], a25: s.slice(1, 1 + w), a26: s.slice(1 + w, 1 + 2 * w), lux: LUX.has(s[0]),
      q1: mn[s[0]]?.[0], q3: mn[s[0]]?.[1] })) };
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
  if (FM === F.usd2) return { unit: "usd2", scale: 1, chart: "line", labelFormat: '"$"#,##0.00', axisFormat: '"$"#,##0', fmt: FM.v, dfmt: FM.d };
  if (FM === F.usd)  return { unit: "usd", scale: 1, chart: "line", fmt: FM.v, dfmt: FM.d };
  if (FM === F.dec2) return { unit: "count", scale: 1, chart: "line", labelFormat: "0.00", axisFormat: "0.0", fmt: FM.v, dfmt: FM.d };
  if (FM === F.dec1) return { unit: "count", scale: 1, chart: "line", labelFormat: "0.0", axisFormat: "0", fmt: FM.v, dfmt: FM.d };
  return { unit: "count", scale: 1, chart: "line", labelFormat: "#,##0", axisFormat: "#,##0", fmt: FM.v, dfmt: FM.d };
}

export function block(pptx, X, M) {
  const { D, tot, G25, G26, st } = X, f = M.fn, FM0 = M.f;
  // M.whole (Noah, Oct 24): store rankers and movers show whole dollars; the group card and chart keep cents
  const FM = M.whole ? F.usd : FM0;
  const g25 = f(G25), g26 = f(G26);
  const up = D.m26.filter((r, i) => f(r) > f(D.m25[i])).length;
  const rose = g26 >= g25;
  const months = rose ? (up === 9 ? "all nine" : `${NUM[up]} of nine`) : (up === 0 ? "all nine" : `${NUM[9 - up]} of nine`);
  const measure = M.measure ?? M.title;
  // a TOTAL's year is the sum of its months (bars); anything else is a rate (a line)
  const total = M.level === false && FM === F.int;   // a count (ROs): bars, totalRanker
  const CS = chartSpec(FM0);

  // 1. group picture
  H.groupYoy(pptx, pptx, {
    eyebrow: M.seg, title: M.title, metric: M.metric, measure,
    // 2025 runs Jan-Dec on the chart (D.q4 = Oct-Dec group months; Noah, Oct 22); YTD stays Jan-Sep
    prior: [...D.m25, ...(D.q4 ?? [])].map(r => f(r) * CS.scale), current: D.m26.map(r => f(r) * CS.scale),
    priorName: "2025", currentName: "2026",
    ytdPrior: total ? undefined : g25 * CS.scale, ytdCurrent: total ? undefined : g26 * CS.scale,
    unit: total ? "count" : CS.unit, chart: "line",   // totals are lines too (Noah, Oct 22: no bar charts)
    labelFormat: total ? undefined : CS.labelFormat, axisFormat: total ? undefined : CS.axisFormat,
    fmt: total ? FM0.v : CS.fmt, dfmt: total ? FM0.d : CS.dfmt,
    chartLabel: `${M.title} by month, 2025 against 2026`,
    benchLines: M.chartBench, benchLabel: M.chartBenchLabel,   // dashed benchmark line(s) (Noah, Oct 24)
    // M.noYoyTake (Noah, Oct 27): no band on the group chart slide (CP ELR)
    takeaway: M.noYoyTake ? null : `${rose ? "Up" : "Down"} ${FM0.d(H.shownDiff(FM0.v(g25), FM0.v(g26), FM0.d, g26 - g25)).replace(/^[+−]/, "")} to ${FM0.v(g26)}, ${rose ? "ahead of" : "behind"} 2025 in ${months} months.${M.extra ? " " + M.extra(X) : ""}`,
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
      // option A (Noah, Oct 23): luxury then non-luxury, each ranked on its own; bars drawn against each segment's
      // average so both averages fall on one solid navy line
      groups: [{ label: `LUXURY · AVERAGE ${FM.v(la)}`, names: lux, ref: la }, { label: `NON-LUXURY · AVERAGE ${FM.v(na)}`, names: new Set(N.map(x => x.n)), ref: na }],
      alignRef: true, refLine: true, segBreak: true,   // Oct 24 (Noah): no legend, no takeaway band, line stops per group, rule under each group
      // the old takeaway sentence now opens the source line, so it stays with the slide
      source: `Against their own segment, ${aL} of ${L.length} luxury stores beat ${FM.v(la)} and ${aN} of ${N.length} others beat ${FM.v(na)}. ${nm(topN.n)} leads non-luxury at ${FM.v(f(topN.a26))}. ${M.src} ${M.def} Jan-Sep 2026 YTD. Green at or above own segment average, red below. Segment averages from segment totals. Luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson; Acura counted as non-luxury.`,
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

  // 3. change. M.pct (Noah, Oct 24): rank by percent change so store size and level do not set the order; the
  // change in the metric's own unit rides alongside (a column here, a smaller line on the movers slide)
  const ch = M.pct ? ((a, b) => f(b) / f(a) - 1) : (M.chg ?? ((a, b) => f(b) - f(a)));
  const CF = M.pct ? F.pctc.d : (M.cf ?? FM.d);
  const absD = (a, b) => FM.d(f(b) - f(a));
  // Oct 28 audit: a points change prints as the shown 2026 rate minus the shown 2025 rate, the same figure the movers
  // slide shows; the grew / declined counts stay on the exact change
  const shownD = (a, b) => H.shownDiff(FM.v(f(a)), FM.v(f(b)), CF, ch(a, b));
  const deltas = st.map(x => [x.n, shownD(x.a25, x.a26)]);
  const raw = st.map(x => [x.n, ch(x.a25, x.a26)]);
  const grew = raw.filter(r => r[1] > 0).sort((p, q) => q[1] - p[1]);
  const fell = raw.filter(r => r[1] < 0).sort((p, q) => p[1] - q[1]);
  const lead = st.find(x => x.n === grew[0][0]);
  H.ranker(pptx, pptx, {
    eyebrow: M.seg, title: M.title, metric: M.metric, measure, deltas, groupD: shownD(G25, G26), fmtD: CF,
    ...(M.pct ? { valLabel: "Change", stripValue: `${CF(ch(G25, G26))}  ·  ${absD(G25, G26)}`,
                  extra: [{ label: M.absLabel ?? "Change", w: 0.95, vals: Object.fromEntries(st.map(x => [x.n, absD(x.a25, x.a26)])) }] } : {}),
    groupLabel: LABEL.change,
    takeaway: `${grew.length} of 33 stores ${M.verb}, led by ${nm(grew[0][0])} at ${CF(shownD(lead.a25, lead.a26))}${M.pct ? ` (${absD(lead.a25, lead.a26)})` : ""} to ${FM.v(f(lead.a26))}. `
            + (fell.length ? `${fell.length} declined; ${nm(fell[0][0])} fell most at ${CF((x => shownD(x.a25, x.a26))(st.find(y => y.n === fell[0][0])))}.` : "None declined."),
    source: `${M.src} Change in ${M.chgDesc ?? M.title}, Jan-Sep 2026 vs 2025, best first.${M.pct ? " Ranked by percent change; the change in the measure's own unit alongside." : ""}`,
  });

  // 4. top and bottom five, a conversation starter (Noah, Oct 22), plus top five moving now within 2026 (Oct 23)
  const pm = a => a.map(v => v / 3);   // a quarter's inputs as a monthly average: totals read per month, rates unchanged
  const nw = st.filter(x => x.q1);
  if (M.movers !== false) moversSlide(pptx, {
    seg: M.seg, title: M.title, metric: M.metric, fmtD: CF,
    // M.d2 (Noah, Oct 28): a gray second line on a block that is not ranked by percent (CP RO count: the RO change)
    rows: st.map(x => ({ n: x.n, v0: FM.v(f(x.a25)), v1: FM.v(f(x.a26)), d: ch(x.a25, x.a26), d2: M.pct ? absD(x.a25, x.a26) + (M.absUnit ?? "") : M.d2 ? M.d2(x.a25, x.a26) : undefined })),
    now: nw.length ? nw.map(x => ({ n: x.n, v0: FM.v(f(pm(x.q1))), v1: FM.v(f(pm(x.q3))), d: ch(x.q1, x.q3), d2: M.pct ? absD(pm(x.q1), pm(x.q3)) + (M.absUnit ?? "") : M.d2 ? M.d2(pm(x.q1), pm(x.q3)) : undefined })) : undefined,
    nowGroup: nw.length ? CF(ch(tot(nw.map(x => x.q1)), tot(nw.map(x => x.q3)))) : undefined,
    perMonth: total || M.chgDesc !== undefined,
    src: `${M.src} ${M.def ?? ""} Top and bottom five stores by change in ${M.chgDesc ?? M.title}, Jan-Sep 2026 vs 2025.`,
  });
}

/* the top-and-bottom-five slide for any metric (also used by the gross and opcode decks). With o.now (Oct 23, Noah: every
   top/bottom slide) it is three columns: top five and bottom five year over year, then top five moving now, Jul-Sep vs
   Jan-Mar 2026, each store against itself (same basis as Who's Moving Now). */
export function moversSlide(pptx, o) {
  const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
  const dir = H.dirOf(o.metric), up = dir !== "down", K = 5;
  // Oct 28 audit: a points change is the shown end minus the shown start, and rows rank on what they show
  const fix = x => ({ ...x, d: H.shownDiff(x.v0, x.v1, o.fmtD, x.d) });
  const srt = rs => [...rs].map(fix).sort((a, b) => up ? b.d - a.d : a.d - b.d).map(x => ({ ...x, key: x.n, n: nm(x.n) }));
  const r = srt(o.rows);
  const top = r.slice(0, 3).map(x => x.n), bot = r.slice(-3).reverse().map(x => x.n);
  const notes = "Conversation starter: each top store says one thing they started doing; each bottom store says what got in the way and what they would try. The room trades ideas.";
  if (!o.now) return H.movers(pptx, pptx, {
    eyebrow: o.seg, title: o.title, metric: o.metric, rows: r, fmtD: o.fmtD, sub: o.sub, valHead: o.valHead,
    takeaway: `Ask ${list(top)} what changed. Ask ${list(bot)} what got in the way.`, tone: "n", source: o.src, notes,
  });
  const nw = srt(o.now).slice(0, K);
  return H.movers3(pptx, pptx, {
    eyebrow: o.seg, title: o.title, k: K, fmtD: o.fmtD, up, neutral: dir === "neutral",
    spark: o.spark, sparkStyle: o.sparkStyle,   // Oct 28 review: (store key, column) -> { prior, cur } monthly series for a sparkline per row
    cols: [
      { head: `TOP ${K} VS LAST YEAR`, sub: "Jan–Sep, this year vs last", short: "Jan–Sep", rows: r.slice(0, K), tone: "good" },
      { head: `BOTTOM ${K} VS LAST YEAR`, sub: "Jan–Sep, this year vs last", short: "Jan–Sep", rows: r.slice(-K).reverse(), tone: "bad", rank: j => r.length - j },
      { head: `TOP ${K} MOVING NOW`, sub: "Q3 vs Q1, this year", rows: nw, tone: "blue", yrs: ["Q1", "Q3"], short: "Q3 vs Q1" },
    ],
    takeaway: `Ask ${list(top)} what changed. Ask ${list(bot)} what got in the way. Ask ${list(nw.slice(0, 3).map(x => x.n))} what they started this year.`,
    tone: "n",
    source: `${o.src} Moving now: Jul-Sep 2026 vs Jan-Mar 2026, each store against itself${o.perMonth ? " (monthly averages)" : ""}; the group moved ${String(o.nowGroup).replace(/^[\u2212-](0\.0+\D)/, "$1")} on the same basis, so part of every store's change is seasonal. Green bar at the right of a row: the store is on both the top five vs last year and the top five moving now. Blue bar: a turnaround, on the bottom five vs last year and the top five moving now.`,
    notes: notes + " Moving now: ask the top three what they started this year.",
  });
}
