/* Most Improved v2 (2026-10-19, for review): "better than last year" (Jan-Sep 2026 vs Jan-Sep 2025) plus
   "who's moving now" (Q3 2026 vs Q1 2026, within this year). Qlik - Closed Repair Orders, customer pay, by close date,
   same-store 33 (INEOS excluded), GMC Prescott body shop labor types excluded. Data: imp/imp.json (compute.py) and
   imp/mom.json (momentum.py). Every name, count, month and number on the slides is derived from those files. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F } = H;
const Y = JSON.parse(fs.readFileSync("./imp/imp.json", "utf8")).ytd;
const MO = JSON.parse(fs.readFileSync("./imp/mom.json", "utf8"));
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };

const BASIS = "Qlik - Closed Repair Orders. Customer pay, by close date. Same-store, 33 stores (INEOS excluded, opened 2026). GMC Prescott body shop labor types excluded. CP gross = CP labor gross + CP parts gross.";
const SRC_Y = BASIS + " Change is Jan-Sep 2026 vs Jan-Sep 2025. Ranked by percent change (points for parts GP %), so store size does not drive the ranking.";
const SRC_M = BASIS + " Change is Jul-Sep 2026 (Q3) vs Jan-Mar 2026 (Q1), each store against itself. Gross and ROs are per-month averages of each quarter. The group's own Q1-to-Q3 change is the seasonal baseline: it is the same for every store, so ranking by a store's change is the same as ranking by how far it beat the group. 'Stepped up in' = the month that best splits the store's year into a lower stretch before and a higher stretch after (at least three months before, two after). It is a conversation starter, not proof of cause.";
const ORDER = ["gross", "ros", "hpro", "elr", "pro", "pgp"];
const GRID = ORDER.filter(k => k !== "pro");   // grids drop parts $ per RO (Noah, Oct 23: that section left the deck)
const MON = ["J", "F", "M", "A", "M", "J", "J", "A", "S"];
const SHORT = { gross: "CP gross", ros: "CP ROs", hpro: "hours/RO", elr: "ELR", pro: "parts $/RO", pgp: "parts GP%" };
const sg = v => v >= 0 ? "+" : "−";
/* Oct 28 audit (Noah): one rule with the rest of the deck - changes round (were truncated, so the Close read 0.1 under
   the same store on the earlier slides), and a points change is the shown end minus the shown start */
const chg = (k, v, v0, v1) => {
  let d = v * 100;
  if (PTS.has(k) && v0 != null && v1 != null) { const a = H.pctOf(val(k, v0)), b = H.pctOf(val(k, v1)); if (a != null && b != null) d = b - a; }
  const t = Math.abs(d).toFixed(1), z = Number(t) === 0 ? "" : sg(d);
  return PTS.has(k) ? `${z}${t} pts` : `${z}${t}%`;
};
const PTS = new Set(["pgp", "align", "bal", "bf", "fuel", "abs", "opc"]);
const val = (k, v) => ({
  gross: v >= 1e6 ? "$" + (v / 1e6).toFixed(2) + "M" : "$" + Math.round(v / 1e3) + "K",
  ros: Math.round(v).toLocaleString("en-US"), hpro: v.toFixed(2), elr: "$" + v.toFixed(2),
  pro: "$" + Math.round(v), pgp: (v * 100).toFixed(1) + "%",
  // opcode rates truncated, as on the opcode slides (Oct 28 audit)
  align: tr(v), bal: tr(v), bf: tr(v), fuel: tr(v), opc: tr(v),
  lpr: "$" + Math.round(v), np: (v < 0 ? "\u2212" : "") + "$" + (Math.abs(v) >= 1e6 ? (Math.abs(v) / 1e6).toFixed(2) + "M" : Math.round(Math.abs(v) / 1e3) + "K"), abs: (v * 100).toFixed(1) + "%" }[k]);
function tr(v) { return (Math.floor(v * 1000 + 1e-9) / 10).toFixed(1) + "%"; }
const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
const NUMW = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"];
const cap = w => w[0].toUpperCase() + w.slice(1);
const EB = "Fixed Ops · Most Improved";

/* both views as the same shape: { label, group: [g0, g1, gchg], rows: [{ n, v0, v1, chg }] } */
const VIEW = {
  yoy: Object.fromEntries(ORDER.map(k => [k, { label: Y[k].label, group: Y[k].group,
    rows: Y[k].rows.map(r => ({ n: r[0], v0: r[1], v1: r[2], chg: r[3] })) }])),
  mom: Object.fromEntries(ORDER.map(k => [k, MO[k]])),
};

/* Oct 28 (Noah): Better Than Last Year on two slides. Extra Jan-Sep 2026 vs 2025 measures, same store basis as the deck:
   CP labor $ per RO (m5.json), fixed ops net profit (ntgs + ntgp department profit; % = $ change / |2025 profit|, as the
   net profit change slide), fixed absorption (abs.json, GMC Prescott body shop in, as on the P&L; points) and the four
   opcodes combined (all.json: alignment + balance + brake fluid + fuel services / op-code-report CP ROs; points). */
const J = f => JSON.parse(fs.readFileSync(f, "utf8"));
const M5 = J("./m5.json").S, NS = J("./ntgs.json").S, NP = J("./ntgp.json").S, AB = J("./abs.json").S, AL = J("./all.json");
const OPK4 = ["align", "bal", "bf", "fuel"];
const sumS = (S, i) => S.reduce((t, r) => t + r[i], 0);
const mk = (label, rows, g0, g1, pct) => {
  const ch = (a, b) => pct ? (b - a) / Math.abs(a) : b - a;
  return { label, group: [g0, g1, ch(g0, g1)], rows: rows.map(([n, a, b]) => ({ n, v0: a, v1: b, chg: ch(a, b) })).sort((p, q) => q.chg - p.chg) };
};
VIEW.yoy.lpr = mk("CP Labor $ per RO", M5.map(r => [r[0], r[2] / r[1], r[8] / r[7]]), sumS(M5, 2) / sumS(M5, 1), sumS(M5, 8) / sumS(M5, 7), true);
VIEW.yoy.np = mk("Fixed Ops Net Profit", NS.map(r => { const p = NP.find(x => x[0] === r[0]); return [r[0], r[1] + p[1], r[3] + p[3]]; }),
  sumS(NS, 1) + sumS(NP, 1), sumS(NS, 3) + sumS(NP, 3), true);
VIEW.yoy.abs = mk("Fixed Absorption", AB.map(r => [r[0], r[1] / r[2], r[3] / r[4]]), sumS(AB, 1) / sumS(AB, 2), sumS(AB, 3) / sumS(AB, 4), false);
const opc = AL.align.S.map(r => { const n = r[0], rr = OPK4.map(k => AL[k].S.find(x => x[0] === n));
  return [n, rr.reduce((t, x) => t + x[1], 0) / r[2], rr.reduce((t, x) => t + x[3], 0) / r[4]]; });
VIEW.yoy.opc = mk("Opcodes, Four Combined", opc, OPK4.reduce((t, k) => t + sumS(AL[k].S, 1), 0) / sumS(AL.align.S, 2), OPK4.reduce((t, k) => t + sumS(AL[k].S, 3), 0) / sumS(AL.align.S, 4), false);

function tops(view, n = 3) {
  const c = {};
  GRID.forEach(k => VIEW[view][k].rows.slice(0, n).forEach(r => (c[r.n] ??= []).push(k)));
  return Object.entries(c).sort((a, b) => b[1].length - a[1].length);
}

/* ---- grids: five measures (parts $ per RO dropped Oct 23), top 3 and bottom 3 each (was 5 and 5; cut Oct 21, Noah: too crowded to read) ---- */
function grid(view, { title, sub, groupLabel, src, keys = GRID }) {
  const GRID = keys;   // Oct 28 (Noah): Better Than Last Year runs on two slides, each with its own measures
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  s.addText(sub, { x: 0.78, y: 1.4, w: 11.77, h: 0.28, fontFace: F, fontSize: 14, color: C.ink600, isTextBox: true, margin: 0 });
  // leaderboard: measures across; top 5 then bottom 5 down (bottom 5 ADDED 2026-10-20, Noah). Each cell = store, change.
  // Before -> after values moved to the speaker notes to make room for ten rows.
  const K = 3, X0 = 1.0, CW = 11.68 / GRID.length, HY = 1.85, rh = 0.58, TOPS = 2.55, BOTS = 4.46, N = VIEW[view][ORDER[0]].rows.length;
  [[TOPS, C.good, C.goodInk, j => j + 1], [BOTS, C.bad, C.badInk, j => N - K + 1 + j]].forEach(([y0, bar, ink, rank]) => {
    s.addShape(pptx.ShapeType.rect, { x: 0.68, y: y0, w: 0.05, h: rh * K, fill: { color: bar } });
    for (let j = 0; j < K; j++) {
      const y = y0 + j * rh;
      if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.75, y, w: 11.9, h: rh, fill: { color: C.zebra } });
      s.addText(String(rank(j)), { x: 0.74, y, w: 0.3, h: rh, fontFace: F, fontSize: 16, bold: true, color: ink, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    }
  });
  const notes = [];
  // thin gray rule between measure columns, so each column reads on its own (Noah, Oct 28)
  for (let i = 1; i < GRID.length; i++)
    s.addShape(pptx.ShapeType.rect, { x: X0 + i * CW - 0.006, y: HY, w: 0.012, h: BOTS + rh * K - HY, fill: { color: C.rule } });
  GRID.forEach((k, i) => {
    const M = VIEW[view][k], x = X0 + i * CW;
    s.addShape(pptx.ShapeType.rect, { x: x + 0.06, y: HY, w: CW - 0.12, h: 0.05, fill: { color: C.good } });
    s.addText(M.label.replace(" per Month", "").toUpperCase(), { x: x + 0.08, y: HY + 0.09, w: CW - 0.16, h: 0.26, fontFace: F, fontSize: 13, bold: true,
      color: C.ink600, charSpacing: 0.3, isTextBox: true, margin: 0, fit: "shrink" });
    s.addText(`${groupLabel} ${chg(k, M.group[2], M.group[0], M.group[1])}`, { x: x + 0.08, y: HY + 0.33, w: CW - 0.16, h: 0.22, fontFace: F, fontSize: 13,
      color: C.ink600, isTextBox: true, margin: 0 });
    const cell = (r, y, top) => {
      // Title Case, not caps (Noah, Oct 22): PowerPoint does not shrink-to-fit until a box is edited, so caps names
      // like HYUNDAI ST GEORGE wrapped onto the change line below them
      s.addText(H.storeName(r.n), { x: x + 0.08, y: y + 0.05, w: CW - 0.1, h: 0.24, fontFace: F, fontSize: 13, bold: true, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
      s.addText(chg(k, r.chg, r.v0, r.v1), { x: x + 0.08, y: y + 0.29, w: CW - 0.12, h: 0.25, fontFace: F, fontSize: 16, bold: true, valign: "middle",
        color: /^[0.]+/.test(chg(k, r.chg, r.v0, r.v1)) ? C.ink600 : r.chg > 0 ? C.goodInk : C.badInk, isTextBox: true, margin: 0 });
    };
    const top = M.rows.slice(0, K), bot = M.rows.slice(-K);
    top.forEach((r, j) => cell(r, TOPS + j * rh, true));
    bot.forEach((r, j) => cell(r, BOTS + j * rh, false));
    notes.push(`${M.label}: top ${top.map(r => `${r.n} ${val(k, r.v0)} \u2192 ${val(k, r.v1)}`).join("; ")}. Bottom ${bot.map(r => `${r.n} ${val(k, r.v0)} \u2192 ${val(k, r.v1)}`).join("; ")}.`);
  });
  s.addNotes("\n\nBefore \u2192 after, by measure:\n" + notes.join("\n"));
  const rep = tops(view, K), hi = rep[0][1].length;
  const lead = rep.filter(e => e[1].length === hi).map(e => e[0]), next = rep.filter(e => e[1].length === hi - 1).map(e => e[0]);
  const bc = {}; GRID.forEach(k => VIEW[view][k].rows.slice(-K).forEach(r => (bc[r.n] ??= []).push(k)));
  const brep = Object.entries(bc).sort((p, q) => q[1].length - p[1].length), bh = brep[0][1].length;
  const blead = brep.filter(e => e[1].length === bh).map(e => e[0]);
  H.takeaway(pptx, s, `Most top-three finishes: ${list(lead)} (${hi} of ${GRID.length}). Most bottom-three: ${list(blead)} (${bh} of ${GRID.length}).`, "n");
  H.foot(pptx, s, src);
  return s;
}

/* ---- momentum, one measure: top 10 with the store's own 9-month line ---- */
function moving(k) {
  const M = MO[k], s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title: `Who's Moving Now: ${M.label}` });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  s.addText(`Q1 → Q3 2026. Group ${val(k, M.group[0])} → ${val(k, M.group[1])}, ${chg(k, M.group[2], M.group[0], M.group[1])}: that is the season. ${M.beat} of ${M.rows.length} stores beat it.`,
    { x: 0.78, y: 1.42, w: 11.77, h: 0.34, fontFace: F, fontSize: 14, color: C.ink600, isTextBox: true, margin: 0 });
  const X = { rk: 0.82, name: 1.2, vals: 4.0, line: 5.85, step: 8.95, chg: 11.15 }, TOP = 2.18, rh = 0.418;
  [["Q1 → Q3", X.vals, 1.7, "left"], ["Jan – Sep 2026, by month", X.line, 2.9, "left"], ["Stepped up in", X.step, 1.9, "left"],
   ["Q1 → Q3", X.chg, 1.5, "right"]].forEach(([t, x, w, al]) =>
    s.addText(t, { x, y: 1.84, w, h: 0.3, fontFace: F, fontSize: 14, bold: true, color: C.ink600, align: al, isTextBox: true, margin: 0 }));
  M.rows.slice(0, 10).forEach((r, j) => {
    const y = TOP + j * rh;
    if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.68, y, w: 11.97, h: rh, fill: { color: C.zebra } });
    s.addText(String(j + 1), { x: X.rk, y, w: 0.34, h: rh, fontFace: F, fontSize: 15, bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(r.n, { x: X.name, y, w: 2.7, h: rh, fontFace: F, fontSize: 14, bold: true, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(`${val(k, r.v0)} → ${val(k, r.v1)}`, { x: X.vals, y, w: 1.8, h: rh, fontFace: F, fontSize: 14, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    const lo = Math.min(...r.m), hi = Math.max(...r.m), pad = (hi - lo) * 0.15 || hi * 0.05;
    s.addChart(pptx.ChartType.line, [{ name: r.n, labels: MON, values: r.m }], {
      x: X.line - 0.08, y: y + 0.02, w: 2.95, h: rh - 0.04, showLegend: false, showTitle: false,
      catAxisHidden: true, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      valAxisMinVal: lo - pad, valAxisMaxVal: hi + pad, chartColors: [C.good], lineSize: 2, lineDataSymbol: "circle",
      lineDataSymbolSize: 3, plotArea: { fill: { color: j % 2 === 0 ? C.zebra : C.panel } } });
    s.addText(r.step, { x: X.step, y, w: 1.9, h: rh, fontFace: F, fontSize: 14, bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(chg(k, r.chg, r.v0, r.v1), { x: X.chg, y, w: 1.5, h: rh, fontFace: F, fontSize: 16, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  const t = M.rows[0];
  H.takeaway(pptx, s, `${t.n} went ${val(k, t.v0)} to ${val(k, t.v1)}, stepping up in ${t.step}. Ask what changed that month.`, "g");
  H.foot(pptx, s, SRC_M);
  s.addNotes(`\n\nReview note: the month line is the store's own monthly ${M.label.toLowerCase()}, so a step up shows as the line moving to a higher level and staying there. Volume measures (gross, ROs) still carry some store-specific seasonality (snowbird markets) the group baseline cannot fully remove; the rate measures are the cleanest read.`);
}

/* ---- two views side by side: who shows up most ---- */
function twoViews() {
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title: "Two Ways to Read Improvement" });
  const A = tops("yoy", 5).slice(0, 6), B = tops("mom", 5).slice(0, 6);
  // green bar down the right edge of a store on both lists, as on the movers slides (Noah, Oct 28)
  const onBoth = new Set(A.map(e => e[0]).filter(n => B.some(e => e[0] === n)));
  [[A, 0.5, "Better than last year", "Jan–Sep 2026 vs Jan–Sep 2025"], [B, 6.75, "Moving now", "Q3 vs Q1, this year"]].forEach(([rows, x, head, sub]) => {
    const w = 6.08;
    H.panel(pptx, s, { x, y: 1.32, w, h: 5.08, label: `${head} · ${sub}` });
    s.addText(`Top-5 finishes, of ${NUMW[GRID.length]} measures`, { x: x + 0.25, y: 1.96, w: w - 0.5, h: 0.3, fontFace: F, fontSize: 14, color: C.ink600, isTextBox: true, margin: 0 });
    const rh = 0.64, mx = rows[0][1].length;
    rows.forEach(([n, ks], j) => {
      const y = 2.38 + j * rh;
      if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: x + 0.18, y, w: w - 0.36, h: rh, fill: { color: C.zebra } });
      s.addText(n, { x: x + 0.3, y: y + 0.05, w: 2.6, h: 0.3, fontFace: F, fontSize: 14, bold: true, color: C.ink, isTextBox: true, margin: 0 });
      s.addText(ks.map(k => SHORT[k]).join(", "), { x: x + 0.3, y: y + 0.33, w: 3.85, h: 0.26, fontFace: F, fontSize: 13, color: C.ink600, isTextBox: true, margin: 0 });
      s.addShape(pptx.ShapeType.rect, { x: x + 4.25, y: y + rh * 0.3, w: ks.length / mx * 1.05, h: rh * 0.4, fill: { color: C.good } });
      s.addText(String(ks.length), { x: x + w - 0.75, y, w: 0.45, h: rh, fontFace: F, fontSize: 18, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      if (onBoth.has(n)) s.addShape(pptx.ShapeType.rect, { x: x + w - 0.25, y: y + 0.04, w: 0.07, h: rh - 0.08, fill: { color: C.good } });
    });
  });
  const lead = rows => { const t = rows.filter(r => r[1].length === rows[0][1].length); return t.length === 1 ? `${t[0][0]} leads` : `${cap(NUMW[t.length])} stores tie`; };
  // Oct 28 (Noah): stores tied with the last row shown but cut by the six-row list go on the band (store names allowed here)
  const tiedOut = (all, shown) => all.filter(e => !shown.some(x => x[0] === e[0]) && e[1].length === shown.at(-1)[1].length).map(e => H.storeName(e[0]));
  const tl = tiedOut(tops("yoy", 5), A), tr = tiedOut(tops("mom", 5), B);
  const tieNote = [tl.length ? `Also tied at ${A.at(-1)[1].length} last year: ${list(tl)}.` : "", tr.length ? `Also tied at ${B.at(-1)[1].length} moving now: ${list(tr)}.` : ""].filter(Boolean).join(" ");
  const band = `Last year: ${lead(A)}. Moving now: ${lead(B)}. ${onBoth.size === 0 ? "No store is on both lists." : onBoth.size === 1 ? "One store is on both lists (green bar)." : cap(NUMW[onBoth.size]) + " stores are on both lists (green bar)."}`;
  if (tieNote) { if (H.namesStore(band)) s.addNotes(`Takeaway with the store call-outs (off the slide): ${band}`); H.takeaway(pptx, s, (H.stripStores(band) + " " + tieNote).trim(), "g", { keep: true }); }
  else H.takeaway(pptx, s, band, "g");
  H.foot(pptx, s, BASIS + ` Left: Jan-Sep 2026 vs Jan-Sep 2025. Right: Q3 2026 vs Q1 2026. Top-5 finishes counted across ${NUMW[GRID.length]} measures (parts $ per RO dropped, Oct 23); top six stores shown in each view.`);
}

/* ---- the ask: one question for each measure's top mover ---- */
function whatChanged() {
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title: "What Changed?" });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  // Oct 28 (Noah): all four opcodes listed, so nine rows at 14pt (was six at 15pt)
  const X = { n: 0.82, m: 3.85, v: 6.2, st: 8.45, q: 10.2 }, rh = 0.48, TOP = 1.92;
  [["Store", X.n], ["Moved on", X.m], ["Q1 → Q3", X.v], ["Stepped up in", X.st], ["Change", X.q]].forEach(([t, x]) =>
    s.addText(t, { x, y: 1.5, w: 2.2, h: 0.32, fontFace: F, fontSize: 14, bold: true, color: C.ink600, isTextBox: true, margin: 0 }));
  const rows = GRID.map(k => [k, MO[k].rows[0], MO[k].label]);   // parts $ per RO out (Noah, Oct 23)
  // opcodes: the store that tops the most of the four opcode lists, shown on its biggest move (added 2026-10-19)
  const OPK = ["align", "bal", "bf", "fuel"], tally = {};
  OPK.forEach(k => { const n = MO[k].rows[0].n; tally[n] = (tally[n] ?? 0) + 1; });
  const opStore = Object.entries(tally).sort((a, b) => b[1] - a[1])[0][0];
  // Oct 28 audit (Noah): the row keeps the slide's rule - an opcode the store is the top mover on (was its biggest move,
  // alignment, where Audi Henderson is the top mover)
  // Oct 28 (Noah): every opcode gets its own row, each with its top Q1-to-Q3 mover
  const OPL = { align: "Alignment", bal: "Balance", bf: "Brake fluid", fuel: "Fuel service" };
  OPK.forEach(k => rows.push([k, MO[k].rows[0], `Opcodes \u00b7 ${OPL[k].toLowerCase()}`]));
  rows.forEach(([k, r, lab], j) => {
    const y = TOP + j * rh;
    if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.68, y, w: 11.97, h: rh, fill: { color: C.zebra } });
    [[r.n, X.n, 2.9, true, C.ink], [lab, X.m, 2.3, false, C.ink], [`${val(k, r.v0)} → ${val(k, r.v1)}`, X.v, 2.2, false, C.ink],
     [r.step, X.st, 1.7, true, C.navy], [chg(k, r.chg, r.v0, r.v1), X.q, 1.4, true, C.goodInk]].forEach(([t, x, w, b, col]) =>
      s.addText(t, { x, y, w, h: rh, fontFace: F, fontSize: 14, bold: b, color: col, valign: "middle", isTextBox: true, margin: 0 }));
  });
  const stores = [...new Set(rows.map(r => r[1].n))];
  H.takeaway(pptx, s, `${cap(NUMW[stores.length])} stores, one question each: what did you start doing that month? One sentence, before you leave.`, "n");
  H.foot(pptx, s, SRC_M + ` Each row is the top Q1-to-Q3 mover on that measure, the four opcodes included. Opcode closing % = ROs with the service \u00f7 customer-pay ROs on the op-code report basis.`);
  s.addNotes(`\n\n${opStore}'s August opcode increase was confirmed with the store as a real change (Noah, Oct 19 2026).`);
}

if (process.env.BTLY === "1") {
  // review file (Noah, Oct 28): Better Than Last Year as two slides
  const SUBY = "Top and bottom three stores on each measure, Jan–Sep 2026 vs Jan–Sep 2025";
  grid("yoy", { title: "Better Than Last Year: Customer Pay", sub: SUBY, groupLabel: "Group", src: SRC_Y + " CP labor $ per RO = CP labor sale / CP ROs.",
    keys: ["gross", "ros", "hpro", "elr", "lpr"] });
  grid("yoy", { title: "Better Than Last Year: Profit and Process", sub: SUBY, groupLabel: "Group",
    src: "Fixed ops net profit: Qlik - Profit & Loss, service + parts department profit (% change = $ change / |2025 profit|). Fixed absorption: P&L, fixed ops gross / total store fixed expenses, GMC Prescott body shop in, as on the P&L. CP parts GP %: Qlik - Closed Repair Orders. Opcodes: alignment + balance + brake fluid + fuel services / customer-pay ROs on the op-code report basis. Same-store, 33 stores. Points for absorption, parts GP % and opcodes.",
    keys: ["np", "abs", "pgp", "opc"] });
  await pptx.writeFile({ fileName: "./Better Than Last Year - Review.pptx" }); console.log("btly"); process.exit(0);
}
grid("yoy", { title: "Better Than Last Year", sub: "Top and bottom three stores on each measure, Jan–Sep 2026 vs Jan–Sep 2025", groupLabel: "Group", src: SRC_Y });
grid("mom", { title: "Who's Moving Now", sub: "Top and bottom three stores on each measure, Q3 2026 vs Q1 2026 (gross and ROs per month). Group change = the season.", groupLabel: "Group", src: SRC_M });
ORDER.forEach(moving);
twoViews();
whatChanged();
await pptx.writeFile({ fileName: "./Most Improved v2 - Review.pptx" });
console.log("ok");
