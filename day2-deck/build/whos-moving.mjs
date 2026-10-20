/* Who's Moving Now, top 10 per measure, for insertion into the CP Engine and Opcodes sections. Built from Most Improved v2 (2026-10-19, for review): "better than last year" (Jan-Sep 2026 vs Jan-Sep 2025) plus
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
const MON = ["J", "F", "M", "A", "M", "J", "J", "A", "S"];
const SHORT = { gross: "CP gross", ros: "CP ROs", hpro: "hours/RO", elr: "ELR", pro: "parts $/RO", pgp: "parts GP%" };
const sg = v => v >= 0 ? "+" : "−";
const tr1 = v => (Math.trunc(Math.abs(v) * 10) / 10).toFixed(1);
const chg = (k, v) => { const t = tr1(v * 100), z = Number(t) === 0 ? "" : sg(v); return PTS.has(k) ? `${z}${t} pts` : `${z}${t}%`; };
const PTS = new Set(["pgp", "align", "bal", "bf", "fuel"]);
const val = (k, v) => ({
  gross: v >= 1e6 ? "$" + (v / 1e6).toFixed(2) + "M" : "$" + Math.round(v / 1e3) + "K",
  ros: Math.round(v).toLocaleString("en-US"), hpro: v.toFixed(2), elr: "$" + v.toFixed(2),
  pro: "$" + Math.round(v), pgp: (v * 100).toFixed(1) + "%",
  align: (v * 100).toFixed(1) + "%", bal: (v * 100).toFixed(1) + "%", bf: (v * 100).toFixed(1) + "%", fuel: (v * 100).toFixed(1) + "%" }[k]);
const SEC = { gross: "The CP Engine", ros: "The CP Engine", hpro: "The CP Engine", elr: "The CP Engine", pro: "The CP Engine", pgp: "The CP Engine",
  align: "Opcodes", bal: "Opcodes", bf: "Opcodes", fuel: "Opcodes" };
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

function tops(view, n = 3) {
  const c = {};
  ORDER.forEach(k => VIEW[view][k].rows.slice(0, n).forEach(r => (c[r.n] ??= []).push(k)));
  return Object.entries(c).sort((a, b) => b[1].length - a[1].length);
}

/* ---- grids: six measures, top 5 each ---- */
function grid(view, { title, sub, groupLabel, src }) {
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  s.addText(sub, { x: 0.78, y: 1.4, w: 11.77, h: 0.28, fontFace: F, fontSize: 14, color: C.ink600, isTextBox: true, margin: 0 });
  // leaderboard: measures across, rank 1-5 down; each cell = store, change, before -> after
  const X0 = 1.0, CW = 11.68 / 6, HY = 1.74, TOP = 2.42, rh = 0.786;
  for (let j = 0; j < 5; j++) {
    const y = TOP + j * rh;
    if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.68, y, w: 11.97, h: rh, fill: { color: C.zebra } });
    s.addText(String(j + 1), { x: 0.7, y, w: 0.3, h: rh, fontFace: F, fontSize: 20, bold: true, color: C.navy, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  }
  ORDER.forEach((k, i) => {
    const M = VIEW[view][k], x = X0 + i * CW;
    s.addShape(pptx.ShapeType.rect, { x: x + 0.06, y: HY, w: CW - 0.12, h: 0.05, fill: { color: C.good } });
    s.addText(M.label.replace(" per Month", "").toUpperCase(), { x: x + 0.08, y: HY + 0.1, w: CW - 0.16, h: 0.3, fontFace: F, fontSize: 13, bold: true,
      color: C.ink600, charSpacing: 0.3, isTextBox: true, margin: 0, fit: "shrink" });
    s.addText(`${groupLabel} ${chg(k, M.group[2])}`, { x: x + 0.08, y: HY + 0.38, w: CW - 0.16, h: 0.24, fontFace: F, fontSize: 13,
      color: C.ink600, isTextBox: true, margin: 0 });
    M.rows.slice(0, 5).forEach((r, j) => {
      const y = TOP + j * rh;
      s.addText(r.n, { x: x + 0.06, y: y + 0.05, w: CW - 0.06, h: 0.24, fontFace: F, fontSize: 13, bold: true, color: C.ink, isTextBox: true, margin: 0, fit: "shrink" });
      s.addText(chg(k, r.chg), { x: x + 0.08, y: y + 0.28, w: CW - 0.12, h: 0.26, fontFace: F, fontSize: 16, bold: true, color: C.goodInk, isTextBox: true, margin: 0 });
      s.addText(`${val(k, r.v0)} \u2192 ${val(k, r.v1)}`, { x: x + 0.08, y: y + 0.53, w: CW - 0.12, h: 0.22, fontFace: F, fontSize: 13, color: C.ink600, isTextBox: true, margin: 0 });
    });
  });
  const rep = tops(view, 5), hi = rep[0][1].length;
  const lead = rep.filter(e => e[1].length === hi).map(e => e[0]), next = rep.filter(e => e[1].length === hi - 1).map(e => e[0]);
  H.takeaway(pptx, s, `${list(lead)} ${lead.length === 1 ? "makes" : "each make"} ${NUMW[hi]} of six lists` +
    (next.length ? `; ${next.length > 2 ? NUMW[next.length] + " stores" : list(next)} make ${NUMW[hi - 1]}.` : "."), "g");
  H.foot(pptx, s, src);
  return s;
}

/* ---- momentum, one measure: top 10 with the store's own 9-month line ---- */
function moving(k) {
  const M = MO[k], s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: `${SEC[k]} \u00b7 Who's Moving Now`, title: `Who's Moving Now: ${M.label}` });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  s.addText(`Q1 → Q3 2026. Group ${val(k, M.group[0])} → ${val(k, M.group[1])}, ${chg(k, M.group[2])}: that is the season. ${M.beat} of ${M.rows.length} stores beat it.`,
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
    s.addText(chg(k, r.chg), { x: X.chg, y, w: 1.5, h: rh, fontFace: F, fontSize: 16, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  const t = M.rows[0];
  H.takeaway(pptx, s, `${t.n} went ${val(k, t.v0)} to ${val(k, t.v1)}, stepping up in ${t.step}. Ask what changed that month.`, "g");
  H.foot(pptx, s, SRC_M + (SEC[k] === "Opcodes" ? " Opcode closing % = ROs with the service \u00f7 customer-pay ROs on the op-code report basis, same as the Opcode Benchmarks slides." : ""));
  s.addNotes(`\n\nReview note: the month line is the store's own monthly ${M.label.toLowerCase()}, so a step up shows as the line moving to a higher level and staying there. Volume measures (gross, ROs) still carry some store-specific seasonality (snowbird markets) the group baseline cannot fully remove; the rate measures are the cleanest read.`);
}

/* ---- two views side by side: who shows up most ---- */
function twoViews() {
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title: "Two Ways to Read Improvement" });
  const A = tops("yoy", 5).slice(0, 6), B = tops("mom", 5).slice(0, 6);
  [[A, 0.5, "Better than last year", "Jan–Sep 2026 vs Jan–Sep 2025"], [B, 6.75, "Moving now", "Q3 vs Q1, this year"]].forEach(([rows, x, head, sub]) => {
    const w = 6.08;
    H.panel(pptx, s, { x, y: 1.32, w, h: 5.08, label: `${head} · ${sub}` });
    s.addText("Top-5 finishes, of six measures", { x: x + 0.25, y: 1.96, w: w - 0.5, h: 0.3, fontFace: F, fontSize: 14, color: C.ink600, isTextBox: true, margin: 0 });
    const rh = 0.64, mx = rows[0][1].length;
    rows.forEach(([n, ks], j) => {
      const y = 2.38 + j * rh;
      if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: x + 0.18, y, w: w - 0.36, h: rh, fill: { color: C.zebra } });
      s.addText(n, { x: x + 0.3, y: y + 0.05, w: 2.6, h: 0.3, fontFace: F, fontSize: 14, bold: true, color: C.ink, isTextBox: true, margin: 0 });
      s.addText(ks.map(k => SHORT[k]).join(", "), { x: x + 0.3, y: y + 0.33, w: 3.85, h: 0.26, fontFace: F, fontSize: 13, color: C.ink600, isTextBox: true, margin: 0 });
      s.addShape(pptx.ShapeType.rect, { x: x + 4.25, y: y + rh * 0.3, w: ks.length / mx * 1.05, h: rh * 0.4, fill: { color: C.good } });
      s.addText(String(ks.length), { x: x + w - 0.75, y, w: 0.45, h: rh, fontFace: F, fontSize: 18, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    });
  });
  const inA = new Set(tops("yoy", 5).map(e => e[0])), both = tops("mom", 5).filter(e => inA.has(e[0])).map(e => e[0]);
  const lead = rows => { const t = rows.filter(r => r[1].length === rows[0][1].length); return t.length === 1 ? `${t[0][0]} leads` : `${cap(NUMW[t.length])} stores tie`; };
  H.takeaway(pptx, s, `Last year: ${lead(A)}. Moving now: ${lead(B)}. ${both.length === 0 ? "No store makes both views." : cap(NUMW[both.length]) + " stores make both views."}`, "g");
  H.foot(pptx, s, BASIS + " Left: Jan-Sep 2026 vs Jan-Sep 2025. Right: Q3 2026 vs Q1 2026. Top-5 finishes counted across six measures; top six stores shown in each view.");
}

/* ---- the ask: one question for each measure's top mover ---- */
function whatChanged() {
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: EB, title: "What Changed?" });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const X = { n: 0.82, m: 3.85, v: 6.2, st: 8.45, q: 10.2 }, rh = 0.66, TOP = 1.98;
  [["Store", X.n], ["Moved on", X.m], ["Q1 → Q3", X.v], ["Stepped up in", X.st], ["Change", X.q]].forEach(([t, x]) =>
    s.addText(t, { x, y: 1.5, w: 2.2, h: 0.32, fontFace: F, fontSize: 14, bold: true, color: C.ink600, isTextBox: true, margin: 0 }));
  const rows = ORDER.map(k => [k, MO[k].rows[0]]);
  rows.forEach(([k, r], j) => {
    const y = TOP + j * rh;
    if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.68, y, w: 11.97, h: rh, fill: { color: C.zebra } });
    [[r.n, X.n, 2.9, true, C.ink], [MO[k].label, X.m, 2.3, false, C.ink], [`${val(k, r.v0)} → ${val(k, r.v1)}`, X.v, 2.2, false, C.ink],
     [r.step, X.st, 1.7, true, C.navy], [chg(k, r.chg), X.q, 1.4, true, C.goodInk]].forEach(([t, x, w, b, col]) =>
      s.addText(t, { x, y, w, h: rh, fontFace: F, fontSize: 15, bold: b, color: col, valign: "middle", isTextBox: true, margin: 0 }));
  });
  const stores = [...new Set(rows.map(r => r[1].n))];
  H.takeaway(pptx, s, `${cap(NUMW[stores.length])} stores, one question each: what did you start doing that month? One sentence, before you leave.`, "n");
  H.foot(pptx, s, SRC_M + " Each row is the top Q1-to-Q3 mover on that measure.");
}

["gross", "ros", "hpro", "elr", "pro", "pgp", "align", "bal", "bf", "fuel"].forEach(moving);
await pptx.writeFile({ fileName: "./Who's Moving Now - Review.pptx" });
console.log("ok");
