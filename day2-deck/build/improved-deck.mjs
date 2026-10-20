/* Most Improved stores (2026-10-19, for review). Qlik - Closed Repair Orders, customer pay, by close date,
   same-store (INEOS excluded), GMC Prescott body shop labor types excluded. Store rows reconcile to the deck:
   CP ROs 279,300 / 279,756, CP hours 504,018.97 / 523,141.86, ELR $158.86 / $162.42 (Jan-Sep 2025 / 2026).
   Data: imp/imp.json from imp/compute.py. Every name, count and number on the slides is derived from it. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F } = H;
const D = JSON.parse(fs.readFileSync("./imp/imp.json", "utf8"));
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };

const SRC = "Qlik - Closed Repair Orders. Customer pay, by close date. Same-store, 33 stores (INEOS excluded, opened 2026). GMC Prescott body shop labor types excluded. Change is 2026 vs the same months of 2025, so seasonality cancels out. Stores ranked by percent change (points for parts GP %), so store size does not drive the ranking. CP gross = CP labor gross + CP parts gross.";
const PER = { ytd: { name: "January–September", short: "Jan–Sep", vs: "vs Jan–Sep 2025" },
              q3:  { name: "Last 90 Days", short: "Jul–Sep", vs: "Jul–Sep vs Jul–Sep 2025" } };
const ORDER = ["gross", "ros", "hpro", "elr", "pro", "pgp"];
const SHORT = { gross: "CP gross", ros: "CP RO count", hpro: "hours/RO", elr: "ELR", pro: "parts $/RO", pgp: "parts GP%" };
const sg = v => v >= 0 ? "+" : "−";
const tr1 = v => (Math.trunc(Math.abs(v) * 10) / 10).toFixed(1);           // truncated, house style
const chg = (k, v) => D.ytd[k].kind === "pts" ? `${sg(v)}${tr1(v * 100)} pts` : `${sg(v)}${tr1(v * 100)}%`;
const val = (k, v) => ({
  gross: v >= 1e6 ? "$" + (v / 1e6).toFixed(2) + "M" : "$" + Math.round(v / 1e3) + "K",
  ros: Math.round(v).toLocaleString("en-US"), hpro: v.toFixed(2), elr: "$" + v.toFixed(2),
  pro: "$" + Math.round(v), pgp: (v * 100).toFixed(1) + "%" }[k]);
const title = s => s;   // store names in caps, same as every ranker in the deck
const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
const NUMW = ["zero", "one", "two", "three", "four", "five", "six"];

/* who shows up most in the top 3 of one period */
function repeaters(per, n = 3) {
  const c = {};
  ORDER.forEach(k => D[per][k].rows.slice(0, n).forEach(r => (c[r[0]] ??= []).push(k)));
  return Object.entries(c).sort((a, b) => b[1].length - a[1].length);
}

/* ---- 1-2: six-metric grid, top 3 each ---- */
function grid(per) {
  const P = PER[per], s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: "Fixed Ops · Most Improved", title: `Most Improved Stores · ${P.name}` });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  s.addText(`Top three stores on each measure, ${P.vs}`, { x: 0.78, y: 1.42, w: 11.77, h: 0.32, fontFace: F, fontSize: 14,
    color: C.ink600, isTextBox: true, margin: 0 });
  const gap = 0.22, W = (11.77 - gap * 2) / 3, Hc = 2.12;
  ORDER.forEach((k, i) => {
    const M = D[per][k], x = 0.78 + (i % 3) * (W + gap), y = 1.84 + Math.floor(i / 3) * (Hc + 0.14);
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: Hc, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: 0.06, fill: { color: C.good } });
    s.addText(M.label.toUpperCase(), { x: x + 0.16, y: y + 0.14, w: W - 1.6, h: 0.3, fontFace: F, fontSize: 14, bold: true,
      color: C.ink600, charSpacing: 0.6, isTextBox: true, margin: 0 });
    s.addText(`Group ${chg(k, M.group[2])}`, { x: x + W - 1.5, y: y + 0.14, w: 1.34, h: 0.3, fontFace: F, fontSize: 13,
      color: C.ink600, align: "right", isTextBox: true, margin: 0 });
    M.rows.slice(0, 3).forEach((r, j) => {
      const ry = y + 0.48 + j * 0.53;
      s.addText(String(j + 1), { x: x + 0.16, y: ry, w: 0.26, h: 0.28, fontFace: F, fontSize: 14, bold: true, color: C.navy, isTextBox: true, margin: 0 });
      s.addText(title(r[0]), { x: x + 0.44, y: ry, w: W - 1.7, h: 0.28, fontFace: F, fontSize: 14, bold: true, color: C.ink, isTextBox: true, margin: 0 });
      s.addText(`${val(k, r[1])} → ${val(k, r[2])}`, { x: x + 0.44, y: ry + 0.27, w: W - 1.7, h: 0.24, fontFace: F, fontSize: 13,
        color: C.ink600, isTextBox: true, margin: 0 });
      s.addText(chg(k, r[3]), { x: x + W - 1.4, y: ry, w: 1.24, h: 0.5, fontFace: F, fontSize: 18, bold: true,
        color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    });
  });
  const rep = repeaters(per), hi = rep[0][1].length;
  const lead = rep.filter(e => e[1].length === hi).map(e => title(e[0])), next = rep.filter(e => e[1].length === hi - 1).map(e => title(e[0]));
  const tk = `${list(lead)} ${lead.length === 1 ? "makes" : "each make"} ${NUMW[hi]} of six lists` +
    (next.length ? `; ${next.length > 2 ? NUMW[next.length] + " stores" : list(next)} make ${NUMW[hi - 1]}.` : ".");
  H.takeaway(pptx, s, tk, "g");
  H.foot(pptx, s, SRC + ` Period: ${P.short} 2026 vs ${P.short} 2025. Stores with under 50 CP ROs in the base period not ranked (none were excluded).`);
  s.addNotes(`\n\nReview note: percent change rewards a store coming off a weak base. The before → after values are on the card so the room sees the starting point.`);
}

/* ---- 3: honor roll, both periods ---- */
function honor() {
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: "Fixed Ops · Most Improved", title: "Most Improved: Who Shows Up Most" });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const c = {};
  ["ytd", "q3"].forEach(per => ORDER.forEach(k => D[per][k].rows.slice(0, 3).forEach(r => (c[r[0]] ??= { ytd: [], q3: [] })[per].push(k))));
  const rows = Object.entries(c).map(([n, v]) => [n, v, v.ytd.length + v.q3.length])
    .sort((a, b) => b[2] - a[2] || b[1].q3.length - a[1].q3.length).slice(0, 8);
  const X = { name: 0.78, bar: 3.2, cnt: 6.1, ytd: 6.75, q3: 9.65 }, TOP = 1.96, rh = 0.53, maxN = rows[0][2];
  [["Top-3 finishes of 12", X.bar], ["Jan–Sep", X.ytd], ["Last 90 days", X.q3]].forEach(([t, x]) =>
    s.addText(t, { x, y: 1.46, w: 2.8, h: 0.34, fontFace: F, fontSize: 14, bold: true, color: C.ink600, isTextBox: true, margin: 0 }));
  rows.forEach(([n, v, t], i) => {
    const y = TOP + i * rh;
    if (i % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 0.68, y, w: 11.97, h: rh, fill: { color: C.zebra } });
    s.addText(title(n), { x: X.name, y, w: 2.3, h: rh, fontFace: F, fontSize: 14, bold: true, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    s.addShape(pptx.ShapeType.rect, { x: X.bar, y: y + rh * 0.25, w: t / maxN * 2.9, h: rh * 0.5, fill: { color: C.good } });
    s.addText(String(t), { x: X.cnt - 0.1, y, w: 0.5, h: rh, fontFace: F, fontSize: 16, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    [["ytd", X.ytd], ["q3", X.q3]].forEach(([p, x]) => s.addText(v[p].map(k => SHORT[k]).join(", ") || "—",
      { x, y, w: 3.0, h: rh, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 }));
  });
  const both = Object.values(c).filter(v => v.ytd.length && v.q3.length).length;
  const lead = rows[0];
  H.takeaway(pptx, s, `${title(lead[0])} has ${lead[2]} top-3 finishes of 12. ${both} stores make a top-3 list in both periods.`, "g");
  H.foot(pptx, s, SRC + " Top-3 finishes counted across six measures in each of two periods (Jan-Sep and Jul-Sep), 12 lists in all. Top 8 stores shown.");
}

/* ---- 4-9: one measure, top 5, both periods side by side ---- */
function metric(k) {
  const s = pptx.addSlide(), M = D.ytd[k];
  H.page(pptx, s, { eyebrow: "Fixed Ops · Most Improved", title: `Most Improved: ${M.label}` });
  [["ytd", 0.5], ["q3", 6.75]].forEach(([per, x]) => {
    const P = PER[per], R = D[per][k], w = 6.08;
    H.panel(pptx, s, { x, y: 1.32, w, h: 5.08, label: `${P.name} · ${P.vs.replace("vs ", "vs ")}` });
    s.addText(`Group ${chg(k, R.group[2])} · ${R.up} of ${R.rows.length} stores improved`,
      { x: x + 0.25, y: 1.98, w: w - 0.5, h: 0.32, fontFace: F, fontSize: 14, color: C.ink600, isTextBox: true, margin: 0 });
    const top = R.rows.slice(0, 5), mx = top[0][3], rh = 0.72;
    top.forEach((r, j) => {
      const y = 2.46 + j * rh;
      if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: x + 0.18, y, w: w - 0.36, h: rh, fill: { color: C.zebra } });
      s.addText(String(j + 1), { x: x + 0.3, y, w: 0.3, h: rh, fontFace: F, fontSize: 18, bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
      s.addText(title(r[0]), { x: x + 0.65, y: y + 0.08, w: 2.5, h: 0.3, fontFace: F, fontSize: 15, bold: true, color: C.ink, isTextBox: true, margin: 0 });
      s.addText(`${val(k, r[1])} → ${val(k, r[2])}`, { x: x + 0.65, y: y + 0.38, w: 2.5, h: 0.26, fontFace: F, fontSize: 13, color: C.ink600, isTextBox: true, margin: 0 });
      s.addShape(pptx.ShapeType.rect, { x: x + 3.2, y: y + 0.24, w: Math.max(r[3] / mx * 1.4, 0.03), h: 0.24, fill: { color: C.good } });
      s.addText(chg(k, r[3]), { x: x + w - 1.42, y, w: 1.12, h: rh, fontFace: F, fontSize: 17, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    });
  });
  const a = D.ytd[k].rows.slice(0, 5).map(r => r[0]), b = D.q3[k].rows.slice(0, 5).map(r => r[0]);
  const both = a.filter(n => b.includes(n)).map(title);
  const lead = a[0] === b[0] ? `${title(a[0])} leads both periods.` : `${title(a[0])} leads Jan–Sep; ${title(b[0])} leads the last 90 days.`;
  H.takeaway(pptx, s, `${lead} ${both.length === 0 ? "No store is top five in both." : both.length > 2 ? `${NUMW[both.length][0].toUpperCase() + NUMW[both.length].slice(1)} stores are top five in both.` : `${list(both)} ${both.length === 1 ? "is" : "are"} top five in both.`}`, "g");
  H.foot(pptx, s, SRC);
}

grid("ytd"); grid("q3"); honor(); ORDER.forEach(metric);
await pptx.writeFile({ fileName: "./Most Improved - Review.pptx" });
console.log("ok");
