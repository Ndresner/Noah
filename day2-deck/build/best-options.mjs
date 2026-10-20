/* Design options for "Your Best Techs Prove It" (2026-10-16). Same data as prof-deck.mjs (prof/opp.json). */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F, sz } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const O = JSON.parse(fs.readFileSync("./prof/opp.json", "utf8"));
const BEN = 0.85, pc = v => Math.floor(v * 100) + "%";
const U = O.S.filter(o => o.prof < BEN).sort((a, b) => a.prof - b.prof);
const proven = U.filter(o => o.top >= BEN);
const SEG = "Fixed Ops Productivity", TITLE = "Your Best Techs Prove It";
const SRC = "Stores under 85% only. Store = sold / paid hours for counted technicians. Best techs = the store's most proficient quarter of counted technicians (at least one), their sold / paid hours combined. Same population and rules as the proficiency slides.";
const TAKE = `At ${proven.length} of the ${U.length} stores under 85%, the best quarter of techs already beat 85%. Same bays, same dispatch, same car count.`;
const note = s => s.addNotes("Source: " + SRC + "\n" + U.map(o => `${o.n}: store ${pc(o.prof)}, best ${o.topn} of ${o.techs} techs ${pc(o.top)}`).join("\n"));
const label = (s, t, o) => s.addText(H.words(t, "label"), { fontFace: F, margin: 0, isTextBox: true, ...o });

/* ---------- OPTION A: two plain bars per store, one column, stores under 85% only ---------- */
{
  const s = pptx.addSlide(); H.page(pptx, s, { eyebrow: SEG + " · Option A", title: TITLE });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const X0 = 0.75, NAMEW = 2.35, BX = X0 + NAMEW + 0.1, BW = 8.2, MAXV = 1.5, TOP = 1.95, rowH = (6.25 - TOP) / U.length;
  const bx = v => BX + Math.min(v, MAXV) / MAXV * BW;
  [[C.bad, "Store today"], [C.good, "Its best techs"]].forEach(([col, t], i) => {
    s.addShape(pptx.ShapeType.rect, { x: X0 + i * 2.6, y: 1.5, w: 0.26, h: 0.22, fill: { color: col } });
    label(s, t, { x: X0 + 0.36 + i * 2.6, y: 1.42, w: 2.2, h: 0.38, fontSize: sz(15, "legend"), color: C.ink, valign: "middle" });
  });
  label(s, "TARGET 85%", { x: bx(BEN) - 1.0, y: 1.42, w: 2.0, h: 0.38, fontSize: sz(14, "target"), bold: true, color: C.navy, align: "center", valign: "middle", charSpacing: 0.8 });
  U.forEach((o, k) => {
    const y = TOP + k * rowH, h = rowH * 0.36;
    label(s, o.n, { x: X0, y, w: NAMEW, h: rowH, fontSize: 13, color: C.ink, valign: "middle" });
    [[o.prof, C.bad, C.badInk, y + rowH * 0.1], [o.top, o.top >= BEN ? C.good : C.bad, o.top >= BEN ? C.goodInk : C.badInk, y + rowH * 0.1 + h + 0.02]].forEach(([v, fill, ink, yy]) => {
      s.addShape(pptx.ShapeType.rect, { x: BX, y: yy, w: bx(v) - BX, h, fill: { color: fill } });
      label(s, pc(v) + (v > MAXV ? "+" : ""), { x: bx(v) + 0.06, y: yy - 0.05, w: 0.8, h: h + 0.1, fontSize: 13, bold: true, color: ink, valign: "middle" });
    });
  });
  s.addShape(pptx.ShapeType.rect, { x: bx(BEN) - 0.01, y: TOP - 0.05, w: 0.02, h: U.length * rowH + 0.05, fill: { color: C.navy } });
  H.takeaway(pptx, s, TAKE, "n"); note(s);
}

/* ---------- OPTION B: one card per store, "today -> best techs" in big type ---------- */
{
  const s = pptx.addSlide(); H.page(pptx, s, { eyebrow: SEG + " · Option B", title: TITLE });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  label(s, "Each store under 85%: where the store is today, and where its best quarter of techs already runs.", { x: 0.75, y: 1.42, w: 11.8, h: 0.4, fontSize: sz(15, "intro"), color: C.ink600, valign: "middle" });
  const COLS = 5, GAP = 0.16, CW = (11.83 - GAP * (COLS - 1)) / COLS, CH = 1.38, Y0 = 1.92;
  U.forEach((o, i) => {
    const c = i % COLS, r = Math.floor(i / COLS), x = 0.75 + c * (CW + GAP), y = Y0 + r * (CH + 0.1), ok = o.top >= BEN;
    s.addShape(pptx.ShapeType.rect, { x, y, w: CW, h: CH, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.07, h: CH, fill: { color: ok ? C.good : C.bad } });
    label(s, o.n, { x: x + 0.2, y: y + 0.08, w: CW - 0.25, h: 0.34, fontSize: 13, bold: true, fit: "shrink", wrap: false, color: C.ink, valign: "middle" });
    const half = (CW - 0.3) / 2;
    [[o.prof, C.badInk, "today", 0], [o.top, ok ? C.goodInk : C.badInk, "best techs", 1]].forEach(([v, ink, t, j]) => {
      label(s, pc(v), { x: x + 0.2 + j * half, y: y + 0.42, w: half, h: 0.55, fontSize: sz(22, "card number"), bold: true, color: ink, valign: "middle", wrap: false });
      label(s, t, { x: x + 0.2 + j * half, y: y + 0.97, w: half, h: 0.32, fontSize: 13, color: C.ink600, valign: "middle" });
    });
    label(s, "\u2192", { x: x + 0.2 + half - 0.32, y: y + 0.42, w: 0.3, h: 0.55, fontSize: 18, color: C.ink300, valign: "middle" });
  });
  H.takeaway(pptx, s, TAKE, "n"); note(s);
}

/* ---------- OPTION C: the headline number, then a plain table ---------- */
{
  const s = pptx.addSlide(); H.page(pptx, s, { eyebrow: SEG + " · Option C", title: TITLE });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  label(s, `${proven.length} of ${U.length}`, { x: 0.8, y: 1.9, w: 4.2, h: 1.3, fontSize: sz(80, "hero"), bold: true, color: C.goodInk, valign: "middle" });
  label(s, "stores under 85% already have techs beating it.", { x: 0.8, y: 3.2, w: 4.1, h: 0.9, fontSize: sz(22, "hero sub"), bold: true, color: C.ink, valign: "top" });
  label(s, "Same bays, same dispatch, same car count. The target is not theory; their own best techs hit it every month.", { x: 0.8, y: 4.25, w: 4.0, h: 1.4, fontSize: sz(16, "hero body"), color: C.ink600, valign: "top" });
  const hdr = ["Store", "Store today", "Best techs"].map((t, i) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy }, align: i ? "right" : "left" } }));
  const rows = U.map((o, k) => [
    { text: o.n, options: { color: C.ink, fill: { color: k % 2 ? C.white : C.zebra } } },
    { text: pc(o.prof), options: { color: C.badInk, bold: true, align: "right", fill: { color: k % 2 ? C.white : C.zebra } } },
    { text: pc(o.top), options: { color: o.top >= BEN ? C.goodInk : C.badInk, bold: true, align: "right", fill: { color: k % 2 ? C.white : C.zebra } } },
  ]);
  s.addTable([hdr, ...rows], { x: 5.4, y: 1.55, w: 7.15, colW: [3.75, 1.7, 1.7], rowH: 0.31, fontFace: F, fontSize: 14, margin: [0, 0.1, 0, 0.1], valign: "middle",
    border: { type: "none" } });
  H.takeaway(pptx, s, TAKE, "n"); note(s);
}
await pptx.writeFile({ fileName: "Best Techs - Design Options.pptx" }); console.log("options built", U.length, proven.length);
