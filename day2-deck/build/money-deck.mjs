/* "Where the Money Is" (2026-10-19, for review). Every figure is labor + parts gross a year, already built on its own
   review slide: proficiency (Technician Proficiency, $1,646,932/mo x 12), 0.1 hr per RO (Hours per RO Value), op codes
   (Op Code Missed Opportunity), one more CP RO a day (CP Hours & RO a Day), one tech minute (Tech Minute). The figures
   overlap (op codes are part of hours per RO; selling hours needs tech time), so they are not added together. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F, sz } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const usd = v => "$" + Math.round(v).toLocaleString("en-US");
const O = JSON.parse(fs.readFileSync("./prof/opp.json", "utf8")).G;
const prof = (O.lab + O.parts) / O.months * 12;
const L = [
  { name: "Every tech under 85% reaches 85%", v: prof, how: "Dispatch, parts staging, approvals, coaching", own: ["Service manager", "GM"], src: "Technician Proficiency" },
  { name: "0.1 more hours on every CP RO", v: 6602564, how: "MPI on every car, the menu, Road to the Sale", own: ["Service manager"], src: "Hours per RO" },
  { name: "Every store at the four op code benchmarks", v: 4895199, how: "Alignment, balance, brake fluid, fuel service", own: ["Service manager"], src: "Op Codes" },
  { name: "One more customer-pay RO a day", v: 4473284, how: "Phones, scheduler, retention, marketing", own: ["GM", "Service manager"], src: "CP RO count" },
  { name: "One minute a day saved for every tech", v: 444380, how: "Parts to the bay, tools and keys in place", own: ["Service manager", "GM"], src: "Tech Minute" },
];
const s = pptx.addSlide(); H.page(pptx, s, { eyebrow: "Fixed Ops", title: "Where the Money Is" });
H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
const t = (txt, o) => s.addText(H.words(txt, "money"), { fontFace: F, margin: 0, isTextBox: true, valign: "middle", ...o });
[["THE LEVER", 0.78, 4.4, "left"], ["GROSS A YEAR", 5.2, 2.3, "right"], ["HOW", 7.85, 2.9, "left"], ["WHO OWNS IT", 10.8, 1.85, "left"]]
  .forEach(([h, x, w, al]) => t(h, { x, y: 1.48, w, h: 0.34, fontSize: sz(14, "head"), bold: true, color: C.navy, charSpacing: 1.0, align: al }));
const TOP = 1.92, RH = 0.86;
L.forEach((l, i) => {
  const y = TOP + i * RH;
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y, w: 12.13, h: RH - 0.08, fill: { color: i % 2 ? C.white : C.zebra } });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y, w: 0.07, h: RH - 0.08, fill: { color: C.good } });
  t(l.name, { x: 0.82, y, w: 4.3, h: RH - 0.08, fontSize: sz(17, "lever"), bold: true, color: C.ink });
  t(usd(l.v), { x: 5.2, y, w: 2.3, h: RH - 0.08, fontSize: sz(24, "lever $"), bold: true, color: C.goodInk, align: "right" });
  t(l.how, { x: 7.85, y, w: 2.85, h: RH - 0.08, fontSize: sz(14, "how"), color: C.ink600 });
  const n = l.own.length, ch = 0.3, gap = 0.06, y0 = y + (RH - 0.08 - (n * ch + (n - 1) * gap)) / 2;
  l.own.forEach((o, k) => {
    const w = o === "GM" ? 0.62 : 1.85, gm = o === "GM", yy = y0 + k * (ch + gap);
    s.addShape(pptx.ShapeType.roundRect, { x: 10.8, y: yy, w, h: ch, rectRadius: 0.06, fill: { color: gm ? C.navy : C.blueTint }, line: { color: C.navy, width: 0.75 } });
    t(o, { x: 10.8, y: yy, w, h: ch, fontSize: sz(14, "owner"), bold: true, color: gm ? C.white : C.navy, align: "center" });
  });
});
H.takeaway(pptx, s, "Every number is gross on cars, bays and techs we already have, except one more RO a day. They overlap, so read each on its own.", "g");
H.foot(pptx, s, "Labor + parts gross a year, each from its own slide: Technician Proficiency (each tech under 85% raised to 85%, monthly x 12), Hours per RO (0.1 hr x CP ROs at each store's gross per hour), Op Codes (stores under each benchmark brought to it), One More CP RO a Day (307 open days), Tech Minute (one minute per tech per day). Jan-Sep 2026, same-store, 33 stores. The levers overlap - op code services are part of hours per RO, and selling more hours needs tech time - so they are not added together.");
s.addNotes("Talk track: the GM owns the conditions - staffing, pay plans, marketing budget, the parts department. The service manager owns the daily process - dispatch, the inspection, the menu, the drive. Each line has its own slide earlier in the day.");
await pptx.writeFile({ fileName: "Where the Money Is - Review.pptx" }); console.log("built", usd(prof));
