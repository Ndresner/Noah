/* "What a Tech Minute Is Worth" (2026-10-19, for review). Per store: value of a minute = the store's counted
   technicians' labor gross + parts gross per sold hour / 60 (same rates as the proficiency opportunity slides).
   Tech-days a month = counted techs' worked hours / 8 / 9.49 months. Assumes each minute saved becomes a minute of
   flagged work. Same population and rules as the proficiency slides (583 counted technicians, INEOS excluded). */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F, sz } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const O = JSON.parse(fs.readFileSync("./prof/opp.json", "utf8")), MO = O.G.months;
const S = O.S.map(o => { const dpm = o.paid / 8 / MO, rate = o.lgh + o.pgh; return { n: o.n, techs: o.techs, dpm, rate, vpm: rate / 60, mo: dpm * rate / 60 }; });
const techs = S.reduce((t, s) => t + s.techs, 0), MOV = S.reduce((t, s) => t + s.mo, 0), YR = MOV * 12;
const rate = O.S.reduce((t, o) => t + (o.lgh + o.pgh) * o.sold, 0) / O.S.reduce((t, o) => t + o.sold, 0);
const usd = v => "$" + Math.round(v).toLocaleString("en-US"), usd2 = v => "$" + v.toFixed(2);
const SEG = "Technician Proficiency", TITLE = "What a Tech Minute Is Worth";
const SRC = "Value of a minute = each store's counted technicians' labor gross + parts gross per sold hour, divided by 60 (Qlik - Closed Repair Orders, all pay types, same rates as the proficiency opportunity slides). Tech-days = counted techs' worked hours (ADP) / 8 per day. Assumes each minute saved becomes a minute of flagged work. 583 counted technicians, same-store, INEOS excluded. Month = Dec 25, 2025 - Oct 9, 2026 / 9.49; year = month x 12.";

// slide 1 - the group
{
  const s = pptx.addSlide(); H.page(pptx, s, { eyebrow: SEG, title: TITLE });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 5.0, h: 5.08, label: null });
  H.panel(pptx, s, { x: 5.65, y: 1.32, w: 7.18, h: 5.08, label: null });
  const t = (txt, o) => s.addText(H.words(txt, "minute"), { fontFace: F, margin: 0, isTextBox: true, ...o });
  t("ONE TECHNICIAN HOUR", { x: 0.78, y: 1.55, w: 4.5, h: 0.32, fontSize: sz(14, "label"), bold: true, color: C.navy, charSpacing: 1.0 });
  t(usd(rate), { x: 0.78, y: 1.9, w: 4.5, h: 0.9, fontSize: sz(54, "hero"), bold: true, color: C.navy });
  t("in labor and parts gross, group average", { x: 0.78, y: 2.8, w: 4.5, h: 0.36, fontSize: sz(15, "sub"), color: C.ink600 });
  s.addShape(pptx.ShapeType.rect, { x: 0.78, y: 3.35, w: 4.45, h: 0.012, fill: { color: C.rule } });
  t("ONE TECHNICIAN MINUTE", { x: 0.78, y: 3.55, w: 4.5, h: 0.32, fontSize: sz(14, "label"), bold: true, color: C.navy, charSpacing: 1.0 });
  t(usd2(rate / 60), { x: 0.78, y: 3.9, w: 4.5, h: 0.9, fontSize: sz(54, "hero"), bold: true, color: C.goodInk });
  t(`About ${usd(YR / techs)} a year for each tech, one minute a day`, { x: 0.78, y: 4.8, w: 4.5, h: 0.7, fontSize: sz(15, "sub"), color: C.ink600, valign: "top" });

  t(`MINUTES SAVED PER TECH, EVERY DAY · ${techs} TECHNICIANS`, { x: 5.93, y: 1.55, w: 6.7, h: 0.32, fontSize: sz(14, "label"), bold: true, color: C.navy, charSpacing: 1.0 });
  const hdr = ["Minutes a day", "A month", "A year"].map((x, i) => ({ text: x, options: { bold: true, color: C.white, fill: { color: C.navy }, align: i ? "right" : "left" } }));
  const rows = [1, 5, 10, 15].map((m, k) => [m, MOV * m, YR * m]).map(([m, mo, yr], k) => [
    { text: m === 1 ? "1 minute" : `${m} minutes`, options: { bold: true, color: C.ink, fill: { color: k % 2 ? C.white : C.zebra } } },
    { text: H.kusd(mo), options: { color: C.ink, align: "right", fill: { color: k % 2 ? C.white : C.zebra } } },
    { text: H.kusd(yr), options: { bold: true, color: C.goodInk, align: "right", fill: { color: k % 2 ? C.white : C.zebra } } }]);
  s.addTable([hdr, ...rows], { x: 5.93, y: 2.0, w: 6.62, colW: [2.42, 2.1, 2.1], rowH: 0.52, fontFace: F, fontSize: sz(20, "minute table"), margin: [0, 0.12, 0, 0.12], valign: "middle", border: { type: "none" } });
  t("Where the minutes go: walking to the parts counter, waiting on dispatch or a customer approval, hunting for tools, chasing a key or a car.",
    { x: 5.93, y: 4.85, w: 6.62, h: 1.35, fontSize: sz(16, "where"), color: C.ink, valign: "top" });
  H.takeaway(pptx, s, `One minute a day for every tech is worth ${usd(MOV)} a month, ${usd(YR)} a year. No new techs, no new bays.`, "g");
  H.foot(pptx, s, SRC);
  s.addNotes("Talk track: staged parts delivered to the bay (Roundtable: 'Stage parts, deliver and retrieve'), dispatch with the next job ready, fast MPI approvals, tools and keys where they belong. Ten minutes a day is less than two parts-counter trips.\nAssumption to say out loud: a minute saved only counts if it turns into flagged time, so the store needs the work waiting.");
}

// slide 2 - by store
{
  const ex1 = Object.fromEntries(S.map(s => [s.n, String(s.techs)])), ex2 = Object.fromEntries(S.map(s => [s.n, usd2(s.vpm)]));
  const top = [...S].sort((a, b) => b.mo - a.mo)[0], hi = [...S].sort((a, b) => b.vpm - a.vpm)[0];
  H.amountRanker(pptx, pptx, {
    eyebrow: SEG, title: TITLE,
    rows: S.map(s => [s.n, s.mo * 12]), fmt: v => H.kusd(v, 1),   // $K, one decimal: every store is under $50K (Noah, Oct 22) color: C.blue, ink: C.navy, valLabel: "A year",
    extra: [{ label: "Techs", w: 0.62, vals: ex1 }, { label: "Per min", w: 0.82, vals: ex2 }],
    takeaway: `One minute a day per tech: ${top.n} ${usd(top.mo * 12)} a year. ${hi.n} techs are worth the most per minute at ${usd2(hi.vpm)}.`,
    source: SRC + " A year = one minute a day for every counted tech at the store, for a year. Per min = gross per tech minute.",
  });
}
await pptx.writeFile({ fileName: "Tech Minute - Review.pptx" }); console.log("built", usd(MOV), usd(YR));
