/* Store records for the opening (Noah, Oct 25): who broke an all-time record this year. Replaces the Open Decisions
   record slides. Slide 1: the headline numbers. Slide 2: every 2026 month a store set an all-time monthly record in
   service, parts or combined. Data: records/records.json (Qlik Fixed Ops Accounting TB, Jan 2021 - Sep 2026, from
   records/compute.py). TB matches the P&L for every store except GMC Prescott service (body shop), see CLAUDE.md. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F, sz, words } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const R = JSON.parse(fs.readFileSync("./records/records.json", "utf8"));
const SEG = "Fixed Ops Records";
const SRC = "Qlik - Fixed Ops Accounting TB, fixed gross (service and parts), monthly by store, January 2021 - September 2026. A record = the store's highest month on file; months before a store opened are ignored. GMC Prescott service on the TB includes body shop gross.";
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep"].map(m => m + " 2026");
const nm = H.storeName, list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
const B = R.broken, Y = R.ytd;
const rec = Object.entries(B).map(([n, c]) => ({ n, svc: c.svc.map(e => e.m), pts: c.pts.map(e => e.m), tot: c.tot.map(e => e.m), E: c }))
  .filter(r => r.svc.length + r.pts.length + r.tot.length);
const months = r => new Set([...r.svc, ...r.pts, ...r.tot]).size;
rec.sort((a, b) => months(b) - months(a) || b.tot.length - a.tot.length || a.n.localeCompare(b.n));
const ytdRec = Object.entries(Y).filter(([, r]) => r.rec.tot.record).map(([n]) => n);
const [bestM, bestV] = R.group.topMonths[0];
const byMonth = MON.map(m => rec.filter(r => r.tot.includes(m)).length), busiest = byMonth.indexOf(Math.max(...byMonth));
const again = rec.filter(r => r.svc.length > 1 || r.pts.length > 1 || r.tot.length > 1);

// 1. headline
H.statTiles(pptx, pptx, { eyebrow: SEG, title: "A Record Year", bigSize: 40,
  tiles: [
    { label: "Stores with an all-time record month", big: String(rec.length), sub: `of 33 set their best month ever in service, parts or combined during 2026.`, tone: "g" },
    { label: "Best January–September ever", big: String(ytdRec.length), sub: `stores had their best January–September ever in combined fixed gross.`, tone: "g" },
    { label: "Best group month ever", big: bestM, sub: `${H.Mmixed(bestV)} of fixed gross across the group, the most in any month on file.`, tone: "g" },
    { label: "Broke it, then broke it again", big: String(again.length), sub: `stores set a record and then beat it again later in the year.`, tone: "g" },
  ],
  takeaway: `${MON[busiest].replace(" 2026", "")} alone brought ${byMonth[busiest]} combined store records. Every one of these stores can tell the room what they changed.`,
  tone: "g", source: SRC });

// 2. record board: one row per store, one column per month
{
  const tk = `${list(again.slice(0, 3).map(r => nm(r.n)))} broke their records more than once this year. Ask them what they changed.`;
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: SEG, title: "Record Months in 2026" });
  const top = H.BODY.top, h = H.bodyH(tk);
  H.panel(pptx, s, { x: 0.5, y: top, w: 12.33, h, label: null });
  const NX = 0.72, NW = 2.45, MX = NX + NW + 0.05, CW = 0.98, HY = top + 0.1, RT = top + 0.62, rh = (h - 0.72) / rec.length;
  s.addText("S = service record   P = parts record   green cell = best combined month ever", { x: NX, y: HY, w: 11.5, h: 0.24, fontFace: F, fontSize: sz(14, "legend"), color: C.ink600, isTextBox: true, margin: 0 });
  MON.forEach((m, j) => s.addText(m.slice(0, 3).toUpperCase(), { x: MX + j * CW, y: HY + 0.26, w: CW, h: 0.26, fontFace: F, fontSize: sz(14, "column header"), bold: true, color: C.ink600, align: "center", isTextBox: true, margin: 0 }));
  rec.forEach((r, i) => {
    const y = RT + i * rh;
    if (i % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: NX - 0.08, y, w: NW + 0.13 + 9 * CW, h: rh, fill: { color: C.zebra } });
    s.addText(r.n, { x: NX, y, w: NW, h: rh, fontFace: F, fontSize: 12, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    MON.forEach((m, j) => {
      const sv = r.svc.includes(m), pt = r.pts.includes(m), to = r.tot.includes(m);
      if (!(sv || pt || to)) return;
      const x = MX + j * CW;
      if (to) s.addShape(pptx.ShapeType.rect, { x: x + 0.08, y: y + 0.015, w: CW - 0.16, h: rh - 0.03, fill: { color: C.goodBg }, line: { color: C.good, width: 0.75 } });
      const t = [sv ? "S" : "", pt ? "P" : ""].filter(Boolean).join(" ") || "•";
      s.addText(t, { x, y, w: CW, h: rh, fontFace: F, fontSize: 12, bold: true, color: to ? C.goodInk : C.navy, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    });
  });
  H.takeaway(pptx, s, tk, "g");
  H.foot(pptx, s, SRC + " A green cell with a dot only: the combined record came without a service or parts record that month.");
  s.addNotes("Records by store (month: service / parts / combined, with the record it beat):\n" + rec.map(r => `${nm(r.n)}: ` + ["svc", "pts", "tot"].flatMap(k => r.E[k].map(e => `${{ svc: "service", pts: "parts", tot: "combined" }[k]} ${e.m} ${H.kusd(e.v)} (beat ${H.kusd(e.prev)}, ${e.prevAt})`)).join("; ")).join("\n"));
}
await pptx.writeFile({ fileName: "Store Records - Opening.pptx" }); console.log("records opening built", rec.length, ytdRec.length, again.length);
