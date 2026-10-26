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

// counts per store: every record set in 2026 (service, parts and combined count separately)
const cnt = r => r.svc.length + r.pts.length + r.tot.length;
const total = rec.reduce((t, r) => t + cnt(r), 0);
const byCnt = [...rec].sort((a, b) => cnt(b) - cnt(a) || b.tot.length - a.tot.length || a.n.localeCompare(b.n));
const mo = a => a.map(m => m.slice(0, 3)).join(", ");
const what = r => [["Service", r.svc], ["Parts", r.pts], ["Combined", r.tot]].filter(([, a]) => a.length).map(([k, a]) => `${k}: ${mo(a)}`).join("  ·  ");
const topLine = (() => { const t = byCnt.slice(0, 5); return list(t.map(r => `${nm(r.n)} ${cnt(r)}`)); })();

// 1. headline (Oct 26, Noah: store records only, so the group-month tile is out; the band names the leaders)
H.statTiles(pptx, pptx, { eyebrow: SEG, title: "A Record Year", bigSize: 40,
  tiles: [
    { label: "Stores with an all-time record month", big: String(rec.length), sub: `of 33 set their best month ever in service, parts or combined during 2026.`, tone: "g" },
    { label: "Records set", big: String(total), sub: `all-time monthly records across those stores: service, parts and combined counted separately.`, tone: "g" },
    { label: "Best January–September ever", big: String(ytdRec.length), sub: `stores had their best January–September ever in combined fixed gross.`, tone: "g" },
    { label: "Broke it, then broke it again", big: String(again.length), sub: `stores set a record and then beat it again later in the year.`, tone: "g" },
  ],
  banner: `Most records set this year: ${topLine}.`,
  takeaway: `${MON[busiest].replace(" 2026", "")} alone brought ${byMonth[busiest]} combined store records. Every one of these stores can tell the room what they changed.`,
  tone: "g", source: SRC });

// 2. OPTION A: ranker of stores by records set, with the split
H.amountRanker(pptx, pptx, { eyebrow: SEG, title: "Records Set in 2026",
  rows: byCnt.map(r => [r.n, cnt(r)]), fmt: v => String(v), color: C.good, ink: C.goodInk, valLabel: "Records",
  extra: [{ label: "Svc", w: 0.55, vals: Object.fromEntries(byCnt.map(r => [r.n, String(r.svc.length || "–")])) },
          { label: "Parts", w: 0.6, vals: Object.fromEntries(byCnt.map(r => [r.n, String(r.pts.length || "–")])) },
          { label: "Comb.", w: 0.65, vals: Object.fromEntries(byCnt.map(r => [r.n, String(r.tot.length || "–")])) }],
  takeaway: `${rec.length} stores set ${total} all-time monthly records in 2026. ${nm(byCnt[0].n)} and ${nm(byCnt[1].n)} lead with ${cnt(byCnt[0])} each.`,
  tone: "g", source: SRC + " Records = each month a store beat its best-ever service, parts or combined month; one month can count up to three." })
  .addNotes("OPTION A. Records by store (month: service / parts / combined, with the record it beat):\n" + rec.map(r => `${nm(r.n)}: ` + ["svc", "pts", "tot"].flatMap(k => r.E[k].map(e => `${{ svc: "service", pts: "parts", tot: "combined" }[k]} ${e.m} ${H.kusd(e.v)} (beat ${H.kusd(e.prev)}, ${e.prevAt})`)).join("; ")).join("\n"));

// 3. OPTION B: a plain list, two columns: store, how many, what they were
{
  const tk = `${rec.length} stores set ${total} records. ${nm(byCnt[0].n)} and ${nm(byCnt[1].n)} set ${cnt(byCnt[0])} each.`;
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: SEG, title: "Who Set Records in 2026" });
  const top = H.BODY.top, h = H.bodyH(tk);
  H.panel(pptx, s, { x: 0.5, y: top, w: 12.33, h, label: null });
  const per = Math.ceil(byCnt.length / 2), rh = (h - 0.25) / per;
  byCnt.forEach((r, i) => {
    const c = i < per ? 0 : 1, k = i - c * per, x = 0.72 + c * 6.1, y = top + 0.12 + k * rh;
    if (k % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: x - 0.08, y, w: 5.95, h: rh, fill: { color: C.zebra } });
    s.addText(String(cnt(r)), { x, y, w: 0.4, h: rh, fontFace: F, fontSize: sz(14, "count"), bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(r.n, { x: x + 0.5, y, w: 1.95, h: rh, fontFace: F, fontSize: 12, bold: true, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(what(r), { x: x + 2.45, y, w: 3.4, h: rh, fontFace: F, fontSize: 12, color: C.ink600, valign: "middle", isTextBox: true, margin: 0 });
  });
  H.takeaway(pptx, s, tk, "g");
  H.foot(pptx, s, SRC + " Count = every record set: service, parts and combined counted separately. Months listed are the months each record was set.");
  s.addNotes("OPTION B: one line per store, how many records and which months.");
}

// 4. OPTION C: spotlight the leaders in words, everyone else on one line
{
  const lead = byCnt.slice(0, 6), rest = byCnt.slice(6);
  const tk = `Ask ${list(lead.slice(0, 3).map(r => nm(r.n)))} what changed this year.`;
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: SEG, title: "Record Setters" });
  const top = H.BODY.top, h = H.bodyH(tk), gap = 0.16, W = (12.33 - 0.44 - gap * 2) / 3, Ht = (h - 1.25 - gap) / 2;
  H.panel(pptx, s, { x: 0.5, y: top, w: 12.33, h, label: null });
  const say = r => {
    const best = k => r.E[k].length ? r.E[k].at(-1) : null, parts = [];
    if (r.tot.length) parts.push(`Best combined month ever ${r.tot.length === 1 ? "in " + mo(r.tot) : r.tot.length + " times (" + mo(r.tot) + ")"}, now ${H.kusd(best("tot").v)}.`);
    if (r.svc.length) parts.push(`Service record ${r.svc.length === 1 ? "in " + mo(r.svc) : r.svc.length + " times"}.`);
    if (r.pts.length) parts.push(`Parts record ${r.pts.length === 1 ? "in " + mo(r.pts) : r.pts.length + " times"}.`);
    return parts.join(" ");
  };
  lead.forEach((r, i) => {
    const x = 0.72 + (i % 3) * (W + gap), y = top + 0.18 + Math.floor(i / 3) * (Ht + gap);
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: Ht, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.06, h: Ht, fill: { color: C.good } });
    s.addText([{ text: nm(r.n) + "   ", options: { bold: true, color: C.ink, fontSize: sz(17, "card head") } }, { text: `${cnt(r)} records`, options: { bold: true, color: C.goodInk, fontSize: sz(15, "card count") } }],
      { x: x + 0.2, y: y + 0.1, w: W - 0.3, h: 0.36, fontFace: F, isTextBox: true, margin: 0 });
    s.addText(say(r), { x: x + 0.2, y: y + 0.5, w: W - 0.3, h: Ht - 0.6, fontFace: F, fontSize: sz(14, "card body"), color: C.ink600, valign: "top", isTextBox: true, margin: 0 });
  });
  s.addText([{ text: "Also set records: ", options: { bold: true, color: C.ink } }, { text: rest.map(r => `${nm(r.n)} (${cnt(r)})`).join(", "), options: { color: C.ink600 } }],
    { x: 0.72, y: top + h - 0.86, w: 11.9, h: 0.78, fontFace: F, fontSize: sz(14, "rest"), valign: "middle", isTextBox: true, margin: 0 });
  H.takeaway(pptx, s, tk, "g");
  H.foot(pptx, s, SRC + " Records counted separately for service, parts and combined.");
  s.addNotes("OPTION C: the six stores with the most records in words; everyone else on one line.");
}
await pptx.writeFile({ fileName: "Store Records - Opening.pptx" }); console.log("records opening built", rec.length, ytdRec.length, again.length);
