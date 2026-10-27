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
    { label: "Stores with all-time record months", big: String(rec.length), sub: `of 33 set their best month ever in service, parts or combined during 2026.`, tone: "g" },
    { label: "Store records set", big: String(total), sub: `all-time monthly records across those stores: service, parts and combined counted separately.`, tone: "g" },
    { label: "Best January–September ever", big: String(ytdRec.length), sub: `stores had their best January–September ever in combined fixed gross.`, tone: "g" },
    { label: "Stores who set a record and then reset it", big: String(again.length), sub: `stores set a record and then topped it again later in the year.`, tone: "g" },   // Noah, Oct 26
  ],
  banner: `Most records set this year: ${topLine}.`,
  takeaway: `${MON[busiest].replace(" 2026", "")} alone brought ${byMonth[busiest]} combined store records. Every one of these stores can tell the room what they changed.`,
  tone: "g", source: SRC });

// 2. ranker of stores by records set, with the split (Noah picked option A, Oct 26; list and spotlight options retired)
H.amountRanker(pptx, pptx, { eyebrow: SEG, title: "Records Set in 2026",
  rows: byCnt.map(r => [r.n, cnt(r)]), fmt: v => String(v), color: C.good, ink: C.goodInk, valLabel: "Total",   // shorter head, wider columns (Noah, Oct 26)
  extra: [{ label: "Svc", w: 0.72, vals: Object.fromEntries(byCnt.map(r => [r.n, String(r.svc.length || "–")])) },
          { label: "Parts", w: 0.78, vals: Object.fromEntries(byCnt.map(r => [r.n, String(r.pts.length || "–")])) },
          { label: "Comb.", w: 0.8, vals: Object.fromEntries(byCnt.map(r => [r.n, String(r.tot.length || "–")])) }],
  takeaway: `${rec.length} stores set ${total} all-time monthly records in 2026. ${nm(byCnt[0].n)} and ${nm(byCnt[1].n)} lead with ${cnt(byCnt[0])} each.`,
  tone: "g", source: SRC + " Records = each month a store beat its best-ever service, parts or combined month; one month can count up to three." })
  .addNotes("Records by store (month: service / parts / combined, with the record it beat):\n" + rec.map(r => `${nm(r.n)}: ` + ["svc", "pts", "tot"].flatMap(k => r.E[k].map(e => `${{ svc: "service", pts: "parts", tot: "combined" }[k]} ${e.m} ${H.kusd(e.v)} (beat ${H.kusd(e.prev)}, ${e.prevAt})`)).join("; ")).join("\n"));

await pptx.writeFile({ fileName: "Store Records - Opening.pptx" }); console.log("records opening built", rec.length, ytdRec.length, again.length);
