import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* ADP Tech Employee List export (paid hours, Dec 25 2025 - Oct 9 2026) + Qlik Closed Repair Orders c3efc739 sold hours by
   Division x tech no for the same close dates, all pay types, pulled 2026-10-16. Population and outlier rules from the
   Technician Proficiency workbook skill (qualifies_for_prof_subset). Labelled Jan-Sep at Noah's direction (Oct 16 2026). */
const P = JSON.parse(fs.readFileSync("./prof/prof.json", "utf8"));
const SEG = "Fixed Ops Productivity", TITLE = "Technician Proficiency";
const ST = P.stores.filter(r => r[0] !== "INEOS");
const SRC = "ADP paid hours (regular + overtime) and Qlik - Closed Repair Orders sold hours, all pay types, matched by tech number. Dec 25, 2025 - Oct 9, 2026 (shown as Jan-Sep). Same-store, INEOS excluded (opened 2026). Service, express, team-leader and shop-foreman titles only; body shop, detail and tint excluded. Outliers dropped: shop foremen at or below 30% (workbook uses 20%; under 30% a foreman is mainly dispatching), team leaders at or below 20%, service and express technicians at or below 10% (workbook uses 2%; under 10% a tech is not working as a line tech). Unassigned codes (MULT, 999, 77777) excluded. Role changes: ADP shows one current title with paid hours for the whole period, so when a tech's paid hours exceed 125% of full time over the months they actually flagged, paid hours are scaled to that span (56 techs). 45 keying-error lines removed from sold hours (5,399 hrs: single lines of 20+ hours with under $25 labor sale per hour, e.g. 940.58 hrs for $0 at Hyundai St George, RO 177970).";
const sum = i => ST.reduce((t, r) => t + r[i], 0);
const G = sum(1) / sum(2), N = sum(3), BEN = 0.85;
const pc = v => (Math.floor(v * 1000) / 10).toFixed(1) + "%";   // truncated so just-under never reads as on target
const rows = ST.map(r => [r[0], r[1] / r[2]]);
const sorted = [...rows].sort((a, b) => b[1] - a[1]);
const on = rows.filter(r => r[1] >= BEN).length, over = rows.filter(r => r[1] >= 1).length;

H.section(pptx, pptx, { kicker: "Fixed Ops", title: TITLE, chips: ["Store proficiency", "Technician status", "The opportunity", "Work or workers"] });

H.levelRanker(pptx, pptx, {
  eyebrow: SEG, title: TITLE, metric: "tech proficiency %", rows, group: G, bench: BEN, fmt: pc, floor: 0.4,
  takeaway: `${on} of 33 stores at or above the 85% target and ${over} above 100%. Group ${pc(G)}, led by ${sorted[0][0]} at ${pc(sorted[0][1])}; ${sorted[32][0]} lowest at ${pc(sorted[32][1])}.`,
  source: `${SRC} Proficiency = sold hours / paid hours, from store totals. Green at or above 85%, red below. Bars start at 40%.`,
});

const T = P.tiers, mix = ST.map(r => [r[0], [r[4] / r[3], r[5] / r[3], r[6] / r[3]]]);
H.mixBars(pptx, pptx, {
  eyebrow: SEG, title: TITLE + " Status",
  parts: [{ label: "Above 100", color: H.C.good, dark: true }, { label: "85 to 99.9", color: H.C.blueLt, dark: true }, { label: "Under 85", color: H.C.bad, dark: true }],
  group: [T.above / N, T.on / N, T.under / N], rows: mix,
  takeaway: `Of ${N} technicians, ${T.above} are above 100% and ${T.under} (${(T.under / N * 100).toFixed(0)}%) are under 85%. ${mix.filter(r => r[1][2] > 0.5).length} stores have more than half their technicians under target.`,
  source: `${SRC} Share of each store's technicians by proficiency band (workbook status: Above Target, On Target, Under Target), most above first.`,
});
for (const sl of pptx._slides.slice(1, 3)) sl.addNotes(ST.map(r => `${r[0]}: ${pc(r[1] / r[2])} (${Math.round(r[1]).toLocaleString()} sold / ${Math.round(r[2]).toLocaleString()} paid, ${r[3]} techs: ${r[4]} above, ${r[5]} on, ${r[6]} under)`).join("\n"));

/* ---------- the opportunity (ADDED 2026-10-16) ----------
   Each qualifying tech under 85% lifted to 85% of their own paid hours; no extra hours, no extra staff.
   Gap hours valued at the store's own labor gross, parts gross and labor cost (tech pay) per sold hour,
   all pay types, same dates (Qlik Closed Repair Orders). Per month = period total / 9.49 months. */
const O = JSON.parse(fs.readFileSync("./prof/opp.json", "utf8")), OS = O.S, OG = O.G, MO = OG.months;
const K = v => v >= 1e6 ? "$" + (v / 1e6).toFixed(2) + "M" : "$" + Math.round(v / 1e3).toLocaleString("en-US") + "K";
const tot = o => o.lab + o.parts;
const OSRC = "Each technician under 85% lifted to 85% of their own paid hours; techs at or above 85% count as zero, so no store offsets one tech with another. Gap hours valued at what the store's own technicians (counted titles) earn per sold hour in labor gross and parts gross, all pay types, Qlik - Closed Repair Orders, same dates; tech pay at their labor cost per sold hour. Store-wide averages are not used because they carry detail, rental and recon sale booked with no hours. Per month = period total / 9.49 months (Dec 25, 2025 - Oct 9, 2026). Parts assumes the extra hours carry the store's average parts per hour. Population as on the proficiency slides.";
const perTech = OG.pay / OG.under / MO;
H.statTiles(pptx, pptx, {
  eyebrow: SEG, title: "The Proficiency Opportunity",
  tiles: [
    { label: "Gross missed a month", big: K((OG.lab + OG.parts) / MO), tone: "b",
      sub: `Labor ${K(OG.lab / MO)} + parts ${K(OG.parts / MO)}: ${(OG.share * 100).toFixed(1)}% of the ${K(OG.rog / MO)} a month our ROs generate.` },
    { label: "Tech pay missed a month", big: K(OG.pay / MO), tone: "b",
      sub: `About $${(Math.round(perTech / 10) * 10).toLocaleString("en-US")} a month for each tech under 85%. Same hours, bigger checks.` },
    { label: "Hours not produced, year to date", big: Math.round(OG.gap / 1000) + "K YTD", tone: "b",
      sub: `About ${(OG.gap / MO / 1000).toFixed(1)}K hours a month on the clock and not flagged, below the 85% line alone. ${Math.round(OG.gap / OG.fte)} full-time techs of capacity.` },
    { label: "Techs under 85%", big: `${OG.under} of ${OG.techs}`, tone: "n",
      sub: `${Math.round(OG.under / OG.techs * 100)}% of technicians. Every store has at least one.` },
  ],
  banner: "No extra hours, no extra techs, no extra bays. The techs, bays and lifts are already here; the hour is spent whether it is flagged or not. An hour not produced today cannot be sold tomorrow.",
  takeaway: `Lifting each tech under 85% to 85% is worth ${K((OG.lab + OG.parts) / MO)} a month in gross and ${K(OG.pay / MO)} a month in tech pay.`,
  tone: "b", source: OSRC,
});
const ex1 = Object.fromEntries(OS.map(o => [o.n, `${o.under} of ${o.techs}`]));
const ex2 = Object.fromEntries(OS.map(o => [o.n, pc(o.prof)]));
const top6 = [...OS].sort((a, b) => tot(b) - tot(a)).slice(0, 6), t6 = top6.reduce((t, o) => t + tot(o), 0);
const hidden = OS.filter(o => o.prof >= BEN && tot(o) / MO >= 50000);
H.amountRanker(pptx, pptx, {
  eyebrow: SEG, title: "Proficiency Opportunity by Store",
  rows: OS.map(o => [o.n, tot(o) / MO]), fmt: K, color: H.C.bad, ink: H.C.badInk, valLabel: "Per mo",
  extra: [{ label: "Under 85%", w: 1.02, vals: ex1 }, { label: "Store", w: 0.74, vals: ex2 }],
  takeaway: `Six stores hold ${K(t6 / MO)} a month, led by ${top6[0].n} at ${K(tot(top6[0]) / MO)}. ${hidden.length} stores at or above 85% still leave ${K(hidden.reduce((t, o) => t + tot(o), 0) / MO)} a month.`,
  tone: "b", source: OSRC + " Per mo = labor + parts gross. Under 85% = technicians under 85% of all counted. Store = store proficiency.",
});
// ADDED 2026-10-16 (Noah): the same opportunity as a share of the gross each store's ROs generate
const sh = [...OS].sort((a, b) => b.share - a.share), big = sh.filter(o => o.share >= 0.2);
H.amountRanker(pptx, pptx, {
  eyebrow: SEG, title: "Opportunity as a Share of Gross",
  rows: OS.map(o => [o.n, o.share]), fmt: v => (v * 100).toFixed(1) + "%", color: H.C.bad, ink: H.C.badInk, valLabel: "Share",
  extra: [{ label: "Missed", w: 0.95, vals: Object.fromEntries(OS.map(o => [o.n, K(tot(o) / MO)])) },
          { label: "Gross", w: 0.95, vals: Object.fromEntries(OS.map(o => [o.n, K(o.rog / MO)])) }],
  takeaway: `Group-wide the gap is ${(OG.share * 100).toFixed(1)}% of RO gross. ${big.length} stores are missing a fifth or more, led by ${sh[0].n} at ${(sh[0].share * 100).toFixed(0)}% (${K(tot(sh[0]) / MO)} of ${K(sh[0].rog / MO)} a month).`,
  tone: "b", source: OSRC + " Missed and Gross are per month. Share = missed labor + parts gross / labor + parts gross on the store's closed ROs, all pay types, same dates (includes work on shared tech codes). Not P&L gross: counter and wholesale parts are excluded because techs do not produce them.",
});
const under = OS.filter(o => o.prof < BEN), proven = under.filter(o => o.top >= BEN);
H.dumbbell(pptx, pptx, {
  eyebrow: SEG, title: "Your Best Techs Prove It",
  rows: OS.map(o => [o.n, o.prof, o.top]), bench: BEN, min: 0.5, max: 1.9, fmt: v => Math.round(v * 100) + "%",
  aLabel: "Store proficiency", bLabel: "Top quarter of that store's techs",
  takeaway: `At ${proven.length} of the ${under.length} stores under 85%, the top quarter of techs already run above 85%. Same bays, same dispatch, same car count.`,
  source: `${SRC} Top quarter = the store's most proficient quarter of counted technicians (at least one), sold / paid hours combined. Bars capped at 190%.`,
});
const ws = under.map(o => ({ ...o, avg: o.lux ? OG.rpdL : OG.rpdN }));
const short = ws.filter(o => o.rpd < o.avg), busy = ws.filter(o => o.rpd >= o.avg);
const exR = Object.fromEntries(ws.map(o => [o.n, o.rpd.toFixed(1)])), exA = Object.fromEntries(ws.map(o => [o.n, o.avg.toFixed(1)]));
H.splitRanker(pptx, pptx, {
  eyebrow: SEG, title: "Is It the Work or the Workers?", metric: "tech proficiency %", bench: BEN, floor: 0.4, fmt: pc, valLabel: "Prof.",
  sides: [{ label: "Fewer ROs", rows: short.map(o => [o.n, o.prof]) }, { label: "Enough ROs", rows: busy.map(o => [o.n, o.prof]) }],
  extra: [{ label: "ROs/day", w: 0.92, vals: exR }, { label: "Seg avg", w: 0.92, vals: exA }],
  takeaway: `${short.length} of the ${under.length} stores under 85% also see fewer ROs per tech than their segment: check car count and staffing first. At ${busy.map(o => o.n).join(", ")} the work is there.`,
  source: `${SRC} Stores under 85% only. ROs/day = distinct ROs, all pay types, per counted tech per 8 paid hours. Seg avg = luxury ${OG.rpdL.toFixed(1)}, non-luxury ${OG.rpdN.toFixed(1)} (luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson). Fewer ROs per tech can mean too few cars for the staff or slow techs turning work away; it starts the conversation rather than settling it.`,
});
const notes = OS.map(o => `${o.n}: ${pc(o.prof)}; ${o.under} of ${o.techs} under 85%; gap ${Math.round(o.gap).toLocaleString()} hrs; labor ${K(o.lab / MO)}/mo + parts ${K(o.parts / MO)}/mo; tech pay ${K(o.pay / MO)}/mo; ${(o.share * 100).toFixed(1)}% of ${K(o.rog / MO)}/mo RO gross; top quarter ${Math.round(o.top * 100)}% (${o.topn} techs); ${o.rpd.toFixed(2)} ROs/tech day; ${o.hpro.toFixed(2)} hrs/RO`).join("\n");
pptx._slides.slice(3).forEach(sl => sl.addNotes(notes));
await pptx.writeFile({ fileName: "Technician Proficiency - House Style.pptx" }); console.log("prof built", N, pc(G), on, over);
