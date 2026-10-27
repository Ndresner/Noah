/* ===================================================================
   STORE PRESCRIPTIONS - Day 2 Fixed Ops (review, Noah Oct 27 2026)
   Built from the Day 1 handout (store-prescriptions.mjs in the Findlay vault)
   and its build guide: same anatomy, colours, sizes, ranking and rules; Day 2
   fills it with the fixed ops sections. A review file only - nothing goes in
   the deck until Noah says (Oct 27).

   Choices Noah approved Oct 27 ("go with your picks"):
     vitals       Fixed Ops Gross, Fixed Ops Net Profit, Absorption (vs the
                  group), CP ROs - YTD with the change on 2025
     benchmarks   alignment 15, balance 15, brake fluid 10, fuel 7.5 (opcode
                  report basis); MPI 90 (Jul-Sep); labor GP 78 (bottom of the
                  78-80 range); ELR 70% of door rate (Jul-Sep); proficiency 85
     peer avg     hours per RO, CP labor $ per RO, CP parts GP % - the store's
                  segment (luxury / non-luxury, as the deck splits them):
                  segment total / segment divisor
     2025         CP ROs, service gross, parts gross (fewer than 2025)
     group        video completion (no benchmark)
     one price    a missed CP RO = the group's CP labor + parts gross per CP
                  RO ($342); an hour = CP labor + parts gross per sold hour
                  ($183). Both Jan-Sep 2026, same-store.
     a year       Jan-Sep x 12/9; Jul-Sep x 4 (ELR, MPI, video)
     counted      MPI and video (missed inspections), fixed gross per head
                  (heads over the segment average: no store personnel cost)
     left off     warranty (OEM-driven), absorption and net to gross as lines
                  (they overlap gross and cost; absorption is a vital),
                  internal GP, body shop and INEOS
   Opcode lines skip a service the manufacturer doesn't call for at that
   store (Noah, Oct 26-27: balance at the Audi stores; fuel at Volvo, Audi
   Henderson, VW Henderson, Audi Reno).
   Every priced gap of $10K a year or more, largest first; then the counted
   lines. Lines overlap and are ranked, never added.
   =================================================================== */
import fs from "fs";
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const { C, F, sz } = H;
const J = f => JSON.parse(fs.readFileSync(f, "utf8"));
const LOGO = "image/png;base64," + fs.readFileSync("../assets/findlay-logo.png").toString("base64");
const EYEBROW = "Store Prescriptions";
const name = H.storeName;
const LUX = new Set(["LEXUS","LINCOLN","CADILLAC","VOLVO","AUDI RENO","LR RENO","AUDI HENDERSON","LR LV","LR HENDERSON"]);
const SRC = "Qlik - Closed Repair Orders (customer pay, opcodes), Profit & Loss (gross, net profit, absorption), Jan-Sep 2026 vs 2025, "
          + "same-store, INEOS and GMC Prescott body shop out; CP ELR and door rate Jul-Sep; MPI and video store-reported Jul-Sep; "
          + "proficiency Qlik hours flagged / ADP clocked hours; headcount ADP. The numbers on the Day 2 slides.";

const sum = a => a.reduce((t, v) => t + v, 0);
const tie = (label, got, want, tol = 1) => { if (Math.abs(got - want) > tol) throw new Error(`Rx: ${label} = ${got}, expected ${want}`); };
const YEAR = 12 / 9, QUARTER = 4, MIN = 10000;
const K = v => (v < 0 ? "−" : "") + (Math.abs(v) >= 1e6 ? "$" + (Math.abs(v) / 1e6).toFixed(2) + "M" : "$" + Math.round(Math.abs(v) / 1e3) + "K");
const DK = v => (v < 0 ? "−" : "+") + K(Math.abs(v)).replace("−", "");
const P = v => H.Pct(v);                       // v in percent
const Pb = v => (Math.round(v * 10) / 10).toString().replace(/\.0$/, "") + "%";   // a benchmark as set: "15%", "7.5%"
const D1 = v => v.toFixed(1), D2 = v => v.toFixed(2);
const ord = n => n + (["th", "st", "nd", "rd"][(n % 100 - n % 10 != 10) * n % 10] || "th");

/* ---------- the data, each tied to the deck's own figures ---------- */
const M5 = J("m5.json").S;          // [n, 25: ROs, labor sale, hours, labor gross, parts sale, parts gross, 26: same]
const ABS = J("abs.json").S;        // [n, fixed gross 25, fixed exp 25, gross 26, exp 26]
const NS = J("ntgs.json").S, NPp = J("ntgp.json").S;   // [n, profit 25, gross 25, profit 26, gross 26]
const Q3 = J("elr_q3.json"), DOOR = J("doorinfo.json");
const OPP = J("prof/opp.json"), OS = Array.isArray(OPP.S) ? OPP.S : Object.values(OPP.S), MO = OPP.G.months;
const MPI = J("mpi/mpi.json").rows;  // [n, CP+W ROs Jul-Sep, mpi, tool, video, tool, photo, cpw]
const FGE = J("prod/fge.json").S;   // [n, fixed gross 25, head-months 25, gross 26, head-months 26]
const OPR = fs.readFileSync("./prof/op_raw.txt", "utf8").trim().split("\n").map(l => { const p = l.split("|"); return [p[0], ...p.slice(1).map(Number)]; });
const STORES = ABS.map(r => r[0]).sort();
const NAMES = STORES.join("|");
[[M5, "CP"], [NS, "service net"], [NPp, "parts net"], [FGE, "productivity"], [OPR, "opcodes"], [MPI, "MPI"], [OS.map(o => [o.n]), "proficiency"]]
  .forEach(([rows, l]) => { if (rows.map(r => r[0]).sort().join("|") !== NAMES) throw new Error(`Rx: the ${l} rows are not the 33 stores`); });
if (Object.keys(Q3).sort().join("|") !== NAMES) throw new Error("Rx: the Jul-Sep ELR rows are not the 33 stores");
const row = (rows, n) => rows.find(r => r[0] === n);
tie("CP ROs 2026", sum(M5.map(r => r[7])), 279756, 0);
tie("CP ROs 2025", sum(M5.map(r => r[1])), 279300, 0);
// fixed gross = service + parts gross (body shop out); abs.json's gross keeps GMC Prescott body shop for absorption, as the absorption slides do
tie("fixed gross 2026", sum(NS.map(r => r[4])) + sum(NPp.map(r => r[4])), 188799724, 2);
tie("fixed gross 2025", sum(NS.map(r => r[2])) + sum(NPp.map(r => r[2])), 181331157, 2);
tie("absorption gross less the body shop", sum(ABS.map(r => r[3])) - (sum(NS.map(r => r[4])) + sum(NPp.map(r => r[4]))), 868202, 2);
tie("fixed ops net profit 2026", sum(NS.map(r => r[3])) + sum(NPp.map(r => r[3])), 50449895, 2);
STORES.filter(n => n !== "GMC PRESCOTT").forEach(n => tie(`${n} gross = service + parts`, row(NS, n)[4] + row(NPp, n)[4], row(ABS, n)[3], 2));

/* ---------- the group, the segments, the two prices ---------- */
const g26 = k => sum(M5.map(r => r[7 + k]));
const ELR = g26(1) / g26(2);
tie("CP ELR 2026 vs the slide's $162.42", Math.round(ELR * 100), 16242, 0);
const RO_PRICE = (g26(3) + g26(5)) / g26(0), HR_PRICE = (g26(3) + g26(5)) / g26(2);
tie("price of a missed RO", Math.round(RO_PRICE), 342, 0);
tie("price of an hour", Math.round(HR_PRICE), 183, 0);
const ABS_G = sum(ABS.map(r => r[3])) / sum(ABS.map(r => r[4])) * 100;
tie("group absorption vs the slide's 78.1%", Math.round(ABS_G * 10), 781, 0);
const seg = n => LUX.has(n) ? 0 : 1, SEGNAME = ["luxury", "non-luxury"];
const segSum = (i, k) => sum(M5.filter(r => seg(r[0]) === i).map(r => r[7 + k]));
const PEER = [0, 1].map(i => ({
  hpro: segSum(i, 2) / segSum(i, 0),
  lpr: segSum(i, 1) / segSum(i, 0),
  pgp: segSum(i, 5) / segSum(i, 4) * 100,
  fge: (() => { const r = FGE.filter(x => seg(x[0]) === i); return sum(r.map(x => x[3])) / sum(r.map(x => x[4])); })(),
}));
const VID_G = (() => { const r = MPI.filter(x => x[4] != null); return sum(r.map(x => x[1] * x[4])) / sum(r.map(x => x[1])) * 100; })();
tie("group video vs the slide's 57.8%", Math.round(VID_G * 10), 578, 0);
const MPI_G = sum(MPI.map(x => x[1] * x[2])) / sum(MPI.map(x => x[1])) * 100;
tie("group MPI vs the slide's 81.3%", Math.round(MPI_G * 10), 813, 0);
tie("group proficiency vs the slide's 94.7%", Math.round(sum(OS.map(o => o.sold)) / sum(OS.map(o => o.paid)) * 1000), 948, 0);   // 94.78% prints truncated as 94.7%

/* the opcodes, as opmiss-deck.mjs values them: the store's own gross per service (group where it sold under 30), capped at 2x the group */
const SV = [["Alignment", "align", 0.15], ["Tire balance", "bal", 0.15], ["Brake fluid", "bf", 0.10], ["Fuel service", "fuel", 0.075]];
const GRAY = { bal: ["AUDI HENDERSON", "AUDI RENO"], fuel: ["VOLVO", "AUDI HENDERSON", "VW HENDERSON", "AUDI RENO"] };
const OPC = SV.map(([area, key, b], i) => {
  const gpj = sum(OPR.map(r => r[6 + i])) / sum(OPR.map(r => r[2 + i]));
  return { area, key, b, rows: Object.fromEntries(OPR.map(r => {
    const D = r[1], n = r[2 + i], raw = n >= 30 ? r[6 + i] / n : gpj;
    return [r[0], { rate: n / D * 100, short: (b * D - n) * YEAR, val: Math.min(raw, 2 * gpj) }];
  })) };
});

/* ---------- one store ---------- */
function storeOf(n) {
  const m = row(M5, n), a = row(ABS, n), sv = row(NS, n), pt = row(NPp, n), f = row(FGE, n), mp = row(MPI, n), o = OS.find(x => x.n === n);
  const c = k => [m[1 + k], m[7 + k]];   // [2025, 2026]
  const s = { n, seg: seg(n),
    gross: [sv[2] + pt[2], sv[4] + pt[4]], np: [sv[1] + pt[1], sv[3] + pt[3]], abs: [a[1] / a[2] * 100, a[3] / a[4] * 100],
    ros: c(0), hpro: c(2).map((h, y) => h / c(0)[y]), lpr: c(1).map((l, y) => l / c(0)[y]),
    elr: c(1).map((l, y) => l / c(2)[y]), lgp: c(3).map((g, y) => g / c(1)[y] * 100), pgp: c(5).map((g, y) => g / c(4)[y] * 100),
    svcG: [sv[2], sv[4]], ptsG: [pt[2], pt[4]],
    q3: { ls: Q3[n].ls, hrs: Q3[n].hrs, elr: Q3[n].ls / Q3[n].hrs, door: DOOR[n].door },
    mpi: mp[2] * 100, vid: mp[4] == null ? null : mp[4] * 100, cpw: mp[1],
    fge: [f[1] / f[2], f[3] / f[4]], hm: f[4],
    prof: o.prof * 100, techs: o.techs, under: o.under, profOpp: (o.lab + o.parts) / MO * 12,
    m };
  s.gaps = gapsOf(s);
  return s;
}

function gapsOf(s) {
  const out = [], n = s.n, P2 = PEER[s.seg], ro = s.ros[1];
  const add = g => { if ((g.worth == null || g.worth >= MIN) && g.yours !== g.bar && !g.same) out.push(g); };

  // the four opcodes against our benchmarks, each valued as its own slide values it
  OPC.forEach(({ area, key, b, rows }) => {
    if ((GRAY[key] ?? []).includes(n)) return;
    const r = rows[n]; if (r.rate >= b * 100) return;
    add({ key, area, yours: P(r.rate), bar: Pb(b * 100), barNote: "benchmark", same: P(r.rate) === P(b * 100),
      worth: r.short * r.val, rx: `Recommend it on every visit: ${H.Int(r.short)} more a year.`,
      calc: `${area} ${P(r.rate)} vs ${Pb(b * 100)}: ${H.Int(r.short)} more a year (Jan–Sep x 12/9) x ${H.dollars(r.val)} gross a service.` });
  });

  // technician proficiency - the opportunity slide's own figure: techs under 85% raised to 85%
  if (s.profOpp > 0 && s.under > 0) add({ key: "prof", area: "Techs under 85%", yours: `${s.under} of ${s.techs}`, bar: "85%", barNote: "benchmark",
    worth: s.profOpp, rx: s.prof < 85 ? "Dispatch, staging and approvals: flag 85%." : `Lift the ${s.under} under 85%: dispatch, parts.`,
    calc: `Store proficiency ${Math.floor(s.prof)}%. Each tech under 85% raised to 85% of their own clocked hours, valued at the store's labor + parts gross per flagged hour: ${K(s.profOpp / 12)} a month x 12.` });

  // ELR against 70% of door rate, Jul-Sep (the door rate slide's basis)
  const elrBar = 0.70 * s.q3.door;
  if (s.q3.elr < elrBar) add({ key: "elr", area: "CP ELR", yours: H.dollars(s.q3.elr), bar: H.dollars(elrBar), barNote: "70% of door",
    same: H.dollars(s.q3.elr) === H.dollars(elrBar), worth: (elrBar - s.q3.elr) * s.q3.hrs * QUARTER,
    rx: "Charge the posted rate; control discounts.",
    calc: `Jul–Sep CP ELR ${H.dollars(s.q3.elr)} vs 70% of the ${H.dollars(s.q3.door)} door rate = ${H.dollars(elrBar)}; gap x ${H.Int(s.q3.hrs)} CP hours x 4.` });

  // hours per RO against the segment average, an hour at the one price
  if (s.hpro[1] < P2.hpro) add({ key: "hpro", area: "CP hours per RO", yours: D2(s.hpro[1]), bar: D2(P2.hpro), barNote: "peer avg",
    same: D2(s.hpro[1]) === D2(P2.hpro), worth: (P2.hpro - s.hpro[1]) * ro * HR_PRICE * YEAR,
    rx: `Present every MPI finding: +${D2(P2.hpro - s.hpro[1])} hr per RO.`,
    calc: `${D2(s.hpro[1])} vs the ${SEGNAME[s.seg]} ${D2(P2.hpro)}: gap x ${H.Int(ro)} CP ROs x ${H.dollars(HR_PRICE)} an hour x 12/9.` });

  // CP labor $ per RO against the segment average, in gross at the store's own labor GP %
  if (s.lpr[1] < P2.lpr) add({ key: "lpr", area: "CP labor $ per RO", yours: H.dollars(s.lpr[1]), bar: H.dollars(P2.lpr), barNote: "peer avg",
    same: H.dollars(s.lpr[1]) === H.dollars(P2.lpr), worth: (P2.lpr - s.lpr[1]) * ro * s.lgp[1] / 100 * YEAR,
    rx: "Menu and MPI on every RO; track it daily.",
    calc: `${H.dollars(s.lpr[1])} vs the ${SEGNAME[s.seg]} ${H.dollars(P2.lpr)}: gap x ${H.Int(ro)} CP ROs x the store's ${P(s.lgp[1])} labor GP x 12/9. Overlaps hours per RO and ELR.` });

  // CP labor GP % under the bottom of the 78-80% range
  if (s.lgp[1] < 78) add({ key: "lgp", area: "CP labor GP %", yours: P(s.lgp[1]), bar: "78%", barNote: "benchmark", same: P(s.lgp[1]) === P(78),
    worth: (78 - s.lgp[1]) / 100 * s.m[8] * YEAR, rx: "Right tech, right work; review discounts.",
    calc: `${P(s.lgp[1])} vs 78% (bottom of the 78-80% range) x ${K(s.m[8])} CP labor sale x 12/9.` });

  // CP parts GP % against the segment average
  if (s.pgp[1] < P2.pgp) add({ key: "pgp", area: "CP parts GP %", yours: P(s.pgp[1]), bar: P(P2.pgp), barNote: "peer avg", same: P(s.pgp[1]) === P(P2.pgp),
    worth: (P2.pgp - s.pgp[1]) / 100 * s.m[11] * YEAR, rx: "Price to the matrix; watch overrides.",
    calc: `${P(s.pgp[1])} vs the ${SEGNAME[s.seg]} ${P(P2.pgp)} x ${K(s.m[11])} CP parts sale x 12/9.` });

  // CP RO count under 2025, a missed RO at the one price
  if (s.ros[1] < s.ros[0]) add({ key: "ros", area: "CP RO count", yours: H.Int(s.ros[1]), bar: H.Int(s.ros[0]), barNote: "in 2025",
    worth: (s.ros[0] - s.ros[1]) * RO_PRICE * YEAR, rx: `Win back lapsed customers: ${H.Int((s.ros[0] - s.ros[1]) * YEAR)} more ROs.`,
    calc: `${H.Int(s.ros[0] - s.ros[1])} fewer CP ROs than Jan–Sep 2025 x ${H.dollars(RO_PRICE)} a RO x 12/9.` });

  // service and parts gross under 2025
  if (s.svcG[1] < s.svcG[0]) add({ key: "svc", area: "Service gross", yours: K(s.svcG[1]), bar: K(s.svcG[0]), barNote: "in 2025",
    worth: (s.svcG[0] - s.svcG[1]) * YEAR, rx: "Sell the MPI; price menus to market.",
    calc: `Service gross ${H.dollars(s.svcG[1])} vs ${H.dollars(s.svcG[0])} Jan–Sep 2025, x 12/9. Overlaps the CP lines.` });
  if (s.ptsG[1] < s.ptsG[0]) add({ key: "pts", area: "Parts gross", yours: K(s.ptsG[1]), bar: K(s.ptsG[0]), barNote: "in 2025",
    worth: (s.ptsG[0] - s.ptsG[1]) * YEAR, rx: "Parts on every RO; wholesale that pays.",
    calc: `Parts gross ${H.dollars(s.ptsG[1])} vs ${H.dollars(s.ptsG[0])} Jan–Sep 2025, x 12/9. Overlaps CP parts GP.` });

  // counted, not priced: missed inspections a year (Jul-Sep x 4), heads over the segment
  if (s.mpi < 90) { const c = (0.9 - s.mpi / 100) * s.cpw * QUARTER;
    add({ key: "mpi", area: "MPI completion", yours: P(s.mpi), bar: "90%", barNote: "benchmark", same: P(s.mpi) === P(90), worth: null, count: c, short: (90 - s.mpi) / 90,
      rx: `MPI every car: ${H.Int(c)} more a year.`, calc: `${P(s.mpi)} of ${H.Int(s.cpw)} CP & W ROs Jul–Sep vs 90%: ${H.Int(c)} more a year (x 4). Not priced.` }); }
  if (s.vid != null && s.vid < VID_G) { const c = (VID_G - s.vid) / 100 * s.cpw * QUARTER;
    add({ key: "vid", area: "Video completion", yours: P(s.vid), bar: P(VID_G), barNote: "group", worth: null, count: c, short: (VID_G - s.vid) / VID_G,
      rx: `Video on every MPI: ${H.Int(c)} more a year.`, calc: `${P(s.vid)} vs the group's ${P(VID_G)} on ${H.Int(s.cpw)} CP & W ROs Jul–Sep: ${H.Int(c)} more a year (x 4). Not priced.` }); }
  if (s.fge[1] < P2.fge) { const heads = (s.hm - row(FGE, n)[3] / P2.fge) / 9;
    if (heads >= 1) add({ key: "fge", area: "Fixed gross/head", yours: H.dollars(s.fge[1]), bar: H.dollars(P2.fge), barNote: "peer avg", worth: null, count: heads,
      short: (P2.fge - s.fge[1]) / P2.fge, rx: `Staff to gross: ${D1(heads)} heads over peer avg.`,
      calc: `${H.dollars(s.fge[1])} a month (Jan–Sep gross ÷ ${H.Int(s.hm)} head-months, ADP) vs the ${SEGNAME[s.seg]} ${H.dollars(P2.fge)}: ${D1(heads)} heads over. Counted, not priced: no store personnel cost in the data.` }); }

  const priced = out.filter(g => g.worth != null).sort((p, q) => q.worth - p.worth);
  const counted = out.filter(g => g.worth == null).sort((p, q) => q.short - p.short);
  return [...priced, ...counted];
}

/* ---------- where the store leads ---------- */
const ok = { up: v => v > 0 };
const LEADS = [
  { label: "Absorption", v: s => s.abs[1], f: P },
  { label: "Hours per RO", v: s => s.hpro[1], f: D2 },
  { label: "CP ELR", v: s => s.elr[1], f: H.dollars },
  { label: "CP labor $/RO", v: s => s.lpr[1], f: H.dollars },
  { label: "Parts GP", v: s => s.pgp[1], f: P },
  { label: "Proficiency", v: s => s.prof, f: v => Math.floor(v) + "%" },
  { label: "MPI", v: s => s.mpi, f: P },
  { label: "Video", v: s => s.vid, f: P },
  { label: "Gross/head", v: s => s.fge[1], f: H.dollars },
  ...OPC.map(o => ({ label: o.area, v: s => o.rows[s.n].rate, f: P })),
  { label: "Fixed gross", v: s => s.gross[1] - s.gross[0], f: DK, ok: ok.up },
  { label: "Net profit", v: s => s.np[1] - s.np[0], f: DK, ok: ok.up },
  { label: "CP ROs", v: s => s.ros[1] - s.ros[0], f: H.Dint, ok: ok.up },
  { label: "Hours per RO", v: s => s.hpro[1] - s.hpro[0], f: v => "+" + D2(v), ok: ok.up },
  { label: "CP ELR", v: s => s.elr[1] - s.elr[0], f: H.Dmixed, ok: ok.up },
];
function leadsOf(s, all) {
  const places = LEADS.map(m => {
    const v = m.v(s); if (v == null || (m.ok && !m.ok(v))) return null;
    const rank = 1 + all.map(m.v).filter(x => x != null && x > v).length;
    return { rank, text: `${ord(rank)} · ${m.label} ${m.f(v)}` };
  }).filter(Boolean).sort((p, q) => p.rank - q.rank);
  const within = [3, 5, 10].find(k => places.some(p => p.rank <= k));
  return within ? places.filter(p => p.rank <= within).slice(0, 3) : [];
}

/* ---------- one line per cell, checked ---------- */
const EM_REG = 0.53, EM_BOLD = 0.6, EM_BAR = 0.6;
const linesOf = (text, w, fs, em) => {
  const perLine = w / (em * fs / 72); let lines = 1, cur = 0;
  for (const word of String(text).split(/\s+/)) {
    if (word.length > perLine) return Infinity;
    const a = word.length + (cur ? 1 : 0);
    if (cur + a > perLine) { lines++; cur = word.length; } else cur += a;
  }
  return lines;
};
const fit = (text, w, fs, em, why) => { if (linesOf(text, w, fs, em) > 1) throw new Error(`Rx: ${why} will not hold one line at ${fs}pt: "${text}"`); };

/* ---------- the slide (Day 1's geometry) ---------- */
const VIT = { y: 1.32, h: 0.78 }, TOP = VIT.y + VIT.h + 0.1, HEAD = 0.28, ROW_MAX = 0.36, ROW_MIN = 0.25;
const COL = { n: [0.72, 0.34], area: [1.12, 2.05], yours: [3.27, 1.1], bar: [4.47, 2.0], worth: [6.55, 0.95], rx: [7.72, 4.89] };
const firstRow = top => top + 0.08 + HEAD + 0.08;
const rowsFit = top => Math.floor((H.BODY.botFull - 0.06 - firstRow(top)) / ROW_MIN + 1e-9);

function vitals(pptx, sl, s) {
  H.panel(pptx, sl, { x: 0.5, y: VIT.y, w: 12.33, h: VIT.h, label: null });
  const chg = (d, f) => ({ text: f(d), color: Math.abs(d) < 0.5 ? C.ink600 : d > 0 ? C.goodInk : C.badInk, bold: true });
  [
    { label: "Fixed Ops Gross", value: H.Mmixed(s.gross[1]), note: chg(s.gross[1] - s.gross[0], DK) },
    { label: "Fixed Ops Net Profit", value: H.Mmixed(s.np[1]), neg: s.np[1] < 0, note: chg(s.np[1] - s.np[0], DK) },
    { label: "Absorption", value: P(s.abs[1]), note: { text: `group ${P(ABS_G)}`, color: C.ink600 } },
    { label: "CP ROs", value: H.Int(s.ros[1]), note: chg(s.ros[1] - s.ros[0], H.Dint) },
  ].forEach((v, i) => {
    const x = 0.75 + i * 3.02, w = 2.9;
    if (i) sl.addShape(pptx.ShapeType.rect, { x: x - 0.1, y: VIT.y + 0.12, w: 0.012, h: VIT.h - 0.24, fill: { color: C.rule } });
    sl.addText(v.label.toUpperCase(), { x, y: VIT.y + 0.08, w, h: 0.26, fontFace: F, fontSize: sz(14, "vital label"), bold: true,
      color: C.ink600, charSpacing: 0.8, isTextBox: true, margin: 0 });
    if (v.value.length * EM_BOLD * 22 / 72 + (v.note.text.length + 2) * EM_REG * 14 / 72 > w) throw new Error(`Rx: ${s.n} vital ${v.label} will not hold one line`);
    sl.addText([
      { text: v.value, options: { fontSize: sz(22, "vital value"), bold: true, color: v.neg ? C.badInk : C.navy } },
      { text: "  " + v.note.text, options: { fontSize: sz(14, "vital change"), bold: !!v.note.bold, color: v.note.color } },
    ], { x, y: VIT.y + 0.34, w, h: 0.38, fontFace: F, valign: "middle", isTextBox: true, margin: 0 });
  });
}

function leadsTag(pptx, sl, s, all) {
  let lead = leadsOf(s, all);
  const LABEL = lead.every(p => p.rank <= 3) ? "WHERE YOU LEAD" : "YOUR BEST PLACES", SEP = "   ·   ", R = 12.83, PAD = 0.14;
  const titleEnd = 0.5 + name(s.n).length * 0.55 * 28 / 72 + 0.3;
  // the label is bold caps with 0.8pt letter spacing: measured with both, plus slack, so a one-place tag never wraps it (Noah, Oct 27: Lincoln, Subaru LV)
  const widthOf = items => Math.max(LABEL.length * (0.72 * 14 + 0.8) / 72 + 0.15, items.map(p => p.text).join(SEP).length * EM_REG * 14 / 72 + 0.1) + 2 * PAD;
  while (lead.length > 1 && R - widthOf(lead) < titleEnd) lead = lead.slice(0, -1);
  if (!lead.length || R - widthOf(lead) < titleEnd) return [];
  const w = widthOf(lead), x = R - w, y = 0.72, h = 0.52;
  sl.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: C.blueTint } });
  sl.addText([
    { text: LABEL, options: { bold: true, color: C.navy, charSpacing: 0.8, breakLine: true } },
    { text: lead.map(p => p.text).join(SEP), options: { color: C.ink } },
  ], { x: x + PAD, y, w: w - 2 * PAD, h, fontFace: F, fontSize: sz(14, "leads tag"), valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 0.95 });
  return lead;
}

function table(pptx, sl, s, rows, from, top) {
  const bot = H.BODY.botFull;
  H.panel(pptx, sl, { x: 0.5, y: top, w: 12.33, h: bot - top, label: null });
  const HY = top + 0.08;
  [["", "n"], ["AREA", "area"], ["YOURS", "yours"], ["THE BAR", "bar"], ["A YEAR", "worth"], ["PRESCRIPTION", "rx"]].forEach(([t, k]) =>
    sl.addText(t, { x: COL[k][0], y: HY, w: COL[k][1], h: HEAD, fontFace: F, fontSize: sz(14, "table head"), bold: true, color: C.ink600,
      charSpacing: 0.8, align: k === "worth" ? "right" : "left", valign: "middle", isTextBox: true, margin: 0 }));
  sl.addShape(pptx.ShapeType.rect, { x: 0.72, y: HY + HEAD + 0.03, w: 11.89, h: 0.012, fill: { color: C.ink } });
  const y0 = firstRow(top), rowH = Math.min(ROW_MAX, (bot - 0.06 - y0) / Math.max(rows.length, 1));
  if (rowH < ROW_MIN) throw new Error(`Rx: ${s.n} - ${rows.length} rows will not fit one slide`);
  if (!rows.length) sl.addText("No gap against any bar.", { x: COL.area[0], y: y0, w: 9, h: ROW_MAX, fontFace: F, fontSize: sz(16, "no gaps"), color: C.ink600, valign: "middle", isTextBox: true, margin: 0 });
  const cells = g => [[g.area, COL.area[1], EM_BOLD, "area"], [g.yours, COL.yours[1], EM_BOLD, "yours"], [`${g.bar} ${g.barNote}`, COL.bar[1], EM_BAR, "bar"],
    [g.worth == null ? "—" : K(g.worth), COL.worth[1], EM_BOLD, "worth"], [g.rx, COL.rx[1], EM_REG, "prescription"]];
  rows.forEach(g => cells(g).forEach(([t, w, em, k]) => fit(t, w, 14, em, `${s.n} ${g.key} ${k}`)));
  const fs = [16, 15].find(f => rowH >= f * 1.2 / 72 + 0.03 && rows.every(g => cells(g).every(([t, w, em]) => linesOf(t, w, f, em) <= 1))) ?? 14;
  rows.forEach((g, i) => {
    const y = y0 + i * rowH;
    if (i && g.worth == null && rows[i - 1].worth != null) sl.addShape(pptx.ShapeType.rect, { x: 0.72, y: y - 0.006, w: 11.89, h: 0.012, fill: { color: C.ink300 } });
    else if (i % 2) sl.addShape(pptx.ShapeType.rect, { x: 0.72, y, w: 11.89, h: rowH, fill: { color: C.zebra } });
    const cell = (k, text, o) => sl.addText(text, { x: COL[k][0], y, w: COL[k][1], h: rowH, fontFace: F, fontSize: sz(fs, `${k} cell`), isTextBox: true, margin: 0, valign: "middle", ...o });
    cell("n", String(from + i + 1), { bold: true, color: C.blueLt });
    cell("area", g.area, { bold: true, color: C.ink });
    cell("yours", g.yours, { bold: true, color: C.ink });
    cell("bar", [{ text: g.bar, options: { bold: true, color: C.ink600 } }, { text: ` ${g.barNote}`, options: { color: C.ink600 } }]);
    cell("worth", g.worth == null ? "—" : K(g.worth), { bold: g.worth != null, color: g.worth == null ? C.ink300 : C.navy, align: "right" });
    cell("rx", g.rx, { color: C.ink });
  });
}

function rxSlides(pptx, s, all) {
  const first = rowsFit(TOP), more = rowsFit(H.BODY.top);
  const pages = [s.gaps.slice(0, first)];
  for (let i = first; i < s.gaps.length; i += more) pages.push(s.gaps.slice(i, i + more));
  pages.forEach((rows, p) => {
    const sl = pptx.addSlide(), from = p ? first + (p - 1) * more : 0;
    H.page(pptx, sl, { eyebrow: EYEBROW, title: p ? `${name(s.n)} · Continued` : name(s.n) });
    let lead = [];
    if (!p) { vitals(pptx, sl, s); lead = leadsTag(pptx, sl, s, all); }
    table(pptx, sl, s, rows, from, p ? H.BODY.top : TOP);
    sl.addNotes([`Source: ${SRC}`,
      ...(!p ? [`Every priced gap worth ${K(MIN)} a year or more, largest first; then the lines counted, not priced. Lines overlap and are never added. Method: the header of store-rx.mjs.`] : []),
      ...rows.map((g, i) => `${from + i + 1}. ${g.area} - ${g.worth == null ? "not priced" : K(g.worth) + " a year"}. ${g.calc}`),
      ...(lead.length ? ["Where you lead: the store's best places among the 33 on the day's measures, levels and changes on 2025 - top three, else top five, else top ten. A change counts only when it went the right way."] : []),
    ].join("\n\n"));
  });
}

/* ---------- how to read it (Tommy doesn't use notes, so the method is on a slide) ---------- */
function howTo(pptx) {
  const sl = pptx.addSlide();
  H.page(pptx, sl, { eyebrow: EYEBROW, title: "How to Read Your Prescription" });
  const take = "Pick the two lines you will own by the next meeting.";
  H.panel(pptx, sl, { x: 0.5, y: H.BODY.top, w: 12.33, h: H.bodyH(take), label: null });
  const items = [
    ["Your vitals", "2026 year to date, with the change on 2025. Absorption reads against the group."],
    ["The bar", "Our benchmarks; else your peers' average (luxury or non-luxury), the group, or your 2025."],
    ["A year", `Jan–Sep gaps x 12/9; ELR, MPI and video Jul–Sep x 4. A CP RO is worth ${H.dollars(RO_PRICE)}, an hour ${H.dollars(HR_PRICE)}.`],
    ["Opcodes", "Valued at your own gross per service, as the benchmark slides show them."],
    ["Ranked", `Biggest first, down to ${K(MIN)} a year. Lines overlap: rank them, don't add them. A dash means counted, not priced.`],
  ];
  const y0 = H.BODY.top + 0.3, step = (H.bodyBot(take) - y0 - 0.2) / items.length;
  items.forEach(([lead, text], i) => {
    const y = y0 + i * step;
    sl.addShape(pptx.ShapeType.rect, { x: 0.78, y: y + 0.06, w: 0.055, h: step - 0.3, fill: { color: C.blue } });
    sl.addText(lead, { x: 1.0, y, w: 2.4, h: step - 0.2, fontFace: F, fontSize: sz(18, "how lead"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
    sl.addText(text, { x: 3.5, y, w: 9.1, h: step - 0.2, fontFace: F, fontSize: sz(17, "how text"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
  });
  H.takeaway(pptx, sl, take, "n");
  sl.addNotes(`Source: ${SRC}`);
}

const ALL = STORES.map(storeOf);
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
H.cover(pptx, pptx, { title: "Store\nPrescriptions", subtitle: "Day 2 Fixed Operations  ·  October 29, 2026", meta: ["January–September 2026 against 2025", "South Point, Las Vegas"], logo: LOGO });
howTo(pptx);
ALL.forEach(s => rxSlides(pptx, s, ALL));
await pptx.writeFile({ fileName: "Store Prescriptions - Day 2 Review.pptx" });
console.log(`${pptx._slides.length} slides; RO ${H.dollars(RO_PRICE)}, hour ${H.dollars(HR_PRICE)}, video ${P(VID_G)}`);
PEER.forEach((p, i) => console.log(SEGNAME[i], D2(p.hpro), H.dollars(p.lpr), P(p.pgp), H.dollars(p.fge)));
ALL.forEach(s => console.log(`  ${s.n.padEnd(19)} ${String(s.gaps.length).padStart(2)}  ` + s.gaps.map(g => `${g.key} ${g.worth == null ? "#" + H.Int(g.count) : K(g.worth)}`).join(" · ")));
