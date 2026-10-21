import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* REVIEW (Noah, Oct 21 2026): store records options. Data: records/records.json from records/compute.py
   (Qlik Fixed Ops - Accting TB, pulled Oct 8 2026, Jan 2021 - Sep 2026). Slides:
   1 A Record Year (group)  2 Record Board  3 Spotlight by %  4 Spotlight by $  5 Most monthly records (ranker). */
const R = JSON.parse(fs.readFileSync("./records/records.json", "utf8"));
const SEG = "Store Records";
const SRC = "Qlik - Fixed Ops Accounting TB (Family FIXED GROSS, service + parts, body shop excluded), pulled Oct 8 2026. History Jan 2021 - Sep 2026; months before a store opened are ignored. Same-store, INEOS excluded. Matches the P&L for every store except GMC PRESCOTT service (TB higher by $153,746 Jan-Sep 2026, $286,430 Jan-Sep 2025; GMC PRESCOTT's combined record holds on both).";
const M = v => "$" + (v / 1e6).toFixed(2) + "M";
const usd = v => "$" + Math.round(v).toLocaleString("en-US");
const tp = x => (Math.trunc(x * 1000) / 10).toFixed(1) + "%";          // truncated, not rounded (improvement slides)
const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
const FULL = { Jan: "January", Feb: "February", Mar: "March", Apr: "April", May: "May", Jun: "June", Jul: "July", Aug: "August", Sep: "September", Oct: "October", Nov: "November", Dec: "December" };
const CAT = { svc: "Service", pts: "Parts", tot: "Combined" };

const stores = Object.entries(R.ytd).map(([n, d]) => {
  const r = d.rec, on = c => !!r[c]?.record;
  return { n, r, on, count: ["svc", "pts", "tot"].filter(on).length, jump: r.tot.v / r.tot.prev - 1, add: r.tot.v - r.tot.prev, short: d.years[0] > 2021 };
});
const setters = stores.filter(s => s.count > 0).sort((a, b) => b.count - a.count || b.jump - a.jump);
const triple = setters.filter(s => s.count === 3);
const combined = stores.filter(s => s.on("tot"));

// ---- 1. A Record Year (group)
{
  const yrs = Object.entries(R.group.janSep).map(([y, v]) => [y, v]);
  const grew = yrs.every((r, i) => i === 0 || r[1] > yrs[i - 1][1]);
  const [bm, bv] = R.group.topMonths[0];
  const top7 = R.group.topMonths.slice(0, 7).filter(m => m[0].endsWith("2026")).length;
  H.yearBars(pptx, pptx, { eyebrow: `${SEG} · A Record Year`, title: "Total Gross Profit - Fixed Ops",
    panelLabel: "Group fixed gross, January–September, by year",
    years: yrs.map(([y, v]) => [y, v]),
    bullets: [
      `Best month ever: ${FULL[bm.split(" ")[0]]} ${bm.split(" ")[1]}, ${M(bv)}.`,
      `${top7} of the group's 7 best months came in 2026.`,
      `${combined.length} of 33 stores had their best January–September ever.`,
    ],
    takeaway: `January–September fixed gross ${grew ? "has grown every year since 2021" : "is the best since 2021"}, to ${M(yrs.at(-1)[1])} in 2026.`,
    source: `${SRC} Group = sum of 33 stores. TB group total differs from the P&L figure used elsewhere in the deck ($188.80M) by the GMC PRESCOTT difference above.` });
}

// ---- 2. Record Board
H.recordBoard(pptx, pptx, { eyebrow: `${SEG} · Record Board`, title: "Best January–September Ever",
  markLabels: ["SVC", "PTS", "TOT"], valLabel: "JAN\u2013SEP", gainLabel: "VS BEST",
  rows: setters.map(s => ({ name: s.n, marks: [s.on("svc"), s.on("pts"), s.on("tot")],
    val: s.on("tot") ? M(s.r.tot.v) : "", gain: s.on("tot") ? "+" + tp(s.jump) : "" })),
  takeaway: `${setters.length} of 33 stores set a January–September record; ${triple.length} set it in service, parts and total.`,
  source: `${SRC} Record = Jan-Sep 2026 above every prior Jan-Sep for that store. % = combined gain over the previous best Jan-Sep. Stores with shorter histories count the same: ${stores.filter(s => s.short).map(s => `${s.n} (from ${R.ytd[s.n].years[0]})`).join(", ")}. Order: records in all three first, then by % gain.` });

// ---- 3 & 4. Spotlight by % and by $
function spotlight(label, pick, big, sub, why) {
  const top = [...combined].sort(pick).slice(0, 4);
  H.statTiles(pptx, pptx, { eyebrow: `${SEG} · Spotlight ${label}`, title: "Best January–September Ever",
    tiles: top.map(s => ({ label: s.n, big: big(s), sub: sub(s), tone: "g" })),
    banner: "What changed? Each of these stores tells the room one thing they did differently this year.",
    takeaway: `${list(top.map(s => s.n))} beat their previous best by the most ${why}.`, tone: "g",
    source: `${SRC} Combined service + parts, Jan-Sep 2026 vs the store's best prior Jan-Sep.` });
}
spotlight("by %", (a, b) => b.jump - a.jump, s => "+" + tp(s.jump),
  s => `${M(s.r.tot.v)} vs ${M(s.r.tot.prev)}, the previous best (${s.r.tot.prevYear}).`, "in percentage terms");
spotlight("by $", (a, b) => b.add - a.add, s => "+" + M(s.add),
  s => `${M(s.r.tot.v)} vs ${M(s.r.tot.prev)}, the previous best (${s.r.tot.prevYear}); +${tp(s.jump)}.`, "in dollars");

// ---- 5. Most monthly records broken in 2026
{
  const rows = Object.entries(R.broken).map(([n, b]) => [n, b.svc.length + b.pts.length + b.tot.length]).filter(r => r[1] > 0);
  rows.sort((a, b) => b[1] - a[1]);
  const lead = rows.filter(r => r[1] === rows[0][1]).map(r => r[0]);
  H.amountRanker(pptx, pptx, { eyebrow: `${SEG} · Option: Ranker`, title: "Monthly Gross Records Set in 2026",
    rows, fmt: v => String(v), valLabel: "RECORDS", valW: 1.15, color: H.C.good,
    takeaway: `${rows.length} of 33 stores set a new monthly record in 2026; ${list(lead)} set the most, ${rows[0][1]}${lead.length > 1 ? " each" : ""}.`, tone: "g",
    source: `${SRC} Count of times a store set a new all-time monthly high in service, parts or combined gross, Jan-Sep 2026. Each new high counts, so a record beaten twice counts twice.` });
}

// ---- 6. Same ranker, split by category (Noah, Oct 21): service, parts and combined counts
{
  const K = ["svc", "pts", "tot"], WORD = { svc: "service", pts: "parts", tot: "combined" };
  const rows = Object.entries(R.broken).map(([n, b]) => [n, K.map(c => b[c].length)]).filter(r => r[1].some(x => x > 0));
  const sum = k => rows.reduce((t, r) => t + r[1][k], 0);
  const all = sum(0) + sum(1) + sum(2);
  const times = n => n === 1 ? "once" : n === 2 ? "twice" : `${n} times`;
  const notes = rows.map(([n]) => n).sort().map(n => {
    const b = R.broken[n];
    return `${n}: ` + K.filter(c => b[c].length).map(c => `${WORD[c]} record reset ${times(b[c].length)} (${b[c].map(h => h.m.slice(0, 3)).join(", ")})`).join("; ") + ".";
  }).join("\n");
  H.stackedCountRanker(pptx, pptx, { eyebrow: `${SEG} · Option: Ranker by Category`, title: "Monthly Gross Records Set in 2026",
    parts: ["SVC", "PTS", "TOT"], legend: ["Service", "Parts", "Combined"], totalLabel: "ALL", rows,
    takeaway: `${rows.length} of 33 stores set ${all} new monthly records in 2026: ${sum(0)} in service, ${sum(1)} in parts and ${sum(2)} in combined gross.`,
    notes: `Month each record was set, by store:\n${notes}`,
    source: `${SRC} Count of times a store set a new all-time monthly high, Jan-Sep 2026. SVC = service gross, PTS = parts gross, TOT = service + parts combined. Each new high counts, so a record beaten twice counts twice.` });
}

await pptx.writeFile({ fileName: "Store Records - Review.pptx" }); console.log("records review built");
