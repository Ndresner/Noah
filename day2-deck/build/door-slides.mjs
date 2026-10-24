import fs from "fs";
import * as H from "./house.mjs";
/* Door-rate option slides for the CP Engine (Jul-Sep 2026 CP ELR). Called from cp-deck.mjs. The option letter is
   in each slide's speaker notes so Noah can pick which to keep. */
export function doorSlides(pptx, SEG, current) {
const I = JSON.parse(fs.readFileSync("./doorinfo.json", "utf8")), Q = JSON.parse(fs.readFileSync("./elr_q3.json", "utf8"));
const st = Object.keys(Q).map(n => ({ n, ls: Q[n].ls, hrs: Q[n].hrs, elr: Q[n].ls / Q[n].hrs, ...I[n] }));
const d2 = v => "$" + Math.round(v).toLocaleString("en-US"),   // whole dollars (Noah, Oct 24)
      p = v => (Math.floor(v * 1000 + 1e-9) / 10).toFixed(1) + "%";
const agg = f => { const r = st.filter(f), h = r.reduce((t, x) => t + x.hrs, 0);
  return { elr: r.reduce((t, x) => t + x.ls, 0) / h, door: r.reduce((t, x) => t + x.door * x.hrs, 0) / h, n: r.length }; };
const ALL = agg(() => true), G = agg(x => x.grid), NG = agg(x => !x.grid);
const vals = (f, extra = {}) => ({ ...Object.fromEntries(st.map(x => [x.n, f(x)])), ...extra });
const at = st.filter(x => x.elr / x.door >= 0.70).length;
const srt = [...st].sort((a, b) => b.elr / b.door - a.elr / a.door), low = srt[srt.length - 1];
const SRC = "Qlik - Closed Repair Orders. CP ELR Jul-Sep 2026 (last three months, to match current door rates). Door rate = standard rate from the store door rate sheet; premium, EV and diesel rates not applied. Same-store, INEOS excluded; GMC Prescott body shop excluded. Averages weighted by Jul-Sep CP sold hours.";
const base = { eyebrow: SEG, title: "CP ELR % of Door Rate", metric: "cp elr % of door rate",
  rows: st.map(x => [x.n, x.elr / x.door]), group: ALL.elr / ALL.door, bench: 0.70, groupLabel: "JULY–SEPTEMBER 2026", fmt: p, floor: 0.4, valLabel: "%" };
const main = `Jul\u2013Sep: ${at} of 33 stores realize at least 70% of door rate; the group is at ${p(ALL.elr / ALL.door)}. ${low.n} is lowest at ${p(low.elr / low.door)}.`;

  // OPTION D / E - prices first
  // Door rate split, no premium columns
H.splitRanker(pptx, pptx, {
  eyebrow: SEG, title: "Door Rate", metric: "door rate", fmt: d2, valLabel: "DOOR", floor: 100,
  sides: [
    { label: "Labor grid", rows: st.filter(x => x.grid).map(x => [x.n, x.door]), avg: G.door, avgLabel: "GRID AVG" },
    { label: "No grid", rows: st.filter(x => !x.grid).map(x => [x.n, x.door]), avg: NG.door, avgLabel: "NO-GRID AVG" },
  ],
  takeaway: `${G.n} stores price labor on a grid, ${NG.n} at a flat door rate. Hours-weighted door rate: ${d2(G.door)} grid, ${d2(NG.door)} no grid.`,
  source: "OPTION D - door rate by store, split by labor grid, standard rate only. Store door rate sheet (Oct 15 2026); averages weighted by Jul-Sep CP sold hours. Bars start at $100. INEOS on the sheet but excluded (same-store).\n"
        + st.filter(x => x.prem).map(x => `${x.n}: premium ${d2(Number(x.prem))} (${x.note})`).join("\n") });
  // Door rate split, with premium columns
const noteTxt = x => (x.note ?? "").replace("Non-Standard Grid", "Non-std").replace("Lambo Non-Grid", "Lambo").replace("Diesel/Corvette", "Diesel");
H.splitRanker(pptx, pptx, {
  eyebrow: SEG, title: "Door Rate", metric: "door rate", fmt: d2, valLabel: "DOOR", floor: 100,
  sides: [
    { label: "Labor grid", rows: st.filter(x => x.grid).map(x => [x.n, x.door]), avg: G.door, avgLabel: "GRID AVG" },
    { label: "No grid", rows: st.filter(x => !x.grid).map(x => [x.n, x.door]), avg: NG.door, avgLabel: "NO-GRID AVG" },
  ],
  extra: [{ label: "PREMIUM", w: 1.08, vals: vals(x => x.prem ? d2(Number(x.prem)) : "") }, { label: "FOR", w: 0.72, vals: vals(noteTxt) }],
  takeaway: `${G.n} stores price labor on a grid, ${NG.n} at a flat door rate. Hours-weighted door rate: ${d2(G.door)} grid, ${d2(NG.door)} no grid.`,
  source: "OPTION E - door rate by store with premium rates, split by labor grid. Store door rate sheet (Oct 15 2026); averages weighted by Jul-Sep CP sold hours. Bars start at $100." });

  // current version (bars + % only)
  current(pptx);
  // OPTION A / B / C - how much of the price stores collect
  // Option A: % + ELR + door
H.levelRanker(pptx, pptx, { ...base, takeaway: main,
  extra: [{ label: "ELR", w: 0.86, vals: vals(x => d2(x.elr), { GROUP: d2(ALL.elr) }) },
          { label: "DOOR", w: 0.86, vals: vals(x => d2(x.door), { GROUP: d2(ALL.door) }) }],
  source: "OPTION A - ELR and door rate columns. " + SRC });
  // Option B: % + ELR + 70% target
H.levelRanker(pptx, pptx, { ...base,
  takeaway: `Jul\u2013Sep: ${at} of 33 stores at or above 70% of door rate. ${low.n} runs ${d2(low.elr)} against a ${d2(low.door * 0.7)} target.`,
  extra: [{ label: "ELR", w: 0.86, vals: vals(x => d2(x.elr), { GROUP: d2(ALL.elr) }) },
          { label: "70% TGT", w: 0.92, vals: vals(x => d2(x.door * 0.7), { GROUP: d2(ALL.door * 0.7) }) }],
  source: "OPTION B - ELR and the ELR needed for 70% of door rate (door x 0.70). " + SRC });
  // Grid split, % + ELR + door
H.splitRanker(pptx, pptx, {
  eyebrow: SEG, title: "CP ELR % of Door Rate", metric: "cp elr % of door rate", fmt: p, valLabel: "%", bench: 0.70, groupLabel: "JULY–SEPTEMBER 2026", floor: 0.4,
  sides: [
    { label: "Labor grid", rows: st.filter(x => x.grid).map(x => [x.n, x.elr / x.door]), avg: G.elr / G.door, avgLabel: "GRID AVG" },
    { label: "No grid", rows: st.filter(x => !x.grid).map(x => [x.n, x.elr / x.door]), avg: NG.elr / NG.door, avgLabel: "NO-GRID AVG" },
  ],
  extra: [{ label: "ELR", w: 0.86, vals: vals(x => d2(x.elr), { "GRID AVG": d2(G.elr), "NO-GRID AVG": d2(NG.elr) }) },
          { label: "DOOR", w: 0.86, vals: vals(x => d2(x.door), { "GRID AVG": d2(G.door), "NO-GRID AVG": d2(NG.door) }) }],
  takeaway: `Jul\u2013Sep: grid stores realize ${p(G.elr / G.door)} of door rate, no-grid stores ${p(NG.elr / NG.door)}. A grid should lift ELR on longer jobs, yet grid stores trail.`,
  source: "OPTION C - split by labor grid. " + SRC });
}
