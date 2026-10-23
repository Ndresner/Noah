import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, block, F } from "./blocks.mjs";
import fs from "fs";
import { doorSlides } from "./door-slides.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* Qlik Closed Repair Orders app c3efc739 (Fixed space), pulled 2026-10-14. a = [ROs, labor sale, hours, labor gross, parts sale, parts gross] */
const X = load("./m5.json"), SEG = "Customer Pay";
const SRC = "Qlik - Closed Repair Orders. Customer pay, Jan-Sep by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded.";
const ELR = a => a[1] / a[2];
H.section(pptx, pptx, { kicker: "Fixed Ops", title: SEG, chips: ["ROs", "Hours per RO", "ELR", "$ per RO", "Margin"] });
const base = { seg: SEG, src: SRC };
block(pptx, X, { ...base, title: "CP RO Count", metric: "ro count", fn: a => a[0], f: F.int, level: false,
  def: "Distinct customer-pay ROs: every RO with a customer-pay line, whether or not labor was charged (governed RO Count).",
  extra: () => "Volume is flat, so every gain here comes from selling more per RO.",
  chg: (a, b) => b[0] / a[0] - 1, cf: F.pctc.d, verb: "grew CP RO count", chgDesc: "CP RO count (percent, so store size does not drive the ranking)" });
block(pptx, X, { ...base, title: "CP Hours per RO", metric: "hpro", fn: a => a[2] / a[0], f: F.dec2, segment: true,
  def: "Customer-pay sold hours / all customer-pay ROs.",
  extra: ({ G26 }) => `Each 0.1 hour per RO is worth ${F.usd.v(0.1 * G26[0] * ELR(G26))} of labor at the 2026 ELR.`,
  verb: "sold more hours per RO" });
block(pptx, X, { ...base, title: "CP ELR", metric: "elr", fn: ELR, f: F.usd2, segment: true,
  def: "Customer-pay labor sale / customer-pay sold hours.", verb: "raised CP ELR" });
/* CP ELR as % of door rate - door rates from Store_Door_Rate_Sheet.xlsx (uploaded Oct 15 2026; standard rate,
   premium/EV/diesel rates not used). ELR is JUL-SEP 2026 (last three months), not YTD: the sheet holds today's
   rates and several stores changed rates mid-year (Noah, Oct 15 2026). Group = group labor sale / hours-weighted
   door rate. Benchmark 70% (May deck). Monthly ELR reconciles to the YTD ELR used elsewhere. */
const currentDoor = (pptx) => { const door = JSON.parse(fs.readFileSync("./door.json", "utf8")), Q = JSON.parse(fs.readFileSync("./elr_q3.json", "utf8"));
  const rows = X.D.S.map(s => { const q = Q[s[0]]; return [s[0], (q.ls / q.hrs) / door[s[0]], q.hrs, q.ls / q.hrs, door[s[0]]]; });
  const grp = rows.reduce((t, r) => t + r[3] * r[2], 0) / rows.reduce((t, r) => t + r[4] * r[2], 0);
  const p = v => (Math.floor(v * 1000 + 1e-9) / 10).toFixed(1) + "%";   // truncated, so a store just under 70% never reads 70.0%
  const at = rows.filter(r => r[1] >= 0.70).length, srt = [...rows].sort((a, b) => b[1] - a[1]);
  const under = rows.filter(r => r[1] < 0.70).sort((a, b) => a[1] - b[1]);
  H.levelRanker(pptx, pptx, {
    eyebrow: SEG, title: "CP ELR % of Door Rate", metric: "cp elr % of door rate",
    rows: rows.map(r => [r[0], r[1]]), group: grp, bench: 0.70, groupLabel: "JULY–SEPTEMBER 2026", fmt: p, floor: 0.4,
    takeaway: `Jul\u2013Sep: ${at} of 33 stores realize at least 70% of door rate; the group is at ${p(grp)}. ${srt[0][0]} leads at ${p(srt[0][1])}; ${under[0][0]} is lowest at ${p(under[0][1])}.`,
    source: `CURRENT VERSION - bars and % only. ${SRC} CP ELR for JUL-SEP 2026 (last three months, to match current door rates) / store door rate (standard rate from the store door rate sheet; premium, EV and diesel rates not applied). Group = group CP labor sale / hours-weighted door rate. Green at or above 70%, red below. Bars start at 40%. Percentages are truncated, not rounded.\n`
          + srt.map(r => `${r[0]}: ELR $${r[3].toFixed(2)} / door $${r[4].toFixed(2)} = ${p(r[1])}`).join("\n"),
  });
};
doorSlides(pptx, SEG, currentDoor);
block(pptx, X, { ...base, title: "CP Labor $ per RO", metric: "cp labor per ro", fn: a => a[1] / a[0], f: F.usd, segment: true,
  def: "Customer-pay labor sale / all customer-pay ROs (governed Labor Sale/RO).",
  extra: ({ G25, G26 }) => `ROs were flat, so the ${F.usd.v(G26[1] - G25[1])} labor gain came from selling more per visit.`,
  verb: "grew CP labor per RO" });
block(pptx, X, { ...base, title: "CP Parts $ per RO", metric: "cp parts per ro", fn: a => a[4] / a[0], f: F.usd, segment: true,
  def: "Customer-pay parts sale / all customer-pay ROs (governed Parts Sale/RO).",
  extra: ({ G25, G26 }) => `Parts sale grew ${F.usd.v(G26[4] - G25[4])} on flat RO count.`,
  verb: "grew CP parts per RO" });
block(pptx, X, { ...base, title: "CP Labor Gross Profit %", metric: "cp labor gp %", fn: a => a[3] / a[1], f: F.pct, floor: 0.6,
  def: "Customer-pay labor gross / customer-pay labor sale. Target range 78-80%.",
  extra: () => "The target range is 78\u201380%.", verb: "improved labor margin" });
block(pptx, X, { ...base, title: "CP Parts Gross Profit %", metric: "cp parts gp %", fn: a => a[5] / a[4], f: F.pct, floor: 0.25,
  def: "Customer-pay parts gross / customer-pay parts sale on repair orders (counter and wholesale not included).",
  verb: "improved parts margin" });
await pptx.writeFile({ fileName: "The CP Engine - House Style.pptx" }); console.log("cp built");
