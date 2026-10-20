/* CP sold hours YOY + one more CP RO a day (2026-10-19, for review). Qlik - Closed Repair Orders, customer pay,
   Jan-Sep by close date, same-store (INEOS excluded), GMC Prescott body shop labor types excluded. Hours reconcile:
   store rows = monthly totals, 504,018.97 (2025) and 523,141.86 (2026). */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
import { load, block, F } from "./blocks.mjs";
const { C } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const SRC = "Qlik - Closed Repair Orders. Customer pay, Jan-Sep by close date. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded.";

// 1-2: CP sold hours, group card + change by store
block(pptx, load("./prof/cphrs.json"), { seg: "The CP Engine", src: SRC, title: "Customer Pay Hours Sold", metric: "cp hours sold",
  fn: a => a[0], f: F.int, level: false,
  def: "Total customer-pay labor hours sold, all CP ROs.",
  extra: () => "CP RO count rose only 0.2%, so the extra hours came from selling more on each visit.",
  chg: (a, b) => b[0] / a[0] - 1, cf: F.pctc.d, verb: "sold more CP hours",
  chgDesc: "customer-pay hours sold (percent, so store size does not drive the ranking)" });

// 3: one more CP RO a day
const HP = JSON.parse(fs.readFileSync("./prof/hpro.json", "utf8"));
/* Days open a year (Noah, Oct 19 2026): every service department is open 6 days a week, never Sundays. 2026 calendar
   less the 6 service holidays (New Year's, Memorial Day, Independence Day, Labor Day, Thanksgiving, Christmas - all fall
   Mon-Sat in 2026): 313 Mon-Sat days - 6 = 307. Not taken from RO close dates, because a store does not close ROs every
   day it is open. */
const DAYS_YR = () => 307;
const R = HP.map(([n, ro, sale, hrs, lg, pg]) => { const days = DAYS_YR(n), gpr = (lg + pg) / ro; return { n, ro, days, perDay: ro * 12 / 9 / days, gpr, yr: gpr * days }; });
const usd = v => "$" + Math.round(v).toLocaleString("en-US");
const tot = R.reduce((t, r) => t + r.yr, 0), gGPR = HP.reduce((t, r) => t + r[4] + r[5], 0) / HP.reduce((t, r) => t + r[1], 0);
const top = [...R].sort((a, b) => b.yr - a.yr);
H.amountRanker(pptx, pptx, {
  eyebrow: "The CP Engine", title: "One More CP RO a Day",
  rows: R.map(r => [r.n, r.yr]), fmt: usd, color: C.blue, ink: C.navy, valLabel: "Gross/yr", valW: 0.98,
  extra: [{ label: "ROs/day", w: 0.86, vals: Object.fromEntries(R.map(r => [r.n, r.perDay.toFixed(0)])) },
          { label: "Per RO", w: 0.8, vals: Object.fromEntries(R.map(r => [r.n, usd(r.gpr)])) }],
  takeaway: `One more customer-pay RO a day at every store is worth ${usd(tot)} a year in labor and parts gross. Each RO averages ${usd(gGPR)}.`,
  tone: "g",
  source: SRC + " Per RO = the store's customer-pay labor gross + parts gross / its CP RO count, Jan-Sep 2026. Days open a year: 307 for every store (open 6 days a week, 2026 calendar, less the 6 service holidays: New Year's, Memorial Day, Independence Day, Labor Day, Thanksgiving, Christmas). ROs/day = Jan-Sep 2026 CP ROs at a yearly pace / days open. Assumes the added RO is an average one.",
});
pptx._slides[pptx._slides.length - 1].addNotes("Talk track: one more customer a day - one more answered call, one appointment kept, one 'close to losing' customer brought back. Retention, phones and the scheduler are the levers.\n" +
  top.map(r => `${r.n}: ${usd(r.gpr)} gross per RO x ${r.days} days open a year = ${usd(r.yr)}; ${r.perDay.toFixed(1)} CP ROs a day today`).join("\n"));
await pptx.writeFile({ fileName: "CP Hours & RO a Day - Review.pptx" }); console.log("built", usd(tot), usd(gGPR), top.slice(0, 3).map(r => r.n + " " + usd(r.yr)).join(", "));
