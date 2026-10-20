/* REBUILT 2026-10-20 on ADP headcount (prod/fge26.txt from hc/build_prod.py). Fixed Gross per Fixed Employee - gap to segment average (2026-10-19, for review). Qlik Productivity app, governed
   measure Fixed Gross per Fixed Employee = P&L Total Gross Profit - Fixed Ops / Fixed Employee headcount, Jan-Sep 2026,
   monthly average. Same-store (INEOS excluded). GMC Prescott: body shop gross ($868,902) and the 4 body shop titles out.
   Group $14,159 (deck slide $14,137 from the Oct 14 pull). Each store compared with its own segment average. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const LUX = new Set(["LEXUS","LINCOLN","CADILLAC","VOLVO","AUDI RENO","LR RENO","AUDI HENDERSON","LR LV","LR HENDERSON"]);
const R = fs.readFileSync("./prod/fge26.txt", "utf8").trim().split("\n").map(l => { const [n, g, e] = l.split("|"); return { n, g: +g, em: +e, lux: LUX.has(n) }; });
const seg = lux => { const S = R.filter(r => r.lux === lux); return S.reduce((t, r) => t + r.g, 0) / S.reduce((t, r) => t + r.em, 0); };
const AVG = { true: seg(true), false: seg(false) };
const usd = v => "$" + Math.round(v).toLocaleString("en-US");
// CHANGED 2026-10-19 (Noah): each store against itself - 5% more gross per fixed employee with the same staff. No
// cross-store average, so brand mix (Audi Henderson vs Lincoln) does not set anyone's target.
const UP = 0.05;
const rows = R.map(r => { const per = r.g / r.em, staff = r.em / 9; return { ...r, per, staff, target: per * (1 + UP), gain: UP * per * staff * 12 }; });
const tot = rows.reduce((t, r) => t + r.gain, 0), top = [...rows].sort((a, b) => b.gain - a.gain);
const grp = R.reduce((t, r) => t + r.g, 0) / R.reduce((t, r) => t + r.em, 0);
const SRC = `Gross: Qlik - Productivity app, P&L Total Gross Profit - Fixed Ops. Headcount: ADP Employee Master extract 10/19/2026 (rebuilt 2026-10-20; the Qlik headcount history undercounted). Fixed gross per fixed employee = gross / fixed employees, Jan-Sep 2026 monthly average. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop gross and body shop staff excluded. Each store is measured against its own number: Target = today + 5%. Gross/yr = 5% x the store's gross per employee x its average fixed employees x 12 (the same staff producing 5% more). Per working day uses 307 days open a year (6 days a week less the 6 service holidays).`;
H.amountRanker(pptx, pptx, {
  eyebrow: "Fixed Ops Productivity", title: "What 5% More per Employee Is Worth",
  rows: rows.map(r => [r.n, r.gain]), fmt: usd, color: C.blue, ink: C.navy, valLabel: "Gross/yr", valW: 1.04,
  extra: [{ label: "Today", w: 0.82, vals: Object.fromEntries(rows.map(r => [r.n, usd(r.per)])) },
          { label: "Target", w: 0.82, vals: Object.fromEntries(rows.map(r => [r.n, usd(r.target)])) }],
  takeaway: `5% more gross from the people already on staff is worth ${usd(tot)} a year. For the average employee that is ${usd(grp * UP)} more a month, about ${usd(grp * UP * 12 / 307)} a working day.`,
  tone: "g", source: SRC + " Today and Target = fixed gross per fixed employee a month.",
});
pptx._slides[0].addNotes("Framing: every store is measured against itself, so an Audi store does not set a Lincoln store's target. 5% is roughly one more customer-pay RO a day for a small store, or a few tenths of an hour per RO.\n" +
  top.map(r => `${r.n}: ${usd(r.per)} today -> ${usd(r.target)} target per employee a month; ${r.staff.toFixed(1)} fixed employees; ${usd(r.gain)} a year`).join("\n"));
await pptx.writeFile({ fileName: "Fixed Gross per Employee 5 Pct - Review.pptx" });
console.log("built", usd(tot), usd(grp*UP), top.slice(0, 3).map(r => r.n + " " + usd(r.gain)).join(", "));
