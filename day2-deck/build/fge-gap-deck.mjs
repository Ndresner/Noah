/* REBUILT 2026-10-20 on ADP headcount (prod/fge26.txt from hc/build_prod.py). Fixed Gross per Fixed Employee - gap to segment average (2026-10-19, for review; restored with a segment-average
   column at Noah's request). Qlik Productivity app, governed Fixed Gross per Fixed Employee = P&L Total Gross Profit -
   Fixed Ops / Fixed Employee headcount, Jan-Sep 2026 monthly average. Same-store (INEOS excluded). GMC Prescott body
   shop gross ($868,902) and the 4 body shop titles out. Each store compared with its own segment average. */
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
const rows = R.map(r => { const per = r.g / r.em, avg = AVG[r.lux], staff = r.em / 9; return { ...r, per, avg, staff, gap: (avg - per) * staff * 12 }; }).filter(r => r.per < r.avg);
const tot = rows.reduce((t, r) => t + r.gap, 0), top = [...rows].sort((a, b) => b.gap - a.gap);
const SRC = `Gross: Qlik - Productivity app, P&L Total Gross Profit - Fixed Ops. Headcount: ADP Employee Master extract 10/19/2026 (rebuilt 2026-10-20; the Qlik headcount history undercounted). Fixed gross per fixed employee = gross / fixed employees, Jan-Sep 2026 monthly average. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop gross and body shop staff excluded. Each store compared with its own segment average a month: luxury ${usd(AVG.true)}, non-luxury ${usd(AVG.false)} (luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson). Gross/yr = (segment average - store) x the store's average fixed employees x 12. Stores at or above their segment average are not shown.`;
H.amountRanker(pptx, pptx, {
  eyebrow: "Fixed Ops Productivity", title: "Fixed Gross per Fixed Employee",
  rows: rows.map(r => [r.n, r.gap]), fmt: usd, color: C.bad, ink: C.badInk, valLabel: "Gross/yr", valW: 1.12,
  extra: [{ label: "Per emp", w: 0.84, vals: Object.fromEntries(rows.map(r => [r.n, usd(r.per)])) },
          { label: "Seg avg", w: 0.84, vals: Object.fromEntries(rows.map(r => [r.n, usd(r.avg)])) }],
  takeaway: `Luxury stores are compared with ${usd(AVG.true)} a month, non-luxury with ${usd(AVG.false)}. ${rows.length} stores below their average leave ${usd(tot)} a year with the same staff.`,
  tone: "b", source: SRC + " Per emp = the store's fixed gross per fixed employee a month. Seg avg = its segment's average.",
});
pptx._slides[0].addNotes("Framing: gross the team already in the building could produce, not a headcount target. Luxury stores are measured against luxury, non-luxury against non-luxury. Audi Henderson ($23,686) lifts the luxury average; Lincoln ($11,127) sits below even the non-luxury average.\n" +
  top.map(r => `${r.n} (${r.lux ? "luxury" : "non-luxury"}): ${usd(r.per)} per employee a month vs ${usd(r.avg)}; ${r.staff.toFixed(1)} fixed employees; ${usd(r.gap)} a year`).join("\n"));
await pptx.writeFile({ fileName: "Fixed Gross per Employee Gap - Review.pptx" });
console.log("built", rows.length, usd(tot), usd(AVG.true), usd(AVG.false));
