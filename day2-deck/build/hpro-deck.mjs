/* "What 0.1 Hour per RO Is Worth" (2026-10-19, for review). Qlik - Closed Repair Orders, customer pay, all CP ROs,
   Jan-Sep 2026 by close date, same-store (INEOS excluded), GMC Prescott body shop labor types excluded. Reconciles to
   the CP Engine slides: 279,756 ROs, $84,966,723.52 labor sale, 1.87 hrs/RO, $162.42 ELR. Each store valued at its
   OWN CP ELR (labor sale) and its own labor gross per hour; year = Jan-Sep x 12/9. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const R = JSON.parse(fs.readFileSync("./prof/hpro.json", "utf8")).map(([n, ro, sale, hrs, gross, pg]) =>
  ({ n, ro, hpro: hrs / ro, elr: sale / hrs, labYr: 0.1 * ro * (gross / hrs) * 12 / 9, partsYr: 0.1 * ro * (pg / hrs) * 12 / 9 }))
  .map(r => ({ ...r, totYr: r.labYr + r.partsYr }));
const usd = v => "$" + Math.round(v).toLocaleString("en-US");
const S = k => R.reduce((t, r) => t + r[k], 0);
const top = [...R].sort((a, b) => b.totYr - a.totYr);
const SRC = "Qlik - Closed Repair Orders. Customer pay, all CP ROs, Jan-Sep 2026 by close date. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded. Each store: 0.1 x its CP RO count x its own customer-pay labor gross and parts gross per sold hour; a year = Jan-Sep x 12 / 9. Gross, not sale, so it will not match the $4,543,691 labor sale on the Hours per RO slide.";
H.amountRanker(pptx, pptx, {
  eyebrow: "The CP Engine", title: "What 0.1 Hour per RO Is Worth",
  rows: R.map(r => [r.n, r.totYr]), fmt: H.kusd, color: C.blue, ink: C.navy, valLabel: "Gross/yr", valW: 1.0,
  extra: [{ label: "Labor", w: 0.98, vals: Object.fromEntries(R.map(r => [r.n, H.kusd(r.labYr)])) },
          { label: "Parts", w: 0.92, vals: Object.fromEntries(R.map(r => [r.n, H.kusd(r.partsYr)])) }],
  takeaway: `0.1 hour more on every CP RO is worth ${usd(S("totYr"))} a year in gross: ${usd(S("labYr"))} labor and ${usd(S("partsYr"))} parts. ${top[0].n} leads at ${usd(top[0].totYr)}.`,
  tone: "g", source: SRC + " Gross/yr = labor gross + parts gross a year. Parts assumes the extra hours carry the store's average customer-pay parts gross per sold hour.",
});
pptx._slides[0].addNotes("Talk track: 0.1 hour is six minutes of sold time on a visit - one more inspection item sold, one tire rotation, one cabin filter. Parts gross assumes the extra hours carry the store's usual parts.\n" + R.map(r => `${r.n}: ${r.ro.toLocaleString()} CP ROs, ${r.hpro.toFixed(2)} hrs/RO; +0.1 = ${usd(r.totYr / 12)} gross a month, ${usd(r.totYr)} a year (labor ${usd(r.labYr)}, parts ${usd(r.partsYr)})`).join("\n"));
await pptx.writeFile({ fileName: "Hours per RO Value - Review.pptx" }); console.log("built", usd(S("totYr")), usd(S("labYr")), usd(S("partsYr")), usd(S("totYr") / 12));
