import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* REVIEW (Noah, Oct 21 2026): internal GP% should closely mirror customer pay. One slide per measure showing CP,
   internal and the gap (internal - CP) by store, Jan-Sep 2026. Talking point beside the internal GP% slides.
   Option A draws the gap as a bar; option B draws both rates as a dumbbell. Green = gap at or better than the group.
   Data: internal/cp_vs_internal.json (Qlik Closed RO; group CP reproduces the deck's 80.8% / 40.7%). */
const D = JSON.parse(fs.readFileSync("./internal/cp_vs_internal.json", "utf8"));
const SEG = "Warranty & Internal";
const SRC = "Qlik - Closed Repair Orders. Jan-Sep 2026 by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded. Gap = internal GP% minus customer-pay GP%, in points.";
const pct = v => (v * 100).toFixed(1) + "%";
const pts = v => (v >= 0 ? "+" : "−") + Math.abs(v * 100).toFixed(1) + " pts";
const list = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);

const M = [
  { key: "L", title: "Internal vs Customer-Pay Labor Gross Profit %", word: "labor", min: 0.55, max: 0.9 },
  { key: "P", title: "Internal vs Customer-Pay Parts Gross Profit %", word: "parts", min: 0.2, max: 0.5 },
];
// Oct 25 (Noah): in the Internal section, each store against its OWN customer-pay rate, not the group: internal at or
// above its CP rate is green, below is red (CP parts 41% and internal 38% shows as -3.0 pts). Option B (dumbbell) retired.
for (const m of M) {
  const rows = D.rows.map(r => [r.n, r["cp" + m.key], r["in" + m.key]]);
  const gG = D.G["in" + m.key] - D.G["cp" + m.key];
  const g = r => r[2] - r[1];
  const sorted = [...rows].sort((p, q) => g(p) - g(q));
  const below = rows.filter(r => g(r) < 0).length;
  const above = rows.filter(r => g(r) >= 0).map(r => r[0]);
  const widest = sorted.slice(0, 3);
  H.gapRanker(pptx, pptx, { eyebrow: "Internal", title: m.title, mode: "bar", okRef: 0,
    rows, group: [D.G["cp" + m.key], D.G["in" + m.key]], fmt: pct, fmtGap: pts,
    aLabel: "CUSTOMER PAY", bLabel: "INTERNAL", aShort: "CP", bShort: "INT",
    plotLabel: "", min: m.min, max: m.max,
    note: "gap = internal minus own CP",
    takeaway: `${below} of 33 stores collect less ${m.word} margin on internal work than on customer pay; the group is ${pts(gG)}. Widest: ${list(widest.map(r => `${H.storeName(r[0])} (${pts(g(r))})`))}.`,
    source: `${SRC} Green: internal at or above the store's own customer-pay rate; red: below. Dashed line: the group gap (${pts(gG)}). ${above.length ? `At or above customer pay: ${list(above.map(H.storeName))}.` : ""} A gap can be a policy (internal detail, used-car recon pricing) rather than a miss; confirm with the store before calling it out.` });
}
await pptx.writeFile({ fileName: "CP vs Internal GP - Review.pptx" }); console.log("cp vs internal built");
