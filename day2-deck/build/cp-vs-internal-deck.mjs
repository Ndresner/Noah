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
for (const [mode, opt] of [["bar", "Option A"], ["dumbbell", "Option B"]]) {
  for (const m of M) {
    const rows = D.rows.map(r => [r.n, r["cp" + m.key], r["in" + m.key]]);
    const gG = D.G["in" + m.key] - D.G["cp" + m.key];
    const g = r => r[2] - r[1];
    const sorted = [...rows].sort((p, q) => g(p) - g(q));
    const atOrBetter = rows.filter(r => g(r) >= gG - 1e-9).length;
    const above = rows.filter(r => g(r) >= 0).map(r => r[0]);
    const widest = sorted.slice(0, 3);
    H.gapRanker(pptx, pptx, { eyebrow: `${SEG} · ${opt}`, title: m.title, mode,
      rows, group: [D.G["cp" + m.key], D.G["in" + m.key]], fmt: pct, fmtGap: pts,
      aLabel: "CUSTOMER PAY", bLabel: "INTERNAL", aShort: "CP", bShort: "INT",
      plotLabel: "", min: m.min, max: m.max,
      note: mode === "bar" ? "green = gap at or better than the group" : "navy dot = CP, green/red dot = internal",
      takeaway: `Internal ${m.word} margin runs ${pts(gG).replace(/^[+−]/, "")} ${gG < 0 ? "below" : "above"} customer pay across the group; ${atOrBetter} of 33 stores are at or better than that gap. Widest: ${list(widest.map(r => `${r[0]} (${pts(g(r))})`))}.`,
      source: `${SRC} Green at or better than the group gap (${pts(gG)}), red below. ${above.length} stores run internal at or above customer pay: ${list(above)}. A wide gap can be a policy (internal detail, used-car recon pricing) rather than a miss; confirm with the store before calling it out.` });
  }
}
await pptx.writeFile({ fileName: "CP vs Internal GP - Review.pptx" }); console.log("cp vs internal built");
