/* Preview (Noah, Oct 24): the luxury / non-luxury slide without the legend line and the takeaway band, the average
   line stopping at each group, and the right column labelled when non-luxury runs over into it. CP ELR as the example. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, LUX } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const X = load("./m5.json"), f = a => a[1] / a[2], fmt = v => "$" + Math.round(v).toLocaleString("en-US");
const L = X.st.filter(x => LUX.has(x.n)), N = X.st.filter(x => !LUX.has(x.n));
const la = f(X.tot(L.map(x => x.a26))), na = f(X.tot(N.map(x => x.a26)));
H.levelRanker(pptx, pptx, { eyebrow: "Customer Pay", title: "CP ELR", metric: "elr", measure: "CP ELR", fmt, groupLabel: "SEPTEMBER 2026 YTD",
  rows: X.st.map(x => [x.n, f(x.a26)]), group: f(X.G26),
  groups: [{ label: `LUXURY · AVERAGE ${fmt(la)}`, names: new Set(L.map(x => x.n)), ref: la }, { label: `NON-LUXURY · AVERAGE ${fmt(na)}`, names: new Set(N.map(x => x.n)), ref: na }],
  alignRef: true, refLine: true, segBreak: true,
  source: "Qlik - Closed Repair Orders. Customer pay, Jan-Sep 2026. Same-store, INEOS excluded. Segment averages from segment totals." });
await pptx.writeFile({ fileName: "../decks/Luxury Slide - Preview.pptx" }); console.log("preview built");
