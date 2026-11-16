import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, block, F } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* Qlik P&L app bc30d66f, pulled 2026-10-15. Jan-Sep, RangeSum per store (null trap avoided), signs flipped in Qlik.
   Absorption uses the P&L's own definition: line 257 "Fixed Ops Absorption %" = line 254 Total Gross Profit - Fixed Ops
   / line 258 Fixed Expenses (the stored line returns a dollar artifact, so it is derived). Net to gross = Detail statement
   DEPARTMENT PROFIT OR LOSS / GROSS PROFIT / INCOME, Mechanical Dept (service) and P&A Dept (parts). Same-store, INEOS
   excluded. Store rows reconcile to the group to the dollar on every line, both years. */
const SEG = "Fixed Ops Profitability";
const SRC = "Qlik - Profit & Loss. Jan-Sep. Same-store, INEOS excluded (opened 2026).";
const R = a => a[0] / a[1];
H.section(pptx, pptx, { kicker: "Fixed Ops", title: "Profitability", chips: ["Absorption", "Service net to gross", "Parts net to gross"] });
const base = { seg: SEG, src: SRC };
const pc = v => (v * 100).toFixed(1) + "%";
block(pptx, load("./abs.json"), { ...base, title: "Fixed Ops Absorption %", metric: "fixed ops absorption %", fn: R, f: F.pct, floor: 0.3,
  def: "P&L line Fixed Ops Absorption % = Total Gross Profit - Fixed Ops / Fixed Expenses. GMC Prescott includes its body shop, as on the P&L (Noah, Oct 28: kept in).",
  extra: ({ G25, G26 }) => `Fixed gross grew ${pc(G26[0] / G25[0] - 1)}; fixed expenses ${pc(G26[1] / G25[1] - 1)}.`,
  verb: "raised absorption" });
block(pptx, load("./ntgs.json"), { ...base, title: "Service Net to Gross %", metric: "service net to gross %", fn: R, f: F.pct, floor: -0.45,
  def: "Mechanical department profit / mechanical department gross (P&L Detail statement). Body shop is its own department and is not included.",
  extra: ({ G25, G26 }) => `Department profit grew ${pc(G26[0] / G25[0] - 1)} on ${pc(G26[1] / G25[1] - 1)} more gross.`,
  verb: "raised service net to gross" });
block(pptx, load("./ntgp.json"), { ...base, title: "Parts Net to Gross %", metric: "parts net to gross %", fn: R, f: F.pct, floor: 0,
  def: "P&A department profit / P&A department gross (P&L Detail statement).",
  extra: ({ G25, G26 }) => `Gross grew ${pc(G26[1] / G25[1] - 1)} but department profit ${G26[0] >= G25[0] ? "grew" : "fell"} ${pc(Math.abs(G26[0] / G25[0] - 1))}.`,
  verb: "raised parts net to gross" });
await pptx.writeFile({ fileName: "Fixed Ops Profitability - House Style.pptx" }); console.log("pl built");
