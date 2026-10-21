import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, block, F } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* REVIEW (Noah, Oct 21 2026): internal labor and parts gross profit % by store, for the Warranty & Internal section.
   Qlik Closed Repair Orders app c3efc739, pulled 2026-10-21. a = [I ROs, I labor sale, I hours, I labor gross, I parts sale, I parts gross].
   Same set as the CP GP% slides (validated: CP labor 80.9/80.8%, CP parts 39.4/40.7%, internal ELR $126.09/$129.84).
   Colour: green at or above the group, red below (Noah). Store rows reconcile on every $ and hours input, both years. */
const X = load("./internal/internal.json"), SEG = "Warranty & Internal";
const SRC = "Qlik - Closed Repair Orders. Internal pay type, Jan-Sep by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded.";
const base = { seg: SEG, src: SRC };
block(pptx, X, { ...base, title: "Internal Labor Gross Profit %", metric: "internal labor gp %", fn: a => a[3] / a[1], f: F.pct, floor: 0.5,
  def: "Internal labor gross / internal labor sale. A higher internal margin adds service gross but is a cost to used cars.",
  verb: "raised internal labor GP%" });
block(pptx, X, { ...base, title: "Internal Parts Gross Profit %", metric: "internal parts gp %", fn: a => a[5] / a[4], f: F.pct, floor: 0.2,
  def: "Internal parts gross / internal parts sale on repair orders. A higher internal margin adds parts gross but is a cost to used cars.",
  verb: "raised internal parts GP%" });
await pptx.writeFile({ fileName: "Internal Gross Profit - Review.pptx" }); console.log("internal gp built");
