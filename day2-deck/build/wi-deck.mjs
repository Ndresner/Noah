import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, block, F } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* Qlik Closed Repair Orders app c3efc739, pulled 2026-10-15. a = [W ROs, W labor sale, W hours, I labor sale, I hours].
   Same-store (INEOS excluded), GMC Prescott body shop labor types excluded. Store rows reconcile on all five fields. */
const X = load("./wi.json"), SEG = "Warranty & Internal";
const SRC = "Qlik - Closed Repair Orders. Jan-Sep by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded.";
H.section(pptx, pptx, { kicker: "Fixed Ops", title: SEG, chips: ["Warranty ROs", "Warranty ELR", "Internal ELR"] });
const base = { seg: SEG, src: SRC };
block(pptx, X, { ...base, title: "Warranty RO Count", metric: "warranty ro count", fn: a => a[0], f: F.int, level: false,
  def: "Distinct warranty ROs: every RO with a warranty line.",
  extra: () => "Warranty volume follows OEM campaigns and recalls as much as store effort.",
  chg: (a, b) => b[0] / a[0] - 1, cf: F.pctc.d, verb: "grew warranty RO count", chgDesc: "warranty RO count (percent, so store size does not drive the ranking)" });
block(pptx, X, { ...base, title: "Warranty ELR", metric: "warranty elr", fn: a => a[1] / a[2], f: F.usd2,
  def: "Warranty labor sale / warranty sold hours.", verb: "raised warranty ELR" });
block(pptx, X, { ...base, title: "Internal ELR", metric: "internal elr", fn: a => a[3] / a[4], f: F.usd2,
  def: "Internal labor sale / internal sold hours. A higher internal rate adds service gross but is a cost to used cars.",
  verb: "raised internal ELR" });
await pptx.writeFile({ fileName: "Warranty & Internal - House Style.pptx" }); console.log("wi built");
