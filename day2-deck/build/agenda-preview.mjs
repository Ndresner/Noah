/* One-slide preview (Noah, Oct 23): the agenda with simpler, plain-language section descriptions. Not merged; once
   approved, the details go into SECTIONS in opening-deck.mjs. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
H.agenda(pptx, pptx, { eyebrow: "Day 2 · General Managers & Service Managers", title: "Agenda", items: [
  { name: "Gross", detail: "Fixed ops, service and parts gross performance." },
  { name: "Profitability", detail: "What we keep, and how much of the store we cover." },
  { name: "Customer Pay", detail: "Car count, sales per visit, pricing and margin." },
  { name: "Opcode Benchmarks", detail: "Inspections and key services against benchmark." },
  { name: "Internal", detail: "Internal labor rate and margins." },
  { name: "Productivity", detail: "Staffing, output and technician proficiency." },
  { name: "Warranty", detail: "Warranty rates, margins and compliance." },
  { name: "Close", detail: "What changed, and what we commit to." },
] });
await pptx.writeFile({ fileName: "../decks/Agenda - Preview.pptx" }); console.log("preview built");
