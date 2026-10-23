/* One-slide preview of the reworked Pay Plans Are Work Instructions talking point (Noah, Oct 23). Not merged: once
   approved, the same content goes into tp-deck.mjs and the deck is rebuilt. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
H.talkingPoints(pptx, pptx, { source: "Talking points: GM meeting decks (Jan and May 2026). Data: Qlik Closed Repair Orders and Productivity apps; ADP headcount. Same-store, 33 stores, Jan-Sep 2026 vs 2025.",
  eyebrow: "Fixed Ops Productivity", title: "Pay Plans Are Work Instructions", saidFrom: null,
  said: ["Clear and easily understood",
         "Focus on the metrics that matter most and the ones you want to improve",
         "Service advisors: pay on customer-pay performance: CP labor or gross per RO, maintenance penetration, total sales or gross",
         "Service advisors: base versus bonus, a 50/50 ratio",
         "Technicians: flat rate, paid on their own production. Flat-rate techs are more proficient because they are paid for what they produce, not for hours on the clock.",
         "Every system is perfectly designed to get the results it gets"],
  now: [{ big: "$304", label: "CP labor per RO, the number the plans should pay on. Up $17.", tone: "g" },
        { big: "166", label: "CP & W ROs per advisor a month, down 7. Advisors grew 1.5%; ROs fell 2.3%.", tone: "b" }],
  ask: "check that every advisor and tech pay plan pays for the numbers on these slides.",
  notes: "PREVIEW (Noah, Oct 23). Two groups of productive employees move gross: service advisors (paid on customer-pay performance, 50/50 base to bonus) and technicians (flat rate on their own production; hourly techs are paid whether they produce or not). \"Drive the results you are looking for\" reworded to \"Focus on the metrics that matter most and the ones you want to improve.\"" });
await pptx.writeFile({ fileName: "../decks/Pay Plans - Preview.pptx" }); console.log("preview built");
