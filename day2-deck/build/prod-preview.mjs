/* One-slide preview (Noah, Oct 23): What Is Productivity?, an explainer in the same layout as the other sections, to
   replace the People and Pay Plans Are Work Instructions talking points at the start of Productivity. Not merged.
   Figures are the Productivity section's own (ADP headcount, Qlik gross and ROs, Jan-Sep 2026 vs 2025, same-store). */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
H.explainer(pptx, pptx, { source: "Gross and ROs: Qlik - Productivity app. Headcount: ADP Employee Master extract, month-end snapshots, Qlik's fixed-employee title list. Monthly averages, Jan-Sep 2026 vs 2025. Same-store, 33 stores.",
  eyebrow: "Fixed Ops Productivity", title: "What Is Productivity?",
  formula: ["Productivity", "=", "gross and ROs produced", "÷", "the people producing them"], formulaSize: 22,
  define: "How much gross and work each person produces: gross per fixed employee and CP & W ROs per fixed employee.",
  example: [["Gross per employee", "$13,617 / mo"], ["ROs per employee", "26.8 / mo"], ["Fixed ops staff", "1,540"]], exampleLabel: "OUR NUMBERS",
  why: ["People are our most valuable asset and the biggest expense in the department.",
        "Being short-handed costs appointments, MPIs and sold hours.",
        "Pay plans are work instructions: every system is perfectly designed to get the results it gets."],
  levers: [["Staffing", "Every seat filled with someone who adds hours"],
           ["Training", "Invest in people; grow your own apprentices"],
           ["Advisor pay plans", "Paid on CP performance; 50/50 base to bonus"],
           ["Tech pay plans", "Flat rate, paid on their own production"],
           ["Clear expectations", "Job descriptions and the metrics you want to improve"],
           ["Car count", "Fewer cars, fewer hours to produce"]],
  ours: "Group: $13,617 gross per fixed employee a month, up $1,207. Gross grew 4.1% with 5.1% fewer people.",
  tone: "g", notes: "PREVIEW (Noah, Oct 23): replaces People and Pay Plans Are Work Instructions at the start of Productivity. Said: Roundtable slide 6 (People); Jan GM slide 57 (Technician Pipeline); Jan GM slide 16 / May GM slide 12 (Pay Plans aka Work Instructions). Flat-rate technicians are paid for what they produce; hourly technicians are paid whether they produce or not. Staff -76 Sep to Sep to 1,540 (technicians -17, advisors -7, managers -3). Headcount from ADP (Employee Master extract 10/19/2026). Advisor plans: CP labor or gross per RO, maintenance penetration, total sales or gross. CP & W ROs fell 2.3%." });
await pptx.writeFile({ fileName: "../decks/Productivity Explainer - Preview.pptx" }); console.log("preview built");
