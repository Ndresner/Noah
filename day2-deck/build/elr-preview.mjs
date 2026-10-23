/* One-slide preview of the reworked What Is Effective Labor Rate? explainer (Noah, Oct 23). Not merged: once Noah
   approves it, the same content goes into explain-deck.mjs and the deck is rebuilt. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const SRC = "Definitions match the data slides: Qlik - Closed Repair Orders (all customer-pay ROs, by close month), Profit & Loss and Productivity apps, Jan-Sep 2026 vs 2025, same-store, 33 stores.";
H.explainer(pptx, pptx, { source: SRC, eyebrow: "Customer Pay", title: "What Is Effective Labor Rate?",
  formula: ["Effective labor rate", "=", "CP labor sale", "÷", "CP hours sold"],
  define: "The average amount we collect for each hour of customer-pay labor we sell.",
  why: ["Every hour sold is worth the ELR, so it sets the value of the whole shop. Know what you are charging.",
        "It shows whether pricing, discounting and mix are under control.",
        "Benchmark: CP ELR at least 70% of door rate."],
  levers: [["Work mix", "Repair is priced above maintenance: repair-heavy shops run a higher ELR, maintenance-heavy lower. Neither is wrong."],
           ["Maintenance pricing", "If most of your hours are maintenance, the maintenance rate moves ELR more than the repair rate"],
           ["ELR by labor operation", "Know what every op code collects per hour, not just the store total"],
           ["Door rate and labor grid", "Posted rate from a market labor study; a grid prices longer repair jobs higher"],
           ["Discounts", "Every discount lowers ELR. Run them through ADC so we can see them."],
           ["Maintenance labor times", "The hours we set and pay the tech on menu jobs; repair times come from labor guides"]],
  leverRows: [1.25, 1, 1],
  ours: "Group: $162.42 CP ELR, up $3.56. Jul-Sep it ran 74.7% of door rate; 25 of 33 stores at or above 70%.",
  tone: "g", notes: "PREVIEW. Talk track (Noah, Oct 23): to understand your ELR you have to understand your work mix. A shop that does more repair than maintenance will run a higher ELR, because the repair rate is higher; a shop that is mostly maintenance will run lower. There is nothing wrong with that, but you need to know your mix to know why your ELR is what it is. Raising the repair rate may do less than raising the maintenance rate if most of your work is maintenance. Know your ELR for every labor operation." });
await pptx.writeFile({ fileName: "../decks/ELR Explainer - Preview.pptx" }); console.log("preview built");
