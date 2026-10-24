/* One-slide preview (Noah, Oct 24): Why We Are Here with a group-level view of what fixed ops is worth to the
   dealership, instead of department numbers that come later in the deck. Data: pulls/net_by_dept.json. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const D = JSON.parse(fs.readFileSync("./pulls/net_by_dept.json", "utf8")).all_stores;   // whole group, INEOS in (Noah, Oct 24)
const fx = y => -(D[y].service + D[y].parts), net = y => -D[y].net_gm;
const M = v => "$" + (v / 1e6).toFixed(2) + "M", usd = v => "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
const share = y => fx(y) / net(y) * 100, dFx = fx("y26") - fx("y25"), dNet = net("y26") - net("y25");
H.talkingPoints(pptx, pptx, { source: "Qlik - Profit & Loss, Jan-Sep 2026 vs 2025, all stores (INEOS included, to match the group P&L). Fixed ops = service (Mechanical) + parts (P&A) department profit: department gross less the department's own expenses. Group net = Net Profit Or Loss on the GM statement.",
  eyebrow: "Fixed Ops", title: "Why We Are Here", saidFrom: null,
  said: ["Foster collaboration, introduce ideas that drive operational excellence, and set a clear path to exceptional service performance",
         "Focus on profit-generating activities: the work that brings in incremental dollars, not just the work that fills the day",
         "Work on the business, not just in it: step out of the daily fires long enough to fix the process behind them",
         "No risk of judgement: your success is our success",
         "Customer first, company second, team third",
         "You get what you tolerate",
         "Leverage the power in this room: collaboration"],
  now: [{ big: `${share("y26").toFixed(0)}%`, label: `of the group's net profit came from service and parts: ${M(fx("y26"))} of ${M(net("y26"))} (${share("y25").toFixed(0)}% in 2025).`, tone: "g" },
        { big: M(fx("y26")), label: `Fixed ops department profit. New, used, F&I and the rest of the store together: ${M(net("y26") - fx("y26"))}.`, tone: "n" },
        { big: "+" + usd(dFx), label: `More fixed ops profit than 2025, while group net profit ${dNet < 0 ? "fell" : "rose"} ${M(Math.abs(dNet))}.`, tone: "g" }],
  ask: "leave today with one or two things you will put in place before year end, and tell the room what they are.",
  notes: `PREVIEW (Noah, Oct 24). Group-level view of fixed ops' weight in the dealership's profit, Jan-Sep. Fixed ops department profit ${usd(fx("y26"))} (2025 ${usd(fx("y25"))}); group net profit ${usd(net("y26"))} (2025 ${usd(net("y25"))}). On the P&L, F&I income, other income and store-level costs are not split by department, so 'the rest of the store' is the group net less fixed ops department profit. Group net = Net Profit Or Loss on the GM statement (P&L summary), all stores; INEOS (opened 2026) is in both lines. Body shop (GMC Prescott) department profit is left out of fixed ops, as everywhere in the deck.` });
await pptx.writeFile({ fileName: "../decks/Why We Are Here - Preview.pptx" }); console.log("preview built");
