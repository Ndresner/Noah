import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { build as gross } from "./gross-mod.mjs";
import { build as pl } from "./pl-mod.mjs";
import { build as cp } from "./cp-mod.mjs";
import { build as mix } from "./labormix-mod.mjs";
import { build as op } from "./opcode-mod.mjs";
import { build as wi } from "./wi-mod.mjs";
import { build as prod } from "./prod-mod.mjs";
/* MASTER REVIEW DECK - every segment in day order, one file. Built from the same segment scripts as the standalone
   decks; section breaks carry the day tracker. Review draft for Noah and Tray, Oct 15 2026. */
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const SEGMENTS = ["Gross", "Profitability", "CP Engine", "Labor Mix", "Opcodes", "Warranty", "Productivity"];
const cov = H.cover(pptx, pptx, {
  title: "Fixed Operations\nBusiness Update",
  subtitle: "GM & Service Manager Meeting  \u00b7  Day 2  \u00b7  October 29, 2026",
  meta: ["Review draft", "Data through September 2026", "Same-store, 33 stores"],
  segments: SEGMENTS,
});
cov.addNotes("Review draft for Noah and Tray. Segment order follows the day: scoreboard (gross, profitability), the CP engine (ROs > hours/RO > ELR > $/RO > margin), labor mix, selling the work (opcodes), warranty & internal, productivity. Door-rate slides in the CP Engine include every option (A-E plus the current version) for selection - option letter is the first line of each slide's notes.");
const parts = [gross, pl, cp, mix, op, wi, prod];
for (let i = 0; i < parts.length; i++) { globalThis.__TRACK = { segments: SEGMENTS, current: i }; await parts[i](pptx); }
delete globalThis.__TRACK;
await pptx.writeFile({ fileName: "Fixed Ops - Master Review Deck.pptx" });
console.log("master built");
