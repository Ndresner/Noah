import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
const { C, F, sz, words } = H;
/* Warranty audit slides (DRAFT, Oct 22 2026, Noah: a warranty-audit focus in the Warranty section, more source
   material to come from Noah). Content: Jan 2026 GM meeting Day 1, slides 13-14 (reference/). Stores are not named:
   the Warranty Compliance talking point already keeps the audit total without store call-outs from the podium.
   The one-line descriptions under each cause and step are drafted from the Jan slide's headings - confirm them
   against OEM policy before the meeting. */
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const SEG = "Warranty & Internal";
const SRC = "Jan 2026 GM meeting (Day 1), slides 13-14: six warranty audits, primary causes and Warranty Compliance steps. Descriptions under each heading drafted for this deck; confirm against each OEM's warranty policy manual.";

function tiles(s, items, { cols, top, bot, accent, headH = 0.4 }) {   // headH: room for a heading that wraps
  const rows = Math.ceil(items.length / cols), gap = 0.18, W = (12.33 - 0.5 - gap * (cols - 1)) / cols, Ht = (bot - top - gap * (rows - 1)) / rows;
  items.forEach(([head, line], i) => {
    const x = 0.75 + (i % cols) * (W + gap), y = top + Math.floor(i / cols) * (Ht + gap);
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: Ht, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.06, h: Ht, fill: { color: accent } });
    s.addText(words(head, "tile head"), { x: x + 0.22, y: y + 0.14, w: W - 0.36, h: headH, fontFace: F, fontSize: sz(18, "tile head"), bold: true,
      color: C.ink, valign: "top", isTextBox: true, margin: 0 });
    s.addText(words(line, "tile line"), { x: x + 0.22, y: y + 0.24 + headH, w: W - 0.36, h: Ht - 0.36 - headH, fontFace: F, fontSize: sz(15, "tile line"),
      color: C.ink600, valign: "top", isTextBox: true, margin: 0 });
  });
}

/* 1. what the auditor debits */
{
  const s = pptx.addSlide(), tk = "Six audits cost over $400,000. Every cause on this list is documentation the store controls.";
  H.page(pptx, s, { eyebrow: SEG, title: "What Gets Claims Debited" });
  H.panel(pptx, s, { x: 0.5, y: H.BODY.top, w: 12.33, h: H.bodyH(tk), label: "Primary causes in the six audits" });
  tiles(s, [
    ["Written documentation", "Complaint, cause and correction written on the RO, and matching what was claimed"],
    ["Manager signature on add-ons", "Any added repair approved and signed by a manager before it is claimed"],
    ["Signatures on in-stock units", "Customer and manager signatures on repairs to units in stock"],
    ["Out mileage", "Recorded on every RO and consistent with the mileage in"],
    ["Punch times", "Technician clock times on the RO support the hours claimed"],
    ["Manufacturer authorization", "Prior approval on file wherever the manufacturer requires it"],
  ], { cols: 3, top: H.BODY.top + 0.72, bot: H.bodyBot(tk) - 0.2, accent: C.bad, headH: 0.7 });
  H.takeaway(pptx, s, tk, "b");
  H.foot(pptx, s, SRC);
  s.addNotes("DRAFT. Causes from Jan 2026 GM meeting slide 13. That slide named six stores and their debits; stores are left off here, as on Warranty Compliance. The store amounts on the Jan slide add to $414,000 (headline said $400,000), so the slide says 'over $400,000'.");
}

/* 2. how a store keeps the claim paid */
{
  const s = pptx.addSlide(), tk = "The General Manager is responsible for making sure warranty policies and procedures are followed.";
  H.page(pptx, s, { eyebrow: SEG, title: "Keep the Claim Paid" });
  H.panel(pptx, s, { x: 0.5, y: H.BODY.top, w: 12.33, h: H.bodyH(tk), label: "Warranty compliance, in five steps" });
  tiles(s, [
    ["1  Documented policies", "Each manufacturer's current warranty policy manual, where the team can reach it"],
    ["2  Defined process", "Who writes, who approves, who submits: every claim handled the same way"],
    ["3  Create awareness", "Advisors, technicians and warranty admins know what an auditor checks"],
    ["4  Safety nets", "Signatures, out mileage, punch times and authorizations checked before the claim goes in"],
    ["5  Spot-check and self-audit", "Pull warranty ROs every month and audit them the way the factory will"],
  ], { cols: 5, top: H.BODY.top + 0.72, bot: H.bodyBot(tk) - 0.2, accent: C.blue, headH: 0.7 });
  H.takeaway(pptx, s, tk, "n");
  H.foot(pptx, s, SRC);
  s.addNotes("DRAFT. Steps from Jan 2026 GM meeting slide 14 (Warranty Compliance); the GM responsibility line from slide 13.");
}

await pptx.writeFile({ fileName: "Warranty Audits - Draft.pptx" }); console.log("warranty built");
