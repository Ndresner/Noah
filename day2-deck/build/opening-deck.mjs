import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
/* ===================================================================
   GM MEETING - OCTOBER 2026 - DAY 2 (fixed ops)
   The section list lives here ONCE, as on Day 1 (Tommy's deck.mjs): the
   agenda numbers, the transition kickers and the tracker are all derived
   from it. Cover, agenda wording and audience line approved by Noah, Oct 21 2026.
   =================================================================== */
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
export const LOGO = "image/png;base64," + fs.readFileSync("../assets/findlay-logo.png").toString("base64");

export const MEETING = {
  title:    "Fixed Operations",
  subtitle: "Day 2  ·  General Managers & Service Managers  ·  October 29, 2026",
  venue:    "South Point, Las Vegas",
};

export const SECTIONS = [
  { short: "Gross",               detail: "Fixed ops, service and parts gross, and the store records.", chips: ["Fixed ops", "Service", "Parts"] },
  { short: "Profitability",       detail: "Absorption and net to gross.",                                chips: ["Absorption", "Service net to gross", "Parts net to gross"] },
  { short: "CP Engine",           detail: "ROs, hours per RO, rate and margin.",                         chips: ["ROs", "Hours per RO", "ELR", "$ per RO", "Margin"], name: "The CP Engine" },
  { short: "Labor Mix",           detail: "Express, maintenance and repair mix.",                        chips: ["Pay type", "Work type", "Op code descriptions"] },
  { short: "Opcodes",             detail: "Alignment, balance, brake fluid and fuel.",                   chips: ["Alignment 15%", "Balance 15%", "Brake fluid 10%", "Fuel service 7.5%"], name: "Opcode Benchmarks" },
  { short: "Warranty & Internal", detail: "Warranty ROs and rate, internal rate and margin.",            chips: ["Warranty ROs", "Warranty ELR", "Internal ELR"] },
  { short: "Productivity",        detail: "Gross per employee, headcount and technician proficiency.",   chips: ["Gross per employee", "ROs per person", "Headcount", "Technician proficiency"] },
  { short: "Close",               detail: "What changed, where the money is, and commitments.",         chips: ["Most improved", "Where the money is", "Before you leave"] },
];
const TRACKER = SECTIONS.map(s => s.short);

H.cover(pptx, pptx, { title: MEETING.title, subtitle: MEETING.subtitle, meta: [MEETING.venue], logo: LOGO });
H.agenda(pptx, pptx, { eyebrow: "Day 2 · General Managers & Service Managers", title: "Agenda",
  items: SECTIONS.map(s => ({ name: s.name ?? s.short, detail: s.detail })) });
SECTIONS.forEach((sec, i) => H.section(pptx, pptx, {
  kicker: `Section ${String(i + 1).padStart(2, "0")}`, title: sec.name ?? sec.short, chips: sec.chips,
  sections: TRACKER, current: i, logo: LOGO,
}));
// the appendix is not a meeting section: no number, no tracker
H.section(pptx, pptx, { kicker: "Appendix", title: "Design Options", chips: ["Door rate", "Your best techs prove it"], logo: LOGO })
  .addNotes("Appendix for review with Tommy. Alternate designs still to choose between: Door Rate / ELR % of Door Rate options, and Your Best Techs Prove It options A-C. Delete whichever we don't pick.");

await pptx.writeFile({ fileName: "Day 2 Opening - Day 1 Style.pptx" }); console.log("opening built", 2 + SECTIONS.length + 1);
