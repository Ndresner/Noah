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

// agenda details simplified to plain language (Noah, Oct 23, from the agenda preview)
export const SECTIONS = [
  { short: "Gross",               detail: "Fixed ops, service and parts gross performance.", chips: ["Fixed ops", "Service", "Parts"] },
  { short: "Profitability",       detail: "What we keep, and how much of the store we cover.",                                chips: ["Absorption", "Service net to gross", "Parts net to gross"] },
  { short: "Customer Pay",        detail: "Car count, sales per visit, pricing and margin.",  chips: ["ROs", "Hours per RO", "Pricing & mix", "ELR", "$ per RO", "Margin"] },
  { short: "Opcodes",             detail: "Inspections and key services against benchmark.", chips: ["MPI 90%", "Alignment 15%", "Balance 15%", "Brake fluid 10%", "Fuel 7.5%"], name: "Opcode Benchmarks" },
  { short: "Internal",            detail: "Internal labor rate and margins.", chips: ["Internal ELR", "Internal labor GP %", "Internal parts GP %"] },   // tracking moved to Opcodes (Oct 23)   // pay type and work type moved to Customer Pay (Oct 23)
  { short: "Productivity",        detail: "Staffing, output and technician proficiency.",   chips: ["Headcount", "Gross per employee", "Technician proficiency"] },   // ROs per person dropped (Noah, Oct 23)
  { short: "Warranty",            detail: "Warranty rates, margins and compliance.", chips: ["Rate increases", "Warranty ELR", "Parts margin", "Warranty audits"] },
  { short: "Close",               detail: "What changed, where the money is, and what we commit to.",         chips: ["Most improved", "Where the money is", "Before you leave"] },
];
// Oct 28 audit, option 8 (review only until Noah approves): CHIPS8=1 builds the proposed Warranty and Close wording
// into "Section Dividers - Review.pptx"; the deck build is unchanged
// approved by Noah Oct 28 - now the deck wording
if (process.env.CHIPS8 !== "0") {
  // Oct 28 (Noah): the old agenda line plus protecting the money; "Warranty audits" chip reworded
  Object.assign(SECTIONS[6], { detail: "Rates, margins and compliance: protecting the money.", chips: ["Rates rising", "Warranty gross profit", "Claim process: protect the money"] });
  Object.assign(SECTIONS[7], { detail: "What changed, where the money is, and each store's prescription.", chips: ["Most improved", "Where the money is", "Store prescriptions"] });
}
const TRACKER = SECTIONS.map(s => s.short);

H.cover(pptx, pptx, { title: MEETING.title, subtitle: MEETING.subtitle, meta: [MEETING.venue], logo: LOGO });
H.agenda(pptx, pptx, { eyebrow: "Day 2 · General Managers & Service Managers", title: "Agenda",
  items: SECTIONS.map(s => ({ name: s.name ?? s.short, detail: s.detail })) });
SECTIONS.forEach((sec, i) => H.section(pptx, pptx, {
  kicker: `Section ${String(i + 1).padStart(2, "0")}`, title: sec.name ?? sec.short, chips: sec.chips,
  sections: TRACKER, current: i, logo: LOGO,
}));
// not a meeting section (no number, no tracker): alternate designs and slides still to be decided on. Not presented.
H.section(pptx, pptx, { kicker: "Not presented · for review", title: "Open Decisions", chips: ["Store records", "Door rate", "CP vs internal", "Photo %", "Productivity", "Who's moving now", "Most improved"], logo: LOGO })
  .addNotes("For review with Tommy, not for the room. Every slide built but not placed in the deck: alternate versions of slides already in it, and slides not yet decided on. In deck order: Store Records (five designs), Door Rate option A, CP vs Internal GP % option A, Photo completion % (optional measure), Productivity slides taken out of the flow (gross per employee gap, 5% more per employee, ROs per advisor, ROs per manager; Oct 23), Who's Moving Now alternate versions and the line-chart versions taken out of the flow (Oct 23), Most Improved v1. Pick what moves into the deck, then delete this section before the meeting.");

await pptx.writeFile({ fileName: "Day 2 Opening - Day 1 Style.pptx" }); console.log("opening built", 2 + SECTIONS.length + 1);
