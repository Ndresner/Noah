/* Review (Noah, Oct 23): two other ways to show luxury vs non-luxury on the segment level slides, using CP ELR as the
   example. Not in the deck; the deck keeps one ranking of all 33 with a tick at each store's segment average.
   A: same two columns, luxury grouped first, then non-luxury, each ranked on its own and scored against its own average.
   A without the line (Oct 23): same, no navy average line; colour alone shows above or below.
   B: two slides, one per segment, each against its own segment average. Segment averages from segment totals. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, LUX } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const X = load("./m5.json"), f = a => a[1] / a[2], fmt = v => "$" + v.toFixed(2), nm = H.storeName;
const L = X.st.filter(x => x.lux), N = X.st.filter(x => !x.lux);
const la = f(X.tot(L.map(x => x.a26))), na = f(X.tot(N.map(x => x.a26)));
const rows = X.st.map(x => [x.n, f(x.a26)]), lv = new Set(L.map(x => x.n)), nv = new Set(N.map(x => x.n));
const SRC = "Qlik - Closed Repair Orders. Customer pay, Jan-Sep 2026 by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded. Segment averages from segment totals. Luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson; Acura counted as non-luxury.";
const at = (S, ref) => S.filter(x => f(x.a26) >= ref).length;
const best = S => [...S].sort((p, q) => f(q.a26) - f(p.a26))[0];
const base = { title: "CP ELR", metric: "elr", measure: "CP ELR", fmt, groupLabel: "SEPTEMBER 2026 YTD" };

H.levelRanker(pptx, pptx, { ...base, eyebrow: "Customer Pay · Review · Option A: grouped", rows, group: f(X.G26),
  groups: [{ label: `LUXURY · AVERAGE ${fmt(la)}`, names: lv, ref: la }, { label: `NON-LUXURY · AVERAGE ${fmt(na)}`, names: nv, ref: na }],
  alignRef: true, legend: "Each segment ranked on its own. Navy line = the segment average; green at or above it.",
  takeaway: `${at(L, la)} of ${L.length} luxury stores beat ${fmt(la)} and ${at(N, na)} of ${N.length} others beat ${fmt(na)}. ${nm(best(L).n)} leads luxury, ${nm(best(N).n)} non-luxury.`,
  source: SRC + " Option A: one slide, two groups." });

H.levelRanker(pptx, pptx, { ...base, eyebrow: "Customer Pay · Review · Option A: grouped, no average line", rows, group: f(X.G26),
  groups: [{ label: `LUXURY · AVERAGE ${fmt(la)}`, names: lv, ref: la }, { label: `NON-LUXURY · AVERAGE ${fmt(na)}`, names: nv, ref: na }],
  alignRef: true, hideRef: true, legend: "Each segment ranked on its own. Green at or above its segment average, red below.",
  takeaway: `${at(L, la)} of ${L.length} luxury stores beat ${fmt(la)} and ${at(N, na)} of ${N.length} others beat ${fmt(na)}. ${nm(best(L).n)} leads luxury, ${nm(best(N).n)} non-luxury.`,
  source: SRC + " Option A without the average line (Noah, Oct 23): bars still drawn against each segment average, colour shows above or below." });

for (const [S, ref, lab, names] of [[L, la, "Luxury", lv], [N, na, "Non-Luxury", nv]]) {
  const b = best(S);
  H.levelRanker(pptx, pptx, { ...base, eyebrow: `Customer Pay · Review · Option B: ${lab}`, title: `CP ELR - ${lab}`, measure: `CP ELR · ${lab.toUpperCase()}`,
    rows: rows.filter(r => names.has(r[0])), group: ref,
    takeaway: `${at(S, ref)} of ${S.length} ${lab.toLowerCase()} stores at or above the ${fmt(ref)} ${lab.toLowerCase()} average, led by ${nm(b.n)} at ${fmt(f(b.a26))}.`,
    source: SRC + ` Option B: one slide per segment; green at or above the ${lab.toLowerCase()} average.` });
}
await pptx.writeFile({ fileName: "../decks/Luxury Split Options - Review.pptx" }); console.log("segment review built");
