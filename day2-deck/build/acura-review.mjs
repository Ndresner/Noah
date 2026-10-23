/* Review (Noah, Oct 23): every luxury / non-luxury slide in the deck as option A (grouped, no average line), twice:
   Acura in non-luxury (as the deck has it today) and Acura in luxury, to decide where Acura belongs. Not in the deck.
   Same data, same Jan-Sep 2026 basis as the deck slides; segment averages from segment totals. Acura row highlighted. */
import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import { load, LUX, F } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const R = a => a[0] / a[1], nm = H.storeName;
const CP = "Qlik - Closed Repair Orders. Customer pay, Jan-Sep 2026 by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded.";
const PR = "Gross and ROs: Qlik - Productivity app; headcount: ADP Employee Master extract, month-end snapshots. Monthly averages, Jan-Sep 2026. Same-store, INEOS excluded.";
const M = [
  { file: "./m5.json", seg: "Customer Pay", title: "CP Hours per RO", metric: "hpro", fn: a => a[2] / a[0], f: F.dec2, src: CP },
  { file: "./m5.json", seg: "Customer Pay", title: "CP ELR", metric: "elr", fn: a => a[1] / a[2], f: F.usd2, src: CP },
  { file: "./m5.json", seg: "Customer Pay", title: "CP Labor $ per RO", metric: "cp labor per ro", fn: a => a[1] / a[0], f: F.usd, src: CP },
  { file: "./prod/rpa.json", seg: "Fixed Ops Productivity", title: "CP & W ROs per Service Advisor", metric: "ros per advisor", fn: R, f: F.int, src: PR },
  { file: "./prod/rpe.json", seg: "Fixed Ops Productivity", title: "CP & W ROs per Fixed Employee", metric: "ros per fixed employee", fn: R, f: F.dec1, src: PR },
  { file: "./prod/rpm.json", seg: "Fixed Ops Productivity", title: "CP & W ROs per Fixed Manager", metric: "ros per fixed manager", fn: R, f: F.int, src: PR },
];
const summary = [];
for (const m of M) {
  const X = load(m.file), f = m.fn, fmt = m.f.v;
  for (const acuraLux of [false, true]) {
    const lux = new Set([...LUX, ...(acuraLux ? ["ACURA"] : [])]);
    const L = X.st.filter(x => lux.has(x.n)), N = X.st.filter(x => !lux.has(x.n));
    const la = f(X.tot(L.map(x => x.a26))), na = f(X.tot(N.map(x => x.a26)));
    const ac = X.st.find(x => x.n === "ACURA"), av = f(ac.a26), seg = acuraLux ? L : N, ref = acuraLux ? la : na;
    const rk = [...seg].sort((p, q) => f(q.a26) - f(p.a26)).findIndex(x => x.n === "ACURA") + 1;
    const at = (S, r) => S.filter(x => f(x.a26) >= r).length;
    const where = `Acura ${fmt(av)}: #${rk} of ${seg.length} ${acuraLux ? "luxury" : "non-luxury"}, ${av >= ref ? "above" : "below"} the ${fmt(ref)} average.`;
    summary.push(`${m.title} | Acura in ${acuraLux ? "luxury" : "non-luxury"}: ${where} Luxury avg ${fmt(la)}, non-luxury avg ${fmt(na)}.`);
    H.levelRanker(pptx, pptx, {
      eyebrow: `${m.seg} · Review · Acura in ${acuraLux ? "luxury" : "non-luxury"}`, title: m.title, metric: m.metric, measure: m.title,
      fmt, groupLabel: "SEPTEMBER 2026 YTD", rows: X.st.map(x => [x.n, f(x.a26)]), group: f(X.G26),
      groups: [{ label: `LUXURY · AVERAGE ${fmt(la)}`, names: new Set(L.map(x => x.n)), ref: la }, { label: `NON-LUXURY · AVERAGE ${fmt(na)}`, names: new Set(N.map(x => x.n)), ref: na }],
      alignRef: true, hideRef: true, mark: new Set(["ACURA"]),
      legend: `Each segment ranked on its own. Green at or above its segment average, red below. Acura highlighted.`,
      takeaway: `${where} ${at(L, la)} of ${L.length} luxury and ${at(N, na)} of ${N.length} non-luxury stores at or above their average.`,
      source: `${m.src} Segment averages from segment totals. Luxury = Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson${acuraLux ? ", plus Acura on this slide" : "; Acura non-luxury, as in the deck"}.`,
    });
  }
}
console.log(summary.join("\n"));
await pptx.writeFile({ fileName: "../decks/Acura Segment - Review.pptx" }); console.log("acura review built");
