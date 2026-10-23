import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
import { block, F, LUX } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* LABOR MIX - one segment: how labor splits by pay type, then how customer-pay labor splits by work type.
   Merged from mix-deck.mjs and wm-deck.mjs, Oct 15 2026. Segment name is the eyebrow on every slide. */
const SEG = "Labor Mix";
H.section(pptx, pptx, { kicker: "Fixed Ops", title: "Labor Mix", chips: ["Pay type", "Work type", "Op code descriptions"] });
/* Qlik Closed Repair Orders app c3efc739, pulled 2026-10-14/15. Labor sale by pay type: C from m5.json a[1],
   W and I from wi.json a[1], a[3]. Same-store, INEOS excluded, GMC Prescott body shop labor types excluded. */

const SRC_PT = "Qlik - Closed Repair Orders. Labor sale by pay type, Jan-Sep by close month. Same-store, INEOS excluded. GMC Prescott body shop labor types excluded.";
const C5 = JSON.parse(fs.readFileSync("./m5.json", "utf8")), WI = JSON.parse(fs.readFileSync("./wi.json", "utf8"));
const wi = Object.fromEntries(WI.S.map(s => [s[0], s]));
// a = [cp, w, i] labor sale
const D = {
  m25: C5.m25.map((r, k) => [r[1], WI.m25[k][1], WI.m25[k][3]]),
  m26: C5.m26.map((r, k) => [r[1], WI.m26[k][1], WI.m26[k][3]]),
  q4: C5.q4.map((r, k) => [r[1], WI.q4[k][1], WI.q4[k][3]]),   // 2025 Oct-Dec, chart only
  S: C5.S.map(s => { const w = wi[s[0]]; return [s[0], s[2], w[2], w[4], s[8], w[7], w[9]]; }),
};
const tot = rows => rows.reduce((t, r) => t.map((x, i) => x + r[i]), [0, 0, 0]);
// moving now (Oct 23): Jan-Mar and Jul-Sep 2026 [cp, w, i] labor sale (m5 mn: labor sale at 2 and 8; wi mn: W at 2 and 7, I at 4 and 9)
const c5m = Object.fromEntries(C5.mn.map(r => [r[0], r])), wim = Object.fromEntries(WI.mn.map(r => [r[0], r]));
const X = { D, tot, G25: tot(D.m25), G26: tot(D.m26),
  st: D.S.map(s => ({ n: s[0], a25: s.slice(1, 4), a26: s.slice(4, 7), lux: LUX.has(s[0]),
    q1: [c5m[s[0]][2], wim[s[0]][2], wim[s[0]][4]], q3: [c5m[s[0]][8], wim[s[0]][7], wim[s[0]][9]] })) };
const share = a => a.map(v => v / (a[0] + a[1] + a[2]));
const pc = v => (v * 100).toFixed(1) + "%";
const g = share(X.G26), g5 = share(X.G25);
const sorted = [...X.st].sort((p, q) => share(q.a26)[0] - share(p.a26)[0]);
H.mixBars(pptx, pptx, {
  eyebrow: SEG, title: "Labor Sale by Pay Type",
  parts: [{ label: "Customer pay", color: "2587C5", dark: true }, { label: "Warranty", color: "8FBEE6" }, { label: "Internal", color: "D9D9D9" }],
  rows: X.st.map(x => [x.n, share(x.a26)]), group: g,
  takeaway: `Customer pay is ${pc(g[0])} of labor sale (${pc(g5[0])} in 2025), warranty ${pc(g[1])}, internal ${pc(g[2])}. ${sorted[0].n} is most customer-pay at ${pc(share(sorted[0].a26)[0])}.`,
  source: `${SRC_PT} Jan-Sep 2026 shares of labor sale; sorted by customer-pay share. Bar labels are percent of labor sale; every value is also listed here.\n` + sorted.map(x => `${x.n}: CP ${pc(share(x.a26)[0])}, W ${pc(share(x.a26)[1])}, I ${pc(share(x.a26)[2])}`).join("\n"),
});
block(pptx, X, { seg: SEG, src: SRC_PT, title: "Customer Pay Share of Labor Sale", metric: "cp labor share", fn: a => a[0] / (a[0] + a[1] + a[2]), f: F.pct, level: false,
  def: "Customer-pay labor sale / total labor sale (customer pay + warranty + internal).",
  extra: ({ G25, G26 }) => `Warranty labor ${G26[1] >= G25[1] ? "rose" : "fell"} ${pc(Math.abs(G26[1] / G25[1] - 1))}, internal rose ${pc(G26[2] / G25[2] - 1)}.`,
  verb: "raised customer-pay share" });

/* Qlik Closed Repair Orders app c3efc739, pulled 2026-10-15. Customer-pay labor sale, Jan-Sep, by op code description.
   Maintenance and repair are keyword matches on the description (rules in cls.py); maintenance wins a tie; known
   descriptions that are ONLY a placeholder (exact match: MISC, DEFAULT, STANDARD, CUSTOMER STATES...) are forced to unclassified. Everything else
   is unclassified. Same-store, INEOS excluded, GMC Prescott body shop labor types excluded. Store totals reconcile
   to the group CP labor sale to the dollar. */

const RULES = "Maintenance = descriptions naming oil and filter, rotation, brake pads, rotors and brake service, scheduled/interval/mileage service, OEM care plans, alignment, balance, fluid exchanges and flushes, fuel induction, transmission/differential/transfer case/coolant service, spark plugs, filters, wipers, multi-point inspection, timing belt, BG and refresh services, smog/emissions inspection, fluid services, tune-ups, tire promotions (buy 3 get 1), CVT, EGR, 4x4, AWD, intermediate and transmission services, and any coolant replacement or change. Repair = repair, replace, R&R, diagnosis, leak, noise, inoperative, check engine, warning light, struts, shocks, alternator, starter, water pump, reseal, gasket, calibration, overhaul, engine light, drive belt, belt and hose kit. Descriptions that are only a placeholder (exactly MISC, DEFAULT, STANDARD, GENERAL, GENERAL REPAIR, CUSTOMER STATES, GENERAL CONCERN, ADD DESCRIPTION, TELL US MORE and similar) are left unclassified; a placeholder followed by real detail is classified by that detail.";
const SRC_WM = "Qlik - Closed Repair Orders. Customer-pay labor sale by op code description, Jan-Sep. Same-store, INEOS excluded. GMC Prescott body shop labor types excluded.";
/* Oct 21 2026: the raw op-code-description pull (wm.json) did not survive the move to Claude Code. The shares
   are recovered from the v1 deck's speaker notes and takeaway (wm_shares.json, to 0.1%). */
const WS = JSON.parse(fs.readFileSync("./wm_shares.json", "utf8"));
const w26 = WS.group26, w25 = [null, null, WS.group25U];
const pcW = v => (v * 100).toFixed(1) + "%";
const rowsW = Object.entries(WS.stores).map(([n, v]) => [n, v]);
const byU = [...rowsW].sort((a, b) => a[1][2] - b[1][2]);
H.mixBars(pptx, pptx, {
  eyebrow: SEG, title: "Customer-Pay Labor by Work Type",
  parts: [{ label: "Maintenance", color: "2587C5", dark: true }, { label: "Repair", color: "0B2C4A", dark: true }, { label: "Unclassified", color: "D9D9D9" }],
  rows: rowsW, group: w26,
  takeaway: `Maintenance is ${pcW(w26[0])} of customer-pay labor, repair ${pcW(w26[1])}, and ${pcW(w26[2])} cannot be classified from its description (${pcW(w25[2])} in 2025).`,
  source: `${SRC_WM} ${RULES}\n` + [...rowsW].sort((a, b) => b[1][0] - a[1][0]).map(r => `${r[0]}: M ${pcW(r[1][0])}, R ${pcW(r[1][1])}, U ${pcW(r[1][2])}`).join("\n"),
});
H.levelRanker(pptx, pptx, {
  eyebrow: SEG, title: "Unclassified Customer-Pay Labor", metric: "cp unclassified work %",
  rows: rowsW.map(r => [r[0], r[1][2]]), group: w26[2], fmt: pcW,
  takeaway: `${byU[byU.length - 1][0]} can't classify ${pcW(byU[byU.length - 1][1][2])} of its customer-pay labor from the description; ${byU[0][0]} only ${pcW(byU[0][1][2])}. Vague op code descriptions hide what the shop sells.`,
  tone: "n",
  source: `${SRC_WM} Share of CP labor sale whose op code description matches neither the maintenance nor the repair rules, Jan-Sep 2026. Lower is better: green at or below the group, red above. ${RULES}`,
});

await pptx.writeFile({ fileName: "Labor Mix - House Style.pptx" }); console.log("labor mix built");
