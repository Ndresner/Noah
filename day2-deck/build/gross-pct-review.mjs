import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
/* REVIEW (Noah, Oct 22 2026): option A for the gross change slides - rank by % change so store size does not decide
   the order, with the dollar change beside it. Same data as gross-deck.mjs (Qlik P&L, ntgs/ntgp). Not in the deck. */
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const N = JSON.parse(fs.readFileSync("./ntgs.json", "utf8")), P = JSON.parse(fs.readFileSync("./ntgp.json", "utf8"));
const SEG = "Fixed Ops Gross";
const SRC = "Qlik - Profit & Loss. Same-store, INEOS excluded (opened 2026). Fixed ops = service + parts; GMC Prescott body shop gross excluded.";
const svS = N.S.map(s => [s[0], s[2], s[4]]), ptS = P.S.map(s => [s[0], s[2], s[4]]);
const fxS = svS.map((r, i) => [r[0], r[1] + ptS[i][1], r[2] + ptS[i][2]]);
const sg = v => v >= 0 ? "+" : "−";
const pc = v => sg(v) + Math.abs(v * 100).toFixed(1) + "%";
const kd = v => sg(v) + H.kusd(Math.abs(v));
const nm = H.storeName;
for (const [title, S, metric] of [["Total Gross Profit - Fixed Ops", fxS, "fixed ops gross"], ["Total Gross Profit - Service", svS, "service gross"], ["Total Gross Profit - Parts", ptS, "parts gross"]]) {
  const g25 = S.reduce((t, r) => t + r[1], 0), g26 = S.reduce((t, r) => t + r[2], 0), gp = g26 / g25 - 1;
  const rows = S.map(([n, a, b]) => ({ n, p: b / a - 1, d: b - a })).sort((x, y) => y.p - x.p);
  const grew = rows.filter(r => r.p > 0).length, top = rows[0], low = rows.at(-1);
  H.ranker(pptx, pptx, {
    eyebrow: `${SEG} · Option A`, title, metric, measure: title.replace("Total Gross Profit - ", "") + " Gross",
    deltas: rows.map(r => [r.n, r.p]), groupD: gp, fmtD: pc, valLabel: "Change", stripValue: `${pc(gp)}  ·  ${kd(g26 - g25)}`,
    extra: [{ label: "$", w: 0.95, vals: Object.fromEntries(rows.map(r => [r.n, kd(r.d)])) }],
    groupLabel: "2026 VS 2025 YTD · **YEAR-OVER-YEAR CHANGE**",
    takeaway: `${grew} of 33 stores grew. ${nm(top.n)} grew most at ${pc(top.p)} (${kd(top.d)}); ${nm(low.n)} fell most at ${pc(low.p)} (${kd(low.d)}).`,
    source: `${SRC} Ranked by percent change, Jan-Sep 2026 vs Jan-Sep 2025, so store size does not set the order; dollar change alongside.`,
  });
}
await pptx.writeFile({ fileName: "Gross Change by Percent - Option A.pptx" }); console.log("built");
