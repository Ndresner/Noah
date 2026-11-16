/* Review (Noah, Oct 23): slide 10 (fixed ops gross movers) as three lists - top and bottom year over year, plus top
   moving now within 2026 - at five and at three stores. Year over year = Jan-Sep 2026 vs Jan-Sep 2025 (ntgs/ntgp.json,
   same as slide 10). Moving now = Jul-Sep 2026 (Q3) vs Jan-Mar 2026 (Q1), per-month average, same basis as Who's Moving
   Now (pulls/gross_m26_store.json; Jan-Sep sums reconcile to slide 10 store by store). */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const N = JSON.parse(fs.readFileSync("./ntgs.json", "utf8")), P = JSON.parse(fs.readFileSync("./ntgp.json", "utf8"));
const Q = JSON.parse(fs.readFileSync("./pulls/gross_m26_store.json", "utf8"));
const nm = H.storeName, sgn = v => v >= 0 ? "+" : "−", pcA = v => sgn(v) + Math.abs(v * 100).toFixed(1) + "%";
const list = a => a.slice(0, -1).join(", ") + " and " + a.at(-1);
const yoy = N.S.map((s, i) => ({ n: nm(s[0]), a: s[2] + P.S[i][2], b: s[4] + P.S[i][4] }))
  .map(r => ({ n: r.n, v0: H.kusd(r.a), v1: H.kusd(r.b), d: r.b / r.a - 1 })).sort((x, y) => y.d - x.d);
const mov = Q.rows.filter(r => r[0] !== "INEOS").map(([n, sv, pt]) => {
  const m = sv.split("|").map((v, i) => -(+v + +pt.split("|")[i]));
  const q1 = (m[0] + m[1] + m[2]) / 3, q3 = (m[6] + m[7] + m[8]) / 3;
  return { n: nm(n), v0: H.kusd(q1), v1: H.kusd(q3), d: q3 / q1 - 1, q1, q3 };
}).sort((x, y) => y.d - x.d);
const g1 = mov.reduce((t, r) => t + r.q1, 0), g3 = mov.reduce((t, r) => t + r.q3, 0), gm = g3 / g1 - 1;
const SRC = "Qlik - Profit & Loss. Same-store, INEOS excluded (opened 2026). Fixed ops = service + parts; GMC Prescott body shop gross excluded.";
for (const K of [5, 3]) {
  const top = yoy.slice(0, K), bot = yoy.slice(-K).reverse(), now = mov.slice(0, K), n3 = a => list(a.slice(0, 3).map(r => r.n));
  H.movers3(pptx, pptx, {
    eyebrow: `Fixed Ops Gross · Review · Top ${K}`, title: "Total Gross Profit - Fixed Ops", k: K, fmtD: pcA,
    cols: [
      { head: `TOP ${K} VS LAST YEAR`, sub: "Jan–Sep 2026 vs Jan–Sep 2025", rows: top, tone: "good" },
      { head: `BOTTOM ${K} VS LAST YEAR`, sub: "Jan–Sep 2026 vs Jan–Sep 2025", rows: bot, tone: "bad", rank: j => yoy.length - j },
      { head: `TOP ${K} MOVING NOW`, sub: "Jul–Sep vs Jan–Mar 2026, per month", rows: now, tone: "blue" },
    ],
    takeaway: `Ask ${n3(top)} what changed. Ask ${n3(bot)} what got in the way. Ask ${n3(now)} what they started this year.`,
    source: `${SRC} Year over year: percent change in fixed ops gross, Jan-Sep 2026 vs Jan-Sep 2025. Moving now: Jul-Sep 2026 vs Jan-Mar 2026, average per month, each store against itself; the group rose ${pcA(gm)} on the same basis (seasonal), so this list is stores that beat the season.`,
    notes: `Review option: top/bottom ${K} year over year plus top ${K} moving now. Group moving-now baseline: ${H.kusd(g1)} to ${H.kusd(g3)} per month (${pcA(gm)}).`,
  });
}
await pptx.writeFile({ fileName: "../decks/Fixed Ops Gross Movers - Review.pptx" }); console.log("movers3 review built");
