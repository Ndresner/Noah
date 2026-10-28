import PptxGenJS from "pptxgenjs";
import * as H from "./house.mjs";
import fs from "fs";
import { load, block, F, moversSlide, sparkFrom } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* Qlik Closed Repair Orders app c3efc739, pulled 2026-10-15. a = [W ROs, W labor sale, W hours, I labor sale, I hours].
   Same-store (INEOS excluded), GMC Prescott body shop labor types excluded. Store rows reconcile on all five fields. */
const X = load("./wi.json"), SEG = "Warranty";   // internal ELR sits in Labor Mix (Noah, Oct 22)
const SRC = "Qlik - Closed Repair Orders. Jan-Sep by close month. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded.";
H.section(pptx, pptx, { kicker: "Fixed Ops", title: SEG, chips: ["Warranty gross", "Warranty ELR", "Internal ELR"] });
const base = { seg: SEG, src: SRC };
// Oct 25 (Noah): warranty RO count is out; volume isn't the point of the section. Warranty gross profit is the money
// the manufacturer legitimately owes us, the dollars that policy and procedure protect. Same four-slide shape as the
// gross blocks (WI 2-5). Data: wg.json, P&L line Gross Profit - Fixed Ops Warranty, reconciled to the cent.
{
  const WG = JSON.parse(fs.readFileSync("./wg.json", "utf8")), T = "Warranty Gross Profit", MET = "fixed ops gross";
  const WSRC = "Qlik - Profit & Loss, Gross Profit - Fixed Ops Warranty (service + parts warranty gross). Same-store, INEOS excluded (opened 2026); no body shop.";
  const sum = a => a.reduce((t, x) => t + x, 0), sgn = v => v >= 0 ? "+" : "\u2212";
  const pcA = v => sgn(v) + Math.abs(v * 100).toFixed(1) + "%", kd = v => sgn(v) + H.kusd(Math.abs(v));
  const m25 = WG.m25.map(r => r[0]), m26 = WG.m26.map(r => r[0]), q4 = WG.q4.map(r => r[0]), v25 = sum(m25), v26 = sum(m26), d = v26 - v25;
  const S = WG.S, MN = Object.fromEntries(WG.mn.map(r => [r[0], [r[1] / 3, r[2] / 3]]));
  const FULL = ["January","February","March","April","May","June","July","August","September"];
  const dm = m26.map((v, i) => v - m25[i]), up = dm.filter(x => x > 0).length, bi = dm.indexOf(Math.max(...dm));
  H.groupYoy(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: "Warranty Gross",
    prior: [...m25, ...q4], current: m26, priorName: "2025", currentName: "2026", expectPrior: v25, expectCurrent: v26, chart: "line", majorUnit: 1,   // chart runs in $M
    chartLabel: "Warranty gross profit by month, 2025 against 2026 ($M)",
    takeaway: `${kd(d)} (${pcA(d / v25)}) of warranty gross year to date; ${up} of nine months ahead of 2025, largest gain in ${FULL[bi]}, ${kd(dm[bi])}. This is money the manufacturer owes us: process keeps it.`,
    source: `${WSRC} Jan-Sep 2026 vs Jan-Sep 2025; 2025 shown January-December. Store rows reconcile to the group to the cent, both years.` });
  H.totalRanker(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: "Warranty Gross", rows: S.map(r => [r[0], r[2]]), expect: v26,
    groupLabel: "SEPTEMBER 2026 YTD",
    takeaway: (() => { const t = [...S].sort((p, q) => q[2] - p[2]); return `${H.storeName(t[0][0])} leads at ${H.Mmixed(t[0][2])} of the group's ${H.Mmixed(v26)}; ${H.storeName(t.at(-1)[0])} is smallest at ${H.Mmixed(t.at(-1)[2])}.`; })(),
    tone: "n", source: `${WSRC} Jan-Sep 2026.` });
  const pr = S.map(([n, a, b]) => ({ n, d: b - a, p: b / a - 1 }));
  const byP = [...pr].sort((x, y) => y.p - x.p), mk = Object.fromEntries([...byP.slice(0, 3).map(r => [r.n, "g"]), ...byP.slice(-3).map(r => [r.n, "r"])]);
  const grew = pr.filter(r => r.d > 0).length, bd = [...pr].sort((x, y) => y.d - x.d);
  H.ranker(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: "Warranty Gross",
    deltas: pr.map(r => [r.n, r.d]), groupD: d, fmtD: kd, valLabel: "$", stripValue: `${kd(d)}  ·  ${pcA(d / v25)}`,
    extra: [{ label: "Change", w: 0.95, vals: Object.fromEntries(pr.map(r => [r.n, pcA(r.p)])), marks: mk }],
    groupLabel: "2026 VS 2025 YTD · **YEAR-OVER-YEAR CHANGE**",
    takeaway: `${grew} of 33 stores grew warranty gross. ${H.storeName(bd[0].n)} added the most at ${kd(bd[0].d)} (${pcA(bd[0].p)}); ${H.storeName(bd.at(-1).n)} fell most at ${kd(bd.at(-1).d)} (${pcA(bd.at(-1).p)}).`,
    source: `${WSRC} Ranked by dollar change, Jan-Sep 2026 vs Jan-Sep 2025; percent change alongside, the top three and bottom three by percent marked. Warranty volume follows campaigns and recalls as well as store effort.` });
  moversSlide(pptx, { seg: SEG, title: T, metric: MET, fmtD: pcA, spark: sparkFrom(WG.sm, a => a[0]),
    // gray second line = the dollar change (Noah, Oct 28)
    rows: S.map(([n, a, b]) => ({ n, v0: H.kusd(a), v1: H.kusd(b), d: b / a - 1, d2: kd(b - a) })),
    now: S.map(([n]) => ({ n, v0: H.kusd(MN[n][0]), v1: H.kusd(MN[n][1]), d: MN[n][1] / MN[n][0] - 1, d2: kd(MN[n][1] - MN[n][0]) })),
    nowGroup: pcA(sum(S.map(([n]) => MN[n][1])) / sum(S.map(([n]) => MN[n][0])) - 1), perMonth: true,
    src: `${WSRC} Top and bottom five stores by percent change in warranty gross, Jan-Sep 2026 vs 2025.` });
}
block(pptx, X, { ...base, title: "Warranty ELR", metric: "warranty elr", fn: a => a[1] / a[2], f: F.usd2,
  def: "Warranty labor sale / warranty sold hours.", verb: "raised warranty ELR" });
block(pptx, X, { ...base, seg: "Internal", title: "Internal ELR", metric: "internal elr", fn: a => a[3] / a[4], f: F.usd2, pct: true, whole: true, absLabel: "$",   // % rank, whole dollars (Noah, Oct 24)
  def: "Internal labor sale / internal sold hours. A higher internal rate adds service gross but is a cost to used cars.",
  verb: "raised internal ELR" });
await pptx.writeFile({ fileName: "Warranty & Internal - House Style.pptx" }); console.log("wi built");
