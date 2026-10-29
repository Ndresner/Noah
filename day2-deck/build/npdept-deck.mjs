/* Oct 29 (Noah): service and parts department net profit, each broken out with the same five slides as the Fixed Ops
   Net Profit block (explainer, group YOY, store total ranker, change ranker with zero-crossing bands, top / bottom /
   moving now with the line style). Sits right after the Fixed Ops Net Profit block (merge key NPD).
   Data: ntgs.json (Mechanical) / ntgp.json (P&A): a = [department profit, department gross], same-store; sm = store
   months for the lines; pulls/net_by_dept.json fx_sales_26 for sales per $1 of profit. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
import { moversSlide, sparkFrom } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const J = f => JSON.parse(fs.readFileSync(f, "utf8"));
const sum = a => a.reduce((t, x) => t + x, 0);
const M = v => "$" + (v / 1e6).toFixed(2) + "M", usd = v => "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
const sgn = v => v >= 0 ? "+" : "−", pcA = v => sgn(v) + Math.abs(v * 100).toFixed(1) + "%", kd = v => sgn(v) + H.kusd(Math.abs(v));
const FULL = ["January","February","March","April","May","June","July","August","September"];
const SL = J("./pulls/net_by_dept.json").fx_sales_26, PA = J("./pulls/pa_expense_lines.json").rows;
const lst = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);

function deptBlock({ D, dept, line, sales, levers, notesExtra, band, bandNotes }) {
  const SEG = `${dept} Net Profit`, T = `${dept} Net Profit`, MET = "net profit", low = dept.toLowerCase();
  const SRC = `Qlik - Profit & Loss, Detail statement. Same-store, INEOS excluded (opened 2026). ${dept} net profit = ${line} department profit: department gross less the department's own expenses; GMC Prescott body shop is its own department and is not included.`;
  const m25 = D.m25.map(r => r[0]), m26 = D.m26.map(r => r[0]), q4 = D.q4.map(r => r[0]);
  const g25 = sum(D.m25.map(r => r[1])), g26 = sum(D.m26.map(r => r[1]));
  const v25 = sum(m25), v26 = sum(m26), d = v26 - v25, ex25 = g25 - v25, ex26 = g26 - v26;
  const S = D.S.map(r => [r[0], r[1], r[3]]);   // [name, profit 2025, profit 2026]
  if (Math.abs(sum(S.map(r => r[1])) - v25) > 1 || Math.abs(sum(S.map(r => r[2])) - v26) > 1) throw new Error(`${dept} net profit stores do not reconcile`);
  const MN = Object.fromEntries(D.mn.map(r => [r[0], [r[1] / 3, r[3] / 3]]));   // Jan-Mar, Jul-Sep monthly
  const per1 = sales / v26;

  H.explainer(pptx, pptx, { source: SRC, eyebrow: SEG, title: T,
    formula: [`${dept} net profit`, "=", `${low} gross`, "−", `${low} expenses`], formulaSize: 22,
    define: `The profit the ${low} department generates after its own expenses are covered.`,
    example: [["Gross", M(g26)], ["− Expenses", M(ex26)], ["= Net profit", M(v26)]], exampleLabel: "OUR NUMBERS",
    why: [`It is what ${low} actually contributes to the store.`,
          "Gross can grow while profit shrinks if expenses grow faster.",
          `Every dollar of expense saved drops straight to the bottom line. Earning that $1 takes $${Math.round(per1)} in ${low} sales.`],
    levers,
    ours: `Group: ${M(v26)} ${low} net profit, ${pcA(v26 / v25 - 1)}. Gross grew ${pcA(g26 / g25 - 1)}; department expenses ${pcA(ex26 / ex25 - 1)}.`,
    tone: v26 >= v25 ? "g" : "r",
    notes: `PLACE: right after the Fixed Ops Net Profit block (Noah, Oct 29: service and parts broken out, same slides as the other categories). DEPARTMENT PROFIT OR LOSS - ${line} Dept as the P&L Detail statement reports it. Jan-Sep 2026 gross ${usd(g26)}, department expenses ${usd(ex26)}, profit ${usd(v26)}; 2025 ${usd(g25)} / ${usd(ex25)} / ${usd(v25)}. Sales per $1 of profit: ${line} net sales ${usd(sales)} / ${usd(v26)} = $${per1.toFixed(2)}. ${notesExtra}` });

  const dm = m26.map((v, i) => v - m25[i]), up = dm.filter(x => x > 0).length, bi = dm.indexOf(Math.max(...dm));
  H.groupYoy(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: T,
    prior: [...m25, ...q4], current: m26, priorName: "2025", currentName: "2026", expectPrior: v25, expectCurrent: v26, chart: "line",
    chartLabel: `${dept} net profit by month, 2025 against 2026 ($M)`,
    // a narrow range (parts: $1.9M-$3.2M) steps in half millions; one decimal so no axis label repeats
    ...(Math.max(...m25, ...q4, ...m26) - Math.min(...m25, ...q4, ...m26) < 2e6 ? { axisFormat: '$0.0"M"' } : {}),
    takeaway: band ? band({ d, v25, g25, g26, ex25, ex26 }) : `${kd(d)} (${pcA(d / v25)}) year to date; ${up} of nine months ahead of 2025, largest gain in ${FULL[bi]}, ${kd(dm[bi])}.`,
    source: `${SRC} Jan-Sep 2026 vs Jan-Sep 2025; 2025 shown January-December. Store rows reconcile to the group, both years.${bandNotes ? " " + bandNotes : ""}` });

  H.totalRanker(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: T, rows: S.map(r => [r[0], r[2]]), expect: v26,
    groupLabel: "SEPTEMBER 2026 YTD", takeaway: null, tone: "n", source: `${SRC} Jan-Sep 2026.` });

  // same rules as fixed ops net profit: $ change ranks, % over the size of 2025 alongside, zero-crossing rows shaded
  const pr = S.map(([n, a, b]) => ({ n, d: b - a, p: (b - a) / Math.abs(a) }));
  const val = [...pr].sort((x, y) => y.p - x.p);
  const mk = Object.fromEntries([...val.slice(0, 3).map(r => [r.n, "g"]), ...val.slice(-3).map(r => [r.n, "r"])]);
  const grew = pr.filter(r => r.d > 0).length, bd = [...pr].sort((x, y) => y.d - x.d);
  const bySwing = (p, q) => Math.abs(q[2] - q[1]) - Math.abs(p[2] - p[1]);
  const toP = S.filter(r => r[1] < 0 && r[2] >= 0).sort(bySwing), toL = S.filter(r => r[1] >= 0 && r[2] < 0).sort(bySwing);
  const nmL = a => lst(a.map(r => H.storeName(r[0]))), ba = r => `${H.storeName(r[0])} ${H.kusd(r[1])} to ${H.kusd(r[2])}`;
  const cross = [toP.length ? `${nmL(toP)} turned a 2025 loss into a profit (green rows)` : "", toL.length ? `${nmL(toL)} went from a profit to a loss (red rows)` : ""].filter(Boolean).join("; ");
  H.ranker(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: T,
    deltas: pr.map(r => [r.n, r.d]), groupD: d, fmtD: kd, valLabel: "$", stripValue: `${kd(d)}  ·  ${pcA(d / v25)}`,
    extra: [{ label: "Change", w: 0.95, vals: Object.fromEntries(pr.map(r => [r.n, pcA(r.p)])), marks: mk }],
    groupLabel: "2026 VS 2025 YTD · **YEAR-OVER-YEAR CHANGE**",
    shade: Object.fromEntries([...toP.map(r => [r[0], "g"]), ...toL.map(r => [r[0], "r"])]), tone: "n",
    takeaway: `${cross ? cross + ". " : ""}${grew} of 33 stores grew ${low} net profit, January through September year to date.`,
    source: `${SRC} Ranked by dollar change, Jan-Sep 2026 vs Jan-Sep 2025; percent change alongside, top three and bottom three by percent marked. Percent change = dollar change / 2025 profit; where 2025 was a loss, divided by the size of the loss, so an improvement reads as positive.`,
    notes: `${H.storeName(bd[0].n)} added the most at ${kd(bd[0].d)}; ${H.storeName(bd.at(-1).n)} fell most at ${kd(bd.at(-1).d)}.${toP.length ? ` Green band: loss in 2025, profit in 2026 (${lst(toP.map(ba))}).` : ""}${toL.length ? ` Red band: profit in 2025, loss in 2026 (${lst(toL.map(ba))}).` : ""}` });

  const nowG = sum(Object.values(MN).map(x => x[1])) - sum(Object.values(MN).map(x => x[0]));
  moversSlide(pptx, { seg: SEG, title: T, metric: MET, fmtD: kd, spark: sparkFrom(D.sm, a => a[0]),
    rows: S.map(([n, a, b]) => ({ n, v0: H.kusd(a), v1: H.kusd(b), d: b - a, d2: pcA((b - a) / Math.abs(a)) })),
    now: S.map(([n]) => ({ n, v0: H.kusd(MN[n][0]), v1: H.kusd(MN[n][1]), d: MN[n][1] - MN[n][0], d2: pcA((MN[n][1] - MN[n][0]) / Math.abs(MN[n][0])) })),
    nowGroup: kd(nowG) + " a month", perMonth: true,
    src: `${SRC} Top and bottom five stores by dollar change in ${low} net profit, Jan-Sep 2026 vs 2025 (dollars, not percent: some stores lost money in 2025).` });
  return { v25, v26 };
}

const sv = deptBlock({ D: J("./ntgs.json"), dept: "Service", line: "Mechanical", sales: SL.service,
  levers: [["Gross", "Car count, hours per RO, effective labor rate, labor margin"],
           ["People", "Technicians and advisors: staffing to the workload, pay plans tied to performance"],
           ["Productivity", "Proficiency: more sold hours from the same techs and stalls"],
           ["Spending", "Supplies, outside services, service policy, advertising, loaners"]],
  notesExtra: "Service sales are labor sales on the Mechanical department; parts sold on service ROs are parts department sales." });
const pt = deptBlock({ D: J("./ntgp.json"), dept: "Parts", line: "P&A", sales: SL.parts,
  levers: [["Gross", "Parts on every RO, matrix pricing, wholesale that covers its own cost"],
           ["People", "Counter, wholesale and drivers: staffing to the workload, pay tied to gross"],
           ["Inventory", "Fill rate, obsolescence and carrying cost"],
           ["Spending", "Freight, delivery vehicles, supplies"]],
  notesExtra: "Parts sales include parts sold to the service department on ROs.",
  // Oct 29 (Noah): the explainer is out, so the chart band says why parts profit is flat: what went up and down in
  // P&A department expenses (pulls/pa_expense_lines.json, P&L Detail, ties to gross - profit)
  band: x => { const L = n => PA[n + " - P&A Dept"], ch = n => L(n)[1] - L(n)[0], ka = v => "$" + Math.round(Math.abs(v) / 1e3) + "K";
    if (Math.abs(L("TOTAL EXPENSES")[1] - x.ex26) > 1 || Math.abs(L("TOTAL EXPENSES")[0] - x.ex25) > 1) throw new Error("P&A expense lines do not tie");
    // Oct 29 (Noah): basics only, the dollar detail is in the notes
    if (!(ch("TOTAL PERSONNEL Expense") < 0 && ch("OUTSIDE SERVICES (OTHER)") > 0 && ch("INFO. TECHNOLOGY SERVICES") > 0 && ch("ADVERTISING REBATES") > 0)) throw new Error("parts expense band no longer matches the data");
    return `Gross grew ${pcA(x.g26 / x.g25 - 1).replace("+", "")}, expenses ${pcA(x.ex26 / x.ex25 - 1).replace("+", "")}. People costs went down; outside services, IT and smaller advertising rebates pushed expenses up.`; },
  bandNotes: (() => { const L = n => PA[n + " - P&A Dept"], f = n => `${n.toLowerCase()} ${usd(L(n)[0])} to ${usd(L(n)[1])}`;
    return `P&A department expenses Jan-Sep, 2025 to 2026 (P&L Detail): total ${f("TOTAL EXPENSES")}; personnel ${f("TOTAL PERSONNEL Expense")}; outside services ${f("OUTSIDE SERVICES (OTHER)")}; IT services ${f("INFO. TECHNOLOGY SERVICES")}; advertising ${f("ADVERTISING")}; advertising rebates (a credit) -${usd(L("ADVERTISING REBATES")[0])} to -${usd(L("ADVERTISING REBATES")[1])}; rent ${f("RENT")}.`; })() });
// the two departments add up to the fixed ops net profit block
const N = J("./ntgs.json"), P = J("./ntgp.json");
const fx26 = sum(N.m26.map((r, i) => r[0] + P.m26[i][0]));
if (Math.abs(sv.v26 + pt.v26 - fx26) > 1) throw new Error("service + parts net profit does not tie to fixed ops net profit");
await pptx.writeFile({ fileName: "Service and Parts Net Profit - House Style.pptx" }); console.log("npdept built", Math.round(sv.v26), Math.round(pt.v26));
