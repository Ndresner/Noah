/* In the deck Oct 24 (Noah; built as a preview first): What Is Fixed Absorption? with what total store fixed expenses include, and a five-slide
   Fixed Ops Net Profit block (explainer, group YOY, store ranker, change ranker, top / bottom / moving now).
   Data: pulls/fixed_exp.json (GM line 258 by account class), ntgs.json / ntgp.json (department profit and gross). */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
import { moversSlide } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const J = f => JSON.parse(fs.readFileSync(f, "utf8"));
const sum = a => a.reduce((t, x) => t + x, 0);
const M = v => "$" + (v / 1e6).toFixed(2) + "M", M1 = v => "$" + (v / 1e6).toFixed(1) + "M",   // one decimal on the fixed expense slide (Noah, Oct 24)
  usd = v => "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
const sgn = v => v >= 0 ? "+" : "−", pcA = v => sgn(v) + Math.abs(v * 100).toFixed(1) + "%", kd = v => sgn(v) + H.kusd(Math.abs(v));

// ---------- 1-2. Fixed absorption ----------
const FX = J("./pulls/fixed_exp.json"), E = FX.y26, fxTot = sum(Object.values(E));
const CAT = [
  { k: "People", c: ["SALARIES SUPERVISION", "OTHER SALARIES AND WAGES", "CLERICAL WAGES", "TAXES PAYROLL", "EMPLOYEE BENEFITS", "ABSENTEE COMP", "RETIREMENT BENEFITS"],
    ex: "Salaries and wages, payroll taxes, benefits, PTO" },
  { k: "Building & equipment", c: ["RENT", "AMORTIZATION LEASEHOLDS", "DEPRECIATION", "TAXES REAL ESTATE", "REPAIRS REAL ESTATE", "REPAIRS EQUIPMENT", "EQUIPMENT RENTAL", "UTILITIES", "INSURANCE"],
    ex: "Rent, depreciation, property taxes, utilities, insurance, building and equipment repairs" },
  { k: "Systems & services", c: ["INFORMATION TECH SVCS", "OUTSIDE SERVICES", "TELEPHONE", "LEGAL AND AUDITING"],
    ex: "DMS and IT, outside services, phones, legal and audit" },
];
const used = new Set(CAT.flatMap(x => x.c));
CAT.push({ k: "Everything else", c: Object.keys(E).filter(k => !used.has(k)),
  ex: "Service and parts advertising, training, company vehicles, freight, supplies, travel, other taxes" });
CAT.forEach(x => { x.v = sum(x.c.map(c => E[c])); });
if (Math.abs(sum(CAT.map(x => x.v)) - fxTot) > 0.01) throw new Error("fixed expense categories do not reconcile");
const ABS = J("./abs.json"), absG = sum(ABS.m26.map(r => r[0])), absE = sum(ABS.m26.map(r => r[1]));
if (Math.abs(absE - fxTot) > 1) throw new Error(`fixed expenses ${fxTot} vs absorption denominator ${absE}`);
const SRCA = "Qlik - Profit & Loss, GM statement. Jan-Sep 2026, same-store, INEOS excluded (opened 2026).";

H.explainer(pptx, pptx, { source: SRCA, eyebrow: "Fixed Ops Profitability", title: "Fixed Absorption",
  formula: ["Fixed absorption", "=", "fixed ops gross", "÷", "total store fixed expenses"],
  // Noah, Oct 27
  define: "How much of the dealership's fixed expenses are covered by parts and service gross profit.",
  example: [["At 100%", "fixed ops covers every fixed expense"], ["Group today", (absG / absE * 100).toFixed(1) + "%"]], exampleLabel: "READING IT",
  why: ["At 100%, the store covers its fixed costs before it sells a car.",
        "It protects the store when car sales slow.",
        "It is the clearest measure of fixed ops' value to the dealership."],
  // Oct 26 (Noah): slides 27-28 combined. The lever box shows what is and isn't in total store fixed expenses
  // (the old breakdown slide is out of the flow); no dollar detail, just what they are.
  leverHead: "WHAT'S IN TOTAL STORE FIXED EXPENSES",
  levers: [["Included: people", "Salaries and wages, payroll taxes, benefits, PTO: every department"],
           ["Included: building", "Rent, depreciation, property taxes, utilities, insurance, repairs"],
           ["Included: systems and other", "DMS and IT, outside services, phones, training, supplies"],
           ["Not included", "Sales and F&I commissions, delivery, floorplan interest, new and used vehicle advertising"]],
  leverRows: [1, 1.2],
  ours: `Group: ${(absG / absE * 100).toFixed(1)}% absorption, up 2.1 pts. Fixed gross grew 4.1%; fixed expenses 1.3%.`,
  tone: "g", notes: `PLACE: Profitability, in front of Fixed Ops Absorption %; the next slide breaks down the fixed expenses. The denominator is the P&L's Fixed Expenses line on the GM statement: the whole store's expenses (new, used, service, parts, body shop) less the costs that move with each car sold: sales and F&I commissions, delivery, policy, floorplan interest, and new and used vehicle advertising. That matches the standard dealer definition: fixed ops gross divided by total store overhead, excluding the variable costs of selling cars. Jan-Sep 2026 same-store: ${usd(fxTot)} of fixed expenses out of ${usd(FX.total_expense_26)} total expense.` });

H.statTiles(pptx, pptx, { eyebrow: "Fixed Ops Profitability", title: "Total Store Fixed Expenses", bigSize: 36, subSize: 14,
  tiles: CAT.map(x => ({ label: x.k, big: M1(x.v), sub: `${(x.v / fxTot * 100).toFixed(0)}% · ${x.ex}`, tone: "n" })),
  banner: `Not included: sales and F&I commissions, delivery, policy, floorplan interest and new and used vehicle advertising, the costs that rise and fall with each car sold.`,
  takeaway: `${M1(fxTot)} of fixed expenses Jan-Sep 2026; fixed ops gross of ${M1(absG)} covers ${(absG / absE * 100).toFixed(1)}% of it. People are ${(CAT[0].v / fxTot * 100).toFixed(0)}% of the total.`,
  tone: "n", source: `${SRCA} Fixed Expenses line by account class, every department of the store; categories reconcile to the line to the cent (${usd(fxTot)}).` });

// ---------- 3-7. Fixed ops net profit ----------
const N = J("./ntgs.json"), P = J("./ntgp.json");
const SEG = "Fixed Ops Net Profit", T = "Fixed Ops Net Profit", MET = "net profit";
const SRC = "Qlik - Profit & Loss, Detail statement. Same-store, INEOS excluded (opened 2026). Fixed ops net profit = service (Mechanical) + parts (P&A) department profit: department gross less the department's own expenses; GMC Prescott body shop is its own department and is not included.";
const m25 = N.m25.map((r, i) => r[0] + P.m25[i][0]), m26 = N.m26.map((r, i) => r[0] + P.m26[i][0]), q4 = N.q4.map((r, i) => r[0] + P.q4[i][0]);
const g25gross = sum(N.m25.map((r, i) => r[1] + P.m25[i][1])), g26gross = sum(N.m26.map((r, i) => r[1] + P.m26[i][1]));
const v25 = sum(m25), v26 = sum(m26), d = v26 - v25;
const PS = Object.fromEntries(P.S.map(r => [r[0], r]));
const S = N.S.map(r => [r[0], r[1] + PS[r[0]][1], r[3] + PS[r[0]][3]]);   // [name, profit 2025, profit 2026]
if (Math.abs(sum(S.map(r => r[1])) - v25) > 1 || Math.abs(sum(S.map(r => r[2])) - v26) > 1) throw new Error("net profit stores do not reconcile");
const PM = Object.fromEntries(P.mn.map(r => [r[0], r]));
const MN = Object.fromEntries(N.mn.map(r => [r[0], [(r[1] + PM[r[0]][1]) / 3, (r[3] + PM[r[0]][3]) / 3]]));   // Jan-Mar, Jul-Sep monthly
const ex25 = g25gross - v25, ex26 = g26gross - v26;
// revenue needed per $1 of net (Noah, Oct 24): service + parts net sales / fixed ops net profit
const SL = J("./pulls/net_by_dept.json").fx_sales_26, sales26 = SL.service + SL.parts, per1 = sales26 / v26;

H.explainer(pptx, pptx, { source: SRC, eyebrow: SEG, title: "Fixed Ops Net Profit",
  formula: ["Fixed ops net profit", "=", "service & parts gross", "−", "service & parts expenses"], formulaSize: 22,
  // Noah, Oct 27
  define: "The profit parts and service generate after their allocated expenses are covered.",
  example: [["Gross", M(g26gross)], ["− Expenses", M(ex26)], ["= Net profit", M(v26)]], exampleLabel: "OUR NUMBERS",
  why: ["It is what fixed ops actually contributes to the store.",
        "Gross can grow while profit shrinks if expenses grow faster.",
        `Every dollar of expense saved drops straight to the bottom line. Earning that $1 takes $${Math.round(per1)} in service and parts sales.`],   // rounded to $7 (Noah, Oct 26)
  levers: [["Gross", "Car count, hours per RO, effective labor rate, parts and labor margin"],
           ["People", "The largest expense: staffing to the workload, pay plans tied to performance"],
           ["Productivity", "More gross from the same team and the same stalls"],
           ["Spending", "Supplies, outside services, service policy, advertising, company vehicles"]],
  ours: `Group: ${M(v26)} fixed ops net profit, ${pcA(v26 / v25 - 1)}. Gross grew ${pcA(g26gross / g25gross - 1)}; department expenses ${pcA(ex26 / ex25 - 1)}.`,
  tone: "g", notes: `PLACE: opening, right after Why We Are Here (Noah, Oct 24): the high-level view of service and parts before the sections. Department profit as the P&L Detail statement reports it (DEPARTMENT PROFIT OR LOSS, Mechanical and P&A). Jan-Sep 2026 gross ${usd(g26gross)}, department expenses ${usd(ex26)}, profit ${usd(v26)}; 2025 ${usd(g25gross)} / ${usd(ex25)} / ${usd(v25)}. Store-level costs the P&L does not split by department (management fees, other income) sit outside it. Same basis as the service and parts net to gross slides. Revenue per $1 of net: service + parts net sales ${usd(sales26)} (P&L Detail NET SALES, Mechanical ${usd(SL.service)} + P&A ${usd(SL.parts)}) / net profit ${usd(v26)} = $${per1.toFixed(2)}; net profit is ${(v26 / sales26 * 100).toFixed(1)} cents of every sales dollar. Parts sales include parts sold to the service department.` });

const FULL = ["January","February","March","April","May","June","July","August","September"];
const dm = m26.map((v, i) => v - m25[i]), up = dm.filter(x => x > 0).length, bi = dm.indexOf(Math.max(...dm));
H.groupYoy(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: "Fixed Ops Net Profit",
  prior: [...m25, ...q4], current: m26, priorName: "2025", currentName: "2026", expectPrior: v25, expectCurrent: v26, chart: "line",
  chartLabel: "Fixed ops net profit by month, 2025 against 2026 ($M)",
  takeaway: `${kd(d)} (${pcA(d / v25)}) year to date; ${up} of nine months ahead of 2025, best ${FULL[bi]} at ${kd(dm[bi])}.`,
  source: `${SRC} Jan-Sep 2026 vs Jan-Sep 2025; 2025 shown January-December. Store rows reconcile to the group, both years.` });

const loss = S.filter(r => r[2] < 0).length;
H.totalRanker(pptx, pptx, { eyebrow: SEG, title: T, metric: MET, measure: "Fixed Ops Net Profit", rows: S.map(r => [r[0], r[2]]), expect: v26,
  groupLabel: "SEPTEMBER 2026 YTD",
  // Oct 27 (Noah): no band - the losing-money count is plain from the bars, and the store call-outs are off every band
  takeaway: null,
  tone: "n", source: `${SRC} Jan-Sep 2026.` });

// $ change ranks (some 2025 bases are losses, so a percent change does not read); % alongside only where 2025 made money
// % change over the size of 2025 (Noah, Oct 24): a store that lost money in 2025 divides by the loss as a positive
// number, so an improvement always reads +, e.g. Lincoln -$389K -> +$94K = +124.3%
const pr = S.map(([n, a, b]) => ({ n, d: b - a, p: (b - a) / Math.abs(a) }));
const val = [...pr].sort((x, y) => y.p - x.p);
const mk = Object.fromEntries([...val.slice(0, 3).map(r => [r.n, "g"]), ...val.slice(-3).map(r => [r.n, "r"])]);
const grew = pr.filter(r => r.d > 0).length, bd = [...pr].sort((x, y) => y.d - x.d);
const chg = (deck, x = {}) => H.ranker(pptx, deck, { eyebrow: SEG, title: T, metric: MET, measure: "Fixed Ops Net Profit",
  deltas: pr.map(r => [r.n, r.d]), groupD: d, fmtD: kd, valLabel: "$", stripValue: `${kd(d)}  ·  ${pcA(d / v25)}`,
  extra: [{ label: "Change", w: 0.95, vals: Object.fromEntries(pr.map(r => [r.n, pcA(r.p)])), marks: mk }],
  groupLabel: "2026 VS 2025 YTD · **YEAR-OVER-YEAR CHANGE**",
  takeaway: `${grew} of 33 stores grew fixed ops net profit. ${H.storeName(bd[0].n)} added the most at ${kd(bd[0].d)}; ${H.storeName(bd.at(-1).n)} fell most at ${kd(bd.at(-1).d)}.`,
  source: `${SRC} Ranked by dollar change, Jan-Sep 2026 vs Jan-Sep 2025; percent change alongside, top three and bottom three by percent marked. Percent change = dollar change / 2025 profit; where 2025 was a loss, divided by the size of the loss, so an improvement reads as positive.`, ...x });
// stores whose fixed ops net profit crossed zero (Noah, Oct 24: option A, shaded rows): green band = 2025 loss to
// 2026 profit, red band = profit to loss; the takeaway names them, biggest swing first
const bySwing = (p, q) => Math.abs(q[2] - q[1]) - Math.abs(p[2] - p[1]);
const toP = S.filter(r => r[1] < 0 && r[2] >= 0).sort(bySwing), toL = S.filter(r => r[1] >= 0 && r[2] < 0).sort(bySwing);
const lst = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a.at(-1);
const nmL = a => lst(a.map(r => H.storeName(r[0])));
const ba = r => `${H.storeName(r[0])} ${H.kusd(r[1])} to ${H.kusd(r[2])}`;
chg(pptx, { shade: Object.fromEntries([...toP.map(r => [r[0], "g"]), ...toL.map(r => [r[0], "r"])]), tone: "n",
  takeaway: `${nmL(toP)} turned a 2025 loss into a profit (green rows)${toL.length ? `; ${nmL(toL)} went from a profit to a loss (red row)` : ""}. ${grew} of 33 stores grew fixed ops net profit, January through September year to date.`,   // Noah, Oct 27
  notes: `Green band: lost money in fixed ops in 2025, made money in 2026 (${lst(toP.map(ba))}). Red band: made money in 2025, lost money in 2026 (${lst(toL.map(ba))}).` });

const nowG = sum(Object.values(MN).map(x => x[1])) - sum(Object.values(MN).map(x => x[0]));
moversSlide(pptx, { seg: SEG, title: T, metric: MET, fmtD: kd,
  rows: S.map(([n, a, b]) => ({ n, v0: H.kusd(a), v1: H.kusd(b), d: b - a })),
  now: S.map(([n]) => ({ n, v0: H.kusd(MN[n][0]), v1: H.kusd(MN[n][1]), d: MN[n][1] - MN[n][0] })),
  nowGroup: kd(nowG) + " a month", perMonth: true,
  src: `${SRC} Top and bottom five stores by dollar change in fixed ops net profit, Jan-Sep 2026 vs 2025 (dollars, not percent: several stores lost money in 2025).` });

await pptx.writeFile({ fileName: "Fixed Absorption & Net Profit - House Style.pptx" }); console.log("np built");
