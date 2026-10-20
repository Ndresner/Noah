import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
import { load, block, F } from "./blocks.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
/* REBUILT 2026-10-20 (Noah): headcount now from the ADP Employee Master extract (10/19/2026) - the Qlik
   Productivity app's headcount history was missing ~170 people a month in 2025 (see hc/ in the engine). Numerators
   unchanged. Data: prod/*.json from hc/build_prod.py. ORIGINAL NOTE: Qlik Productivity app 948ab4f4 (Accounting space), pulled 2026-10-14. Fixed gross = P&L line Total Gross
   Profit - Fixed Ops. Headcount = monthly snapshot. Fixed Employee job group INCLUDES managers. CP & W ROs =
   distinct CP/W ROs (matches the Closed Repair Orders app exactly). GMC Prescott: body shop gross
   (P&L Gross Profit - Body Shop), body shop titles and body-only ROs removed. */
const SEG = "Fixed Ops Productivity";
const SRC = "Gross and ROs: Qlik - Productivity app (P&L Total Gross Profit - Fixed Ops; CP & W ROs). Headcount: ADP Employee Master extract, 10/19/2026, month-end snapshots, Qlik's fixed-employee title list. Each person counted once: someone open at two stores is counted at the old store until that record closes. Monthly averages, Jan-Sep. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop gross, staff and body-only ROs excluded.";
globalThis.__TRACK = { segments: ["Gross", "Profitability", "CP Engine", "Labor Mix", "Opcodes", "Warranty", "Productivity"], current: 6 };
const R = a => a[0] / a[1];
H.section(pptx, pptx, { kicker: "Fixed Ops", title: "Productivity", chips: ["Gross per employee", "ROs per person", "Headcount"] });
const base = { seg: SEG, src: SRC, basis: "monthly avg \u00b7 same-store, 33 stores" };
const fge = load("./prod/fge.json"), rpa = load("./prod/rpa.json"), rpe = load("./prod/rpe.json"), rpm = load("./prod/rpm.json");
const pc = v => (Math.abs(v) * 100).toFixed(1) + "%";
const mv = (label, a, b) => `${label} ${b >= a ? "grew" : "fell"} ${pc(b / a - 1)}`;   // direction from the data, never typed
block(pptx, fge, { ...base, title: "Fixed Gross per Fixed Employee", metric: "fixed gross per fixed employee", fn: R, f: F.usd,
  def: "Monthly fixed ops gross / fixed employees.",
  extra: ({ G25, G26 }) => `${mv("Gross", G25[0], G26[0])}; ${mv("headcount", G25[1], G26[1])}.`, verb: "raised gross per employee" });
block(pptx, rpa, { ...base, title: "CP & W ROs per Service Advisor", metric: "ros per advisor", fn: R, f: F.int, segment: true,
  def: "Monthly CP & W ROs / service advisors.",
  extra: ({ G25, G26 }) => `${mv("Advisors", G25[1], G26[1])}; ${mv("ROs", G25[0], G26[0])}.`, verb: "raised ROs per advisor" });
block(pptx, rpe, { ...base, title: "CP & W ROs per Fixed Employee", metric: "ros per fixed employee", fn: R, f: F.dec1, segment: true,
  def: "Monthly CP & W ROs / fixed employees.",
  extra: ({ G25, G26 }) => `${mv("Headcount", G25[1], G26[1])}; ${mv("ROs", G25[0], G26[0])}.`, verb: "raised ROs per fixed employee" });
block(pptx, rpm, { ...base, title: "CP & W ROs per Fixed Manager", metric: "ros per fixed manager", fn: R, f: F.int, segment: true,
  def: "Monthly CP & W ROs / fixed managers. Small stores carry one to three managers, so one hire moves the ratio sharply.",
  extra: ({ G25, G26 }) => `${mv("Managers", G25[1], G26[1])}.`, verb: "raised ROs per manager" });

// headcount
const T = JSON.parse(fs.readFileSync("./prod/trend.json", "utf8"));
const i25 = 8, i26 = 20, d = a => a[i26] - a[i25], sg = v => v === 0 ? "0" : (v > 0 ? "+" : "\u2212") + Math.abs(v).toLocaleString("en-US");
const K = 9;   // trailing 12 months: Oct 2025 - Sep 2026
const low = Math.min(...T.fe.slice(K)), lowAt = T.months[K + T.fe.slice(K).indexOf(low)];
H.trend(pptx, pptx, {
  eyebrow: SEG, title: "Fixed Ops Headcount", panelLabel: "Fixed employees, trailing 12 months",
  series: T.months.slice(K).map((m, i) => [m.replace(" ", " '"), T.fe[K + i]]), scale: 1, labelFormat: "#,##0", axisFormat: "#,##0",
  min: Math.floor((Math.min(...T.fe.slice(K)) - 10) / 20) * 20, max: Math.ceil((Math.max(...T.fe.slice(K)) + 10) / 20) * 20, majorUnit: 20, pad: 0,
  bullets: [
    `${sg(d(T.fe))} fixed staff, Sep to Sep, to ${T.fe[i26].toLocaleString("en-US")}.`,
    `Technicians ${sg(d(T.tech))}, advisors ${sg(d(T.sa))}, managers ${sg(d(T.mgr))}.`,
    `Low of ${low.toLocaleString("en-US")} in ${lowAt}; up ${(T.fe[i26] - low).toLocaleString("en-US")} since.`,
  ],
  takeaway: (() => { const g = rpe.G25[0], h = rpe.G26[0];
    return `Fixed ops headcount ${d(T.fe) >= 0 ? "rose" : "fell"} ${Math.abs(d(T.fe))} to ${T.fe[i26].toLocaleString("en-US")} (${sg(d(T.fe)).replace(/\d.*/, "")}${pc(d(T.fe) / T.fe[i25])}), Sep to Sep; CP & W ROs ${h >= g ? "rose" : "fell"} ${pc(h / g - 1)} year to date.`; })(),
  tone: "n",
  source: `${SRC} Fixed employees include managers; technicians = service, express, foreman and team-leader titles.\n`
        + T.months.map((m, i) => `${m}: ${T.fe[i]} fixed, ${T.tech[i]} techs, ${T.sa[i]} advisors, ${T.mgr[i]} managers`).join("\n"),
});
const add = T.S.filter(r => r[2] > r[1]).sort((p, q) => (q[2] - q[1]) - (p[2] - p[1]));
const cut = T.S.filter(r => r[2] < r[1]).sort((p, q) => (p[2] - p[1]) - (q[2] - q[1]));
H.ranker(pptx, pptx, {
  eyebrow: SEG, title: "Fixed Ops Headcount", metric: "fixed headcount",
  rows: T.S, expect: d(T.fe), fmtD: v => sg(v), groupLabel: "SEP 2026 VS SEP 2025",
  takeaway: `${add.length} stores added fixed staff, led by ${add[0][0]} (${sg(add[0][2] - add[0][1])}); ${cut.length} reduced, led by ${cut[0][0]} (${sg(cut[0][2] - cut[0][1])}).`,
  source: `${SRC} Change in fixed ops headcount, Sep 2026 vs Sep 2025, most added first. Neutral colour: added staff can be growth or excess.`,
});
await pptx.writeFile({ fileName: "Fixed Ops Productivity - House Style.pptx" }); console.log("prod built");
