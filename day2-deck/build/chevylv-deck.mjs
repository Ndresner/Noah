/* Chevy LV: What's Behind the Numbers (Noah, Nov 3 2026: store follow-up after the Day 2 GM meeting). Jan-Sep 2026 vs 2025.
   Every figure is computed from the saved Qlik pulls in pulls/chevylv/ (Closed Repair Orders, labor type and op code detail);
   deck figures (ranks, door rate) come from the deck's own data files. Same basis as the deck: same-store, INEOS out,
   GMC Prescott body shop labor types out. Companion workbook: reports/Chevy LV - What's Behind the Numbers (Jan-Sep 2026).xlsx. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const J = f => JSON.parse(fs.readFileSync(f, "utf8"));
const ST = "CHEVY LV", EB = "Chevy LV · Jan–Sep 2026";
const SRC = "Qlik - Closed Repair Orders, Jan-Sep 2026 vs 2025 by close date, labor type and op code detail. Same basis as the Day 2 deck (same-store, GMC Prescott body shop labor types out).";

// ---- data
const LT = J("./pulls/chevylv/lt_rows.json");                       // [store, paytype, labortype, laborsale, hrs, laborcost, partssale, partsgross, ROs] 2026
const rows = (o) => o.data.data.map(r => ({ pt: r[0], lt: r[1], ls: r[2], hrs: r[3], ps: r[r.length - 3], pg: r[r.length - 2], ro: r[r.length - 1] }));
const Y25 = rows(J("./pulls/chevylv/chevylv_labortype_2025.json")), Y26 = rows(J("./pulls/chevylv/chevylv_labortype_2026.json"));
const PO = J("./pulls/chevylv/chevylv_prepaid_only.json").data.data[0];   // CP ROs 26, prepaid-only 26, CP ROs 25, prepaid-only 25, any prepaid 26
const DET = J("./pulls/chevylv/internal_detail_by_store_2026.json");       // [store, int sale, int hrs, detail sale, detail hrs]
const DOOR = J("./door.json")[ST];
const M5 = J("./m5.json").S.find(r => r[0] === ST);                   // [.., 2026: ROs 7, labor 8, hrs 9, .., parts sale 11, parts gross 12]
const sum = (a, f) => a.reduce((t, x) => t + f(x), 0);
const g = (Y, pt, lt) => Y.find(r => r.pt === pt && r.lt === lt);

// reconcile to the deck before anything is shown
const cp26 = LT.filter(r => r[0] === ST && r[1] === "C");
const cpLS = sum(cp26, r => r[3]), cpH = sum(cp26, r => r[4]), cpPS = sum(cp26, r => r[6]), cpPG = sum(cp26, r => r[7]);
if (Math.abs(cpLS - M5[8]) > 1 || Math.abs(cpH - M5[9]) > 0.5 || PO[0] !== M5[7]) throw new Error("Chevy LV CP does not tie to m5.json");

// ---- figures
const pm = g(Y26, "C", "CPPM"), em = g(Y26, "C", "CEMP"), rt26 = g(Y26, "C", "CP"), rt25 = g(Y25, "C", "CP"), ex26 = g(Y26, "C", "CEXT"), ex25 = g(Y25, "C", "CEXT");
const elr = cpLS / cpH, elrAdj = (cpLS - pm.ls - em.ls) / (cpH - pm.hrs - em.hrs), rtElr = rt26.ls / rt26.hrs;
const pgp = cpPG / cpPS, rtPgp = rt26.pg / rt26.ps, exShare = ex26.ps / cpPS;
const lpr = cpLS / PO[0], lprAdj = (cpLS - pm.ls) / (PO[0] - PO[1]), hpr = cpH / PO[0], hprAdj = (cpH - pm.hrs) / (PO[0] - PO[1]);
const pmHrsShare = (pm.hrs + em.hrs) / cpH;
// retail ELR rank: each store's largest-hours CP labor type
const stores = [...new Set(LT.map(r => r[0]))];
const retail = stores.map(s => { const r = LT.filter(x => x[0] === s && x[1] === "C").sort((a, b) => b[4] - a[4])[0]; return [s, r[3] / r[4]]; }).sort((a, b) => b[1] - a[1]);
const rtRank = retail.findIndex(r => r[0] === ST) + 1;
const ann = 12 / 9, per10 = rt26.hrs * ann * 10, to70 = (0.7 * DOOR - rtElr) * rt26.hrs * ann;
// internal detail
const d = DET.find(r => r[0] === ST), gS = sum(DET, r => r[1]), gH = sum(DET, r => r[2]), gDS = sum(DET, r => r[3]), gDH = sum(DET, r => r[4]);
const intElr = d[1] / d[2], detShare = d[4] / d[2], detElr = d[3] / d[4], intAdj = (d[1] - d[3]) / (d[2] - d[4]);
const rankOf = (f) => [...DET].sort((a, b) => f(b) - f(a)).findIndex(r => r[0] === ST) + 1;
const intRank = rankOf(r => r[1] / r[2]), intAdjRank = rankOf(r => (r[1] - r[3]) / (r[2] - r[4])), detRank = rankOf(r => r[4] / r[2]);
// retail volume
const rtHrsChg = rt26.hrs / rt25.hrs - 1, rtRoChg = rt26.ro / rt25.ro - 1, rtLsChg = rt26.ls - rt25.ls;
const wg26 = g(Y26, "W", "WGM"), wg25 = g(Y25, "W", "WGM");

const $ = v => "$" + Math.round(v).toLocaleString("en-US"), $2 = v => "$" + v.toFixed(2), pc = v => (v * 100).toFixed(1) + "%", n0 = v => Math.round(v).toLocaleString("en-US");
const K = v => "$" + Math.round(v / 1000) + "K";
const ord = n => n + (["th", "st", "nd", "rd"][((n % 100) - 20) % 10] || ["th", "st", "nd", "rd"][n % 100] || "th");

// 1. overview
H.talkingPoints(pptx, pptx, {
  eyebrow: EB, title: "What's Behind the Numbers", saidFrom: null, source: SRC,
  said: [
    "Some numbers are shaped by how the store runs, not how well it runs",
    `Detail runs through internal: ${pc(detShare)} of internal hours, the most in the group`,
    `Prepaid maintenance and employee work: ${pc(pmHrsShare)} of customer-pay hours at low rates`,
    `Service contracts: ${pc(exShare)} of customer-pay parts sales, at contract margin`,
    `Take those out and CP ELR is ${pc(elrAdj / DOOR)} of door, above the 70% benchmark`,
    "The real gaps: the retail labor rate and retail repair volume",
  ],
  now: [
    { big: $(rtElr), label: `Retail customer-pay ELR: ${pc(rtElr / DOOR)} of the ${$2(DOOR)} door, ${ord(rtRank)} of 33.`, tone: "b" },
    { big: (rtHrsChg * 100).toFixed(1) + "%", label: `Retail customer-pay hours vs 2025; retail labor sale ${rtLsChg < 0 ? "−" : "+"}${K(Math.abs(rtLsChg))}.`, tone: "b" },
    { big: K(per10), label: "A year for every $10 an hour on retail customer-pay hours.", tone: "n" },
  ],
  ask: "pick the retail rate and retail repair volume levers, put an owner on each, and track them weekly.",
  notes: "Talk track: before we work the opportunities, separate what is process from what is performance. Detail in internal, prepaid maintenance and service contracts all move the numbers without being a problem. Once they are out, two gaps remain: what we charge on retail work and how much retail repair work we are writing.",
});

// 2. internal
H.statTiles(pptx, pptx, { eyebrow: EB, title: "Internal ELR: Detail Runs Through Internal", bigSize: 40, subSize: 15,
  tiles: [
    { label: "Internal ELR", big: $(intElr), sub: `${ord(intRank)} of 33. Group ${$(gS / gH)}.`, tone: "b" },
    { label: "Detail share of internal hours", big: pc(detShare), sub: `${n0(d[4])} hours at about ${$(detElr)} an hour. Group ${pc(gDH / gH)}; ${detRank === 1 ? "highest in the group" : ord(detRank) + " highest"}.`, tone: "n" },
    { label: "Internal ELR without detail", big: $(intAdj), sub: `${ord(intAdjRank)} of 33. Group without detail ${$((gS - gDS) / (gH - gDH))}.`, tone: "n" },
  ],
  takeaway: "In-house detail explains most of the internal ELR gap. The rest is the rate charged to the used-car department on mechanical recon.",
  tone: "n", source: SRC + " Detail = internal lines whose op code description contains DETAIL, WASH, BUFF, CLEAN UP, CLEANUP, SHAMPOO or POLISH.",
});

// 3. customer pay process effects
H.statTiles(pptx, pptx, { eyebrow: EB, title: "Customer Pay: Process Effects on the Numbers", bigSize: 36, subSize: 14,
  tiles: [
    { label: "CP ELR without prepaid & employee", big: $2(elrAdj), sub: `${pc(elrAdj / DOOR)} of door (deck: ${$2(elr)}). Prepaid: ${n0(pm.hrs)} hrs at ${$(pm.ls / pm.hrs)}/hr; employee: ${n0(em.hrs)} hrs at ${$(em.ls / em.hrs)}/hr.`, tone: "g" },
    { label: "Retail CP parts GP %", big: pc(rtPgp), sub: `Deck: ${pc(pgp)}. Service contracts are ${pc(exShare)} of CP parts sales at ${pc(ex26.pg / ex26.ps)}.`, tone: "n" },
    { label: "CP ROs without prepaid-only visits", big: n0(PO[0] - PO[1]), sub: `${pc((PO[0] - PO[1]) / (PO[2] - PO[3]) - 1)} vs 2025. Prepaid-only visits fell ${n0(PO[3] - PO[1])} of the ${n0(PO[2] - PO[0])} fewer CP ROs.`, tone: "n" },
    { label: "CP labor $ per RO without prepaid-only", big: $(lprAdj), sub: `Deck: ${$(lpr)}; hours per RO ${hprAdj.toFixed(2)} (deck ${hpr.toFixed(2)}). Service-contract jobs run ${(ex26.hrs / ex26.ro).toFixed(1)} hrs each.`, tone: "n" },
  ],
  takeaway: "Prepaid maintenance, employee work and service contracts move these numbers. None of them is a problem; knowing it keeps the focus on the real gaps.",
  tone: "n", source: SRC + ` Prepaid = labor type CPPM, employee = CEMP, service contract = CEXT. Door rate ${$2(DOOR)} (deck). ROs per labor type are ROs with at least one line of that type.`,
});

// 4. retail rate
H.statTiles(pptx, pptx, { eyebrow: EB, title: "Opportunity: The Retail Labor Rate", bigSize: 40, subSize: 15,
  tiles: [
    { label: "Retail customer-pay ELR", big: $(rtElr), sub: `${pc(rtElr / DOOR)} of the ${$2(DOOR)} door; ${ord(rtRank)} of 33 retail rates. 2025: ${$(rt25.ls / rt25.hrs)}.`, tone: "b" },
    { label: "70% of door", big: $(0.7 * DOOR), sub: `Closing the gap on retail hours is about ${K(to70)} a year.`, tone: "n" },
    { label: "Every $10 an hour", big: K(per10), sub: `A year, on ${n0(rt26.hrs * ann)} retail customer-pay hours.`, tone: "n" },
  ],
  takeaway: "The mix doesn't explain the retail rate. Maintenance and menu pricing, use of the labor grid and discounts decide it.",
  tone: "n", source: SRC + " Retail = labor type CP, the store's largest customer-pay labor type; each store ranked on its own largest. Annualized = Jan-Sep x 12/9.",
});

// 5. retail volume
H.statTiles(pptx, pptx, { eyebrow: EB, title: "Opportunity: Retail Repair Volume", bigSize: 40, subSize: 15,
  tiles: [
    { label: "Retail customer-pay hours", big: (rtHrsChg * 100).toFixed(1) + "%", sub: `${n0(rt25.hrs)} to ${n0(rt26.hrs)} hours.`, tone: "b" },
    { label: "ROs with retail labor", big: (rtRoChg * 100).toFixed(1) + "%", sub: `${n0(rt25.ro)} to ${n0(rt26.ro)}. Hours per retail RO ${(rt25.hrs / rt25.ro).toFixed(2)} to ${(rt26.hrs / rt26.ro).toFixed(2)}.`, tone: "b" },
    { label: "Retail labor sale", big: (rtLsChg < 0 ? "−" : "+") + K(Math.abs(rtLsChg)), sub: `Service contract hours ${((ex26.hrs / ex25.hrs - 1) * 100).toFixed(1)}%. GM warranty labor ${wg26.ls - wg25.ls < 0 ? "−" : "+"}${K(Math.abs(wg26.ls - wg25.ls))}.`, tone: "b" },
  ],
  takeaway: "This is where the service gross went: fewer retail repair ROs and less on each. Road to the Sale, MPI on every car, declined-service follow-up.",
  tone: "b", source: SRC + " Retail = labor type CP.",
});

// 6. questions
H.talkingPoints(pptx, pptx, {
  eyebrow: EB, title: "Questions and Next Steps", saidFrom: null, source: SRC,
  said: [
    `What is internal labor type ISP? ${n0(g(Y26, "I", "ISP").hrs)} hours at ${$(g(Y26, "I", "ISP").ls / g(Y26, "I", "ISP").hrs)} an hour`,
    "What does the used-car department pay for mechanical recon, against the door rate?",
    `Prepaid visits fell ${n0(PO[3] - PO[1])}: fewer plans sold with vehicles, or plans expiring?`,
    `${n0(PO[4])} prepaid visits: track brake fluid, fuel and balance closing on them separately`,
    "Retail rate: menu pricing, labor grid and discounts reviewed weekly",
    "Retail volume: MPI on every car, declined-service follow-up",
  ],
  now: [
    { big: "$" + (wg26.ls / 1e6).toFixed(2) + "M", label: `GM warranty labor, up ${K(wg26.ls - wg25.ls)} on 2025 at ${$(wg26.ls / wg26.hrs)} an hour.`, tone: "g" },
    { big: $(rtElr) + " → " + $(0.7 * DOOR), label: "Retail ELR today and at 70% of door.", tone: "n" },
  ],
  ask: "agree an owner and a weekly number for the retail rate and retail repair volume.",
});

await pptx.writeFile({ fileName: "Chevy LV - What's Behind the Numbers.pptx" }); console.log("chevy lv built");
