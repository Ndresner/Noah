/* Op code missed opportunity (2026-10-19, for review). Qlik - Closed Repair Orders, customer pay, Jan-Sep 2026,
   op code report basis (flagExcludeOpCodeReport = 0), same-store (INEOS excluded). Reconciles to the Opcodes slides:
   279,352 CP ROs; group shortfalls 4,821 / 18,493 / 1,985 / 4,731. Method as in the Opcode Benchmark Report:
   services short = benchmark % x CP ROs - services sold; valued at the store's own labor + parts gross per service (group
   average where the store sold fewer than 30). Stores at or above a benchmark are left off that slide. Year = x 12/9. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const R = fs.readFileSync("./prof/op_raw.txt", "utf8").trim().split("\n").map(l => { const p = l.split("|"); return [p[0], ...p.slice(1).map(Number)]; });
const SV = [["Alignment", 0.15], ["Tire Balance", 0.15], ["Brake Fluid", 0.10], ["Fuel System Service", 0.075]];
const YR = 12 / 9, MIN = 30, CAP = 2;   // gross per service capped at 2x the group average: some flagged lines carry bundled work (LR Reno $1,631 per fuel service)
const usd = v => "$" + Math.round(v).toLocaleString("en-US"), int = v => Math.round(v).toLocaleString("en-US"), pc = v => (Math.floor(v * 1000) / 10).toFixed(1) + "%";
const SRC = "Qlik - Closed Repair Orders. Customer pay, Jan-Sep 2026 by close date, op code report basis. Same-store, INEOS excluded (opened 2026). Services short = benchmark x the store's CP ROs, minus the services it sold. Each service valued at the store's own labor + parts gross per service sold (group average where the store sold fewer than 30), capped at twice the group average because some flagged lines carry other work bundled on them. A year = Jan-Sep x 12 / 9. Stores at or above the benchmark are not shown and do not offset anyone.";
const res = SV.map(([name, b], i) => {
  const groupGPJ = R.reduce((t, r) => t + r[6 + i], 0) / R.reduce((t, r) => t + r[2 + i], 0);
  const rows = R.map(r => { const D = r[1], n = r[2 + i], raw = n >= MIN ? r[6 + i] / n : groupGPJ, gpj = Math.min(raw, CAP * groupGPJ), capped = raw > CAP * groupGPJ, short = Math.max(0, b * D - n);
    return { n: r[0], rate: n / D, short: short * YR, gpj, raw, capped, own: n >= MIN, miss: short * gpj * YR }; }).filter(x => x.rate < b);
  return { name, b, groupGPJ, rows, miss: rows.reduce((t, x) => t + x.miss, 0), short: rows.reduce((t, x) => t + x.short, 0) };
});
const total = res.reduce((t, s) => t + s.miss, 0);

// the one big number (ADDED 2026-10-19)
{
  const s = pptx.addSlide(); H.page(pptx, s, { eyebrow: "Opcodes", title: "Every Store at Benchmark" });
  H.panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const t = (txt, o) => s.addText(H.words(txt, "hero"), { fontFace: H.F, margin: 0, isTextBox: true, align: "center", ...o });
  t("IF EVERY STORE REACHED ALL FOUR BENCHMARKS", { x: 0.8, y: 1.62, w: 11.73, h: 0.36, fontSize: H.sz(16, "hero kicker"), bold: true, color: C.navy, charSpacing: 1.4 });
  t(usd(total), { x: 0.8, y: 2.05, w: 11.73, h: 1.35, fontSize: H.sz(80, "hero number"), bold: true, color: C.goodInk });
  t(`a year in labor and parts gross \u00b7 about ${usd(total / 12)} a month`, { x: 0.8, y: 3.38, w: 11.73, h: 0.45, fontSize: H.sz(20, "hero sub"), color: C.ink600 });
  const W = 2.78, G = 0.2, X0 = 0.8 + (11.73 - (W * 4 + G * 3)) / 2;
  res.forEach((r, i) => {
    const x = X0 + i * (W + G);
    s.addShape(pptx.ShapeType.rect, { x, y: 4.15, w: W, h: 1.85, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y: 4.15, w: W, h: 0.07, fill: { color: C.good } });
    t(`${r.name.replace("Fuel System Service", "Fuel Service").toUpperCase()} \u00b7 ${(r.b * 100).toFixed(r.b === 0.075 ? 1 : 0)}%`, { x: x + 0.1, y: 4.32, w: W - 0.2, h: 0.36, fontSize: H.sz(14, "chip label"), bold: true, color: C.ink600, charSpacing: 0.6 });
    t(usd(r.miss), { x: x + 0.1, y: 4.72, w: W - 0.2, h: 0.6, fontSize: H.sz(26, "chip number"), bold: true, color: C.navy });
    t(`${int(r.short)} more services a year`, { x: x + 0.1, y: 5.35, w: W - 0.2, h: 0.4, fontSize: H.sz(14, "chip sub"), color: C.ink });
  });
  H.takeaway(pptx, s, `${int(res.reduce((a, r) => a + r.short, 0))} more services a year on cars already in our drives. No new customers needed.`, "g");
  H.foot(pptx, s, SRC);
  s.addNotes("Talk track: every one of these cars is already on a customer-pay RO in our drive. The inspection and the menu find the need; the advisor presents it.");
}

// total by store (ADDED 2026-10-19): the four services added together, only where the store is under that benchmark
{
  const T = {};
  res.forEach(r => r.rows.forEach(x => { const t = (T[x.n] ??= { miss: 0, short: 0, k: 0 }); t.miss += x.miss; t.short += x.short; t.k += 1; }));
  const rows = Object.entries(T), top = [...rows].sort((a, b) => b[1].miss - a[1].miss), all4 = rows.filter(([, t]) => t.k === 4).length;
  H.amountRanker(pptx, pptx, {
    eyebrow: "Opcodes", title: "Total Missed Opportunity by Store",
    rows: rows.map(([n, t]) => [n, t.miss]), fmt: usd, color: C.bad, ink: C.badInk, valLabel: "Gross/yr", valW: 0.98,
    extra: [{ label: "Below", w: 0.66, vals: Object.fromEntries(rows.map(([n, t]) => [n, `${t.k} of 4`])) },
            { label: "Services", w: 0.92, vals: Object.fromEntries(rows.map(([n, t]) => [n, int(t.short)])) }],
    takeaway: `${top[0][0]} leaves the most at ${usd(top[0][1].miss)} a year. ${all4} stores are under all four benchmarks.`,
    tone: "b", source: SRC + " Alignment, tire balance, brake fluid and fuel system service added together. Below = how many of the four benchmarks the store is under. Services = services short a year.",
  });
  pptx._slides[pptx._slides.length - 1].addNotes(top.map(([n, t]) => `${n}: ${usd(t.miss)} a year, ${int(t.short)} services, under ${t.k} of 4`).join("\n"));
}

// one ranker per service
res.forEach(s => {
  const top = [...s.rows].sort((a, b) => b.miss - a.miss);
  H.amountRanker(pptx, pptx, {
    eyebrow: "Opcodes", title: `${s.name} Missed Opportunity`,
    rows: s.rows.map(x => [x.n, x.miss]), fmt: usd, color: C.bad, ink: C.badInk, valLabel: "Gross/yr", valW: 0.98,
    extra: [{ label: "Rate", w: 0.62, vals: Object.fromEntries(s.rows.map(x => [x.n, pc(x.rate)])) },
            { label: "Services", w: 0.96, vals: Object.fromEntries(s.rows.map(x => [x.n, int(x.short)])) }],
    takeaway: `${s.rows.length} stores under ${(s.b * 100).toFixed(s.b === 0.075 ? 1 : 0)}% leave ${usd(s.miss)} a year: ${int(s.short)} services. ${top[0].n} has the most at ${usd(top[0].miss)}.`,
    tone: "b", source: SRC + ` ${s.name}: group gross per service ${usd(s.groupGPJ)}. Rate = Jan-Sep closing %. Services = services short a year.`,
  });
  pptx._slides[pptx._slides.length - 1].addNotes("Near-zero stores are included on purpose (Noah, Oct 19 2026): some stores sell little of a service by OEM recommendation or local habit (e.g. Spokane customers swap seasonal tires, so fewer balances), but not every customer fits that pattern, so the opportunity is real. A store with a high RO count can show more opportunity than one selling at zero.\n" + s.rows.map(x => `${x.n}: ${pc(x.rate)}, ${int(x.short)} services short a year x ${usd(x.gpj)}/service${x.own ? "" : " (group avg)"} = ${usd(x.miss)}${x.capped ? ` (capped; raw ${usd(x.raw)}/service)` : ""}`).join("\n"));
});
await pptx.writeFile({ fileName: "Op Code Missed Opportunity - Review.pptx" });
console.log("total", usd(total), res.map(s => `${s.name} ${s.rows.length} ${usd(s.miss)} ${int(s.short)} gpj ${usd(s.groupGPJ)}`).join(" | "));
res.forEach(s => console.log(s.name, s.rows.filter(x => x.own).sort((a, b) => b.gpj - a.gpj).slice(0, 3).map(x => `${x.n} ${usd(x.gpj)}`).join(", ")));
