import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F, sz, words } = H;
/* Warranty slides: rate increases (vendor report) then audits (DRAFT, Oct 22 2026, Noah: a warranty-audit focus in the Warranty section, more source
   material to come from Noah). Content: Jan 2026 GM meeting Day 1, slides 13-14 (reference/). Stores are not named:
   the Warranty Compliance talking point already keeps the audit total without store call-outs from the podium.
   The one-line descriptions under each cause and step are drafted from the Jan slide's headings - confirm them
   against OEM policy before the meeting. */
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const SEG = "Warranty";
const SRC = "Jan 2026 GM meeting (Day 1), slides 13-14: six warranty audits, primary causes and Warranty Compliance steps. Descriptions under each heading drafted for this deck; confirm against each OEM's warranty policy manual.";

function tiles(s, items, { cols, top, bot, accent, headH = 0.4 }) {   // headH: room for a heading that wraps
  const rows = Math.ceil(items.length / cols), gap = 0.18, W = (12.33 - 0.5 - gap * (cols - 1)) / cols, Ht = (bot - top - gap * (rows - 1)) / rows;
  items.forEach(([head, line], i) => {
    const x = 0.75 + (i % cols) * (W + gap), y = top + Math.floor(i / cols) * (Ht + gap);
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: Ht, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.06, h: Ht, fill: { color: accent } });
    s.addText(words(head, "tile head"), { x: x + 0.22, y: y + 0.14, w: W - 0.36, h: headH, fontFace: F, fontSize: sz(18, "tile head"), bold: true,
      color: C.ink, valign: "top", isTextBox: true, margin: 0 });
    s.addText(words(line, "tile line"), { x: x + 0.22, y: y + 0.24 + headH, w: W - 0.36, h: Ht - 0.36 - headH, fontFace: F, fontSize: sz(15, "tile line"),
      color: C.ink600, valign: "top", isTextBox: true, margin: 0 });
  });
}

/* 0. the money: approved warranty rate increases so far this year (vendor YTD report, Oct 22 2026) */
{
  const W = JSON.parse(fs.readFileSync("./warranty/rwr.json", "utf8")), T = W.total;
  const usd = v => "$" + Math.round(v).toLocaleString("en-US"), dash = v => v ? H.kusd(v) : "–";   // table in $K (Noah, Oct 22)
  const top = [...W.rows].sort((a, b) => b.total - a.total)[0];
  H.amountRanker(pptx, pptx, {
    eyebrow: SEG, title: "Warranty Rate Increases",
    rows: W.rows.map(r => [r.n, r.total]), fmt: H.kusd, color: C.blue, ink: C.navy, valLabel: "GP added", valW: 1.0,
    extra: [{ label: "Labor", w: 0.92, vals: Object.fromEntries(W.rows.map(r => [r.n, dash(r.labor)])) },
            { label: "Parts", w: 0.92, vals: Object.fromEntries(W.rows.map(r => [r.n, dash(r.parts)])) }],
    takeaway: `${usd(T.total)} in added gross profit from warranty rate increases approved this year at ${W.rows.length} stores. Higher rates bring more audits: the process is how we keep it.`,
    tone: "n",   // display, not a comparison (Noah, Oct 22)
    source: `YTD RWR Financials 2026 (warranty rate submission vendor), "YTD revenue secured": the added gross profit from approved increases, as more dollars per warranty labor hour sold or more parts markup. ${W.rows.length} stores have an approval this year; the other ${33 - W.rows.length} have not had one yet. Labor ${usd(T.labor)}, parts ${usd(T.parts)}. ${W.note}`,
  }).addNotes(`Talk track: these approvals are new money on every warranty claim, and they are why warranty ELR is up. The factory notices a higher rate too: more money in means more scrutiny. Every claim debited in an audit gives part of it back.\n\nVendor store names: ${W.rows.map(r => `${r.src} = ${r.n}`).join("; ")}.`);
}

/* 0b. warranty ELR and warranty parts gross profit % on one slide (Noah, Oct 24: replaces the warranty ELR block and
   the parts GP % chart). The point is direction, not detail: both are rising, so documentation has to hold up. */
{
  const Wd = JSON.parse(fs.readFileSync("./wi.json", "utf8")), P = JSON.parse(fs.readFileSync("./pulls/q4_2025.json", "utf8")).w_parts;
  const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const elr = r => r[1] / r[2], gp = r => r[1] / r[0] * 100;
  const ytd = (rows, n, d) => rows.reduce((t, r) => t + r[n], 0) / rows.reduce((t, r) => t + r[d], 0);
  const E = { p: [...Wd.m25, ...Wd.q4].map(elr), c: Wd.m26.map(elr), y25: ytd(Wd.m25, 1, 2), y26: ytd(Wd.m26, 1, 2) };
  const G = { p: P.m25.map(gp), c: P.m26.map(gp), y25: ytd(P.m25.slice(0, 9), 1, 0) * 100, y26: ytd(P.m26, 1, 0) * 100 };
  const ahead = X => X.c.filter((v, i) => v > X.p[i]).length;
  const tk = `Warranty ELR is up $${Math.round(E.y26 - E.y25)} to $${Math.round(E.y26)} and warranty parts margin up ${(G.y26 - G.y25).toFixed(1)} pts to ${G.y26.toFixed(1)}%. Higher rates draw more audits: every claim's documentation has to hold up.`;
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: SEG, title: "Warranty Rates Are Rising" });
  [[E, "WARRANTY ELR", v => "$" + Math.round(v), '"$"#,##0', 0.5, 10], [G, "WARRANTY PARTS GROSS PROFIT %", v => v.toFixed(1) + "%", '0"%"', 6.72, 2]].forEach(([X, lab, fmt, axis, x, step]) => {
    const w = 6.11, top = H.BODY.top, h = H.bodyH(tk);
    H.panel(pptx, s, { x, y: top, w, h, label: null });
    s.addText(lab, { x: x + 0.22, y: top + 0.14, w: w - 0.44, h: 0.32, fontFace: F, fontSize: sz(14, "twin head"), bold: true, color: C.navy, charSpacing: 1.0, isTextBox: true, margin: 0 });
    s.addText([{ text: "2025  ", options: { color: C.ink600, fontSize: sz(16, "twin label") } }, { text: fmt(X.y25) + "     ", options: { color: C.ink600, bold: true, fontSize: sz(24, "twin value") } },
               { text: "2026  ", options: { color: C.ink600, fontSize: sz(16, "twin label") } }, { text: fmt(X.y26), options: { color: C.goodInk, bold: true, fontSize: sz(24, "twin value") } }],
      { x: x + 0.22, y: top + 0.5, w: w - 0.44, h: 0.5, fontFace: F, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(`Jan–Sep year to date · ahead of 2025 in ${ahead(X) === 9 ? "all nine" : ahead(X) + " of nine"} months`, { x: x + 0.22, y: top + 1.0, w: w - 0.44, h: 0.3, fontFace: F, fontSize: sz(14, "twin sub"), color: C.ink600, isTextBox: true, margin: 0 });
    const all = [...X.p, ...X.c], lo = Math.floor(Math.min(...all) / step) * step - step, hi = Math.ceil(Math.max(...all) / step) * step + step;
    s.addChart(pptx.ChartType.line, [{ name: "2025", labels: MON, values: X.p }, { name: "2026", labels: MON, values: [...X.c, null, null, null] }], {
      x: x + 0.1, y: top + 1.35, w: w - 0.2, h: h - 1.45, chartColors: [C.ink300, C.blue], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showLegend: true, legendPos: "t", legendFontFace: "+mn-lt", legendFontSize: sz(14, "legend"), legendColor: C.ink,
      catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(14, "cat axis"), catAxisLabelColor: C.ink,
      valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: sz(14, "val axis"), valAxisLabelColor: C.ink600, valAxisLabelFormatCode: axis,
      valAxisMinVal: lo, valAxisMaxVal: hi, valAxisMajorUnit: step, valGridLine: { color: C.rule, size: 0.75 }, catGridLine: { style: "none" }, valAxisLineShow: false,
    });
  });
  H.takeaway(pptx, s, tk, "n");
  H.foot(pptx, s, "Qlik - Closed Repair Orders. Warranty pay type, by close month; 2025 shown January-December, 2026 January-September. Warranty ELR = warranty labor sale / warranty hours sold; parts margin = warranty parts gross / parts sale. Same-store, INEOS excluded; GMC Prescott body shop labor types excluded.");
  s.addNotes("Talk track (Noah, Oct 24): the point is the direction. Rate increases are lifting both warranty ELR and warranty parts margin. The factory notices higher rates: more money in means more scrutiny, so documentation and process on every claim have to be right. Store detail for warranty ELR is no longer in the flow.");
}

/* 1. what the auditor debits */
{
  const s = pptx.addSlide(), tk = "Six audits took back over $400,000, more than a tenth of what rate increases brought in this year. Every cause here is documentation the store controls.";
  H.page(pptx, s, { eyebrow: SEG, title: "What Gets Claims Debited" });
  H.panel(pptx, s, { x: 0.5, y: H.BODY.top, w: 12.33, h: H.bodyH(tk), label: "What auditors check: the causes in the six audits, plus supporting evidence" });
  tiles(s, [
    ["Written documentation", "Complaint, cause and correction written on the RO, and matching what was claimed"],
    ["Required signatures", "Manager sign-off on added repairs; customer and manager signatures on in-stock units"],
    ["Supporting evidence", "Photos of the failed part, test results and proof the required procedure was followed, where the manufacturer asks for it"],
    ["Out mileage", "Recorded on every RO and consistent with the mileage in"],
    ["Punch times", "Technician clock times on the RO support the hours claimed"],
    ["Manufacturer authorization", "Prior approval on file wherever the manufacturer requires it"],
  ], { cols: 3, top: H.BODY.top + 0.72, bot: H.bodyBot(tk) - 0.2, accent: C.bad, headH: 0.7 });
  H.takeaway(pptx, s, tk, "b");
  H.foot(pptx, s, SRC);
  s.addNotes("DRAFT. Causes from Jan 2026 GM meeting slide 13; the two signature causes combined and Supporting evidence added (Noah, Oct 22: some manufacturers require photos of the failed component or proof a procedure was followed). That slide named six stores and their debits; stores are left off here, as on Warranty Compliance. The store amounts on the Jan slide add to $414,000 (headline said $400,000), so the slide says 'over $400,000'.");
}

/* 2. how a store keeps the claim paid */
{
  const s = pptx.addSlide(), tk = "The General Manager is responsible for making sure warranty policies and procedures are followed.";
  H.page(pptx, s, { eyebrow: SEG, title: "Keep the Claim Paid" });
  H.panel(pptx, s, { x: 0.5, y: H.BODY.top, w: 12.33, h: H.bodyH(tk), label: "Warranty compliance, in five steps" });
  tiles(s, [
    ["1  Documented policies", "Each manufacturer's current warranty policy manual, where the team can reach it"],
    ["2  Defined process", "Who writes, who approves, who submits: every claim handled the same way"],
    ["3  Create awareness", "Advisors, technicians and warranty admins know what an auditor checks"],
    ["4  Safety nets", "Signatures, out mileage, punch times and authorizations checked before the claim goes in"],
    ["5  Spot-check and self-audit", "Pull warranty ROs every month and audit them the way the factory will"],
  ], { cols: 5, top: H.BODY.top + 0.72, bot: H.bodyBot(tk) - 0.2, accent: C.blue, headH: 0.7 });
  H.takeaway(pptx, s, tk, "n");
  H.foot(pptx, s, SRC);
  s.addNotes("DRAFT. Steps from Jan 2026 GM meeting slide 14 (Warranty Compliance); the GM responsibility line from slide 13.");
}

await pptx.writeFile({ fileName: "Warranty Audits - Draft.pptx" }); console.log("warranty built");
