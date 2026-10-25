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

/* 0. the money and where it shows, on one slide (Noah, Oct 25: combines Warranty Rate Increases and Warranty Rates
   Are Rising). Left: approved rate increases by store (vendor YTD report), a compact bar list. Right: warranty ELR and
   warranty parts gross profit %, 2025 against 2026. The point is direction: rates are up, so documentation has to hold up. */
{
  const W = JSON.parse(fs.readFileSync("./warranty/rwr.json", "utf8")), T = W.total;
  const usd = v => "$" + Math.round(v).toLocaleString("en-US");
  const Wd = JSON.parse(fs.readFileSync("./wi.json", "utf8")), P = JSON.parse(fs.readFileSync("./pulls/q4_2025.json", "utf8")).w_parts;
  const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const elr = r => r[1] / r[2], gp = r => r[1] / r[0] * 100;
  const ytd = (rows, n, d) => rows.reduce((t, r) => t + r[n], 0) / rows.reduce((t, r) => t + r[d], 0);
  const E = { p: [...Wd.m25, ...Wd.q4].map(elr), c: Wd.m26.map(elr), y25: ytd(Wd.m25, 1, 2), y26: ytd(Wd.m26, 1, 2) };
  const G = { p: P.m25.map(gp), c: P.m26.map(gp), y25: ytd(P.m25.slice(0, 9), 1, 0) * 100, y26: ytd(P.m26, 1, 0) * 100 };
  const tk = `${usd(T.total)} in gross profit added by rate increases approved at ${W.rows.length} stores this year. Warranty ELR is up $${Math.round(E.y26 - E.y25)} and parts margin ${(G.y26 - G.y25).toFixed(1)} pts: higher rates draw more audits, so every claim has to hold up.`;
  const s = pptx.addSlide();
  H.page(pptx, s, { eyebrow: SEG, title: "Warranty Rates Are Rising" });
  const top = H.BODY.top, h = H.bodyH(tk);
  // left: rate increases by store
  const lx = 0.5, lw = 5.75;
  H.panel(pptx, s, { x: lx, y: top, w: lw, h, label: null });
  s.addText([{ text: "RATE INCREASES IN 2026", options: { color: C.navy, bold: true } }],
    { x: lx + 0.22, y: top + 0.12, w: 3.6, h: 0.34, fontFace: F, fontSize: sz(14, "twin head"), charSpacing: 1.0, isTextBox: true, margin: 0 });
  s.addText(H.kusd(T.total) + " GP", { x: lx + lw - 2.0, y: top + 0.08, w: 1.78, h: 0.4, fontFace: F, fontSize: sz(20, "twin value"), bold: true, color: C.navy, align: "right", isTextBox: true, margin: 0 });
  const rows = [...W.rows].sort((a, b) => b.total - a.total), RT = top + 0.58, rh = (h - 0.7) / rows.length, mx = rows[0].total;
  rows.forEach((r, i) => {
    const y = RT + i * rh, bx = lx + 2.45, bw = 2.15;
    s.addText(r.n, { x: lx + 0.22, y, w: 2.2, h: rh, fontFace: F, fontSize: 12, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    s.addShape(pptx.ShapeType.rect, { x: bx, y: y + rh * 0.22, w: Math.max(r.total / mx * bw, 0.02), h: rh * 0.56, fill: { color: C.blue } });
    s.addText(H.kusd(r.total), { x: bx + bw + 0.05, y, w: lw - (bx - lx) - bw - 0.27, h: rh, fontFace: F, fontSize: 12, bold: true, color: C.navy, align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  // right: the two rates, stacked
  const rx = 6.42, rw = 6.41, ch = (h - 0.12) / 2;
  [[E, "WARRANTY ELR", v => "$" + Math.round(v), '"$"#,##0', 10], [G, "WARRANTY PARTS GP %", v => v.toFixed(1) + "%", '0"%"', 2]].forEach(([X, lab, fmt, axis, step], k) => {
    const y0 = top + k * (ch + 0.12);
    H.panel(pptx, s, { x: rx, y: y0, w: rw, h: ch, label: null });
    s.addText([{ text: lab + "    ", options: { color: C.navy, bold: true, fontSize: sz(14, "twin head") } },
               { text: "2025 ", options: { color: C.ink300, fontSize: sz(14, "twin label") } }, { text: fmt(X.y25) + "    ", options: { color: C.ink300, bold: true, fontSize: sz(16, "twin value") } },
               { text: "2026 ", options: { color: C.blue, fontSize: sz(14, "twin label") } }, { text: fmt(X.y26), options: { color: C.blue, bold: true, fontSize: sz(16, "twin value") } }],
      { x: rx + 0.2, y: y0 + 0.08, w: rw - 0.4, h: 0.36, fontFace: F, valign: "middle", isTextBox: true, margin: 0 });
    const all = [...X.p, ...X.c], lo = Math.floor(Math.min(...all) / step) * step - step, hi = Math.ceil(Math.max(...all) / step) * step + step;
    s.addChart(pptx.ChartType.line, [{ name: "2025", labels: MON, values: X.p }, { name: "2026", labels: MON, values: [...X.c, null, null, null] }], {
      x: rx + 0.05, y: y0 + 0.42, w: rw - 0.1, h: ch - 0.48, chartColors: [C.ink300, C.blue], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 5,
      showLegend: false,   // the header carries 2025 (gray) and 2026 (green)
      catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(14, "cat axis"), catAxisLabelColor: C.ink,
      valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: sz(14, "val axis"), valAxisLabelColor: C.ink600, valAxisLabelFormatCode: axis,
      valAxisMinVal: lo, valAxisMaxVal: hi, valAxisMajorUnit: step * 2, valGridLine: { color: C.rule, size: 0.75 }, catGridLine: { style: "none" }, valAxisLineShow: false,
    });
  });
  H.takeaway(pptx, s, tk, "n");
  H.foot(pptx, s, `Rate increases: YTD RWR Financials 2026 (warranty rate submission vendor), "YTD revenue secured" = added gross profit from approved increases (labor ${usd(T.labor)}, parts ${usd(T.parts)}); ${33 - W.rows.length} stores have no approval yet. Rates: Qlik - Closed Repair Orders, warranty pay type by close month; ELR = warranty labor sale / hours sold; parts margin = warranty parts gross / parts sale. 2025 January-December, 2026 January-September. Same-store.`);
  s.addNotes(`Talk track (Noah, Oct 24-25): rate increases are new money on every warranty claim, and they show up as a higher warranty ELR and parts margin. The factory notices higher rates too: more money in means more scrutiny, so documentation and process on every claim is how we keep what we are owed.\n\nVendor store names: ${W.rows.map(r => `${r.src} = ${r.n}`).join("; ")}.`);
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
