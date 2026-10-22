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

/* 0b. warranty parts gross profit %: the rate increases showing up as margin (Noah, Oct 22 2026: an example) */
{
  const P = JSON.parse(fs.readFileSync("./pulls/q4_2025.json", "utf8")).w_parts;
  const gp = ([s, g]) => g / s * 100, ytd = rows => rows.reduce((t, r) => t + r[1], 0) / rows.reduce((t, r) => t + r[0], 0) * 100;
  const y25 = ytd(P.m25.slice(0, 9)), y26 = ytd(P.m26);
  const ahead = P.m26.filter((r, i) => gp(r) > gp(P.m25[i])).length;
  H.groupYoy(pptx, pptx, {
    eyebrow: SEG, title: "Warranty Parts Gross Profit %", metric: "warranty parts gp %", measure: "Warranty Parts GP %",
    prior: P.m25.map(gp), current: P.m26.map(gp), priorName: "2025", currentName: "2026",
    ytdPrior: y25, ytdCurrent: y26, unit: "pct", chart: "line",
    chartLabel: "Warranty parts gross profit % by month, 2025 against 2026",
    takeaway: `Up ${(y26 - y25).toFixed(1)} pts to ${y26.toFixed(1)}%, ahead of 2025 in ${ahead === 9 ? "all nine" : ahead + " of nine"} months: approved parts markup increases show up as margin.`,
    source: "Qlik - Closed Repair Orders. Warranty pay type, parts gross / parts sale, by close month; 2025 January-December, 2026 January-September. Same-store, INEOS excluded (opened 2026). GMC Prescott body shop labor types excluded. Pulled Oct 22 2026.",
  }).addNotes("Example for the rate-increase story. The link to the approvals is the point we are making, not something the data proves on its own: 17 of the 19 stores with approvals got a parts markup increase (vendor YTD report), and warranty parts margin rose in every month. Mix of parts and campaigns also moves this line.");
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
