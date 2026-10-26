/* Phone handling preview (Noah, Oct 26): Flai voice AI, Service view, standings export Jul 26 - Oct 26, 2026
   (flai/standings_*.csv; store rows reconcile to Flai's Group row on every count). Flai's own location names and grain
   (26 locations, not the P&L's 33 stores): never on a slide with P&L numbers. Est. $ columns are Flai's model and are
   not used. Pickup rate and Resolution are read as Flai labels them; definitions to confirm with Noah. */
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import * as H from "./house.mjs";
const { C, F, sz, words } = H;
const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE"; pptx.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };
const raw = fs.readFileSync("./flai/standings_2026-07-26_to_2026-10-26.csv", "utf8").replace(/^﻿/, "").trim().split(/\r?\n/);
const head = raw[0].split(","), rows = raw.slice(1).map(l => Object.fromEntries(l.split(",").map((v, i) => [head[i], v])));
const num = v => v === "" ? null : Number(v);
const R = rows.map(r => ({ n: r.Store.replace(/^Findlay /, "").replace("Chevrolet", "Chevy").replace("Volkswagen", "VW"), calls: num(r.Calls), book: num(r["Booking rate (%)"]), pick: num(r["Pickup rate (%)"]),
  res: num(r["Resolution (%)"]), open: num(r["Open callbacks"]), bookable: num(r["Bookable calls"]), booked: num(r.Booked), tr: num(r.Transfers),
  trr: num(r["Transfer rate (%)"]), avg: num(r["Avg. call (seconds)"]) }));
const G = R.find(r => r.n === "Group"), S = R.filter(r => r.n !== "Group" && r.calls > 0);
for (const k of ["calls", "bookable", "booked", "tr", "open"]) {
  const t = S.reduce((a, r) => a + r[k], 0);
  if (t !== G[k]) throw new Error(`Flai ${k}: stores ${t} vs group ${G[k]}`);
}
const DAYS = 93;   // Jul 26 - Oct 26 inclusive
const MIN = 300;   // rankings: locations with at least 300 calls, so a handful of calls can't top a list
const fmt = v => Math.round(v).toLocaleString("en-US");
const SEG = "Voice AI · Flai · Last 90 Days";   // says what and whose data (Noah, Oct 26)
const SRC = `Flai voice AI dashboard, Service view, standings export, Jul 26 - Oct 26, 2026 (${DAYS} days). Flai's location names (26 with calls). Pickup rate and resolution as Flai reports them. Flai's estimated savings, revenue and profit are not used.`;
const notPicked = Math.round(G.tr * (1 - G.pick / 100));

// 1. the group picture
H.statTiles(pptx, pptx, { eyebrow: SEG, title: "Voice AI: Every Call Is a Customer", bigSize: 36, subSize: 14,
  tiles: [
    { label: "Calls handled", big: fmt(G.calls), sub: `About ${fmt(G.calls / DAYS)} a day across ${S.length} locations, average call ${Math.floor(G.avg / 60)}:${String(G.avg % 60).padStart(2, "0")}.`, tone: "n" },
    // handled 100% by the AI (Noah, Oct 26): Flai's export has no such column, so it is calls not transferred
    { label: "Handled by the AI alone", big: fmt(G.calls - G.tr), sub: `${(100 - G.trr).toFixed(1)}% of calls, start to finish. The other ${fmt(G.tr)} (${G.trr}%) were transferred to the store.`, tone: "g" },
    { label: "Appointments booked", big: fmt(G.booked), sub: `${G.book}% of the ${fmt(G.bookable)} calls that could book an appointment.`, tone: "g" },
    { label: "Transfers picked up", big: `${G.pick}%`, sub: `About ${fmt(notPicked)} transferred calls were not picked up by a person.`, tone: "b" },
  ],
  banner: `Service calls answered by Flai, our voice AI, over the last 90 days (Jul 26 – Oct 26, 2026). The AI books and routes the call; the store still has to answer the transfer and close every callback in the Flai dashboard.`,   // open-callback count off this slide (Noah, Oct 26)
  takeaway: `The AI finishes ${(100 - G.trr).toFixed(1)}% of calls on its own. The rest come back to the store, and fewer than half of those reach a person. Every unanswered transfer is a customer we already had.`,
  tone: "b", source: SRC + " Handled by the AI alone = calls - transfers (the export has no separate column for it). Not picked up = transfers x (1 - pickup rate), from Flai's group figures." });

// 2. top and bottom five on the three things a store controls
{
  const R2 = S.filter(r => r.calls >= MIN);
  const cols = [
    { head: "BOOKING RATE", sub: "booked ÷ bookable", key: "book", f: r => r.book.toFixed(1) + "%", up: true, g: G.book },
    { head: "TRANSFER PICKUP RATE", sub: "transfers picked up", key: "pick", f: r => r.pick.toFixed(1) + "%", up: true, g: G.pick },
    { head: "CALLBACK RESOLUTION", sub: "callbacks closed", key: "res", f: r => r.res.toFixed(1) + "%", up: true, g: G.res, x: r => String(r.open), xh: "Open" },
  ];
  const s = pptx.addSlide();
  const best = [...R2].sort((a, b) => b.pick - a.pick)[0], worst = [...R2].sort((a, b) => a.pick - b.pick)[0], cb = [...R2].sort((a, b) => b.open - a.open);
  const tk = `${best.n} picks up ${best.pick}% of transfers; ${worst.n} picks up ${worst.pick}%. ${cb[0].n} and ${cb[1].n} hold ${cb[0].open + cb[1].open} of the group's ${G.open} open callbacks.`;
  H.page(pptx, s, { eyebrow: SEG, title: "Store Voice AI Metrics" });
  const top = H.BODY.top, h = H.bodyH(tk), W = (12.33 - 0.24) / 3;
  // Oct 26 (Noah): callback resolution lists every store at 100% instead of a top 5 (too many ties), and a green bar
  // marks a store that shows up in more than one top list
  const sortedOf = c => [...R2].sort((a, b) => b[c.key] - a[c.key] || a.open - b.open);
  const topOf = c => c.key === "res" ? sortedOf(c).filter(r => r.res === 100) : sortedOf(c).slice(0, 5);
  const seen = {}; cols.forEach(c => topOf(c).forEach(r => { seen[r.n] = (seen[r.n] ?? 0) + 1; }));
  const multi = new Set(Object.keys(seen).filter(n => seen[n] > 1));
  // red bar: in more than one bottom 5 (Noah, Oct 26)
  const seenB = {}; cols.forEach(c => sortedOf(c).slice(-5).forEach(r => { seenB[r.n] = (seenB[r.n] ?? 0) + 1; }));
  const multiB = new Set(Object.keys(seenB).filter(n => seenB[n] > 1));
  const maxTop = Math.max(...cols.map(c => topOf(c).length));
  cols.forEach((c, k) => {
    const x = 0.5 + k * (W + 0.12);
    H.panel(pptx, s, { x, y: top, w: W, h, label: null });
    s.addText(c.head, { x: x + 0.18, y: top + 0.1, w: W - 0.36, h: 0.3, fontFace: F, fontSize: sz(14, "col head"), bold: true, color: C.navy, charSpacing: 1, isTextBox: true, margin: 0 });
    s.addText(`Group ${c.g}%`,   // header says what it is; just the group (Noah, Oct 26)
      { x: x + 0.18, y: top + 0.4, w: W - 0.36, h: 0.26, fontFace: F, fontSize: sz(14, "col sub"), color: C.ink600, isTextBox: true, margin: 0 });
    const sorted = sortedOf(c), T = topOf(c);
    const lists = [[c.key === "res" ? `AT 100% (${T.length})` : "TOP 5", T, C.goodInk, C.good, true], ["BOTTOM 5", sorted.slice(-5).reverse(), C.badInk, C.bad, false]];
    const rh = (h - 0.85 - 0.6) / (maxTop + 5);
    let y = top + 0.75;
    lists.forEach(([lab, L, ink, bar, isTop]) => {
      if (!isTop) y = top + 1.05 + maxTop * rh;   // BOTTOM 5 starts on the same line in all three columns
      s.addShape(pptx.ShapeType.rect, { x: x + 0.18, y: y + 0.02, w: W - 0.36, h: 0.025, fill: { color: bar } });
      s.addText(lab, { x: x + 0.18, y: y + 0.05, w: 2, h: 0.24, fontFace: F, fontSize: sz(14, "list head"), bold: true, color: ink, isTextBox: true, margin: 0 });
      // column heads on the TOP 5 / BOTTOM 5 line instead of "calls" / "open" on every row (Noah, Oct 26)
      s.addText("Rate", { x: x + W - 2.0, y: y + 0.05, w: 0.8, h: 0.24, fontFace: F, fontSize: sz(14, "col head"), bold: true, color: C.ink600, align: "right", isTextBox: true, margin: 0 });
      s.addText(c.xh ?? "Calls", { x: x + W - 1.18, y: y + 0.05, w: 1.0, h: 0.24, fontFace: F, fontSize: sz(14, "col head"), bold: true, color: C.ink600, align: "right", isTextBox: true, margin: 0 });
      y += 0.3;
      L.forEach(r => {
        s.addText(r.n, { x: x + 0.18, y, w: W - 2.2, h: rh, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
        s.addText(c.f(r), { x: x + W - 2.0, y, w: 0.8, h: rh, fontFace: F, fontSize: 13, bold: true, color: ink, align: "right", valign: "middle", isTextBox: true, margin: 0 });
        s.addText(c.x ? c.x(r) : fmt(r.calls), { x: x + W - 1.18, y, w: 1.0, h: rh, fontFace: F, fontSize: 12, color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
        if (isTop ? multi.has(r.n) : multiB.has(r.n)) s.addShape(pptx.ShapeType.rect, { x: x + W - 0.13, y: y + 0.04, w: 0.06, h: rh - 0.08, fill: { color: isTop ? C.good : C.bad } });
        y += rh;
      });
    });
  });
  H.takeaway(pptx, s, tk, "n");
  H.foot(pptx, s, SRC + ` Green bar: the store is in more than one top list; red bar: in more than one bottom 5. Ranked among the ${R2.length} locations with ${MIN}+ calls (${S.filter(r => r.calls < MIN).map(r => r.n).join(", ")} have fewer). Honda Spokane shows no calls in the period.`);
  s.addNotes("Talk track: the voice AI answers every call, books what it can and hands the rest to the store. What happens next is ours: someone has to pick up the transfer, and every callback has to be closed. A store that doesn't pick up is sending a customer it already has to voicemail.\n\nDefinitions to confirm against Flai: pickup rate is read as the share of transferred calls a person picked up; resolution as the share of callback requests closed. Resolve time in the export runs to weeks at some locations and is not shown.");
}
await pptx.writeFile({ fileName: "../decks/Phone Handling - Preview.pptx" }); console.log("phone preview built");
