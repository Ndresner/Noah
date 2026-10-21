/* ===================================================================
   FINDLAY MEETING DECK - HOUSE STYLE
   Locked October 2026 with Tray. Decisions this file encodes:

     font        Inter          (must be installed on any machine that opens the deck)
     field       grey EDEDED, content on white panels
     titles      the SOURCE LINE NAME and nothing else; the finding goes in the takeaway
     colour      green good, red bad, always. blue is identity, never judgment
     stores      ALL CAPS, all 34 on one page, GROUP as its own strip
     naming      always "stores", never "rooftops" - enforced (Noah, Oct 14 2026)
     floor       nothing carrying meaning below 14pt - enforced, not remembered
     numbers     takeaways and counts DERIVED from the series, never typed

   Built for a projector in a large room, so density loses to legibility
   every time the two conflict.
   =================================================================== */

export const F = "Inter";

export const C = {
  field:    "EDEDED",
  panel:    "FFFFFF",
  ink:      "1A1A1A",
  ink600:   "595959",
  ink300:   "9AA3AD",
  rule:     "D9D9D9",
  zebra:    "F7F7F7",
  blue:     "2587C5",
  blueLt:   "5B9BD5",
  blueTint: "EDF4FA",
  navy:     "0B2C4A",
  good:     "12A566", goodInk: "0B7A4A", goodBg: "E6F6EE",
  bad:      "C2352F", badInk:  "A82B26", badBg:  "FBEBEA",
  white:    "FFFFFF",
};

/* ---------- naming: "stores", never "rooftops" ----------
   Added 2026-10-14 (Noah). Every piece of slide text passes through words(),
   which stops the build rather than ship the old term. */
export function words(text, where) {
  if (typeof text === "string" && /rooftop/i.test(text)) {
    throw new Error(`"rooftop" found in ${where || "slide text"} - say "store": ${text}`);
  }
  return text;
}

/* ---------- the font floor ---------- */
export const FLOOR = 14;
const EXEMPT = {
  13: "ranker store names - every store on one page, signed off Oct 2026",
  12: "source line - reference material, not read from the back of the room",
};
export function sz(size, why) {
  if (size < FLOOR && !EXEMPT[size]) {
    throw new Error(`font ${size}pt below the ${FLOOR}pt floor (${why || "unlabelled"})`);
  }
  return size;
}

/* ---------- the metric registry ---------- */
export const METRICS = {
  "service gross":        { dir: "up",   fmt: "money" },
  "parts gross":          { dir: "up",   fmt: "money" },
  "fixed ops gross":      { dir: "up",   fmt: "money" },
  "elr":                  { dir: "up",   fmt: "money" },
  "hpro":                 { dir: "up",   fmt: "dec"   },
  "ro count":             { dir: "up",   fmt: "int"   },
  "service absorption":   { dir: "up",   fmt: "pct"   },
  "mpi video %":          { dir: "up",   fmt: "pct"   },
  "cp share":             { dir: "up",   fmt: "pct"   },   // customer-pay share of service gross
  // 2026-10-14 (Noah, opcode test build) - PROVISIONAL until Tray confirms.
  // Share of customer-pay ROs carrying the service; more is better.
  "alignment %":          { dir: "up",   fmt: "pct"   },
  "balance %":            { dir: "up",   fmt: "pct"   },
  "brake flush %":        { dir: "up",   fmt: "pct"   },
  "fuel service %":       { dir: "up",   fmt: "pct"   },
  // 2026-10-15 (Noah, CP Engine + Productivity rebuild) - PROVISIONAL until Tray confirms.
  "cp labor per ro":      { dir: "up",   fmt: "money" },
  "cp labor gp %":        { dir: "up",   fmt: "pct"   },   // Roundtable target 78-80%; above can mean techs underpaid
  "cp parts gp %":        { dir: "up",   fmt: "pct"   },
  "fixed gross per fixed employee": { dir: "up", fmt: "money" },
  "ros per advisor":      { dir: "up",   fmt: "int"   },   // fewer can be a deliberate staffing choice
  "ros per fixed employee": { dir: "up", fmt: "dec"   },
  "ros per fixed manager":  { dir: "up", fmt: "int"   },
  "fixed headcount":      { dir: "neutral", fmt: "int" }, // added staff can be growth or excess - never green/red
  "cp parts per ro":      { dir: "up",   fmt: "money" },
  "warranty ro count":    { dir: "up",   fmt: "int"   },   // largely OEM/recall-driven - Tray may prefer neutral
  "warranty elr":         { dir: "up",   fmt: "money" },
  "internal elr":         { dir: "up",   fmt: "money" },   // more service gross, but a cost to used cars - confirm
  "fixed ops absorption %": { dir: "up", fmt: "pct"   },   // P&L line 257 = line 254 / line 258
  "service net to gross %": { dir: "up", fmt: "pct"   },
  "parts net to gross %":   { dir: "up", fmt: "pct"   },
  "cp labor share":         { dir: "up", fmt: "pct"   },
  "cp unclassified work %": { dir: "down", fmt: "pct" },
  "cp elr % of door rate":  { dir: "up", fmt: "pct" },
  "door rate":              { dir: "neutral", fmt: "money" },
  // 2026-10-19 (Noah): total customer-pay sold hours - the volume check behind $ per RO. More is better.
  "cp hours sold":          { dir: "up", fmt: "int" },
  // 2026-10-19: gross from one more CP RO a day - an opportunity amount, shown as a size, not a score
  "one more ro a day":      { dir: "neutral", fmt: "money" },
  // 2026-10-16 (Noah, proficiency slide test) - PROVISIONAL until Tray confirms. Sold hours / paid hours (ADP reg + OT); workbook targets 85% on, 100% above.
  "tech proficiency %":     { dir: "up", fmt: "pct" },   // a posted price, not a score     // May benchmark: CP ELR at least 70% of door rate   // CP labor whose op code description says nothing usable - a description-quality measure   // customer-pay share of labor sale; higher = less dependence on OEM/internal - confirm
};
export function dirOf(key) {
  const m = METRICS[key];
  if (!m) throw new Error(`metric "${key}" is not in the registry - ask Tray, then add it with a dated comment`);
  return m.dir;
}

/* ---------- formatters ---------- */
export const M     = v => "$" + (v / 1e6).toFixed(2) + "M";
export const money = v => v === null ? "\u2014"
  : (v < 0 ? "\u2212" : "+") + "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
export const dollars = v => "$" + Math.round(v).toLocaleString("en-US");
export const pct   = (a, b) => (a / b >= 0 ? "+" : "\u2212") + Math.abs(a / b * 100).toFixed(1) + "%";

/* ---------- page furniture ---------- */
export function page(pptx, s, { eyebrow, title }) {
  s.background = { color: C.field };
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.44, w: 0.055, h: 0.21, fill: { color: C.blue } });
  words(eyebrow, "eyebrow"); words(title, "title");
  s.addText(eyebrow.toUpperCase(), {
    x: 0.68, y: 0.37, w: 9, h: 0.3, fontFace: F, fontSize: sz(14, "eyebrow"), bold: true,
    color: C.ink600, charSpacing: 1.1, isTextBox: true, margin: 0,
  });
  s.addText(title, {   // the source line name, exactly
    x: 0.5, y: 0.7, w: 12.4, h: 0.55, fontFace: F, fontSize: sz(28, "title"),
    color: C.ink, isTextBox: true, margin: 0,
  });
}

/* Source and page number are deliberately NOT drawn (Tray, Oct 2026) - they ate
   space and the room does not read them. Provenance is kept in the speaker
   notes so it still travels with the deck and can be answered from the podium. */
export function foot(pptx, s, source, n) {
  if (source) s.addNotes(`Source: ${words(source, "speaker notes")}`);
}

export function panel(pptx, s, { x, y, w, h, label }) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: C.panel } });
  if (label) {
    s.addText(label, {
      x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.32, fontFace: F, fontSize: sz(15, "panel label"),
      color: C.ink, isTextBox: true, margin: 0,
    });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.25, y: y + 0.53, w: w - 0.5, h: 0.012, fill: { color: C.ink } });
  }
}

export function takeaway(pptx, s, text, tone = "n", y = 6.52) {
  words(text, "takeaway");
  const bg = tone === "g" ? C.goodBg : tone === "b" ? C.badBg : C.blueTint;
  const fg = tone === "g" ? C.goodInk : tone === "b" ? C.badInk : C.navy;
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.33, h: 0.6, fill: { color: bg } });
  s.addText(text, {
    x: 0.72, y: y + 0.02, w: 11.95, h: 0.56, fontFace: F, fontSize: sz(16, "takeaway"), bold: true,
    color: fg, isTextBox: true, margin: 0, valign: "middle",
  });
}

/* ---------- THE DAY TRACKER -------------------------------------------
   Shared by the cover and every section break. Completed segments read
   light blue, the current one is lit and bold, the rest sit grey. In a
   two-day meeting this is the thing that tells 70 people where they are
   and how much is left.

   Labels auto-size from 17pt down to the 14pt floor. If the names still
   do not fit, the build stops rather than shrink past the floor or clip -
   that is a signal to shorten a segment name, not to squeeze the type.
   --------------------------------------------------------------------- */
function dayTracker(pptx, s, segments, current, y = 5.75) {
  const TRACK_X = 0.75, TRACK_W = 11.8;
  const segW = TRACK_W / segments.length;
  const slot = segW - 0.22;

  let fsLabel = 17;
  const longest = Math.max(...segments.map(n => n.length));
  while (fsLabel > FLOOR && (longest * 0.52 * fsLabel) / 72 > slot) fsLabel -= 0.5;
  if ((longest * 0.52 * fsLabel) / 72 > slot) {
    const n = segments.find(x => x.length === longest);
    throw new Error(`segment name "${n}" will not fit the day tracker at the ${FLOOR}pt floor - shorten it or use fewer segments`);
  }

  segments.forEach((name, i) => {
    const x = TRACK_X + i * segW;
    const on = i === current, done = current != null && i < current;
    s.addShape(pptx.ShapeType.rect, {
      x, y, w: slot, h: on ? 0.1 : 0.05,
      fill: { color: on ? C.blue : (done ? C.blueLt : C.rule) },
    });
    s.addText(name, {
      x, y: y + 0.2, w: slot, h: 0.4, fontFace: F, fontSize: sz(fsLabel, "tracker label"),
      bold: on, color: on ? C.ink : C.ink600, isTextBox: true, margin: 0,
    });
  });
}

/* ---------- SECTION BREAK - light, with the day tracker ---------- */
export function section(pptx, deck, { kicker, title, chips, segments, current, logo }) {
  // ADDED 2026-10-15: when a master deck is assembled, globalThis.__TRACK = {segments, current} lights the day tracker
  if (!segments && globalThis.__TRACK) ({ segments, current } = globalThis.__TRACK);
  [kicker, title, ...(chips || [])].forEach(t => words(t, "section slide"));
  const s = deck.addSlide();
  s.background = { color: C.white };
  if (logo) s.addImage({ data: logo, x: 0.75, y: 0.62, w: 2.0, h: 0.608 });
  s.addText(kicker.toUpperCase(), {
    x: 0.78, y: 2.6, w: 10, h: 0.4, fontFace: F, fontSize: sz(20, "section kicker"), bold: true,
    color: C.blue, charSpacing: 2.6, isTextBox: true, margin: 0,
  });
  /* ADDED 2026-10-16: a long title (e.g. "Technician Proficiency") wrapped to two lines at 82pt and ran into
     the kicker and chips. Step down until it fits one line (Inter bold ~0.6em per character); stop at 48pt. */
  let tfs = 82;
  while (tfs > 48 && title.length * 0.6 * tfs / 72 > 11.6) tfs -= 2;
  if (title.length * 0.6 * tfs / 72 > 11.6) throw new Error(`section title "${title}" will not fit one line at 48pt - shorten it`);
  s.addText(title, {
    x: 0.73, y: 3.0, w: 11.8, h: 1.35, fontFace: F, fontSize: sz(tfs, "section title"), bold: true, valign: "middle", wrap: false,
    color: C.ink, isTextBox: true, margin: 0,
  });
  if (chips?.length) s.addText(chips.join("   \u00b7   "), {
    x: 0.76, y: 4.42, w: 11.8, h: 0.45, fontFace: F, fontSize: sz(22, "section chips"),
    color: C.ink600, isTextBox: true, margin: 0,
  });
  if (segments?.length) dayTracker(pptx, s, segments, current);
  return s;
}

/* ---------- COVER - same light field, tracker with nothing lit yet ---------- */
export function cover(pptx, deck, { title, subtitle, meta, segments, logo }) {
  const s = deck.addSlide();
  s.background = { color: C.white };
  if (logo) s.addImage({ data: logo, x: 0.75, y: 0.62, w: 2.4, h: 0.73 });
  s.addText(title, {
    x: 0.73, y: 2.35, w: 11.8, h: 1.9, fontFace: F, fontSize: sz(62, "cover title"), bold: true,
    color: C.ink, isTextBox: true, margin: 0,
  });
  s.addText(subtitle, {
    x: 0.76, y: 4.28, w: 11.8, h: 0.5, fontFace: F, fontSize: sz(26, "cover subtitle"),
    color: C.blue, isTextBox: true, margin: 0,
  });
  if (meta?.length) s.addText(meta.join("   \u00b7   "), {
    x: 0.76, y: 4.82, w: 11.8, h: 0.45, fontFace: F, fontSize: sz(20, "cover meta"),
    color: C.ink600, isTextBox: true, margin: 0,
  });
  if (segments?.length) dayTracker(pptx, s, segments, null);
  return s;
}

/* ---------- TREND - trailing 12, terminus label, no per-point labels ---------- */
export function trend(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 8.3, h: 5.05, label: o.panelLabel });

  // ADDED 2026-10-15: o.scale (default 1e6), o.axisFormat, o.majorUnit for non-dollar series
  const SC = o.scale ?? 1e6;
  const vals = o.series.map(r => +(r[1] / SC).toFixed(2));
  const lo = o.min ?? Math.floor(Math.min(...vals) - 1);
  const hi = o.max ?? Math.ceil(Math.max(...vals) + 1);
  const pad = (hi - lo) * (o.pad ?? 0.12);   // headroom so point labels clear the line (ADDED: o.pad)

  s.addChart(pptx.ChartType.line, [{ name: o.panelLabel, labels: o.series.map(r => r[0]), values: vals }], {
    x: 0.68, y: 1.98, w: 7.5, h: 4.22,
    chartColors: [C.blue], lineSize: 3, lineSmooth: false,
    lineDataSymbol: "circle", lineDataSymbolSize: 8,
    lineDataSymbolLineColor: C.blue, lineDataSymbolLineSize: 1,
    showValue: true, dataLabelPosition: "t",
    dataLabelFontFace: "+mn-lt", dataLabelFontSize: sz(14, "data label"), dataLabelColor: C.ink,
    dataLabelFormatCode: o.labelFormat ?? '0.0',      // bare number; the axis says $M
    showLegend: false, showTitle: false,
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(14, "cat axis"), catAxisLabelColor: C.ink,
    valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: sz(14, "val axis"), valAxisLabelColor: C.ink600,
    valAxisLabelFormatCode: o.axisFormat ?? '$0"M"',
    valAxisMinVal: lo - pad, valAxisMaxVal: hi + pad, valAxisMajorUnit: o.majorUnit ?? 1,
    valGridLine: { color: C.rule, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: true, valAxisLineShow: false,
  });

  panel(pptx, s, { x: 9.0, y: 1.32, w: 3.83, h: 5.05, label: "What it says" });
  let by = 2.02;
  o.bullets.forEach((b, i) => {
    s.addShape(pptx.ShapeType.ellipse, { x: 9.25, y: by + 0.1, w: 0.11, h: 0.11, fill: { color: i === 0 ? C.blue : C.ink } });
    s.addText(b, {
      x: 9.52, y: by - 0.04, w: 3.1, h: 1.15, fontFace: F, fontSize: sz(15, "bullet"),
      color: i === 0 ? C.blue : C.ink, bold: i === 0, isTextBox: true, margin: 0, valign: "top",
    });
    by += 1.24;
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- TREND YOY - this year vs last year, one line each ----------
   ADDED 2026-10-20 (Noah): replaces the trailing-12 line with the reporting period,
   Jan-Sep 2026 over Jan-Sep 2025. Last year grey, this year blue (identity, not
   judgment). o.months = labels, o.a = last year, o.b = this year (o.a may run longer
   than o.b to show what is still ahead). o.card = true puts the YOY card on top and
   drops the bullets panel - the merged slide. Values in $M unless o.scale/o.axisFormat. */
export function trendYoy(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const SC = o.scale ?? 1e6;
  const r2 = v => v == null ? null : +(v / SC).toFixed(2);
  const all = [...o.a, ...o.b].map(v => v / SC);
  const lo = o.min ?? Math.floor(Math.min(...all) - 0.5), hi = o.max ?? Math.ceil(Math.max(...all) + 0.5);
  const yA = o.labels?.[0] ?? "2025", yB = o.labels?.[1] ?? "2026";
  const data = [
    { name: yA, labels: o.months, values: o.a.map(r2) },
    { name: yB, labels: o.months, values: o.months.map((_, i) => r2(o.b[i])) },
  ];

  let cy = 1.32;
  if (o.card) {
    const n = o.b.length, a = o.a.slice(0, n).reduce((t, x) => t + x, 0), b = o.b.reduce((t, x) => t + x, 0), v = b - a;
    const up = dirOf(o.metric) === "up" ? v >= 0 : v < 0;
    const FV = o.fmt?.v ?? M, FD = o.fmt?.d ?? money, FP = o.fmt?.p ?? pct;
    panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 1.3, label: null });
    s.addText(words(o.cardLabel, "card label").toUpperCase(), {
      x: 0.75, y: 1.42, w: 11.8, h: 0.3, fontFace: F, fontSize: sz(14, "card label"),
      color: C.ink600, charSpacing: 1.0, isTextBox: true, margin: 0,
    });
    [[yA, FV(a), C.ink], [yB, FV(b), C.ink], ["VARIANCE", FD(v), up ? C.goodInk : C.badInk], ["VAR %", FP(v, a), up ? C.goodInk : C.badInk]]
      .forEach(([l, val, col], i) => {
        const x = 0.75 + i * 3.05;
        s.addText(l, { x, y: 1.74, w: 2.9, h: 0.26, fontFace: F, fontSize: sz(14, "card sublabel"), color: C.ink600, charSpacing: 0.9, isTextBox: true, margin: 0 });
        s.addText(val, { x, y: 1.98, w: 2.95, h: 0.56, fontFace: F, fontSize: sz(28, "card value"), bold: true, color: col, isTextBox: true, margin: 0 });
      });
    cy = 2.76;
  }
  const W = o.card ? 12.33 : 8.3, H = 6.37 - cy;
  panel(pptx, s, { x: 0.5, y: cy, w: W, h: H, label: o.panelLabel });

  s.addChart(pptx.ChartType.line, data, {
    x: 0.68, y: cy + 0.62, w: W - 0.36, h: H - 0.72,
    chartColors: [C.ink300, C.blue], lineSize: 3, lineSmooth: false,
    lineDataSymbol: "circle", lineDataSymbolSize: 8, displayBlanksAs: "gap",
    showLegend: true, legendPos: "t", legendFontFace: "+mn-lt", legendFontSize: sz(14, "legend"), legendColor: C.ink,
    showValue: false, showTitle: false,
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(14, "cat axis"), catAxisLabelColor: C.ink,
    valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: sz(14, "val axis"), valAxisLabelColor: C.ink600,
    valAxisLabelFormatCode: o.axisFormat ?? '$0"M"',
    valAxisMinVal: lo, valAxisMaxVal: hi, valAxisMajorUnit: o.majorUnit ?? 1,
    valGridLine: { color: C.rule, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: true, valAxisLineShow: false,
  });

  if (!o.card) {
    panel(pptx, s, { x: 9.0, y: 1.32, w: 3.83, h: 5.05, label: "What it says" });
    let by = 2.02;
    o.bullets.forEach((b, i) => {
      s.addShape(pptx.ShapeType.ellipse, { x: 9.25, y: by + 0.1, w: 0.11, h: 0.11, fill: { color: i === 0 ? C.blue : C.ink } });
      s.addText(words(b, "bullet"), {
        x: 9.52, y: by - 0.04, w: 3.1, h: 1.15, fontFace: F, fontSize: sz(15, "bullet"),
        color: i === 0 ? C.blue : C.ink, bold: i === 0, isTextBox: true, margin: 0, valign: "top",
      });
      by += 1.24;
    });
  }

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- YEAR BARS - one bar per year, the latest lit ----------
   ADDED 2026-10-21 (records review). o.years = [[label, value]]; the last bar is blue,
   earlier years grey. Bullets panel on the right, as in trend(). Values in $M. */
export function yearBars(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 8.3, h: 5.05, label: o.panelLabel });
  const n = o.years.length, v = o.years.map(r => +(r[1] / 1e6).toFixed(2));
  const labels = o.years.map(r => r[0]);
  s.addChart(pptx.ChartType.bar, [
    { name: "Earlier", labels, values: v.map((x, i) => i < n - 1 ? x : 0) },
    { name: "Latest",  labels, values: v.map((x, i) => i === n - 1 ? x : 0) },
  ], {
    x: 0.68, y: 1.98, w: 7.95, h: 4.25, barDir: "col", barGrouping: "stacked", barGapWidthPct: 25,
    chartColors: [C.ink300, C.blue], showLegend: false, showTitle: false,
    showValue: true, dataLabelPosition: "inEnd", dataLabelColor: C.white,
    dataLabelFontFace: "+mn-lt", dataLabelFontSize: sz(14, "bar label"), dataLabelFontBold: true,
    dataLabelFormatCode: '$0.0"M";;',
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(16, "cat axis"), catAxisLabelColor: C.ink,
    valAxisHidden: true, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  });
  panel(pptx, s, { x: 9.0, y: 1.32, w: 3.83, h: 5.05, label: "What it says" });
  let by = 2.02;
  o.bullets.forEach((b, i) => {
    s.addShape(pptx.ShapeType.ellipse, { x: 9.25, y: by + 0.1, w: 0.11, h: 0.11, fill: { color: i === 0 ? C.blue : C.ink } });
    s.addText(words(b, "bullet"), { x: 9.52, y: by - 0.04, w: 3.1, h: 1.15, fontFace: F, fontSize: sz(15, "bullet"),
      color: i === 0 ? C.blue : C.ink, bold: i === 0, isTextBox: true, margin: 0, valign: "top" });
    by += 1.24;
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "g");
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- RECORD BOARD - every record-setting store, two columns ----------
   ADDED 2026-10-21 (records review). o.rows = [{ name, marks: [bool], val, gain }],
   o.markLabels = column heads for the marks. A record is a filled green dot; a miss
   is a pale outline. Up to 26 stores (13 a column). */
export function recordBoard(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const per = Math.ceil(o.rows.length / 2);
  if (per > 13) throw new Error(`recordBoard: ${o.rows.length} rows will not fit`);
  const COLX = [0.72, 6.79], NAMEW = 2.17, MW = 0.64, VALW = 0.92, GAINW = 0.86, TOP = 1.86, BOT = 6.26, rowH = (BOT - TOP) / per;
  const heads = [...o.markLabels, o.valLabel, o.gainLabel];
  [0, 1].forEach(c => {
    let hx = COLX[c] + NAMEW;
    heads.forEach((h, i) => {
      const w = i < o.markLabels.length ? MW : i === o.markLabels.length ? VALW : GAINW;
      s.addText(h, { x: hx, y: 1.44, w, h: 0.36, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: i < o.markLabels.length ? "center" : "right", valign: "middle", isTextBox: true, margin: 0 });
      hx += w;
    });
  });
  o.rows.forEach((r, i) => {
    const c = i < per ? 0 : 1, k = i - c * per, x0 = COLX[c], y = TOP + k * rowH;
    if (k % 2) s.addShape(pptx.ShapeType.rect, { x: x0 - 0.08, y, w: 6.0, h: rowH, fill: { color: C.zebra } });
    s.addText(words(r.name, "store"), { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    r.marks.forEach((on, m) => {
      const d = 0.17, cx = x0 + NAMEW + m * MW + MW / 2;
      s.addShape(pptx.ShapeType.ellipse, { x: cx - d / 2, y: y + (rowH - d) / 2, w: d, h: d,
        fill: { color: on ? C.good : C.white }, line: { color: on ? C.good : C.ink300, width: 1 } });
    });
    const vx = x0 + NAMEW + o.markLabels.length * MW;
    s.addText(r.val ?? "", { x: vx, y, w: VALW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(r.gain ?? "", { x: vx + VALW, y, w: GAINW, h: rowH, fontFace: F, fontSize: 13, bold: true, color: C.goodInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "g");
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- YOY - group card over a month table ---------- */
export function yoy(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  /* ADDED 2026-10-14: o.fmt = { v, d, p } formats value / variance / variance %
     for rate and count metrics (default stays $M). o.ytd = [a, b] overrides the
     YTD figures - required for any ratio, where YTD is not the sum of months. */
  const FV = o.fmt?.v ?? M, FD = o.fmt?.d ?? money, FP = o.fmt?.p ?? pct;
  const a = o.ytd ? o.ytd[0] : o.months.reduce((t, r) => t + r[1], 0);
  const b = o.ytd ? o.ytd[1] : o.months.reduce((t, r) => t + r[2], 0);
  const v = b - a;
  if (!o.ytd && o.expect != null && Math.abs(v - o.expect) > 1) {
    throw new Error(`${o.title}: months sum to ${Math.round(v)}, expected ${o.expect}`);
  }
  const up = dirOf(o.metric) === "up" ? v >= 0 : v < 0;
  const tone = up ? "g" : "b";

  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 1.66, label: null });
  const cells = [
    ["2025", FV(a), C.ink], ["2026", FV(b), C.ink],
    ["VARIANCE", FD(v), up ? C.goodInk : C.badInk],
    ["VAR %", FP(v, a), up ? C.goodInk : C.badInk],
  ];
  words(o.cardLabel, "card label");
  s.addText(o.cardLabel.toUpperCase(), {
    x: 0.75, y: 1.44, w: 11.8, h: 0.3, fontFace: F, fontSize: sz(14, "card label"),
    color: C.ink600, charSpacing: 1.0, isTextBox: true, margin: 0,
  });
  cells.forEach(([l, val, col], i) => {
    const x = 0.75 + i * 3.05;
    s.addText(l, { x, y: 1.78, w: 2.9, h: 0.26, fontFace: F, fontSize: sz(14, "card sublabel"), color: C.ink600, charSpacing: 0.9, isTextBox: true, margin: 0 });
    s.addText(val, { x, y: 2.04, w: 2.95, h: 0.66, fontFace: F, fontSize: sz(30, "card value"), bold: true, color: col, isTextBox: true, margin: 0 });
  });

  const hdr = ["MONTH", "2025", "2026", "VARIANCE", "VAR %"].map(h => ({
    text: h, options: { fontFace: F, fontSize: sz(14, "table header"), bold: true, color: C.white,
                        fill: { color: C.blueLt }, align: h === "MONTH" ? "left" : "right", valign: "middle" },
  }));
  const body = o.months.map(([m, x, y], i) => {
    const d = y - x, good = dirOf(o.metric) === "up" ? d >= 0 : d < 0;
    const col = good ? C.goodInk : C.badInk, bg = i % 2 ? C.zebra : C.white;
    const cell = (t, opts = {}) => ({ text: t, options: { fontFace: F, fontSize: sz(14, "table body"), color: C.ink, fill: { color: bg }, valign: "middle", ...opts } });
    return [cell(m), cell(FV(x), { align: "right" }), cell(FV(y), { align: "right" }),
            cell(FD(d), { align: "right", color: col, bold: true }),
            cell(FP(d, x), { align: "right", color: col, bold: true })];
  });
  s.addTable([hdr, ...body], {
    x: 0.5, y: 3.08, w: 12.33, colW: [3.41, 2.23, 2.23, 2.23, 2.23],
    rowH: 0.3, valign: "middle", border: { type: "solid", pt: 0.5, color: C.rule },
  });

  takeaway(pptx, s, o.takeaway, tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- RANKER - every store on one page ---------- */
export function ranker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  // a store with no prior-year base cannot have a variance: dash, last, and
  // out of the group total, the bar scale and every count
  /* ADDED 2026-10-15: o.deltas = [[name, change]] for rate metrics, where the
     group change is the group's own rate change (o.groupD), not a sum of stores.
     o.fmtD formats the change (default money). A "neutral" metric sorts largest
     first and draws every bar in blue - no green/red judgment. */
  const FDm    = o.fmtD ?? money;
  const dir    = dirOf(o.metric), neutral = dir === "neutral";
  const noBase = o.deltas ? [] : o.rows.filter(r => r[1] === 0).map(r => r[0]);
  const based  = o.deltas ? o.deltas.map(([n, d]) => ({ n, d }))
                          : o.rows.filter(r => r[1] !== 0).map(([n, x, y]) => ({ n, d: y - x }));
  const good   = dir === "up" || neutral;
  based.sort((p, q) => good ? q.d - p.d : p.d - q.d);
  const all    = [...based, ...noBase.map(n => ({ n, d: null, noBase: true }))];
  const groupD = o.groupD ?? based.reduce((t, r) => t + r.d, 0);
  if (o.groupD == null && o.expect != null && Math.abs(groupD - o.expect) > 1) {
    throw new Error(`${o.title}: store rows sum to ${Math.round(groupD)}, expected ${o.expect}`);
  }

  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });

  // GROUP as its own strip - its bar is the sum, so it never joins the scale
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.46, w: 11.89, h: 0.46, fill: { color: C.blueTint } });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.46, w: 0.065, h: 0.46, fill: { color: C.navy } });
  s.addText("GROUP", { x: 0.94, y: 1.46, w: 2.2, h: 0.46, fontFace: F, fontSize: sz(15, "group strip"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
  words(o.groupLabel, "group label");
  s.addText(o.groupLabel ?? "GROUP TOTAL", {
    x: 3.1, y: 1.46, w: 4.6, h: 0.46, fontFace: F, fontSize: sz(14, "group strip sub"),
    color: C.ink600, charSpacing: 0.9, valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText(FDm(groupD), { x: 8.9, y: 1.46, w: 3.6, h: 0.46, fontFace: F, fontSize: sz(20, "group strip value"), bold: true, color: C.navy, align: "right", valign: "middle", isTextBox: true, margin: 0 });

  const N = all.length, perCol = Math.ceil(N / 2);
  const TOP = 2.02, BOT = 6.3, rowH = (BOT - TOP) / perCol;
  const maxAbs = Math.max(...based.map(r => Math.abs(r.d)));
  const COLX = [0.72, 6.79], NAMEW = 2.45, BARW = 1.8, VALW = 1.39;   // 5.82 per column + 0.25 gutter

  all.forEach((r, i) => {
    const col = i < perCol ? 0 : 1, k = i - col * perCol;
    const x0 = COLX[col], y = TOP + k * rowH;
    const isUp = r.d >= 0, favourable = good ? isUp : !isUp;
    s.addText(r.n, {
      x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13,   // exempt: 34 on one page
      color: r.noBase ? C.ink600 : C.ink, valign: "middle", isTextBox: true, margin: 0,
    });
    const bx = x0 + NAMEW + 0.1, mid = bx + BARW / 2;
    if (!r.noBase) {
      const w = Math.abs(r.d) / maxAbs * (BARW / 2), bh = rowH * 0.46;
      s.addShape(pptx.ShapeType.rect, {
        x: isUp ? mid : mid - w, y: y + (rowH - bh) / 2, w: Math.max(w, 0.014), h: bh,
        fill: { color: neutral ? C.blue : (favourable ? C.good : C.bad) },
      });
    }
    s.addText(r.noBase ? money(null) : FDm(r.d), {
      x: bx + BARW + 0.08, y, w: VALW, h: rowH, fontFace: F, fontSize: 13,
      color: r.noBase ? C.ink300 : (neutral ? C.ink : (favourable ? C.goodInk : C.badInk)),
      align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
  });
  [0, 1].forEach(c => {
    const mid = COLX[c] + NAMEW + 0.1 + BARW / 2;
    s.addShape(pptx.ShapeType.rect, { x: mid - 0.006, y: TOP, w: 0.012, h: BOT - TOP, fill: { color: C.ink300 } });
  });

  // takeaway derived, never typed
  const best = based[0], worst = based[based.length - 1];
  const downN = based.filter(r => good ? r.d < 0 : r.d > 0).length;
  const t = o.takeaway ?? `${best.n} leads at ${money(best.d)}. ${downN} stores declined, worst ${worst.n} at ${money(worst.d)}.`;
  takeaway(pptx, s, t, o.tone ?? (neutral ? "n" : (groupD >= 0 === good ? "g" : "b")));
  foot(pptx, s, o.source + (noBase.length ? `  ${noBase.join(", ")} \u2014 no prior-year base, excluded from the group.` : ""), o.page);
  return s;
}

/* ---------- MATH CHAIN ---------- */
export function chain(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const sum = o.steps.reduce((t, st) => t + st.v, 0);
  if (Math.abs(sum - o.total) > 2) throw new Error(`${o.title}: chain sums to ${Math.round(sum)}, total says ${o.total}`);

  panel(pptx, s, { x: 0.5, y: 1.32, w: 6.0, h: 5.08, label: o.panelLabel });
  let cy = 1.86;
  o.steps.forEach((st, i) => {
    if (i > 0) {
      s.addShape(pptx.ShapeType.ellipse, { x: 1.08, y: cy, w: 0.32, h: 0.32, fill: { color: C.blue } });
      s.addText(st.v >= 0 ? "+" : "\u2212", { x: 1.08, y: cy, w: 0.32, h: 0.32, fontFace: F, fontSize: sz(18, "operator"), bold: true, color: C.white, align: "center", valign: "middle", isTextBox: true, margin: 0 });
      cy += 0.34;
    }
    s.addShape(pptx.ShapeType.rect, { x: 0.78, y: cy, w: 5.44, h: 0.68, fill: { color: i === 0 ? C.zebra : C.goodBg } });
    s.addText(i === 0 ? M(st.v) : money(st.v), { x: 0.98, y: cy + 0.03, w: 3.0, h: 0.5, fontFace: F, fontSize: sz(26, "chain value"), bold: true, color: i === 0 ? C.ink : C.goodInk, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(st.l, { x: 4.0, y: cy + 0.03, w: 2.05, h: 0.64, fontFace: F, fontSize: sz(14, "chain label"), color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    cy += 0.76;
  });
  s.addShape(pptx.ShapeType.ellipse, { x: 1.08, y: cy, w: 0.32, h: 0.32, fill: { color: C.navy } });
  s.addText("=", { x: 1.08, y: cy, w: 0.32, h: 0.32, fontFace: F, fontSize: sz(18, "operator"), bold: true, color: C.white, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  cy += 0.34;
  s.addShape(pptx.ShapeType.rect, { x: 0.78, y: cy, w: 5.44, h: 0.78, fill: { color: C.blueTint } });
  if (cy + 0.78 > 6.4) throw new Error("math chain overflows its panel - too many steps for one slide");
  s.addText(M(o.total), { x: 0.98, y: cy + 0.03, w: 3.0, h: 0.56, fontFace: F, fontSize: sz(30, "chain total"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
  s.addText(o.totalLabel, { x: 4.0, y: cy + 0.03, w: 2.05, h: 0.72, fontFace: F, fontSize: sz(14, "chain label"), color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });

  // movers rail
  panel(pptx, s, { x: 6.7, y: 1.32, w: 6.13, h: 5.08, label: o.moversLabel });
  const maxM = Math.max(...o.movers.map(m => Math.abs(m[1])));
  let my = 2.0;
  o.movers.forEach(([n, d]) => {
    const up = d > 0;
    s.addText(n, { x: 6.95, y: my, w: 2.35, h: 0.42, fontFace: F, fontSize: sz(14, "mover name"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    const bx = 9.4, bw = 1.6, w = Math.abs(d) / maxM * bw;
    s.addShape(pptx.ShapeType.rect, { x: up ? bx : bx - w, y: my + 0.11, w: Math.max(w, 0.02), h: 0.2, fill: { color: up ? C.good : C.bad } });
    s.addText(money(d), { x: 11.1, y: my, w: 1.5, h: 0.42, fontFace: F, fontSize: sz(14, "mover value"), bold: true, color: up ? C.goodInk : C.badInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    my += 0.47;
  });
  s.addShape(pptx.ShapeType.rect, { x: 9.394, y: 2.0, w: 0.012, h: my - 2.06, fill: { color: C.ink300 } });

  takeaway(pptx, s, o.takeaway, "g");
  foot(pptx, s, o.source, o.page);
  return s;
}


/* ---------- LEVEL RANKER ----------------------------------------------
   For averages, rates and percentages - anything where the question is
   "who is above the group and who is below", not "who moved the most".

   GROUP sits INSIDE the ranking as a navy bar at its own rank. That works
   here precisely because a group AVERAGE lands among the stores; a group
   TOTAL would dwarf them, which is why the variance ranker keeps GROUP in
   its own strip instead. Stores read green above the group line and red
   below it, direction-aware from the registry.
   --------------------------------------------------------------------- */
export function levelRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  const up = dirOf(o.metric) === "up";
  const rows = o.rows.map(([n, v]) => ({ n, v }));
  rows.sort((a, b) => up ? b.v - a.v : a.v - b.v);
  const G = o.group;
  /* ADDED 2026-10-14: o.bench scores stores against a published benchmark
     instead of the group; GROUP still sits at its own rank, and the line is
     drawn at the benchmark and labelled. */
  const REF = o.bench ?? G;
  const all = [...rows];
  let gi = all.findIndex(r => up ? r.v < G : r.v > G);
  if (gi < 0) gi = all.length;
  all.splice(gi, 0, { n: "GROUP", v: G, isGroup: true });

  const above = rows.filter(r => up ? r.v >= REF : r.v <= REF).length;

  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });

  const N = all.length, perCol = Math.ceil(N / 2);
  const TOP = o.bench != null ? 1.86 : 1.52, BOT = 6.24, rowH = (BOT - TOP) / perCol;
  const maxV = Math.max(...all.map(r => r.v), REF * 1.05);
  /* ADDED 2026-10-15: o.extra = [{ label, w, vals: {NAME: text, GROUP: text} }] adds plain-ink
     reference columns after the value (e.g. the ELR and door rate behind a %). Bars shrink to make room;
     column headers sit on the benchmark line's row. */
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);
  const COLX = [0.72, 6.79];
  const NAMEW = EX.length ? 2.12 : 2.45, VALW = EX.length ? 0.74 : 1.19;
  const BARW = EX.length ? 5.64 - NAMEW - VALW - EXW - 0.14 : 2.0;

  all.forEach((r, i) => {
    const col = i < perCol ? 0 : 1, k = i - col * perCol;
    const x0 = COLX[col], y = TOP + k * rowH;
    const better = up ? r.v >= REF : r.v <= REF;
    // ADDED 2026-10-15: o.neutralBelow draws below-reference stores light blue, not red. NOT USED in any deck: Tray's rule (red below) applies everywhere (Noah, Oct 15 2026)
    const fill = r.isGroup ? C.navy : (better ? C.good : (o.neutralBelow ? C.blueLt : C.bad));
    const txt  = r.isGroup ? C.navy : (better ? C.goodInk : (o.neutralBelow ? C.ink : C.badInk));

    if (r.isGroup) {
      s.addShape(pptx.ShapeType.rect, {
        x: x0 - 0.08, y: y - 0.005, w: NAMEW + BARW + VALW + EXW + 0.28, h: rowH - 0.01,
        fill: { color: C.blueTint },
      });
    }
    s.addText(r.n, {
      x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13,
      color: r.isGroup ? C.navy : C.ink, bold: !!r.isGroup,
      valign: "middle", isTextBox: true, margin: 0,
    });
    const bx = x0 + NAMEW + 0.08;
    // ADDED 2026-10-15: o.floor starts the bars above zero so close rates are readable
    const FL = o.floor ?? 0;
    const w = (r.v - FL) / (maxV - FL) * BARW, bh = rowH * 0.5;
    s.addShape(pptx.ShapeType.rect, {
      x: bx, y: y + (rowH - bh) / 2, w: Math.max(w, 0.02), h: bh, fill: { color: fill },
    });
    s.addText(o.fmt(r.v), {
      x: bx + BARW + 0.06, y, w: VALW, h: rowH, fontFace: F, fontSize: 13,
      color: txt, bold: !!r.isGroup, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
    let ex = bx + BARW + 0.06 + VALW;
    EX.forEach(e => {
      s.addText(e.vals[r.n] ?? "", { x: ex, y, w: e.w, h: rowH, fontFace: F, fontSize: 13,
        color: r.isGroup ? C.navy : C.ink600, bold: !!r.isGroup, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      ex += e.w;
    });
  });
  if (EX.length) [0, 1].forEach(c => {
    let ex = COLX[c] + NAMEW + 0.08 + BARW + 0.06;
    [{ label: o.valLabel ?? "", w: VALW }, ...EX].forEach(e => {
      s.addText(e.label, { x: ex, y: 1.44, w: e.w, h: 0.32, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      ex += e.w;
    });
  });

  // the group line (or benchmark line), drawn through both columns so above/below is visible
  [0, 1].forEach(c => {
    const bx = COLX[c] + NAMEW + 0.08;
    const gx = bx + (REF - (o.floor ?? 0)) / (maxV - (o.floor ?? 0)) * BARW;
    s.addShape(pptx.ShapeType.rect, { x: gx - 0.008, y: TOP, w: 0.016, h: BOT - TOP, fill: { color: C.navy } });
    if (o.bench != null) s.addText(EX.length ? `${o.fmt(REF)} LINE` : `BENCHMARK ${o.fmt(REF)}`, {
      x: EX.length ? gx - 0.6 : gx - 1.5, y: 1.44, w: EX.length ? 1.2 : 3.0, fit: "none", wrap: !EX.length, h: 0.32, fontFace: F, fontSize: sz(14, "benchmark label"), bold: true,
      color: C.navy, align: "center", charSpacing: EX.length ? 0 : 0.9, isTextBox: true, margin: 0,
    });
  });

  const t = o.takeaway ?? `${all[0].n} leads at ${o.fmt(all[0].v)}. ${above} stores sit above the group at ${o.fmt(G)}; ${rows.length - above} below.`;
  takeaway(pptx, s, t, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  return s;
}


/* ---------- SEGMENT RANKER ---------------------------------------------
   ADDED 2026-10-15. For metrics where luxury stores sit on a different
   scale from everyone else (a luxury average well above the rest), each
   store is scored against its OWN segment average. Luxury on the left with
   the two segment averages beneath it; non-luxury across two columns on a
   shared scale. Every store on one page, names ALL CAPS, 13pt exempt.
   Green at or above own segment average, red below (house rule: green good, red bad).
   o.segments = [{ label, rows: [[name, value]], avg, prev }, { ... }]
   --------------------------------------------------------------------- */
export function segmentRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const up = dirOf(o.metric) === "up";
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });

  const TOP = 1.92, BOT = 6.28;
  const [L, N] = o.segments;
  const rank = sg => {
    const r = sg.rows.map(([n, v]) => ({ n, v })).sort((a, b) => up ? b.v - a.v : a.v - b.v);
    let gi = r.findIndex(x => up ? x.v < sg.avg : x.v > sg.avg); if (gi < 0) gi = r.length;
    r.splice(gi, 0, { n: sg.label, v: sg.avg, isGroup: true });
    return r;
  };
  const lr = rank(L), nr = rank(N);
  const perCol = Math.ceil(nr.length / 2), rowH = (BOT - TOP) / perCol;

  const header = (t, x, w) => s.addText(t.toUpperCase(), {
    x, y: 1.44, w, h: 0.36, fontFace: F, fontSize: sz(14, "segment header"), bold: true,
    color: C.navy, charSpacing: 0.9, valign: "middle", isTextBox: true, margin: 0,
  });
  const column = (rows, x0, y0, NAMEW, BARW, VALW, avg, maxV) => {
    rows.forEach((r, k) => {
      const y = y0 + k * rowH;
      const better = up ? r.v >= avg : r.v <= avg;
      if (r.isGroup) s.addShape(pptx.ShapeType.rect, { x: x0 - 0.06, y, w: NAMEW + BARW + VALW + 0.2, h: rowH - 0.01, fill: { color: C.blueTint } });
      s.addText(r.n, { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, bold: !!r.isGroup,
        color: r.isGroup ? C.navy : C.ink, valign: "middle", isTextBox: true, margin: 0 });
      const bx = x0 + NAMEW + 0.06, bh = rowH * 0.5;
      s.addShape(pptx.ShapeType.rect, { x: bx, y: y + (rowH - bh) / 2, w: Math.max(r.v / maxV * BARW, 0.02), h: bh,
        fill: { color: r.isGroup ? C.navy : (better ? C.good : C.bad) } });   // red below own segment average (Tray's rule, confirmed by Noah Oct 15 2026)
      s.addText(o.fmt(r.v), { x: bx + BARW + 0.04, y, w: VALW, h: rowH, fontFace: F, fontSize: 13, bold: !!r.isGroup,
        color: r.isGroup ? C.navy : (better ? C.goodInk : C.badInk), align: "right", valign: "middle", isTextBox: true, margin: 0 });
    });
  };

  // luxury, left
  const LX = 0.72, LN = 1.7, LB = 1.15, LV = 0.85;
  header(`${L.label} \u00b7 ${L.rows.length} stores`, LX, LN + LB + LV);
  column(lr, LX, TOP, LN, LB, LV, L.avg, Math.max(...lr.map(r => r.v)) * 1.02);
  // the two segment averages beneath it
  let ty = TOP + lr.length * rowH + 0.14;
  [L, N].forEach(sg => {
    const d = sg.avg - sg.prev, good = up ? d >= 0 : d < 0;
    s.addText([
      { text: `${sg.label} AVG  `, options: { fontSize: sz(14, "avg label"), bold: true, color: C.ink600 } },
      { text: `${o.fmt(sg.avg)}  `, options: { fontSize: sz(18, "avg value"), bold: true, color: C.navy } },
      { text: `(${o.fmtD(d)})`, options: { fontSize: sz(14, "avg delta"), bold: true, color: good ? C.goodInk : C.badInk } },
    ], { x: LX, y: ty, w: LN + LB + LV, h: 0.4, fontFace: F, valign: "middle", isTextBox: true, margin: 0 });
    ty += 0.42;
  });
  if (ty > BOT + 0.05) throw new Error(`${o.title}: segment averages overflow the luxury column`);

  // non-luxury, two columns on a shared scale
  const NX = [4.72, 8.66], NN = 2.1, NB = 0.7, NV = 0.86;
  header(`${N.label} \u00b7 ${N.rows.length} stores`, NX[0], 7.8);
  const nMax = Math.max(...nr.map(r => r.v)) * 1.02;
  column(nr.slice(0, perCol), NX[0], TOP, NN, NB, NV, N.avg, nMax);
  column(nr.slice(perCol), NX[1], TOP, NN, NB, NV, N.avg, nMax);

  takeaway(pptx, s, o.takeaway, "n");
  foot(pptx, s, o.source, o.page);
  return s;
}


/* ---------- MIX BARS ----------------------------------------------------
   ADDED 2026-10-15. 100% stacked bar per store - how a total splits across
   a few parts (pay type, department). Parts are identity, so they wear the
   blue ramp and grey, never green/red. GROUP sits in its own strip on top.
   o.parts = [{ label, color }], o.rows = [[name, [share, share, ...]]],
   o.group = [share, ...]. Sorted by the first part, largest first.
   --------------------------------------------------------------------- */
export function mixBars(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const rows = [...o.rows].sort((a, b) => b[1][0] - a[1][0]);
  rows.forEach(r => { const t = r[1].reduce((x, y) => x + y, 0); if (Math.abs(t - 1) > 0.001) throw new Error(`${o.title}: ${r[0]} shares sum to ${t}`); });

  // legend + GROUP strip
  let lx = 0.72;
  o.parts.forEach(p => {
    s.addShape(pptx.ShapeType.rect, { x: lx, y: 1.5, w: 0.22, h: 0.22, fill: { color: p.color } });
    s.addText(p.label + " %", { x: lx + 0.3, y: 1.42, w: 1.85, h: 0.38, fontFace: F, fontSize: sz(14, "legend"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    lx += 2.15;   // three legend items end before the GROUP strip at 7.3in
  });
  const bar = (x, y, w, h, shares, big) => {
    let cx = x;
    shares.forEach((v, i) => {
      const sw = v * w;
      s.addShape(pptx.ShapeType.rect, { x: cx, y, w: Math.max(sw, 0.001), h, fill: { color: o.parts[i].color }, line: { color: C.white, width: 0.5 } });
      /* CHANGED 2026-10-15: bare numbers (the legend carries the %), and a label is
         drawn whenever its measured width fits - Inter bold digits ~0.6em, +0.06in pad.
         Anything that still cannot fit is left blank and read from the speaker notes. */
      const txt = (v * 100).toFixed(0), fs = big ? 14 : 13;
      if (sw >= txt.length * 0.6 * fs / 72 + 0.06) s.addText(txt, { x: cx, y, w: sw, h, fontFace: F, fontSize: big ? sz(14, "mix label") : 13, bold: true,
        color: o.parts[i].dark ? C.white : C.ink, align: "center", valign: "middle", wrap: false, isTextBox: true, margin: 0 });
      cx += sw;
    });
  };
  s.addShape(pptx.ShapeType.rect, { x: 7.3, y: 1.42, w: 5.31, h: 0.42, fill: { color: C.blueTint } });
  s.addText("GROUP", { x: 7.4, y: 1.42, w: 1.0, h: 0.42, fontFace: F, fontSize: sz(15, "group strip"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
  bar(8.45, 1.47, 4.06, 0.32, o.group, true);

  const perCol = Math.ceil(rows.length / 2), TOP = 2.0, BOT = 6.28, rowH = (BOT - TOP) / perCol;
  const COLX = [0.72, 6.79], NAMEW = 2.3, BARW = 3.45;
  rows.forEach((r, i) => {
    const c = i < perCol ? 0 : 1, k = i - c * perCol, x0 = COLX[c], y = TOP + k * rowH;
    s.addText(r[0], { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    bar(x0 + NAMEW + 0.05, y + rowH * 0.12, BARW, rowH * 0.76, r[1], false);
  });
  takeaway(pptx, s, o.takeaway, "n");
  foot(pptx, s, o.source, o.page);
  return s;
}


/* ---------- SPLIT RANKER ------------------------------------------------
   ADDED 2026-10-15. Two groups of stores side by side (e.g. labor grid vs no
   grid), each ranked on its own with its own average row. Scored against a
   benchmark (o.bench) or neutral blue when there is no judgment to make.
   o.sides = [{ label, rows: [[name, value]], avg, avgLabel }]
   o.extra = [{ label, w, vals: {NAME: text} }] plain-ink reference columns.
   --------------------------------------------------------------------- */
export function splitRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const dir = dirOf(o.metric), up = dir === "up" || dir === "neutral";   // neutral lists largest first
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);
  const COLX = [0.72, 6.79], NAMEW = 2.1, VALW = 0.82, BARW = 5.7 - NAMEW - VALW - EXW - 0.14;
  const sides = o.sides.map(sd => {
    const r = sd.rows.map(([n, v]) => ({ n, v })).sort((a, b) => up ? b.v - a.v : a.v - b.v);
    if (sd.avg != null) { let gi = r.findIndex(x => up ? x.v < sd.avg : x.v > sd.avg); if (gi < 0) gi = r.length;
      r.splice(gi, 0, { n: sd.avgLabel ?? "AVERAGE", v: sd.avg, isGroup: true }); }
    return r;
  });
  const perCol = Math.max(...sides.map(r => r.length)), TOP = 1.92, BOT = 6.28, rowH = (BOT - TOP) / perCol;
  const FL = o.floor ?? 0, maxV = Math.max(...sides.flat().map(r => r.v), (o.bench ?? 0) * 1.05);
  sides.forEach((rows, c) => {
    const x0 = COLX[c], bx = x0 + NAMEW + 0.08;
    s.addText(`${o.sides[c].label} \u00b7 ${o.sides[c].rows.length} stores`.toUpperCase(), { x: x0, y: 1.44, w: NAMEW + BARW, h: 0.36,
      fontFace: F, fontSize: sz(14, "side header"), bold: true, color: C.navy, charSpacing: 0.9, valign: "middle", isTextBox: true, margin: 0 });
    let hx = bx + BARW + 0.06;
    [{ label: o.valLabel ?? "", w: VALW }, ...EX].forEach(e => {
      s.addText(e.label, { x: hx, y: 1.44, w: e.w, h: 0.36, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      hx += e.w;
    });
    rows.forEach((r, k) => {
      const y = TOP + k * rowH;
      const better = o.bench == null ? true : (up ? r.v >= o.bench : r.v <= o.bench);
      const fill = r.isGroup ? C.navy : (o.bench == null ? C.blue : (better ? C.good : C.bad));
      const txt = r.isGroup ? C.navy : (o.bench == null ? C.ink : (better ? C.goodInk : C.badInk));
      if (r.isGroup) s.addShape(pptx.ShapeType.rect, { x: x0 - 0.06, y, w: NAMEW + BARW + VALW + EXW + 0.22, h: rowH - 0.01, fill: { color: C.blueTint } });
      s.addText(r.n, { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, bold: !!r.isGroup, color: r.isGroup ? C.navy : C.ink, valign: "middle", isTextBox: true, margin: 0 });
      const bh = rowH * 0.5;
      s.addShape(pptx.ShapeType.rect, { x: bx, y: y + (rowH - bh) / 2, w: Math.max((r.v - FL) / (maxV - FL) * BARW, 0.02), h: bh, fill: { color: fill } });
      s.addText(o.fmt(r.v), { x: bx + BARW + 0.06, y, w: VALW, h: rowH, fontFace: F, fontSize: 13, bold: !!r.isGroup, color: txt, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      let ex = bx + BARW + 0.06 + VALW;
      EX.forEach(e => { s.addText(e.vals[r.n] ?? "", { x: ex, y, w: e.w, h: rowH, fontFace: F, fontSize: 13, bold: !!r.isGroup,
        color: r.isGroup ? C.navy : C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 }); ex += e.w; });
    });
    if (o.bench != null) {
      const gx = bx + (o.bench - FL) / (maxV - FL) * BARW;
      s.addShape(pptx.ShapeType.rect, { x: gx - 0.008, y: TOP, w: 0.016, h: rows.length * rowH, fill: { color: C.navy } });
    }
  });
  s.addShape(pptx.ShapeType.rect, { x: 6.6, y: 1.5, w: 0.012, h: 4.78, fill: { color: C.rule } });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  return s;
}


/* ---------- STAT TILES ---------------------------------------------------
   ADDED 2026-10-16 (proficiency opportunity). Up to four big-number tiles in
   one row. o.tiles = [{ big, label, sub, tone }] tone g/b/n colours the number.
   o.banner = one plain line under the tiles.
   --------------------------------------------------------------------- */
export function statTiles(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const n = o.tiles.length, gap = 0.22, X0 = 0.78, W = (11.77 - gap * (n - 1)) / n;
  o.tiles.forEach((t, i) => {
    const x = X0 + i * (W + gap), ink = t.tone === "b" ? C.badInk : t.tone === "g" ? C.goodInk : C.navy;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.62, w: W, h: 3.3, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.62, w: W, h: 0.07, fill: { color: t.tone === "b" ? C.bad : t.tone === "g" ? C.good : C.blue } });
    s.addText(words(t.label, "tile label").toUpperCase(), { x: x + 0.2, y: 1.85, w: W - 0.4, h: 0.62, fontFace: F, fontSize: sz(14, "tile label"), bold: true,
      color: C.ink600, charSpacing: 0.8, valign: "top", isTextBox: true, margin: 0 });
    s.addText(words(t.big, "tile number"), { x: x + 0.2, y: 2.5, w: W - 0.4, h: 0.95, fontFace: F, fontSize: sz(o.bigSize ?? 40, "tile number"), bold: true,   // o.bigSize ADDED 2026-10-19
     
      color: ink, valign: "middle", fit: "shrink", isTextBox: true, margin: 0 });
    s.addText(words(t.sub, "tile sub"), { x: x + 0.2, y: 3.5, w: W - 0.4, h: 1.3, fontFace: F, fontSize: sz(15, "tile sub"),
      color: C.ink, valign: "top", isTextBox: true, margin: 0 });
  });
  if (o.banner) s.addText(words(o.banner, "banner"), { x: 0.78, y: 5.15, w: 11.77, h: 1.0, fontFace: F, fontSize: sz(17, "banner"),
    color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- AMOUNT RANKER ------------------------------------------------
   ADDED 2026-10-16. Every store ranked by a dollar amount that is all one
   colour (o.color, e.g. red for money left on the table) - no group line,
   because the bars are sized, not scored. o.rows = [[name, value]],
   o.extra = [{ label, w, vals: {NAME: text} }] plain-ink columns.
   --------------------------------------------------------------------- */
export function amountRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const rows = [...o.rows].sort((a, b) => b[1] - a[1]);
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);
  const COLX = [0.72, 6.79], NAMEW = 2.12, VALW = o.valW ?? 0.86, BARW = 5.7 - NAMEW - VALW - EXW - 0.14;   // o.valW ADDED 2026-10-19 for 7-digit dollars
  const perCol = Math.ceil(rows.length / 2), TOP = 1.86, BOT = 6.26, rowH = (BOT - TOP) / perCol, maxV = rows[0][1];
  [0, 1].forEach(c => {
    let hx = COLX[c] + NAMEW + 0.08 + BARW + 0.06;
    [{ label: o.valLabel ?? "", w: VALW }, ...EX].forEach(e => {
      s.addText(e.label, { x: hx, y: 1.44, w: e.w, h: 0.36, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      hx += e.w;
    });
  });
  rows.forEach((r, i) => {
    const c = i < perCol ? 0 : 1, k = i - c * perCol, x0 = COLX[c], y = TOP + k * rowH, bx = x0 + NAMEW + 0.08;
    s.addText(r[0], { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    const bh = rowH * 0.55;
    s.addShape(pptx.ShapeType.rect, { x: bx, y: y + (rowH - bh) / 2, w: Math.max(r[1] / maxV * BARW, 0.02), h: bh, fill: { color: o.color ?? C.blue } });
    s.addText(o.fmt(r[1]), { x: bx + BARW + 0.06, y, w: VALW, h: rowH, fontFace: F, fontSize: 13, bold: true,
      color: o.ink ?? C.ink, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    let ex = bx + BARW + 0.06 + VALW;
    EX.forEach(e => { s.addText(e.vals[r[0]] ?? "", { x: ex, y, w: e.w, h: rowH, fontFace: F, fontSize: 13,
      color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 }); ex += e.w; });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- DUMBBELL -----------------------------------------------------
   ADDED 2026-10-16. Two values per store on one scale: o.rows = [[name, a, b]]
   a = the store (red below o.bench, green at/above), b = the comparison
   (navy). Line joins them; o.bench draws the target line. Sorted by a, best first.
   --------------------------------------------------------------------- */
export function dumbbell(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 12.33, h: 5.08, label: null });
  const rows = [...o.rows].sort((p, q) => q[1] - p[1]);
  const COLX = [0.72, 6.79], NAMEW = 2.12, VALW = 1.45, PLOTW = 5.7 - NAMEW - VALW - 0.3;
  const perCol = Math.ceil(rows.length / 2), TOP = 1.9, BOT = 6.26, rowH = (BOT - TOP) / perCol;
  const lo = o.min, hi = o.max, X = (bx, v) => bx + (Math.min(Math.max(v, lo), hi) - lo) / (hi - lo) * PLOTW;
  // legend
  [[C.bad, o.aLabel + " (green at or above " + o.fmt(o.bench) + ")"], [C.navy, o.bLabel]].forEach(([col, lab], i) => {
    const lx = 0.72 + i * 6.07;
    s.addShape(pptx.ShapeType.ellipse, { x: lx, y: 1.5, w: 0.2, h: 0.2, fill: { color: col } });
    s.addText(lab, { x: lx + 0.3, y: 1.42, w: 5.6, h: 0.36, fontFace: F, fontSize: sz(14, "legend"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
  });
  [0, 1].forEach(c => {
    const bx = COLX[c] + NAMEW + 0.1, gx = X(bx, o.bench);
    const n = c === 0 ? perCol : rows.length - perCol;
    s.addShape(pptx.ShapeType.rect, { x: gx - 0.008, y: TOP, w: 0.016, h: n * rowH, fill: { color: C.ink300 } });
  });
  rows.forEach((r, i) => {
    const c = i < perCol ? 0 : 1, k = i - c * perCol, x0 = COLX[c], y = TOP + k * rowH, bx = x0 + NAMEW + 0.1, cy = y + rowH / 2;
    const ok = r[1] >= o.bench, d = 0.17;
    s.addText(r[0], { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    const xa = X(bx, r[1]), xb = X(bx, r[2]);
    s.addShape(pptx.ShapeType.rect, { x: Math.min(xa, xb), y: cy - 0.015, w: Math.max(Math.abs(xb - xa), 0.01), h: 0.03, fill: { color: C.rule } });
    s.addShape(pptx.ShapeType.ellipse, { x: xb - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: C.navy } });
    s.addShape(pptx.ShapeType.ellipse, { x: xa - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: ok ? C.good : C.bad } });
    s.addText([{ text: o.fmt(r[1]), options: { color: ok ? C.goodInk : C.badInk, bold: true } }, { text: "  " + o.fmt(r[2]), options: { color: C.navy } }],
      { x: bx + PLOTW + 0.15, y, w: VALW, h: rowH, fontFace: F, fontSize: 13, align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  return s;
}


/* ---------- TALKING POINTS ----------------------------------------------
   ADDED 2026-10-19. The "talking points" step of the segment rhythm. Left:
   what we have already told this room (o.said = bullets, o.saidFrom = where it
   came from). Right: up to three data points that say where the group is now
   (o.now = [{ big, label, tone }], tone g/b/n). Takeaway band = the ask.
   --------------------------------------------------------------------- */
export function talkingPoints(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: 1.32, w: 7.0, h: 5.08, label: null });
  panel(pptx, s, { x: 7.68, y: 1.32, w: 5.15, h: 5.08, label: null });
  const head = (x, w, t) => s.addText(words(t, "tp head").toUpperCase(), { x, y: 1.5, w, h: 0.34, fontFace: F, fontSize: sz(14, "tp head"),
    bold: true, color: C.navy, charSpacing: 1.0, valign: "middle", isTextBox: true, margin: 0 });
  head(0.78, 6.5, "What we said · " + o.saidFrom);
  head(7.95, 4.7, o.nowLabel ?? "Where we are · Jan–Sep 2026");
  s.addShape(pptx.ShapeType.rect, { x: 0.78, y: 1.9, w: 6.45, h: 0.012, fill: { color: C.rule } });
  s.addShape(pptx.ShapeType.rect, { x: 7.95, y: 1.9, w: 4.62, h: 0.012, fill: { color: C.rule } });
  const fs = o.said.length > 6 ? 16 : 18;
  s.addText(o.said.map(t => ({ text: words(t, "tp bullet"), options: { bullet: { indent: 18 }, breakLine: true, paraSpaceAfter: 8 } })), {
    x: 0.78, y: 2.05, w: 6.45, h: 4.2, fontFace: F, fontSize: sz(fs, "tp bullet"), color: C.ink, valign: "top", isTextBox: true, margin: 0 });
  const n = o.now.length, rowH = 4.25 / n;
  o.now.forEach((d, i) => {
    const y = 2.05 + i * rowH, ink = d.tone === "g" ? C.goodInk : d.tone === "b" ? C.badInk : C.navy;
    s.addShape(pptx.ShapeType.rect, { x: 7.95, y: y + 0.05, w: 0.06, h: rowH - 0.25, fill: { color: d.tone === "g" ? C.good : d.tone === "b" ? C.bad : C.blue } });
    s.addText(words(d.big, "tp number"), { x: 8.15, y: y, w: 4.4, h: 0.62, fontFace: F, fontSize: sz(30, "tp number"), bold: true, color: ink,
      valign: "middle", fit: "shrink", isTextBox: true, margin: 0 });
    s.addText(words(d.label, "tp label"), { x: 8.15, y: y + 0.62, w: 4.4, h: rowH - 0.8, fontFace: F, fontSize: sz(14, "tp label"), color: C.ink,
      valign: "top", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, "The ask: " + o.ask, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  if (o.notes) s.addNotes(words(o.notes, "notes"));
  return s;
}


/* ---------- METRIC EXPLAINER ---------------------------------------------
   ADDED 2026-10-19. "What this metric is" slide, shown before a metric's data.
   o.formula = [name, '=', term, '÷', term, ...]  (operators drawn in blue)
   o.define  = one plain sentence under the formula
   o.example = optional [[label, value], ...] worked-example chips
   o.why     = bullets;  o.levers = [[name, line], ...] up to 6
   takeaway band = our number (o.ours).
   --------------------------------------------------------------------- */
export function explainer(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const OPS = new Set(["=", "÷", "×", "−", "+"]);
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.32, w: 12.33, h: 1.12, fill: { color: C.blueTint } });
  s.addText(o.formula.map((t, i) => ({ text: words(t, "formula") + (i < o.formula.length - 1 ? "  " : ""),
      options: { color: OPS.has(t) ? C.blue : C.navy, bold: true } })),
    { x: 0.75, y: 1.36, w: 11.85, h: 0.62, fontFace: F, fontSize: sz(o.formulaSize ?? 24, "formula"), valign: "middle", fit: "shrink", isTextBox: true, margin: 0 });
  s.addText(words(o.define, "define"), { x: 0.75, y: 1.95, w: 11.85, h: 0.46, fontFace: F, fontSize: sz(o.define.length > 115 ? 14 : 15, "define"), color: C.ink600, valign: "middle", isTextBox: true, margin: 0 });
  let Y = 2.56;
  if (o.example) {
    s.addText(words(o.exampleLabel ?? "EXAMPLE", "example label"), { x: 0.5, y: Y, w: 2.0, h: 0.5, fontFace: F, fontSize: sz(14, "example label"), bold: true, color: C.ink600, charSpacing: 0.8, valign: "middle", isTextBox: true, margin: 0 });
    const n = o.example.length, gap = 0.12, W = (12.33 - 2.0 - gap * (n - 1)) / n;
    o.example.forEach(([l, v], i) => {
      const x = 2.5 + i * (W + gap);
      s.addShape(pptx.ShapeType.rect, { x, y: Y, w: W, h: 0.5, fill: { color: C.panel }, line: { color: C.rule, width: 0.75 } });
      s.addText([{ text: words(l, "example") + "  ", options: { color: C.ink } }, { text: words(v, "example"), options: { color: C.navy, bold: true } }],
        { x: x + 0.08, y: Y, w: W - 0.16, h: 0.5, fontFace: F, fontSize: sz(14, "example"), align: "center", valign: "middle", fit: "shrink", isTextBox: true, margin: 0 });
    });
    Y += 0.62;
  }
  const H2 = 6.4 - Y;
  panel(pptx, s, { x: 0.5, y: Y, w: 4.75, h: H2, label: null });
  panel(pptx, s, { x: 5.4, y: Y, w: 7.43, h: H2, label: null });
  const head = (x, w, t) => s.addText(t, { x, y: Y + 0.14, w, h: 0.32, fontFace: F, fontSize: sz(14, "explainer head"), bold: true, color: C.navy, charSpacing: 1.0, valign: "middle", isTextBox: true, margin: 0 });
  head(0.72, 4.3, o.whyHead ?? "WHY IT MATTERS"); head(5.62, 7.0, o.leverHead ?? "WHAT MOVES IT");
  s.addText(o.why.map(t => ({ text: words(t, "why"), options: { bullet: { indent: 16 }, breakLine: true, paraSpaceAfter: 6 } })),
    { x: 0.72, y: Y + 0.55, w: 4.35, h: H2 - 0.7, fontFace: F, fontSize: sz(o.why.length > 3 || H2 < 3.4 ? 15 : 16, "why"), color: C.ink, valign: "top", isTextBox: true, margin: 0 });
  const L = o.levers, COLS = L.length > 4 && H2 < 3.6 ? 3 : 2, rows = Math.ceil(L.length / COLS), cw = (7.0 - 0.12 * (COLS - 1)) / COLS, ch = (H2 - 0.65) / rows;
  L.forEach(([name, line], i) => {
    const x = 5.62 + (i % COLS) * (cw + 0.12), y = Y + 0.55 + Math.floor(i / COLS) * ch;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.05, h: ch - 0.12, fill: { color: C.blue } });
    s.addText([{ text: words(name, "lever"), options: { bold: true, color: C.ink, breakLine: true } }, { text: words(line, "lever"), options: { color: C.ink600 } }],
      { x: x + 0.14, y, w: cw - 0.2, h: ch - 0.12, fontFace: F, fontSize: sz(14, "lever"), valign: "top", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, o.ours, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  if (o.notes) s.addNotes(words(o.notes, "notes"));
  return s;
}
