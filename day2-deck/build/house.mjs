/* ===================================================================
   FINDLAY MEETING DECK - HOUSE STYLE
   Locked October 2026 with Tommy. Decisions this file encodes:

     font        Inter          (must be installed on any machine that opens the deck)
     field       grey EDEDED, content on white panels
     titles      the SOURCE LINE NAME and nothing else; the finding goes in the takeaway.
                 Department blocks (New Vehicle, Used Vehicle, F&I): the title is
                 the department and the ranker strip names the measure - "New
                 Vehicle" / "UNITS" (Tommy, Oct 18 2026)
     colour      green good, red bad, always. blue is identity, never judgment
     stores      ALL CAPS, all 34 on one page, GROUP as its own strip.
                 say "stores", never "rooftops"
     floor       nothing carrying meaning below 14pt - enforced, not remembered
     numbers     counts and totals DERIVED from the series, never typed

     takeaway    OPTIONAL on every slide type. A band renders only when you pass
                 an explicit string. There is no "auto" - a derived sentence that
                 nobody read was worse than no sentence at all. With no band the
                 body panels grow into the space, so the frame still reads as
                 finished rather than cut short.

     tracker     transition slides only, NEVER the cover. Ten sections over two
                 rows. "Sections", not "segments".

   DAY 2 (fixed ops, Oct 29 2026) - Noah, Oct 21 2026: Day 2 is rebuilt on
   this Day 1 engine so both days read as one meeting. Added on top, nothing
   of Day 1's look changed: the Day 2 renderers (segmentRanker, mixBars,
   splitRanker, statTiles, amountRanker, dumbbell, talkingPoints, explainer,
   trendYoy, yearBars, recordBoard, stackedCountRanker, gapRanker), the fixed
   ops metric registry, "neutral" metrics (blue, no judgment), benchmark and
   extra columns on levelRanker, rate options on yoy/trend, and the words()
   guard that stops the build on "rooftop". The v1 engine is build/v1/house.v1.mjs.

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
   Added 2026-10-14 (Noah). Slide text passes through words(), which stops
   the build rather than ship the old term. */
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
  12: "ranker rows once rank numbers joined them - Tommy traded size for the\n       numbering, Oct 2026. Also the source line, which is reference\n       material and not read from the back of the room",
};
export function sz(size, why) {
  if (size < FLOOR && !EXEMPT[size]) {
    throw new Error(`font ${size}pt below the ${FLOOR}pt floor (${why || "unlabelled"})`);
  }
  return size;
}

/* ---------- body geometry -------------------------------------------
   One pair of numbers decides how tall every content panel is. A slide
   with a band stops short of it; a slide without one runs on down to a
   bottom margin that matches the sides.
   --------------------------------------------------------------------- */
export const BODY = { top: 1.32, botFull: 6.95 };

/* ---------- the takeaway band (Tommy, Oct 18 2026) --------------------
   The band's TOP edge is fixed, so a chart is the same height however long
   the line is - "I don't want any of the charts to be shrunk to
   accommodate the takeaway."

   One row is the rule. The font steps down from 16pt to the 14pt floor to
   hold one row. Only if it still will not fit does the band take a second
   row, and that row grows DOWN into the grey bottom margin, not up into
   the chart. Three rows is refused: shorten the line.

   Widths are estimated from character count (Inter bold averages about
   0.54 em). Deliberately on the generous side - an estimate that is too
   tight lets PowerPoint wrap a "one-row" band out of its box.
   --------------------------------------------------------------------- */
const BAND = { top: 6.70, w: 11.95, em: 0.54, pad: 0.14 };
export function bandFit(t) {
  if (!t || namesStore(t)) return null;
  for (let lines = 1; lines <= 2; lines++) {
    for (let fs = 16; fs >= FLOOR; fs -= 0.5) {
      if (Math.ceil(t.length * BAND.em * fs / 72 / BAND.w) <= lines) return { fs, lines };
    }
  }
  throw new Error(`takeaway will not fit two rows at the ${FLOOR}pt floor - shorten it: "${t}"`);
}
export function bandH(t) {
  const f = bandFit(t);
  return f ? f.lines * f.fs * 1.2 / 72 + BAND.pad : 0;
}
export const bodyBot = t => (t && !namesStore(t) ? BAND.top - 0.15 : BODY.botFull);
export const bodyH   = t => bodyBot(t) - BODY.top;

/* ---------- the metric registry ---------- */
export const METRICS = {
  "service gross":        { dir: "up",   fmt: "money" },
  "parts gross":          { dir: "up",   fmt: "money" },
  "fixed ops gross":      { dir: "up",   fmt: "money" },
  "net profit":           { dir: "up",   fmt: "money" },
  "elr":                  { dir: "up",   fmt: "money" },
  "hpro":                 { dir: "up",   fmt: "dec"   },
  "ro count":             { dir: "up",   fmt: "int"   },
  "service absorption":   { dir: "up",   fmt: "pct"   },
  "mpi video %":          { dir: "up",   fmt: "pct"   },
  "cp share":             { dir: "up",   fmt: "pct"   },   // customer-pay share of service gross
  // Oct 18 2026 - variable ops. Units, gross and PVR up-good per the house
  // rule ("income, gross, PVR, units - up is good"); no new judgment made.
  "new units":            { dir: "up",   fmt: "int"   },
  "new gross":            { dir: "up",   fmt: "money" },
  "new pvr":              { dir: "up",   fmt: "money" },
  "used units":           { dir: "up",   fmt: "int"   },
  "used gross":           { dir: "up",   fmt: "money" },
  "used pvr":             { dir: "up",   fmt: "money" },
  "fi gross":             { dir: "up",   fmt: "money" },
  "fi pvr":               { dir: "up",   fmt: "money" },
  // chargebacks are a cost: up is bad (house rule - "chargebacks ... up is bad")
  "chargebacks pvr":      { dir: "down", fmt: "money" },
  // Oct 18 2026 - wholesale & writedowns, inside Used Vehicle. Lower is
  // better for both - CONFIRMED by Tommy, Oct 18 2026.
  // Oct 18 2026 - net profit ratios (January's bonus KPIs). Profit per unit
  // and per dollar - up is good under the house rule; no new judgment.
  "net per unit":         { dir: "up",   fmt: "money" },
  "net to gross":         { dir: "up",   fmt: "pct"   },
  "net to sale":          { dir: "up",   fmt: "pct"   },
  "wholesale %":          { dir: "down", fmt: "pct"   },
  "writedowns pvr":       { dir: "down", fmt: "money" },
  // ---- Day 2, fixed ops (Noah / Tommy, Oct 14-21 2026). "neutral" = blue, no green/red judgment.
  // 2026-10-14 (Noah, opcode test build) - PROVISIONAL until Tray confirms.
  // Share of customer-pay ROs carrying the service; more is better.
  "alignment %":          { dir: "up",   fmt: "pct"   },
  "balance %":            { dir: "up",   fmt: "pct"   },
  "brake fluid %":        { dir: "up",   fmt: "pct"   },
  "warranty parts gp %":  { dir: "up",   fmt: "pct"   },
  "mpi completion %":     { dir: "up",   fmt: "pct"   },   // Oct 22 2026: store-reported, Jul-Sep
  "video completion %":   { dir: "up",   fmt: "pct"   },
  "photo completion %":   { dir: "up",   fmt: "pct"   },   // Oct 22 2026 (Noah): example for the warranty rate-increase story
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
  // 2026-10-21 (Noah): internal GP% coloured against the group like every other level ranker - above group green, below red.
  "internal labor gp %":  { dir: "up",   fmt: "pct"   },
  "internal parts gp %":  { dir: "up",   fmt: "pct"   },
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
  // 2026-10-16 (Noah, proficiency slide test) - PROVISIONAL until Tray confirms. Sold hours / worked hours (ADP reg + OT); workbook targets 85% on, 100% above.
  "tech proficiency %":     { dir: "up", fmt: "pct" },   // a posted price, not a score     // May benchmark: CP ELR at least 70% of door rate   // CP labor whose op code description says nothing usable - a description-quality measure   // customer-pay share of labor sale; higher = less dependence on OEM/internal - confirm
};
export function dirOf(key) {
  const m = METRICS[key];
  if (!m) throw new Error(`metric "${key}" is not in the registry - ask Tommy, then add it with a dated comment`);
  return m.dir;
}

/* ---------- formatters ---------- */
export const M     = v => "$" + (v / 1e6).toFixed(2) + "M";
/* Millions, with the minus where a reader expects it - never "$-13.42M".
   Msigned marks only negatives; Mdelta always carries a sign, because a
   change column is read for direction first and magnitude second. */
export const Msigned = v => (v < 0 ? "−" : "") + "$" + Math.abs(v / 1e6).toFixed(2) + "M";
export const Mdelta  = v => (v < 0 ? "−" : "+") + "$" + Math.abs(v / 1e6).toFixed(2) + "M";
// a change of nothing carries no sign - "+0" reads as a gain
const sign = v => Math.abs(v) < 0.5 ? "" : (v < 0 ? "−" : "+");

/* Mixed scale (Tommy, Oct 2026): millions read as $3.64M, anything under a
   million reads in full dollars - $319,857, not $0.32M. Rounding a store to
   two decimals of a million hides the difference between $319,857 and
   $324,000, and under a million that difference is the whole conversation.
   Mixed always signs the negative; Dmixed always signs both ways. */
const mixed = v => Math.abs(v) >= 1e6
  ? "$" + Math.abs(v / 1e6).toFixed(2) + "M"
  : "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
export const Mmixed = v => (v < 0 ? "−" : "") + mixed(v);
export const Dmixed = v => sign(v) + mixed(v);
/* counts - units, ROs - never with a dollar sign */
export const Int  = v => (v < 0 ? "−" : "") + Math.round(Math.abs(v)).toLocaleString("en-US");
export const Dint = v => sign(v) + Math.round(Math.abs(v)).toLocaleString("en-US");
/* shares, passed in as percent (30.9, not 0.309). A change in a share is in
   points - "−0.8 pts" - because "−0.8%" reads as a relative change. Its own
   zero test: sign() rounds to whole units and would drop the sign on 0.3. */
export const Pct  = v => (v < 0 ? "−" : "") + Math.abs(v).toFixed(1) + "%";
export const Dpts = v => (Math.abs(v) < 0.05 ? "" : v < 0 ? "−" : "+") + Math.abs(v).toFixed(1) + " pts";

/* Store names run ALL CAPS in ranker rows and chart data, and Title Case in
   prose - a takeaway is a sentence, and a sentence that SHOUTS A STORE NAME
   reads as an accusation (Tommy, Oct 2026). These stay upper: they are
   initialisms, not words. */
const KEEP_CAPS = new Set(["LR", "VW", "GMC", "CDJR", "LV", "INEOS", "BMW", "CJDR"]);
export const storeName = s => String(s).split(/\s+/)
  .map(w => KEEP_CAPS.has(w) ? w : w.charAt(0) + w.slice(1).toLowerCase())
  .join(" ");
/* Day 2: a takeaway is prose, so any store name written in ranker caps is
   turned to Title Case on the way in - "Audi Reno is lowest", never "AUDI RENO". */
export const STORES = ["TOYOTA HENDERSON","HONDA HENDERSON","CHEVY LV","AUDI HENDERSON","CADILLAC","SUBARU LV",
  "HONDA NORTH","LR LV","KIA LV","TOYOTA SPOKANE","TOYOTA PRESCOTT","VW HENDERSON","NISSAN","TOYOTA FLAGSTAFF",
  "CHEVY GMC BULLHEAD","AUDI RENO","LEXUS","HONDA SPOKANE","VOLVO","CDJR POST FALLS","SUBARU ST GEORGE",
  "LR HENDERSON","MAZDA","LINCOLN","ACURA","SUBARU PRESCOTT","HYUNDAI ST GEORGE","LR RENO","KIA ST GEORGE",
  "GMC PRESCOTT","HONDA FLAGSTAFF","HYUNDAI PRESCOTT","VW ST GEORGE","INEOS"];
const STORE_RE = new RegExp("\\b(" + [...STORES].sort((a, b) => b.length - a.length).join("|") + ")\\b", "g");
export const proseNames = t => String(t).replace(STORE_RE, m => storeName(m));
// Oct 27 (Noah): no takeaway band that names a store (who's on top or bottom is on the slide already).
// The band is dropped, the body takes its room, and the sentence moves to the speaker notes.
const NAMED_RE = new RegExp("\\b(" + [...STORES].sort((a, b) => b.length - a.length).join("|") + ")\\b", "i");
export const namesStore = t => NAMED_RE.test(String(t));
export const money = v => v === null ? "—"
  : (v < 0 ? "−" : "+") + "$" + Math.round(Math.abs(v)).toLocaleString("en-US");
export const dollars = v => "$" + Math.round(v).toLocaleString("en-US");
export const pct   = (a, b) => (a / b >= 0 ? "+" : "−") + Math.abs(a / b * 100).toFixed(1) + "%";

/* ---------- page furniture ---------- */
export function page(pptx, s, { eyebrow, title }) {
  words(eyebrow, "eyebrow"); words(title, "title");
  s.background = { color: C.field };
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.44, w: 0.055, h: 0.21, fill: { color: C.blue } });
  s.addText(eyebrow.toUpperCase(), {
    x: 0.68, y: 0.37, w: 9, h: 0.3, fontFace: F, fontSize: sz(14, "eyebrow"), bold: true,
    color: C.ink600, charSpacing: 1.1, isTextBox: true, margin: 0,
  });
  s.addText(title, {   // the source line name, exactly
    x: 0.5, y: 0.7, w: 12.4, h: 0.55, fontFace: F, fontSize: sz(28, "title"),
    color: C.ink, isTextBox: true, margin: 0,
  });
}

/* Source and page number are deliberately NOT drawn (Tommy, Oct 2026) - they ate
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

/* The strip across the top of every store ranker: the measure in bold, the
   period or comparison beside it, the group figure on the right. Named for
   the measure - "NET PROFIT", "NEW FRONT PVR" - not "GROUP" (Tommy, Oct 18
   2026): the strip IS the group, so the word added nothing and the measure
   was missing. */
/* Table dollars without the last three digits (Noah, Oct 22 2026): $296K, $1.16M. d = decimals on K
   ($14.2K) where whole thousands would hide the difference. Takeaways and notes keep the exact figure. */
export const kusd = (v, d = 0) => (v < 0 ? "−" : "") + (Math.abs(v) >= 1e6 ? "$" + (Math.abs(v) / 1e6).toFixed(2) + "M"
  : Math.abs(v) < 1e3 ? "$" + Math.round(Math.abs(v)) : "$" + (Math.abs(v) / 1e3).toFixed(d) + "K");   // under $1K stays exact, never "$0K"

export function measureStrip(pptx, s, measure, label, value, valueColor) {
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.46, w: 11.89, h: 0.46, fill: { color: C.blueTint } });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.46, w: 0.065, h: 0.46, fill: { color: C.navy } });
  // Day 2: longer measure names ("CP & W ROS PER FIXED EMPLOYEE") push the label right
  // rather than run under it - Inter bold caps ~0.62em at 15pt
  const mtxt = (measure ?? "GROUP").toUpperCase(), mw = Math.max(2.6, mtxt.length * 0.62 * 15 / 72 + 0.15);
  s.addText(mtxt, {
    x: 0.94, y: 1.46, w: mw, h: 0.46, fontFace: F, fontSize: sz(15, "strip measure"), bold: true,
    color: C.navy, valign: "middle", isTextBox: true, margin: 0,
  });
  // a period or a comparison, never a store count. **part** renders bold -
  // "YEAR-OVER-YEAR CHANGE" is always bold on a change ranker (Tommy, Oct 18 2026)
  // Day 2: a label too long for what is left of the strip drops its period
  // prefix ("2026 VS 2025 YTD · ") and keeps the bold comparison
  const room = 8.8 - (0.94 + mw + 0.06), fits = t => t.replace(/\*\*/g, "").length * 0.7 * 14 / 72 <= room;
  if (!fits(String(label)) && String(label).includes("·")) label = String(label).split("·").slice(1).join("·").trim();
  if (!fits(String(label))) label = String(label).replace("YEAR-OVER-YEAR CHANGE", "YOY CHANGE");
  const runs = String(label).split("**")
    .map((t, i) => ({ text: t, options: i % 2 ? { bold: true, color: C.navy } : {} }))
    .filter(r => r.text);
  s.addText(runs, {
    x: 0.94 + mw + 0.06, y: 1.46, w: 8.8 - (0.94 + mw + 0.06), h: 0.46, fontFace: F, fontSize: sz(14, "strip label"),
    color: C.ink600, charSpacing: 0.9, valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText(value, {
    x: 8.9, y: 1.46, w: 3.6, h: 0.46, fontFace: F, fontSize: sz(20, "strip value"), bold: true,
    color: valueColor, align: "right", valign: "middle", isTextBox: true, margin: 0,
  });
}

/* A band only when there is something worth saying. No text, no band.
   The band names where the opportunity is, with the figure attached -
   never flattery, never a scolding (Tommy, Oct 2026). */
export function takeaway(pptx, s, text, tone = "n") {
  if (!text) return;
  if (namesStore(text)) { s.addNotes(`Takeaway (off the slide): ${proseNames(text)}`); return; }
  words(text, "takeaway");
  text = proseNames(text);
  const bg = tone === "g" ? C.goodBg : tone === "b" ? C.badBg : C.blueTint;
  const fg = tone === "g" ? C.goodInk : tone === "b" ? C.badInk : C.navy;
  const { fs } = bandFit(text);
  const h = bandH(text), y = BAND.top;   // top pinned; a second row grows down
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.33, h, fill: { color: bg } });
  s.addText(text, {
    x: 0.72, y: y + 0.02, w: BAND.w, h: h - 0.04, fontFace: F, fontSize: sz(fs, "takeaway"), bold: true,
    color: fg, isTextBox: true, margin: 0, valign: "middle",
  });
}

/* ---------- THE SECTION TRACKER ---------------------------------------
   Transition slides only. Completed sections read light blue, the current
   one is lit and bold, the rest sit grey. In a ten-section day this is the
   thing that tells the room where they are and how much is left.

   Ten sections do not fit one readable row, so it wraps at five. Labels
   auto-size from 17pt down to the 14pt floor. If the names still do not
   fit, the build stops rather than shrink past the floor or clip - that is
   a signal to shorten a section name, not to squeeze the type.
   --------------------------------------------------------------------- */
const PER_ROW = 5;
export function sectionTracker(pptx, s, sections, current, y = 5.45) {
  const TRACK_X = 0.75, TRACK_W = 11.8, ROW_H = 0.78;
  const segW = TRACK_W / PER_ROW;
  const slot = segW - 0.22;

  let fs = 17;
  const longest = Math.max(...sections.map(n => n.length));
  while (fs > FLOOR && (longest * 0.52 * fs) / 72 > slot) fs -= 0.5;
  if ((longest * 0.52 * fs) / 72 > slot) {
    const n = sections.find(x => x.length === longest);
    throw new Error(`section name "${n}" will not fit the tracker at the ${FLOOR}pt floor - shorten it`);
  }

  sections.forEach((name, i) => {
    const row = Math.floor(i / PER_ROW), col = i % PER_ROW;
    const x = TRACK_X + col * segW, yy = y + row * ROW_H;
    const on = i === current, done = current != null && i < current;
    s.addShape(pptx.ShapeType.rect, {
      x, y: yy, w: slot, h: on ? 0.1 : 0.05,
      fill: { color: on ? C.blue : (done ? C.blueLt : C.rule) },
    });
    s.addText(name, {
      x, y: yy + 0.2, w: slot, h: 0.4, fontFace: F, fontSize: sz(fs, "tracker label"),
      bold: on, color: on ? C.ink : C.ink600, isTextBox: true, margin: 0,
    });
  });
}

/* ---------- SECTION TRANSITION - light, with the tracker ---------- */
export function section(pptx, deck, o) {
  const s = deck.addSlide();
  const sections = o.sections ?? o.segments;
  s.background = { color: C.white };
  if (o.logo) s.addImage({ data: o.logo, x: 0.75, y: 0.62, w: 2.0, h: 0.608 });
  s.addText(o.kicker.toUpperCase(), {
    x: 0.78, y: 2.32, w: 10, h: 0.4, fontFace: F, fontSize: sz(20, "section kicker"), bold: true,
    color: C.blue, charSpacing: 2.6, isTextBox: true, margin: 0,
  });
  [o.kicker, o.title, ...(o.chips || [])].forEach(t => words(t, "section slide"));
  // Day 2 (Oct 16): a long title steps down until it holds one line, never below 48pt
  let tfs = 76;
  while (tfs > 48 && o.title.length * 0.6 * tfs / 72 > 11.6) tfs -= 2;
  if (o.title.length * 0.6 * tfs / 72 > 11.6) throw new Error(`section title "${o.title}" will not fit one line at 48pt - shorten it`);
  s.addText(o.title, {
    x: 0.73, y: 2.72, w: 11.8, h: 1.3, fontFace: F, fontSize: sz(tfs, "section title"), bold: true, valign: "middle", wrap: false,
    color: C.ink, isTextBox: true, margin: 0,
  });
  if (o.chips?.length) s.addText(o.chips.join("   ·   "), {
    x: 0.76, y: 4.12, w: 11.8, h: 0.45, fontFace: F, fontSize: sz(22, "section chips"),
    color: C.ink600, isTextBox: true, margin: 0,
  });
  if (sections?.length) sectionTracker(pptx, s, sections, o.current);
  return s;
}

/* ---------- COVER - no tracker (Tommy, Oct 2026) ---------- */
export function cover(pptx, deck, { title, subtitle, meta, logo }) {
  const s = deck.addSlide();
  s.background = { color: C.white };
  if (logo) s.addImage({ data: logo, x: 0.75, y: 0.62, w: 2.4, h: 0.73 });
  s.addText(title, {
    x: 0.73, y: 2.6, w: 11.8, h: 1.9, fontFace: F, fontSize: sz(62, "cover title"), bold: true,
    color: C.ink, isTextBox: true, margin: 0,
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.78, y: 4.62, w: 1.5, h: 0.055, fill: { color: C.blue } });
  s.addText(subtitle, {
    x: 0.76, y: 4.88, w: 11.8, h: 0.5, fontFace: F, fontSize: sz(26, "cover subtitle"),
    color: C.blue, isTextBox: true, margin: 0,
  });
  if (meta?.length) s.addText(meta.join("   ·   "), {
    x: 0.76, y: 5.42, w: 11.8, h: 0.45, fontFace: F, fontSize: sz(20, "cover meta"),
    color: C.ink600, isTextBox: true, margin: 0,
  });
  return s;
}

/* ---------- AGENDA - the sections as numbered cards -------------------
   Numbers are the position in the section list, never typed, so the
   agenda cannot drift out of step with the tracker.
   --------------------------------------------------------------------- */
export function agenda(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow ?? "Day 1", title: o.title ?? "Agenda" });

  const items = o.items;
  const rows = Math.ceil(items.length / 2);
  const BOT = bodyBot(o.takeaway);
  const COLX = [0.5, 6.72], CW = 6.11;
  const rowH = (BOT - BODY.top) / rows;

  items.forEach((it, i) => {
    const col = i < rows ? 0 : 1, k = i - col * rows;
    const x = COLX[col], y = BODY.top + k * rowH;
    s.addShape(pptx.ShapeType.rect, { x, y, w: CW, h: rowH - 0.1, fill: { color: C.panel } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.055, h: rowH - 0.1, fill: { color: C.blue } });
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.26, y, w: 0.7, h: rowH - 0.1, fontFace: F, fontSize: sz(20, "agenda number"), bold: true,
      color: C.blueLt, valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(it.name, {
      x: x + 1.0, y, w: it.detail ? 2.5 : CW - 1.3, h: rowH - 0.1,
      fontFace: F, fontSize: sz(17, "agenda name"), bold: true,
      color: C.ink, valign: "middle", isTextBox: true, margin: 0,
    });
    if (it.detail) s.addText(it.detail, {
      x: x + 3.55, y, w: CW - 3.8, h: rowH - 0.1, fontFace: F, fontSize: sz(14, "agenda detail"),
      color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source);
  return s;
}

/* ---------- STATEMENT - one thing, said plainly ---------- */
export function statement(pptx, deck, o) {
  const s = deck.addSlide();
  s.background = { color: C.white };
  // block sits high so a seven-line statement still clears the attribution -
  // at the old y the last line landed on "Tommy Rayl"
  s.addShape(pptx.ShapeType.rect, { x: 0.75, y: 1.1, w: 0.065, h: 0.9, fill: { color: C.blue } });
  if (o.kicker) s.addText(o.kicker.toUpperCase(), {
    x: 1.08, y: 1.1, w: 10, h: 0.4, fontFace: F, fontSize: sz(20, "statement kicker"), bold: true,
    color: C.blue, charSpacing: 2.6, isTextBox: true, margin: 0,
  });
  // **phrase** marks the one line that carries the slide - drawn in Findlay
  // blue. One per statement, the way Tommy bolds one thing per email.
  const runs = o.text.split("**")
    .map((t, i) => ({ text: t, options: i % 2 ? { color: C.blue } : {} }))
    .filter(r => r.text);

  // Fit the font to the space, not the words to the font (Tommy, Oct 18
  // 2026): "it's a really big projector ... a couple more sentences in there,
  // and the font is smaller." Simulates word wrap at each size and takes the
  // largest that fits, never below the floor.
  const W = 11.4, H = o.attr ? 4.35 : 5.05;
  const plain = o.text.replace(/\*\*/g, "");
  const linesAt = fs => {
    const per = W / (0.54 * fs / 72);
    let lines = 1, cur = 0;
    for (const w of plain.split(/\s+/)) {
      const add = w.length + (cur ? 1 : 0);
      if (cur + add > per) { lines++; cur = w.length; } else cur += add;
    }
    return lines;
  };
  let fs = o.size ?? 34;
  while (fs > FLOOR && linesAt(fs) * fs * 1.22 * 1.2 / 72 > H) fs--;
  if (linesAt(fs) * fs * 1.22 * 1.2 / 72 > H) throw new Error(`statement will not fit at the ${FLOOR}pt floor - cut it`);

  s.addText(runs, {
    x: 1.05, y: 1.85, w: W, h: H, fontFace: F, fontSize: sz(fs, "statement"), bold: true,
    color: C.ink, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22, valign: "top",
  });
  if (o.attr) s.addText(o.attr, {
    x: 1.08, y: 6.45, w: 11.4, h: 0.4, fontFace: F, fontSize: sz(18, "statement attribution"),
    color: C.ink600, isTextBox: true, margin: 0,
  });
  // The talk track lives in the speaker notes: the slide says one thing,
  // the notes carry what Tommy says around it.
  const notes = [o.notes, o.source && `Source: ${o.source}`].filter(Boolean).join("\n\n");
  if (notes) s.addNotes(notes);
  return s;
}

/* ---------- GRID - a lot of short items, meant to be read at a glance ----
   The values slide. Chips, not bullets: the room is scanning for the three
   it wants, not reading a list top to bottom.
   --------------------------------------------------------------------- */
export function grid(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  let top = BODY.top;
  if (o.sub) {
    s.addText(o.sub, {
      x: 0.5, y: 1.26, w: 12.33, h: 0.34, fontFace: F, fontSize: sz(16, "grid sub"),
      color: C.ink600, isTextBox: true, margin: 0,
    });
    top = 1.72;
  }

  const cols = o.cols ?? 5;
  const rows = Math.ceil(o.items.length / cols);
  const GAP = 0.12, FULL = 12.33;
  const cw = (FULL - GAP * (cols - 1)) / cols;
  const rh = (bodyBot(o.takeaway) - top) / rows;

  o.items.forEach((t, i) => {
    const r = Math.floor(i / cols), c = i % cols;
    const x = 0.5 + c * (cw + GAP), y = top + r * rh;
    s.addShape(pptx.ShapeType.rect, { x, y, w: cw, h: rh - GAP, fill: { color: C.panel } });
    s.addText(t, {
      x: x + 0.14, y, w: cw - 0.28, h: rh - GAP, fontFace: F, fontSize: sz(15, "grid item"),
      color: C.ink, align: "center", valign: "middle", isTextBox: true, margin: 0,
    });
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source);
  return s;
}

/* ---------- PHRASE CLOUD - statements, scattered and sized by weight ----
   For the values workshop (Tommy, Oct 18 2026): "less of a grid and tile and
   more of an artistic word cloud ... phrases or statements that we could
   then distill. Single words are not the answer."

   Three tiers: 1 is largest and lands in the middle, 3 is the supporting
   behaviour around the edge. Placement is a deterministic elliptical
   spiral out from the centre - biggest first, so the heaviest statements
   hold the middle - and nothing overlaps. Everything stays horizontal:
   rotated phrases are unreadable from the back of a ballroom. A phrase that
   will not fit steps down a point at a time to the 14pt floor; past that
   the build stops rather than overlap or clip.

   Blue and ink only - this slide judges nothing, so no green or red.
   --------------------------------------------------------------------- */
export function phraseCloud(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  let top = BODY.top;
  if (o.sub) {
    s.addText(o.sub, {
      x: 0.5, y: 1.26, w: 12.33, h: 0.34, fontFace: F, fontSize: sz(16, "cloud sub"),
      color: C.ink600, isTextBox: true, margin: 0,
    });
    top = 1.72;
  }
  const bot = bodyBot(o.takeaway);
  panel(pptx, s, { x: 0.5, y: top, w: 12.33, h: bot - top, label: null });

  const X0 = 0.72, X1 = 12.61, Y0 = top + 0.16, Y1 = bot - 0.16;
  const cx = (X0 + X1) / 2, cy = (Y0 + Y1) / 2, ASPECT = (X1 - X0) / (Y1 - Y0);
  const TIER = {
    1: { fs: o.sizes?.[0] ?? 22, bold: true,  colors: [C.navy, C.blue] },
    2: { fs: o.sizes?.[1] ?? 17, bold: true,  colors: [C.blue, C.navy] },
    3: { fs: o.sizes?.[2] ?? 14, bold: false, colors: [C.ink600, C.ink] },
  };
  const MAXW = o.maxLine ?? 4.4, PAD = 0.10;
  // Inter measured ~0.49 em bold on the rendered bands; 0.53 leaves a margin
  // without wasting the room the first, more generous estimate did
  const width = (t, fs, bold) => t.length * (bold ? 0.53 : 0.50) * fs / 72;
  const shape = (t, fs, bold) => {
    let lines = [t];
    if (width(t, fs, bold) > MAXW) {          // break at the space nearest the middle
      const sp = [...t.matchAll(/ /g)].map(m => m.index);
      const at = sp.reduce((b, i) => Math.abs(i - t.length / 2) < Math.abs(b - t.length / 2) ? i : b, sp[0]);
      lines = [t.slice(0, at), t.slice(at + 1)];
    }
    return {
      lines,
      w: Math.max(...lines.map(l => width(l, fs, bold))) + 0.1,
      h: lines.length * fs * 1.18 / 72 + 0.04,
    };
  };
  // Every candidate centre on a fine grid. A spiral was tried first; it got
  // coarse toward the edges and missed gaps a phrase would have fitted.
  const STEP = 0.06, grid = [];
  for (let y = Y0; y <= Y1; y += STEP) for (let x = X0; x <= X1; x += STEP) grid.push({ x, y });

  // Each phrase gets a home and settles in the free spot nearest it. With
  // one shared centre the big statements stacked into a column down the
  // middle and the slide read as three lists. Tier 1 sits like the five on a
  // die; tiers 2 and 3 sit on rings around it, so the cloud reaches the
  // corners.
  const RX = (X1 - X0) / 2, RY = (Y1 - Y0) / 2;
  const DIE = [[0, 0], [-0.62, -0.55], [0.62, -0.55], [-0.6, 0.58], [0.62, 0.58]];
  const order = [...o.items].sort((a, b) => a.tier - b.tier || b.t.length - a.t.length);
  const ring = order.filter(it => it.tier > 1).length;
  let r = 0;
  order.forEach((it, i) => {
    if (it.tier === 1 && i < DIE.length) {
      it.home = { x: cx + DIE[i][0] * RX, y: cy + DIE[i][1] * RY };
    } else {
      const a = (r / ring) * 2 * Math.PI + 0.35, k = it.tier === 2 ? 0.48 : 0.86;
      it.home = { x: cx + Math.cos(a) * RX * k, y: cy + Math.sin(a) * RY * k };
      r++;
    }
    it.cand = grid.map(g => ({ ...g, d: ((g.x - it.home.x) / ASPECT) ** 2 + (g.y - it.home.y) ** 2 }))
                  .sort((p, q) => p.d - q.d);
    // a short statement earns a little more size - it is the one that lands
    it.boost = it.t.length <= 20 ? 1.2 : 1;
  });

  // Try the whole cloud at the largest scale first and shrink it as one,
  // so the tiers keep their proportions and the panel is filled rather than
  // the last few phrases being squeezed. Tier 3 never drops below the floor.
  const attempt = scale => {
    const placed = [];
    const free = b => b.x >= X0 && b.x + b.w <= X1 && b.y >= Y0 && b.y + b.h <= Y1 &&
      placed.every(p => b.x + b.w + PAD <= p.x || p.x + p.w + PAD <= b.x ||
                        b.y + b.h + PAD <= p.y || p.y + p.h + PAD <= b.y);
    for (const it of order) {
      const T = TIER[it.tier], fs = Math.max(FLOOR, Math.round(T.fs * scale * it.boost));
      const L = shape(it.t, fs, T.bold);
      const c = it.cand.find(c => free({ x: c.x - L.w / 2, y: c.y - L.h / 2, w: L.w, h: L.h }));
      if (!c) return null;
      placed.push({ x: c.x - L.w / 2, y: c.y - L.h / 2, w: L.w, h: L.h, L, fs, it });
    }
    return placed;
  };
  let layout = null;
  for (let sc = 1.5; sc >= 1.0 && !layout; sc -= 0.05) layout = attempt(sc);
  if (!layout) throw new Error(`the cloud will not fit at the ${FLOOR}pt floor - cut a phrase or shorten one`);

  const used = { 1: 0, 2: 0, 3: 0 };
  layout.forEach(b => {
    const T = TIER[b.it.tier];
    s.addText(b.L.lines.join("\n"), {
      x: b.x, y: b.y, w: b.w, h: b.h, fontFace: F, fontSize: sz(b.fs, "cloud phrase"), bold: T.bold,
      color: T.colors[used[b.it.tier]++ % T.colors.length],
      align: "center", valign: "middle", isTextBox: true, margin: 0, wrap: false,
    });
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source);
  return s;
}

/* ---------- TREND - trailing 12, no per-point labels beyond the fit ---------- */
export function trend(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const H = bodyH(o.takeaway);
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 8.3, h: H, label: o.panelLabel });

  // Day 2 (Oct 15): o.scale (default 1e6), o.axisFormat, o.majorUnit, o.pad for non-dollar series
  const SC = o.scale ?? 1e6;
  const vals = o.series.map(r => +(r[1] / SC).toFixed(2));
  const lo = o.min ?? Math.floor(Math.min(...vals) - 1);
  const hi = o.max ?? Math.ceil(Math.max(...vals) + 1);
  const pad = (hi - lo) * (o.pad ?? 0.12);   // headroom so point labels clear the line

  s.addChart(pptx.ChartType.line, [{ name: o.panelLabel, labels: o.series.map(r => r[0]), values: vals }], {
    x: 0.68, y: 1.98, w: 7.5, h: H - 0.83,
    chartColors: [C.blue], lineSize: 3, lineSmooth: false,
    lineDataSymbol: "circle", lineDataSymbolSize: 8,
    lineDataSymbolLineColor: C.blue, lineDataSymbolLineSize: 1,
    showValue: o.showValue ?? true, dataLabelPosition: "t",   // o.showValue false for long series (Oct 22 2026)
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

  panel(pptx, s, { x: 9.0, y: BODY.top, w: 3.83, h: H, label: "What it says" });
  let by = 2.02;
  const gap = Math.min(1.24, (bodyBot(o.takeaway) - 2.1) / o.bullets.length);
  o.bullets.forEach((b, i) => {
    s.addShape(pptx.ShapeType.ellipse, { x: 9.25, y: by + 0.1, w: 0.11, h: 0.11, fill: { color: i === 0 ? C.blue : C.ink } });
    s.addText(b, {
      x: 9.52, y: by - 0.04, w: 3.1, h: gap - 0.09, fontFace: F, fontSize: sz(15, "bullet"),
      color: i === 0 ? C.blue : C.ink, bold: i === 0, isTextBox: true, margin: 0, valign: "top",
    });
    by += gap;
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- YOY - group card over a month table ---------- */
export function yoy(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  // Day 2 (Oct 14): o.fmt = { v, d, p } for rate and count metrics; o.ytd = [a, b]
  // overrides the YTD figures - required for any ratio, where YTD is not the sum of months
  const FV = o.fmt?.v ?? M, FD = o.fmt?.d ?? money, FP = o.fmt?.p ?? pct;
  const a = o.ytd ? o.ytd[0] : o.months.reduce((t, r) => t + r[1], 0);
  const b = o.ytd ? o.ytd[1] : o.months.reduce((t, r) => t + r[2], 0);
  const v = b - a;
  if (!o.ytd && o.expect != null && Math.abs(v - o.expect) > 1) {
    throw new Error(`${o.title}: months sum to ${Math.round(v)}, expected ${o.expect}`);
  }
  const neu = dirOf(o.metric) === "neutral";   // Day 2: neutral metrics carry no judgment
  const up = dirOf(o.metric) === "up" ? v >= 0 : v < 0;
  const tone = neu ? "n" : (up ? "g" : "b");

  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: 1.66, label: null });
  const cells = [
    ["2025", FV(a), C.ink], ["2026", FV(b), C.ink],
    ["VARIANCE", FD(v), neu ? C.navy : (up ? C.goodInk : C.badInk)],
    ["VAR %", FP(v, a), neu ? C.navy : (up ? C.goodInk : C.badInk)],
  ];
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
  const rowH = Math.max(0.3, (bodyBot(o.takeaway) - 3.08) / (o.months.length + 1));
  s.addTable([hdr, ...body], {
    x: 0.5, y: 3.08, w: 12.33, colW: [3.41, 2.23, 2.23, 2.23, 2.23],
    rowH, valign: "middle", border: { type: "solid", pt: 0.5, color: C.rule },
  });

  takeaway(pptx, s, o.takeaway, tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- GROUP YOY - the whole organisation, before any store -------
   Opens every KPI (Tommy, Oct 18 2026): the year-to-date card - prior,
   current, variance, var % - over a monthly chart that spans the 24-month
   window, last year January to December against this year to date, month
   beside month. The room sees what the group did before it is cut into 34
   store stories.

   Bars start at zero, because a bar that does not is lying about size.
   Last year is grey and this year is Findlay blue: that is identity, not
   judgment, so it does not break the green-good / red-bad rule. Months not
   yet closed are left empty, never zero - a zero bar would read as a month
   that earned nothing.
   --------------------------------------------------------------------- */
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

/* How a metric's numbers read on the chart and the card.
     M     dollar totals, charted in millions ($3.2M)
     count units, ROs - charted in thousands (3.3K) so a pair of labels fits
           over a month without running together
     usd   per-unit dollars (PVR) - charted as whole dollars
     pct   shares, passed as percent - charted to one decimal, changes in pts */
const UNIT = {
  M:     { to: x => +(x / 1e6).toFixed(2), label: "0.0",               axis: '$0"M"',     fmt: Mmixed, dfmt: Dmixed },
  count: { to: x => x,                     label: '#,##0.0,"K"',       axis: '#,##0',     fmt: Int,    dfmt: Dint   },
  usd:   { to: x => Math.round(x),         label: '"$"#,##0',          axis: '"$"#,##0',  fmt: Mmixed, dfmt: Dmixed },
  usd2:  { to: x => +x.toFixed(2),         label: '"$"#,##0.00',       axis: '"$"#,##0',  fmt: Mmixed, dfmt: Dmixed },   // rates in dollars and cents (ELR), Oct 22 2026
  pct:   { to: x => +x.toFixed(1),         label: '0.0"%"',            axis: '0"%"',      fmt: Pct,    dfmt: Dpts   },
};
const niceStep = span => {
  const raw = span / 5, p = 10 ** Math.floor(Math.log10(raw));
  return [1, 2, 2.5, 5, 10].map(m => m * p).find(s => s >= raw);
};

export function groupYoy(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  // A total's year to date is the sum of its months. A RATE's is not - a
  // year-to-date PVR is year-to-date gross over year-to-date units, never
  // the sum or average of twelve monthly PVRs - so a rate passes its
  // year-to-date figures in as ytdPrior / ytdCurrent.
  const n = o.current.length;   // months closed this year
  const a = o.ytdPrior   ?? o.prior.slice(0, n).reduce((t, v) => t + v, 0);
  const b = o.ytdCurrent ?? o.current.reduce((t, v) => t + v, 0);
  if (o.expectPrior != null && Math.abs(a - o.expectPrior) > 1) {
    throw new Error(`${o.title}: prior months sum to ${Math.round(a)}, expected ${o.expectPrior}`);
  }
  if (o.expectCurrent != null && Math.abs(b - o.expectCurrent) > 1) {
    throw new Error(`${o.title}: current months sum to ${Math.round(b)}, expected ${o.expectCurrent}`);
  }
  const v = b - a;
  const neu = dirOf(o.metric) === "neutral";   // Day 2: neutral metrics carry no judgment
  const up = dirOf(o.metric) === "up" ? v >= 0 : v < 0;
  const U = UNIT[o.unit ?? "M"];
  const fmt = o.fmt ?? U.fmt, dfmt = o.dfmt ?? U.dfmt;

  // the card - no header line above the four figures (Tommy, Oct 18 2026):
  // "2025 YTD" and "2026 YTD" already say what the period is, and "GROUP ·
  // JANUARY-SEPTEMBER YTD" on top of them said it twice
  const CARD_H = 1.10;
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: CARD_H, label: null });
  [
    [o.priorName + " YTD",   fmt(a),    C.ink],
    [o.currentName + " YTD", fmt(b),    C.ink],
    ["VARIANCE",             dfmt(v),   neu ? C.navy : (up ? C.goodInk : C.badInk)],
    ["VAR %",                pct(v, a), neu ? C.navy : (up ? C.goodInk : C.badInk)],
  ].forEach(([l, val, col], i) => {
    const x = 0.75 + i * 3.05;
    s.addText(l, { x, y: BODY.top + 0.14, w: 2.9, h: 0.26, fontFace: F, fontSize: sz(14, "card sublabel"), color: C.ink600, charSpacing: 0.9, isTextBox: true, margin: 0 });
    s.addText(val, { x, y: BODY.top + 0.38, w: 2.95, h: 0.6, fontFace: F, fontSize: sz(30, "card value"), bold: true, color: col, isTextBox: true, margin: 0 });
  });

  // the chart
  const cy = BODY.top + CARD_H + 0.12, ch = bodyBot(o.takeaway) - cy;
  panel(pptx, s, { x: 0.5, y: cy, w: 12.33, h: ch, label: o.chartLabel });
  // Day 2: when neither year runs past a month (a Jan-Sep prior year), the axis stops there
  const NM   = Math.max(o.prior.length, o.current.length);
  const pad  = arr => [...arr, ...Array(NM - arr.length).fill(null)];
  // o.benchLines = [value, ...] in chart units, o.benchLabel (Noah, Oct 24): a dashed gray benchmark line (two for a range)
  const BL = o.benchLines ?? [];
  const all  = [...o.prior, ...o.current, ...BL].map(U.to);
  // Totals are bars - each month is an amount. A rate is a line: bars would
  // imply the months add up, and they do not.
  const line = (o.chart ?? "bar") === "line";
  const dMin = Math.min(...all), dMax = Math.max(...all);
  let step, lo, hi;
  if (!line) {
    // Bars always start at zero - a bar that does not is lying about size.
    const top0 = Math.max(0, dMax) * 1.15, bot0 = Math.min(0, dMin);
    step = o.majorUnit ?? niceStep(top0 - bot0);
    hi = Math.ceil(top0 / step) * step;
    lo = Math.floor(bot0 / step) * step;
  } else {
    // A line encodes position, not length, so it does NOT need a zero
    // baseline - and forcing one crushed F&I's two years ($2,060 vs $2,195)
    // into a single band where every label collided. The axis fits the
    // data, with room above and below for a row of labels each side. If
    // that room dips under zero for data that never does, the negative tick
    // is blanked by the number format.
    const span = Math.max(dMax - dMin, Math.abs(dMax) * 0.05, 1);
    step = o.majorUnit ?? niceStep(span * 1.7);
    lo = Math.floor((dMin - span * 0.35) / step) * step;
    hi = Math.ceil((dMax + span * 0.35) / step) * step;
    // o.tightBench (Noah, Oct 26): the benchmark line stretched the opcode axes and flattened the gain.
    // Pad from the data's own span, and reach only just past the benchmark line.
    if (o.tightBench && BL.length) {
      const D = [...o.prior, ...o.current].map(U.to);
      const sD = Math.max(Math.max(...D) - Math.min(...D), Math.abs(Math.max(...D)) * 0.05, 0.5);
      const bl = BL.map(U.to);
      const rLo = Math.min(Math.min(...D) - sD * 0.3, Math.min(...bl) - sD * 0.1);
      const rHi = Math.max(Math.max(...D) + sD * 0.3, Math.max(...bl) + sD * 0.1);
      step = o.majorUnit ?? niceStep(rHi - rLo);
      lo = Math.floor(rLo / step) * step;
      hi = Math.ceil(rHi / step) * step;
    }
  }
  const baseFmt = o.axisFormat ?? U.axis;
  const axisFmt = baseFmt + (lo < 0 && dMin >= 0 ? (baseFmt.includes("$") ? ';;"$"0' : ";;0") : "");
  const common = {
    showValue: true,
    dataLabelFontFace: "+mn-lt", dataLabelFontSize: sz(14, "data label"), dataLabelColor: C.ink,
    dataLabelFormatCode: o.labelFormat ?? U.label,
    showTitle: false,
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(14, "cat axis"), catAxisLabelColor: C.ink,
    valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: sz(14, "val axis"), valAxisLabelColor: C.ink600,
    valAxisLabelFormatCode: axisFmt,
    valAxisMinVal: lo, valAxisMaxVal: hi, valAxisMajorUnit: step,
    valGridLine: { color: C.rule, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: true, valAxisLineShow: false,
    // the month axis sits at the bottom of the plot whatever the floor is
    // (pptxgenjs writes this option onto the category axis's <c:crosses>)
    valAxisCrossesAt: "min", catAxisLabelPos: "low",
  };
  const prior   = { name: o.priorName,   labels: MONTHS.slice(0, NM), values: pad(o.prior.map(U.to)) };
  const current = { name: o.currentName, labels: MONTHS.slice(0, NM), values: pad(o.current.map(U.to)) };

  if (!line) {
    s.addChart(pptx.ChartType.bar, [prior, current], {
      ...common, x: 0.68, y: cy + 0.62, w: 11.97, h: ch - 0.72,
      // Each label is wider than its bar, so the gap inside a month's pair
      // and the gap between months must be EQUAL - that gives every label
      // half a month of room either side. Unequal gaps ran labels together
      // either inside the pair ("14.112.6") or across months ("2.7K2.8K").
      barDir: "col", barGrouping: "clustered", barGapWidthPct: 35, barOverlapPct: -35,
      chartColors: [C.ink300, C.blue], dataLabelPosition: "outEnd",
      showLegend: true, legendPos: "t", legendFontFace: "+mn-lt", legendFontSize: sz(14, "legend"), legendColor: C.ink,
    });
  } else {
    // Two lines on one chart can only take ONE label position, so where the
    // lines run close (February) the labels landed on each other. Instead,
    // two charts on identical geometry and axes: last year labelled above
    // its line, this year below. The masked chart's axis text is drawn in
    // the panel colour, not hidden, so both plot areas lay out to the same
    // pixel - and it goes UNDERNEATH, because white text laid over the real
    // axis text washed it out. The legend is drawn by hand because neither
    // chart carries one.
    const g = { x: 0.68, y: cy + 0.98, w: 11.97, h: ch - 1.08 };
    const ln = { lineSize: 3, lineSmooth: false, lineDataSymbol: "circle", lineDataSymbolSize: 8, showLegend: false };
    const clear = { plotArea: { fill: { color: C.panel, transparency: 100 } },
                    chartArea: { fill: { color: C.panel, transparency: 100 } } };
    // Whichever line runs HIGHER over the months both years have carries its
    // labels above; the lower one carries them below. A fixed "last year
    // above" rule held for new and used PVR, where this year runs lower, and
    // put every label on the wrong line in F&I, where this year runs higher.
    //
    // That rule still collided where the lines CROSS (wholesale Jan-Apr,
    // writedowns in May, Oct 18 2026), so the side is now chosen month by
    // month: in each month the higher point is labelled above, the lower
    // below. One series can only take one label position, so the labels are
    // their own invisible series - no line, no marker - one per side, laid
    // over the two drawn lines. A month only one year has keeps the
    // whole-year rule.
    const meanOf = a => a.reduce((t, v) => t + v, 0) / a.length;
    const curHigh = meanOf(current.values.slice(0, n)) >= meanOf(prior.values.slice(0, n));
    // A month only one year has (Oct-Dec 2025): label above a peak, below a dip, judged against its neighbours, so
    // the label clears the line running in and out of it (Noah, Oct 24: the Oct and Nov 2025 labels sat on the line)
    const side = (vals, i) => {
      const nb = [vals[i - 1], vals[i + 1]].filter(v => v != null);
      return nb.length ? vals[i] >= nb.reduce((t, v) => t + v, 0) / nb.length : true;
    };
    const curTop = current.values.map((c, i) => {
      const p = prior.values[i];
      if (c != null && p != null) return c >= p;
      if (p != null) return !side(prior.values, i);
      if (c != null) return side(current.values, i);
      return curHigh;
    });
    const only = (ser, keep) => ({ ...ser, values: ser.values.map((v, i) => keep(i) ? v : null) });
    const labelSets = [
      [only(current, i =>  curTop[i]), C.blue,   "t"],
      [only(current, i => !curTop[i]), C.blue,   "b"],
      [only(prior,   i => !curTop[i]), C.ink300, "t"],
      [only(prior,   i =>  curTop[i]), C.ink300, "b"],
    ].filter(([ser]) => ser.values.some(v => v != null));
    const masked = { catAxisLabelColor: C.panel, valAxisLabelColor: C.panel, catAxisLineShow: false,
                     valGridLine: { style: "none" } };
    s.addChart(pptx.ChartType.line, [current], {
      ...common, ...ln, ...g, ...masked, showValue: false, chartColors: [C.blue],
      valGridLine: { color: C.rule, size: 0.75 },   // gridlines drawn here, first, so a benchmark line can sit on top
    });
    labelSets.forEach(([ser, col, pos]) => {
      s.addChart(pptx.ChartType.line, [ser], {
        ...common, ...ln, ...g, ...masked, ...clear,
        lineSize: 0, lineDataSymbol: "none", chartColors: [col], dataLabelPosition: pos,
      });
    });
    BL.forEach((b, i) => s.addChart(pptx.ChartType.line, [{ name: "bench" + i, labels: MONTHS.slice(0, NM), values: Array(NM).fill(U.to(b)) }], {
      ...common, ...ln, ...g, ...masked, ...clear, showValue: false, lineSize: 2, lineDataSymbol: "none", lineDash: "dash",
      chartColors: [C.ink600],
    }));   // one chart per line: a second series in one chart lost its dash
    // the real axes go on last, so no masked axis text sits over them
    s.addChart(pptx.ChartType.line, [prior], {
      ...common, ...ln, ...g, ...clear, showValue: false, chartColors: [C.ink300], valGridLine: { style: "none" },
    });
    const ly = cy + 0.66, lx = 6.665 - 1.25 - (BL.length ? 1.3 : 0);
    [[o.priorName, C.ink300], [o.currentName, C.blue]].forEach(([t, col], i) => {
      const x = lx + i * 1.35;
      s.addShape(pptx.ShapeType.rect, { x, y: ly + 0.135, w: 0.34, h: 0.05, fill: { color: col } });
      s.addShape(pptx.ShapeType.ellipse, { x: x + 0.11, y: ly + 0.1, w: 0.12, h: 0.12, fill: { color: col } });
      s.addText(t, { x: x + 0.42, y: ly, w: 0.85, h: 0.32, fontFace: F, fontSize: sz(14, "legend"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    });
    if (BL.length) {   // dashed swatch + label for the benchmark
      const x = lx + 2 * 1.35;
      [0, 0.13, 0.26].forEach(d => s.addShape(pptx.ShapeType.rect, { x: x + d, y: ly + 0.135, w: 0.08, h: 0.04, fill: { color: C.ink600 } }));
      s.addText(o.benchLabel ?? "Benchmark", { x: x + 0.42, y: ly, w: 2.6, h: 0.32, fontFace: F, fontSize: sz(14, "legend"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    }
  }

  takeaway(pptx, s, o.takeaway, o.tone ?? (neu ? "n" : (up ? "g" : "b")));
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- RANKER - every store on one page ---------- */
export function ranker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  // a store with no prior-year base cannot have a variance: dash, last, and
  // out of the group total, the bar scale and every count
  // Day 2 (Oct 15): o.deltas = [[name, change]] for rates, with the group's own change in
  // o.groupD (alias of groupDelta) and its format in o.fmtD (alias of dfmt). A "neutral"
  // metric sorts largest first and draws every bar blue - no green/red judgment.
  const dir    = dirOf(o.metric), neutral = dir === "neutral";
  const noBase = o.deltas ? [] : o.rows.filter(r => r[1] === 0).map(r => r[0]);
  const based  = o.deltas ? o.deltas.map(([n, d]) => ({ n, d }))
                          : o.rows.filter(r => r[1] !== 0).map(([n, x, y]) => ({ n, d: y - x }));
  const good   = dir === "up" || neutral;
  if (o.groupD != null && o.groupDelta == null) o.groupDelta = o.groupD;
  if (o.fmtD && !o.dfmt) o.dfmt = o.fmtD;
  based.sort((p, q) => good ? q.d - p.d : p.d - q.d);
  const all    = [...based, ...noBase.map(n => ({ n, d: null, noBase: true }))];
  // For a total the group change is the sum of the store changes. For a
  // RATE it is not - group PVR moves with the mix of who sold the cars - so
  // a rate passes its own group change in as groupDelta.
  const summed = based.reduce((t, r) => t + r.d, 0);
  if (o.groupDelta == null && o.expect != null && Math.abs(summed - o.expect) > 1) {
    throw new Error(`${o.title}: store rows sum to ${Math.round(summed)}, expected ${o.expect}`);
  }
  const groupD = o.groupDelta ?? summed;

  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });

  // the group figure as its own strip - its bar is the sum, so it never joins the scale
  const gFav = good ? groupD >= 0 : groupD < 0;
  measureStrip(pptx, s, o.measure ?? o.title, o.groupLabel ?? "**YEAR-OVER-YEAR CHANGE**",
               o.stripValue ?? (o.dfmt ?? Dmixed)(groupD), neutral ? C.navy : (gFav ? C.goodInk : C.badInk));

  const N = all.length, perCol = Math.ceil(N / 2);
  // Oct 22 2026 (Noah, option A on the gross change slides): o.extra = [{ label, w, vals: {NAME: text} }]
  // coloured columns after the value (e.g. the dollar change beside a % ranking), o.valLabel heads the value
  // column, o.stripValue overrides the strip figure.
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);
  const TOP = 2.02 + (EX.length ? 0.3 : 0), BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
  const maxAbs = Math.max(...based.map(r => Math.abs(r.d)));
  const dfmt = o.dfmt ?? Dmixed;
  const COLX = [0.72, 6.79];
  const FS = 12;   // traded down from 13 to carry the rank numbers
  const RANKW = 0.34, NAMEW = 2.10, VALW = EX.length ? 0.85 : 1.35, BARW = 1.70 + (1.35 - VALW) - EXW;
  if (EX.length) [0, 1].forEach(c => {
    let hx = COLX[c] + RANKW + 0.05 + NAMEW + 0.08 + BARW + 0.06;
    [{ label: o.valLabel ?? "", w: VALW }, ...EX].forEach(e => {
      s.addText(e.label, { x: hx, y: 2.0, w: e.w, h: 0.28, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      hx += e.w;
    });
  });

  all.forEach((r, i) => {
    const col = i < perCol ? 0 : 1, k = i - col * perCol;
    const x0 = COLX[col], y = TOP + k * rowH;
    const isUp = r.d >= 0, favourable = good ? isUp : !isUp;
    // o.shade = { NAME: "g" | "r" } (Noah, Oct 24): a pale band behind the whole row, e.g. a store that went from a loss to a profit
    const sh = o.shade?.[r.n];
    if (sh) s.addShape(pptx.ShapeType.rect, { x: x0 - 0.05, y: y + 0.01, w: 5.62, h: rowH - 0.02, fill: { color: sh === "g" ? C.goodBg : C.badBg } });

    // A store with no prior year has no variance, so it has no rank either.
    s.addText(r.noBase ? "—" : String(i + 1), {
      x: x0, y, w: RANKW, h: rowH, fontFace: F, fontSize: FS,
      color: C.ink300, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(r.n, {
      x: x0 + RANKW + 0.05, y, w: NAMEW, h: rowH, fontFace: F, fontSize: FS,
      color: r.noBase ? C.ink600 : C.ink, valign: "middle", isTextBox: true, margin: 0,
    });
    const bx = x0 + RANKW + 0.05 + NAMEW + 0.08, mid = bx + BARW / 2;
    if (!r.noBase) {
      const w = Math.abs(r.d) / maxAbs * (BARW / 2), bh = rowH * 0.46;
      s.addShape(pptx.ShapeType.rect, {
        x: isUp ? mid : mid - w, y: y + (rowH - bh) / 2, w: Math.max(w, 0.014), h: bh,
        fill: { color: neutral ? C.blue : (favourable ? C.good : C.bad) },
      });
    }
    s.addText(r.noBase ? "—" : dfmt(r.d), {
      x: bx + BARW + 0.06, y, w: VALW, h: rowH, fontFace: F, fontSize: FS, bold: true,   // the ranked column bold (Noah, Oct 24)
      // grey only when the PRINTED change carries no sign - a dollar test
      // (|d| < 0.5) greyed "−0.3 pts" and coloured "−0.5 pts" beside it
      color: r.noBase ? C.ink300 : !/^[+−]/.test(dfmt(r.d)) ? C.ink600 : neutral ? C.ink : (favourable ? C.goodInk : C.badInk),
      align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
    let ex = bx + BARW + 0.06 + VALW;
    EX.forEach(e => {
      s.addText(e.vals[r.n] ?? "", { x: ex, y, w: e.w, h: rowH, fontFace: F, fontSize: FS,
        color: neutral ? C.ink : (favourable ? C.goodInk : C.badInk), align: "right", valign: "middle", isTextBox: true, margin: 0 });
      // e.marks = { NAME: "g" | "r" } (Noah, Oct 24): a thin vertical bar just right of the value, e.g. the top and bottom three by %
      const mk = e.marks?.[r.n];
      if (mk) s.addShape(pptx.ShapeType.rect, { x: ex + e.w + 0.06, y: y + 0.05, w: 0.06, h: rowH - 0.1, fill: { color: mk === "g" ? C.good : C.bad } });
      ex += e.w;
    });
  });
  [0, 1].forEach(c => {
    const mid = COLX[c] + RANKW + 0.05 + NAMEW + 0.08 + BARW / 2;
    s.addShape(pptx.ShapeType.rect, { x: mid - 0.006, y: TOP, w: 0.012, h: BOT - TOP, fill: { color: C.ink300 } });
  });
  s.addShape(pptx.ShapeType.rect, { x: 6.655, y: TOP, w: 0.012, h: BOT - TOP, fill: { color: C.rule } });

  // The band takes the tone of what it SAYS when the spec sets one: a line
  // naming a decline on a slide where the group is up must not read green.
  takeaway(pptx, s, o.takeaway, o.tone ?? (neutral ? "n" : (groupD >= 0 === good ? "g" : "b")));
  foot(pptx, s, o.source + (noBase.length ? `  ${noBase.join(", ")} — no prior-year base, excluded from the group.` : ""), o.page);
  if (o.notes) s.addNotes(words(o.notes, "notes"));   // ADDED 2026-10-24
  return s;
}

/* ---------- TOTAL RANKER ----------------------------------------------
   Every store ranked best to worst on a dollar TOTAL - net profit, gross.
   One question: who is on top. Year-over-year movement is its own slide,
   the variance ranker, which follows this one.

   Tried and rejected (Tommy, Oct 18 2026): a change chip on every row of
   this slide, to show standing and direction at once. "Too busy" - it gave
   him nothing clean to talk to for either question. He presents to the
   slide, so one idea per slide, and slide count is not a cost. Do not put
   the chip back.

   GROUP rides in its own strip, not in the ranking: a group TOTAL dwarfs
   every store and would flatten the bar scale. (Compare levelRanker, where
   a group AVERAGE does belong inside the list.)
   --------------------------------------------------------------------- */
export function totalRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  const good = dirOf(o.metric) === "up";
  const fmt  = o.fmt ?? Mmixed;
  const rows = o.rows.map(([n, b]) => ({ n, b }));
  rows.sort((p, q) => good ? q.b - p.b : p.b - q.b);

  const gB = rows.reduce((t, r) => t + r.b, 0);
  if (o.expect != null && Math.abs(gB - o.expect) > 1) {
    throw new Error(`${o.title}: store rows sum to ${Math.round(gB)}, expected ${o.expect}`);
  }

  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });

  measureStrip(pptx, s, o.measure ?? o.title, o.groupLabel ?? "YEAR TO DATE", fmt(gB), C.navy);

  const perCol = Math.ceil(rows.length / 2);
  const TOP = 2.02, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
  const maxAbs = Math.max(...rows.map(r => Math.abs(r.b)));
  const COLX = [0.72, 6.79];
  // 12pt, not 13: Tommy traded the size for rank numbers (Oct 2026). NAMEW
  // still holds "CHEVY GMC BULLHEAD" on one line - it wrapped at 2.05/13pt
  // and collided with the row beneath it. The width the change chip used
  // went back to the bar.
  const FS = 12;
  const RANKW = 0.34, NAMEW = 2.10, BARW = 2.00, VALW = 1.20;

  rows.forEach((r, i) => {
    const col = i < perCol ? 0 : 1, k = i - col * perCol;
    const x0 = COLX[col], y = TOP + k * rowH;
    const fav = good ? r.b >= 0 : r.b < 0;   // green above the line, red below it

    s.addText(String(r.rk ?? i + 1), {
      x: x0, y, w: RANKW, h: rowH, fontFace: F, fontSize: FS,
      color: C.ink300, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });

    let x = x0 + RANKW + 0.05;
    s.addText(r.n, {
      x, y, w: NAMEW, h: rowH, fontFace: F, fontSize: FS,
      color: C.ink, valign: "middle", isTextBox: true, margin: 0,
    });

    // Magnitude bar. Day 2 (Noah, Oct 22): blue - a store total is shown, not compared, and green/red
    // is kept for slides that compare against something. A total below zero still reads red.
    x += NAMEW + 0.05;
    const bh = rowH * 0.46, w = Math.abs(r.b) / maxAbs * BARW;
    s.addShape(pptx.ShapeType.rect, {
      x, y: y + (rowH - bh) / 2, w: Math.max(w, 0.02), h: bh,
      fill: { color: fav ? C.blue : C.bad },
    });

    x += BARW + 0.05;
    s.addText(fmt(r.b), {
      x, y, w: VALW, h: rowH, fontFace: F, fontSize: FS,
      color: fav ? C.ink : C.badInk, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
  });

  // the rule between the two stacks, so the eye does not read across them
  s.addShape(pptx.ShapeType.rect, { x: 6.655, y: TOP, w: 0.012, h: BOT - TOP, fill: { color: C.rule } });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- MATRIX - four boxes, names in them -------------------------
   The improvement slide. A scatter would be the obvious choice and the
   wrong one: net profit is so skewed that one store sits at $19.8M and
   thirty cluster under $5M, so the dots bunch into the left quarter and
   the labels collide. Naming the stores in four boxes reads instantly from
   the back of a room and says the same thing.
   --------------------------------------------------------------------- */
export function matrix(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  const BOT = bodyBot(o.takeaway);
  const GAP = 0.13, W = (12.33 - GAP) / 2, H = (BOT - BODY.top - GAP) / 2;

  o.boxes.forEach((b, i) => {
    const x = 0.5 + (i % 2) * (W + GAP), y = BODY.top + Math.floor(i / 2) * (H + GAP);
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: H, fill: { color: C.panel } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: W, h: 0.055, fill: { color: b.tone === "g" ? C.good : b.tone === "b" ? C.bad : C.blueLt } });

    s.addText(b.title, {
      x: x + 0.25, y: y + 0.18, w: W - 1.2, h: 0.36, fontFace: F, fontSize: sz(17, "matrix title"), bold: true,
      color: b.tone === "g" ? C.goodInk : b.tone === "b" ? C.badInk : C.navy, isTextBox: true, margin: 0,
    });
    s.addText(String(b.names.length), {
      x: x + W - 0.95, y: y + 0.14, w: 0.7, h: 0.4, fontFace: F, fontSize: sz(20, "matrix count"), bold: true,
      color: C.ink300, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
    if (b.sub) s.addText(b.sub, {
      x: x + 0.25, y: y + 0.54, w: W - 0.5, h: 0.3, fontFace: F, fontSize: sz(14, "matrix sub"),
      color: C.ink600, isTextBox: true, margin: 0,
    });

    const top = y + (b.sub ? 0.92 : 0.66);
    const cols = b.names.length > 7 ? 2 : 1;
    const per  = Math.ceil(b.names.length / cols);
    const cw   = (W - 0.5) / cols;
    const lh   = Math.min(0.28, (y + H - 0.16 - top) / per);
    const fs   = lh < 0.21 ? 12 : 13;   // a full box drops a point rather than clip
    b.names.forEach((n, j) => {
      const c = Math.floor(j / per), k = j - c * per;
      s.addText(n, {
        x: x + 0.25 + c * cw, y: top + k * lh, w: cw - 0.1, h: lh,
        fontFace: F, fontSize: fs, color: C.ink, valign: "middle", isTextBox: true, margin: 0,
      });
    });
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- MATH CHAIN ---------- */
export function chain(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const sum = o.steps.reduce((t, st) => t + st.v, 0);
  if (Math.abs(sum - o.total) > 2) throw new Error(`${o.title}: chain sums to ${Math.round(sum)}, total says ${o.total}`);
  const H = bodyH(o.takeaway), BOT = bodyBot(o.takeaway);

  panel(pptx, s, { x: 0.5, y: BODY.top, w: 6.0, h: H, label: o.panelLabel });
  let cy = 1.86;
  o.steps.forEach((st, i) => {
    if (i > 0) {
      s.addShape(pptx.ShapeType.ellipse, { x: 1.08, y: cy, w: 0.32, h: 0.32, fill: { color: C.blue } });
      s.addText(st.v >= 0 ? "+" : "−", { x: 1.08, y: cy, w: 0.32, h: 0.32, fontFace: F, fontSize: sz(18, "operator"), bold: true, color: C.white, align: "center", valign: "middle", isTextBox: true, margin: 0 });
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
  if (cy + 0.78 > BOT) throw new Error("math chain overflows its panel - too many steps for one slide");
  s.addText(M(o.total), { x: 0.98, y: cy + 0.03, w: 3.0, h: 0.56, fontFace: F, fontSize: sz(30, "chain total"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
  s.addText(o.totalLabel, { x: 4.0, y: cy + 0.03, w: 2.05, h: 0.72, fontFace: F, fontSize: sz(14, "chain label"), color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });

  // movers rail
  panel(pptx, s, { x: 6.7, y: BODY.top, w: 6.13, h: H, label: o.moversLabel });
  const maxM = Math.max(...o.movers.map(m => Math.abs(m[1])));
  const mGap = Math.min(0.47, (BOT - 2.06) / o.movers.length);
  let my = 2.0;
  o.movers.forEach(([n, d]) => {
    const up = d > 0;
    s.addText(n, { x: 6.95, y: my, w: 2.35, h: mGap - 0.05, fontFace: F, fontSize: sz(14, "mover name"), color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    const bx = 9.4, bw = 1.6, w = Math.abs(d) / maxM * bw;
    s.addShape(pptx.ShapeType.rect, { x: up ? bx : bx - w, y: my + (mGap - 0.05) / 2 - 0.1, w: Math.max(w, 0.02), h: 0.2, fill: { color: up ? C.good : C.bad } });
    s.addText(money(d), { x: 11.1, y: my, w: 1.5, h: mGap - 0.05, fontFace: F, fontSize: sz(14, "mover value"), bold: true, color: up ? C.goodInk : C.badInk, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    my += mGap;
  });
  s.addShape(pptx.ShapeType.rect, { x: 9.394, y: 2.0, w: 0.012, h: my - 2.06, fill: { color: C.ink300 } });

  takeaway(pptx, s, o.takeaway, o.tone ?? "g");
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- LEVEL RANKER ----------------------------------------------
   For averages, rates and percentages - anything where the question is
   "who is above the group and who is below", not "who moved the most".

   The group figure sits in the strip at the top, like every other ranker,
   and a navy line runs through both columns at the group's value so above
   and below reads at a glance. (It used to sit inside the list as a row at
   its own rank; Tommy, Oct 18 2026, could not find it there - "why is there
   no group number" - so it moved to the strip.) Stores read green above the
   group line and red below it, direction-aware from the registry.
   --------------------------------------------------------------------- */
export function levelRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  const dir = dirOf(o.metric), neutral = dir === "neutral", up = dir !== "down";
  const rows = o.rows.map(([n, v]) => ({ n, v }));
  rows.sort((a, b) => up ? b.v - a.v : a.v - b.v);
  const G = o.group;
  /* Day 2 (Oct 14): o.bench scores stores against a published benchmark instead of
     the group - the navy line moves to the benchmark and the strip names it.
     o.floor starts the bars above zero so close rates are readable (Oct 15).
     o.extra = [{ label, w, vals: {NAME: text} }] plain-ink reference columns after
     the value, with o.valLabel heading the value column (Oct 15).
     o.refOf(name) scores each store against its own reference (Oct 21, Noah: one
     ranking for luxury and non-luxury instead of two side-by-side lists) - each row
     gets a short navy tick at its own reference instead of the one line through the
     column. o.mark = Set of names drawn on a tinted row with the name in bold navy;
     o.legend = one line above the rows saying what the tint and tick mean. */
  const REF = o.bench ?? G, mark = o.mark ?? new Set();
  /* o.groups = [{ label, names: Set, ref }] (review, Oct 23): one list per group, each sorted and ranked on its own,
     under a header row, in the same two columns; each store is scored against its group's ref. */
  const grp = o.groups ? n => o.groups.find(g => g.names.has(n)) : null;
  const refOf = o.refOf ?? (grp ? n => grp(n).ref : () => REF);
  const all = !o.groups ? rows : o.groups.flatMap((g, gi) => [{ hdr: g.label, gi }, ...rows.filter(r => g.names.has(r.n)).map((r, i) => ({ ...r, rk: i + 1, gi }))]);
  /* o.segBreak (Noah, Oct 24): when a group runs over into the right column, that column opens with the group's
     label again ("... (CONTINUED)"), and the average line stops at each group instead of running through the headers */
  let perColFix = null;
  if (o.groups && o.segBreak) {
    const pc = Math.ceil((all.length + 1) / 2), at = all[pc];
    if (at && !at.hdr) all.splice(pc, 0, { hdr: o.groups[at.gi].label + "  (CONTINUED)", gi: at.gi });
    perColFix = pc;
  }
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);

  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const lab = (o.groupLabel ?? "SEPTEMBER 2026 YTD") + (o.bench != null ? `   ·   **BENCHMARK ${o.fmt(o.bench)}**` : "");
  measureStrip(pptx, s, o.measure ?? o.title, lab, o.fmt(G), C.navy);

  const N = all.length, perCol = perColFix ?? Math.ceil(N / 2);
  const HEAD = EX.length ? 0.3 : o.legend ? 0.36 : 0;
  if (o.legend) s.addText(words(o.legend, "ranker legend"), { x: 0.94, y: 2.02, w: 11.6, h: 0.3, fontFace: F, fontSize: sz(14, "ranker legend"),
    color: C.ink600, valign: "middle", isTextBox: true, margin: 0 });
  const TOP = 2.02 + HEAD, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
  // The scale runs from the lowest value to the highest, through zero, so a
  // negative store draws left of a zero line instead of vanishing into a sliver.
  const lo = o.floor ?? Math.min(0, ...rows.map(r => r.v)), hi = Math.max(0, REF * 1.02, ...rows.map(r => r.v), ...rows.map(r => refOf(r.n) * 1.02));
  const span = hi - lo || 1;
  const COLX = [0.72, 6.79];
  const FS = 12;   // traded down from 13 to carry the rank numbers
  const RANKW = 0.34, NAMEW = 2.10, VALW = EX.length ? 0.74 : 1.25;
  const BARW = EX.length ? 5.62 - RANKW - NAMEW - VALW - EXW - 0.19 : 1.85;
  const X = (bx, v) => bx + (Math.max(v, lo) - lo) / span * BARW;
  /* o.alignRef (Oct 23, Noah): each bar is drawn against its own reference, so every store's segment average lands
     on the same line and a store at its average reaches it; values printed stay the real ones. */
  const hiR = Math.max(1.02, ...rows.map(r => r.v / refOf(r.n) * 1.02));
  const XR = (bx, v, n) => o.alignRef ? bx + Math.max(v / refOf(n), 0) / hiR * BARW : X(bx, v);

  all.forEach((r, i) => {
    const col = i < perCol ? 0 : 1, k = i - col * perCol;
    const x0 = COLX[col], y = TOP + k * rowH;
    if (r.hdr) {
      s.addText(r.hdr, { x: x0, y, w: 5.8, h: rowH, fontFace: F, fontSize: FS + 1, bold: true, color: C.navy, charSpacing: 1.0, valign: "bottom", isTextBox: true, margin: 0 });
      s.addShape(pptx.ShapeType.rect, { x: x0, y: y + rowH - 0.02, w: 5.8, h: 0.02, fill: { color: C.navy } });
      return;
    }
    const ref = refOf(r.n), better = up ? r.v >= ref : r.v <= ref, mk = mark.has(r.n);
    const fill = neutral ? C.blue : (better ? C.good : C.bad);
    const txt  = neutral ? C.ink : (better ? C.goodInk : C.badInk);
    if (mk) s.addShape(pptx.ShapeType.rect, { x: x0 - 0.06, y, w: 5.86, h: rowH, fill: { color: C.blueTint } });

    s.addText(String(r.rk ?? i + 1), {
      x: x0, y, w: RANKW, h: rowH, fontFace: F, fontSize: FS,
      color: C.ink300, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(r.n, {
      x: x0 + RANKW + 0.05, y, w: NAMEW, h: rowH, fontFace: F, fontSize: FS,
      color: mk ? C.navy : C.ink, bold: mk, valign: "middle", isTextBox: true, margin: 0,
    });
    const bx = x0 + RANKW + 0.05 + NAMEW + 0.08;
    const zx = o.alignRef ? bx : X(bx, Math.max(0, lo)), vx = XR(bx, r.v, r.n), bh = rowH * 0.5;
    s.addShape(pptx.ShapeType.rect, {
      x: Math.min(zx, vx), y: y + (rowH - bh) / 2, w: Math.max(Math.abs(vx - zx), 0.02), h: bh, fill: { color: fill },
    });
    if ((o.refOf || o.groups) && !o.hideRef && !o.refLine) s.addShape(pptx.ShapeType.rect, { x: XR(bx, ref, r.n) - 0.012, y: y + rowH * 0.1, w: 0.024, h: rowH * 0.8, fill: { color: C.navy } });
    s.addText(o.fmt(r.v), {
      x: bx + BARW + 0.06, y, w: VALW, h: rowH, fontFace: F, fontSize: FS,
      color: txt, align: "right", valign: "middle", isTextBox: true, margin: 0,
    });
    let ex = bx + BARW + 0.06 + VALW;
    EX.forEach(e => {
      s.addText(e.vals[r.n] ?? "", { x: ex, y, w: e.w, h: rowH, fontFace: F, fontSize: FS,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      ex += e.w;
    });
    // segBreak (Noah, Oct 24): a solid rule under a group's last store, the same width as the header rule, closing
    // the group off from the next one in the same column
    if (o.segBreak && all[i + 1]?.hdr && (i + 1 < perCol) === (i < perCol))
      s.addShape(pptx.ShapeType.rect, { x: x0, y: y + rowH - 0.02, w: 5.8, h: 0.02, fill: { color: C.navy } });
  });
  if (EX.length) [0, 1].forEach(c => {
    let ex = COLX[c] + RANKW + 0.05 + NAMEW + 0.08 + BARW + 0.06;
    [{ label: o.valLabel ?? "", w: VALW }, ...EX].forEach(e => {
      s.addText(e.label, { x: ex, y: 2.0, w: e.w, h: 0.28, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      ex += e.w;
    });
  });

  // the group (or benchmark) line, drawn through both columns so above/below is visible
  [0, 1].forEach(c => {
    const bx = COLX[c] + RANKW + 0.05 + NAMEW + 0.08;
    if (lo < 0) {   // a faint zero line, so above and below zero reads at a glance
      const zx = X(bx, 0);
      s.addShape(pptx.ShapeType.rect, { x: zx - 0.005, y: TOP, w: 0.01, h: BOT - TOP, fill: { color: C.ink300 } });
    }
    if (o.refLine && o.alignRef) {   // one solid line: every store's own segment average sits here (Noah, Oct 23)
      const gx = bx + 1 / hiR * BARW;
      if (!o.segBreak) { s.addShape(pptx.ShapeType.rect, { x: gx - 0.008, y: TOP, w: 0.016, h: BOT - TOP, fill: { color: C.navy } }); return; }
      // segBreak: one line per run of store rows in this column, so it never crosses a group header
      const idx = all.map((r, i) => i).filter(i => (c === 0 ? i < perCol : i >= perCol));
      let run = null;
      const flush = () => { if (run) { const k0 = run[0] - c * perCol, k1 = run[1] - c * perCol;
        s.addShape(pptx.ShapeType.rect, { x: gx - 0.008, y: TOP + k0 * rowH, w: 0.016, h: (k1 - k0 + 1) * rowH, fill: { color: C.navy } }); run = null; } };
      idx.forEach(i => { if (all[i].hdr) flush(); else run = run ? [run[0], i] : [i, i]; });
      flush();
      return;
    }
    if (o.refOf || o.groups) return;   // per-row ticks instead of one line (groups: no group-wide line, Oct 23)
    const gx = X(bx, REF);
    s.addShape(pptx.ShapeType.rect, { x: gx - 0.008, y: TOP, w: 0.016, h: BOT - TOP, fill: { color: C.navy } });
  });
  s.addShape(pptx.ShapeType.rect, { x: 6.655, y: TOP, w: 0.012, h: BOT - TOP, fill: { color: C.rule } });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- P&L TABLE - the GM statement's own lines ------------------
   The spine of the business overview (Oct 19 2026): the same tables, in the
   same order, with the same rows that January and May ran as Qlik
   screenshots - redrawn natively so they read like the rest of the deck.
   A native PowerPoint table, so Tommy can still click a cell and edit it.

   Values arrive STATEMENT-SIGNED, the way a GM reads his own statement:
   income positive, expense negative. That makes "up is good" true of every
   dollar line - less expense is a smaller negative - so a row only says
   otherwise when it is a share where higher is worse (wholesale %), or
   "none" where Tommy has not ruled. There is no default: a row with no
   direction stops the build, same as an unregistered metric.

   Prior year left, this year right - the order the group card reads in.
   A share's change is in points and carries no Var % (a percent of a
   percent reads as nonsense from the back of a room).

   rows    [{ head: "VARIABLE OPERATIONS SUMMARY" }
            | { line, a, b, kind: "M"|"usd"|"count"|"pct", dir: "up"|"down"|"none", strong }]
   checks  [{ label, sum: [line, ...], equals: line }] - the statement's own
           arithmetic, re-run at build so a mistyped row cannot ship
   --------------------------------------------------------------------- */
const PL_KIND = {
  M:     { fmt: Mmixed, dfmt: Dmixed },
  usd:   { fmt: Mmixed, dfmt: Dmixed },
  count: { fmt: Int,    dfmt: Dint   },
  pct:   { fmt: Pct,    dfmt: Dpts   },
};
export function plTable(pptx, deck, o) {
  const lines = o.rows.filter(r => !r.head);
  const byLine = Object.fromEntries(lines.map(r => [r.line, r]));
  lines.forEach(r => {
    if (!PL_KIND[r.kind]) throw new Error(`${o.title}: "${r.line}" has no kind (M, usd, count, pct)`);
    if (!["up", "down", "none"].includes(r.dir)) {
      throw new Error(`${o.title}: "${r.line}" has no direction - ask Tommy before colouring it`);
    }
    if (!Number.isFinite(r.a) || !Number.isFinite(r.b)) throw new Error(`${o.title}: "${r.line}" is missing a value`);
  });
  (o.checks ?? []).forEach(c => {
    [["a", "2025"], ["b", "2026"]].forEach(([k, yr]) => {
      const parts = c.sum.map(l => {
        if (!byLine[l] && !o.extra?.[l]) throw new Error(`${o.title}: check "${c.label}" names "${l}", not a row`);
        return (byLine[l] ?? o.extra[l])[k];
      });
      const got = parts.reduce((t, v) => t + v, 0);
      const want = (byLine[c.equals] ?? o.extra?.[c.equals])?.[k];
      if (want == null) throw new Error(`${o.title}: check "${c.label}" names "${c.equals}", not a row`);
      if (Math.abs(got - want) > (c.tol ?? 1)) {
        throw new Error(`${o.title}: ${c.label} ${yr} - parts sum to ${got.toFixed(2)}, statement says ${want}`);
      }
    });
  });

  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });

  // Every row holds ONE line of type - a statement line that wraps grows
  // its row, and the table runs out of its frame and under the band (it
  // did, Oct 19 2026: the totals vanished on F&I, Wholesale and Expenses).
  // So the type is the smallest of three fits: the row height, the longest
  // line name in its column, and the widest figure in a number column.
  // Width is estimated from character count, on the generous side: Inter
  // letters run ~0.54 em (0.57 bold), but its figures are wider - ~0.62 em,
  // 0.66 bold - and a first estimate that treated them like letters let
  // "−$14.10M" wrap to two rows. Below the 14pt floor the build stops -
  // shorten a label or split the table; never let it wrap.
  const COLW = [5.73, 1.65, 1.65, 1.65, 1.65], PADW = 0.3;
  const n = o.rows.length + 1;
  const avail = bodyBot(o.takeaway) - BODY.top;
  const rowH = Math.min(o.maxRowH ?? 0.85, avail / n);

  const cells = o.rows.map(r => {
    if (r.head) return null;
    const K = PL_KIND[r.kind];
    const d = r.b - r.a;
    const judged = r.dir !== "none" && Math.abs(d) > (r.kind === "pct" ? 0.005 : 0.5);
    const good = r.dir === "up" ? d >= 0 : d <= 0;
    const vp = r.kind === "pct" || !r.a ? "—"
      : (Math.abs(d / r.a) < 0.0005 ? "" : d < 0 ? "−" : "+") + Math.abs(d / Math.abs(r.a) * 100).toFixed(1) + "%";
    return { r, judged, vcol: !judged ? C.ink600 : good ? C.goodInk : C.badInk,
             t: [r.line, K.fmt(r.a), K.fmt(r.b), K.dfmt(d), vp] };
  });
  const em = (col, bold) => col === 0 ? (bold ? 0.57 : 0.54) : (bold ? 0.66 : 0.62);
  const fitW = (txt, col, bold) => (COLW[col] - PADW) * 72 / (Math.max(txt.length, 1) * em(col, bold));
  let fit = Math.min(22, rowH * 72 / 1.7);
  cells.forEach(c => {
    if (!c) return;
    c.t.forEach((txt, i) => { fit = Math.min(fit, fitW(txt, i, i === 0 ? c.r.strong : (c.r.strong || (i > 2 && c.judged)))); });
  });
  const fs = Math.floor(fit * 2) / 2;
  if (fs < FLOOR) {
    const worst = cells.filter(Boolean).map(c => c.r.line).sort((p, q) => q.length - p.length)[0];
    throw new Error(`${o.title}: will not hold one row per line at the ${FLOOR}pt floor - `
                  + `too many rows for the frame, or a label too long ("${worst}"). Split the table.`);
  }
  const hfs = Math.min(fs, 16);

  const cols = o.cols ?? ["2025 YTD", "2026 YTD", "VARIANCE", "VAR %"];
  const base = { fontFace: F, valign: "middle" };
  const hdr = ["P&L LINE", ...cols].map((t, i) => ({
    text: t, options: { ...base, fontSize: sz(hfs, "table header"), bold: true, color: C.white,
                        fill: { color: C.navy }, align: i ? "right" : "left", charSpacing: 0.6 },
  }));
  o.rows.filter(r => r.head).forEach(r => {
    if (fitW(r.head, 0, true) * 2.2 < hfs) throw new Error(`${o.title}: group heading "${r.head}" is too long for one row`);
  });

  let z = 0;
  const body = o.rows.map((r, k) => {
    if (r.head) {
      z = 0;
      return [{ text: r.head, options: { ...base, colspan: 5, fontSize: sz(hfs, "table group"), bold: true,
                                         color: C.navy, fill: { color: C.blueTint }, charSpacing: 0.8 } }];
    }
    const { judged, vcol, t } = cells[k];
    const bg = z++ % 2 ? C.zebra : C.white;
    const cell = (txt, opts = {}) => ({ text: txt, options: {
      ...base, fontSize: sz(fs, "table body"), color: C.ink, bold: !!r.strong, fill: { color: bg }, align: "right", ...opts } });
    return [
      cell(t[0], { align: "left" }),
      cell(t[1]), cell(t[2]),
      cell(t[3], { color: vcol, bold: judged || !!r.strong }),
      cell(t[4], { color: vcol, bold: judged || !!r.strong }),
    ];
  });

  s.addTable([hdr, ...body], {
    x: 0.5, y: BODY.top, w: 12.33, colW: COLW,
    rowH, valign: "middle", border: { type: "solid", pt: 0.5, color: C.rule },
  });

  takeaway(pptx, s, o.takeaway, o.tone);
  foot(pptx, s, o.source, o.page);
  return s;
}

/* ======================================================================
   DAY 2 RENDERERS - built Oct 14-21 2026 for the fixed ops day, refitted
   to the Day 1 body geometry (BODY / bodyBot / the band at 6.70).
   ====================================================================== */

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
  const W = o.card ? 12.33 : 8.3, H = bodyBot(o.takeaway) - cy;
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
    panel(pptx, s, { x: 9.0, y: BODY.top, w: 3.83, h: bodyH(o.takeaway), label: "What it says" });
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 8.3, h: bodyH(o.takeaway), label: o.panelLabel });
  const n = o.years.length, v = o.years.map(r => +(r[1] / 1e6).toFixed(2));
  const labels = o.years.map(r => r[0]);
  s.addChart(pptx.ChartType.bar, [
    { name: "Earlier", labels, values: v.map((x, i) => i < n - 1 ? x : 0) },
    { name: "Latest",  labels, values: v.map((x, i) => i === n - 1 ? x : 0) },
  ], {
    x: 0.68, y: 1.98, w: 7.95, h: bodyH(o.takeaway) - 0.8, barDir: "col", barGrouping: "stacked", barGapWidthPct: 25,
    chartColors: [C.ink300, C.blue], showLegend: false, showTitle: false,
    showValue: true, dataLabelPosition: "inEnd", dataLabelColor: C.white,
    dataLabelFontFace: "+mn-lt", dataLabelFontSize: sz(14, "bar label"), dataLabelFontBold: true,
    dataLabelFormatCode: '$0.0"M";;',
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: sz(16, "cat axis"), catAxisLabelColor: C.ink,
    valAxisHidden: true, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  });
  panel(pptx, s, { x: 9.0, y: BODY.top, w: 3.83, h: bodyH(o.takeaway), label: "What it says" });
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const per = Math.ceil(o.rows.length / 2);
  if (per > 13) throw new Error(`recordBoard: ${o.rows.length} rows will not fit`);
  const COLX = [0.72, 6.79], NAMEW = 2.17, MW = 0.64, VALW = 0.92, GAINW = 0.86, TOP = 1.86, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / per;
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

/* ---------- STACKED COUNT RANKER - a count split into parts, every store on one page ----------
   ADDED 2026-10-21 (records review). o.rows = [[name, [n1, n2, n3]]], o.parts = column heads,
   o.colors = one per part (default three greens: records are good news). Bar = the stacked
   total; a count column per part, then the total. o.notes = speaker notes. */
export function stackedCountRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const COLS = o.colors ?? [C.goodInk, C.good, "8ED3B0"];
  const tot = r => r[1].reduce((t, x) => t + x, 0);
  const rows = [...o.rows].sort((a, b) => tot(b) - tot(a) || b[1][2] - a[1][2]);
  const NP = o.parts.length, CW = 0.55, TW = 0.6, NAMEW = 2.12;
  const COLX = [0.72, 6.79], BARW = 5.9 - NAMEW - CW * NP - TW - 0.2;
  const perCol = Math.ceil(rows.length / 2), TOP = 1.86, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
  const maxV = Math.max(...rows.map(tot));
  [0, 1].forEach(c => {
    // colour key above the store names and bars; the count columns carry short heads
    let lx = COLX[c];
    (o.legend ?? o.parts).forEach((p, i) => {
      s.addShape(pptx.ShapeType.rect, { x: lx, y: 1.55, w: 0.14, h: 0.14, fill: { color: COLS[i] } });
      const w = 0.22 + p.length * 0.105;
      s.addText(p, { x: lx + 0.2, y: 1.44, w, h: 0.36, fontFace: F, fontSize: sz(14, "legend"),
        color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
      lx += 0.2 + w + 0.12;
    });
    let hx = COLX[c] + NAMEW + 0.08 + BARW + 0.08;
    [...o.parts, o.totalLabel ?? "ALL"].forEach((p, i) => {
      const w = i < NP ? CW : TW;
      s.addText(p, { x: hx, y: 1.44, w, h: 0.36, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: i < NP ? C.ink600 : C.ink, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      hx += w;
    });
  });
  rows.forEach((r, i) => {
    const c = i < perCol ? 0 : 1, k = i - c * perCol, x0 = COLX[c], y = TOP + k * rowH, bh = rowH * 0.55;
    s.addText(words(r[0], "store"), { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    let bx = x0 + NAMEW + 0.08;
    r[1].forEach((v, p) => {
      if (!v) return;
      const w = v / maxV * BARW;
      s.addShape(pptx.ShapeType.rect, { x: bx, y: y + (rowH - bh) / 2, w, h: bh, fill: { color: COLS[p] }, line: { color: C.white, width: 0.75 } });
      bx += w;
    });
    let vx = x0 + NAMEW + 0.08 + BARW + 0.08;
    r[1].forEach(v => {
      s.addText(v ? String(v) : "–", { x: vx, y, w: CW, h: rowH, fontFace: F, fontSize: 13, color: v ? C.ink : C.ink300,
        align: "right", valign: "middle", isTextBox: true, margin: 0 });
      vx += CW;
    });
    s.addText(String(tot(r)), { x: vx, y, w: TW, h: rowH, fontFace: F, fontSize: 13, bold: true, color: C.ink,
      align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "g");
  if (o.notes) s.addNotes(words(o.notes, "notes") + (o.source ? `\n\nSource: ${words(o.source, "speaker notes")}` : ""));
  else foot(pptx, s, o.source, o.page);
  return s;
}

/* ---------- GAP RANKER - two rates per store and the gap between them ----------
   ADDED 2026-10-21 (Noah: internal vs customer-pay GP%). o.rows = [[name, a, b]], gap = b - a.
   o.group = [a, b]. Green when a store's gap is at or above the group gap, red below - the
   group is the differentiator, as on every level ranker. o.mode "bar" draws the gap as a
   bar around zero; "dumbbell" draws both rates on one scale (a navy, b green/red).
   Columns: name, plot, a, b, gap. Sorted by gap, best first. */
export function gapRanker(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const gap = r => r[2] - r[1], gG = o.group[1] - o.group[0];
  const rows = [...o.rows].sort((p, q) => gap(q) - gap(p));
  const FV = o.fmt, FG = o.fmtGap;
  // group strip
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.44, w: 11.89, h: 0.42, fill: { color: C.blueTint } });
  s.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.44, w: 0.06, h: 0.42, fill: { color: C.navy } });
  s.addText([
    { text: "GROUP", options: { bold: true, color: C.navy } },
    { text: `      ${o.aLabel} ${FV(o.group[0])}      ${o.bLabel} ${FV(o.group[1])}      GAP `, options: { color: C.ink } },
    { text: FG(gG), options: { bold: true, color: C.navy } },
    { text: "      " + (o.note ?? "green = gap at or better than the group"), options: { color: C.ink600 } },
  ], { x: 0.92, y: 1.44, w: 11.6, h: 0.42, fontFace: F, fontSize: sz(15, "group strip"), valign: "middle", isTextBox: true, margin: 0 });

  const COLX = [0.72, 6.79], NAMEW = 2.2, PLOTW = 1.18, VW = 0.62, GW = 0.94;
  const per = Math.ceil(rows.length / 2), TOP = 2.3, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / per;
  const gaps = rows.map(gap), gmax = Math.max(...gaps.map(Math.abs), Math.abs(gG));
  const lo = o.min ?? Math.min(...rows.flatMap(r => [r[1], r[2]])), hi = o.max ?? Math.max(...rows.flatMap(r => [r[1], r[2]]));
  [0, 1].forEach(c => {
    const x0 = COLX[c], px = x0 + NAMEW + 0.05;
    const heads = [[o.plotLabel ?? "", px, PLOTW, "center"], [o.aShort, px + PLOTW + 0.05, VW, "right"],
                   [o.bShort, px + PLOTW + 0.05 + VW, VW, "right"], ["GAP", px + PLOTW + 0.05 + 2 * VW, GW, "right"]];
    heads.forEach(([t, x, w, al]) => s.addText(t, { x, y: 1.94, w, h: 0.3, fontFace: F, fontSize: sz(14, "column header"), bold: true,
      color: C.ink600, align: al, valign: "middle", isTextBox: true, margin: 0 }));
    const n = c === 0 ? per : rows.length - per;
    if (o.mode === "bar") {   // zero line, then the group gap as a dashed reference
      const zx = px + PLOTW / 2;
      s.addShape(pptx.ShapeType.rect, { x: zx - 0.006, y: TOP, w: 0.012, h: n * rowH, fill: { color: C.ink300 } });
      const gx = zx + gG / gmax * PLOTW / 2;
      s.addShape(pptx.ShapeType.line, { x: gx, y: TOP, w: 0, h: n * rowH, line: { color: C.navy, width: 1, dashType: "dash" } });
    }
  });
  rows.forEach((r, i) => {
    const c = i < per ? 0 : 1, k = i - c * per, x0 = COLX[c], y = TOP + k * rowH, cy = y + rowH / 2, px = x0 + NAMEW + 0.05;
    const g = gap(r), ok = g >= (o.okRef ?? gG) - 1e-9;   // o.okRef (Noah, Oct 25): 0 scores each store against its own CP rate
    if (k % 2) s.addShape(pptx.ShapeType.rect, { x: x0 - 0.08, y, w: 5.9, h: rowH, fill: { color: C.zebra } });
    s.addText(words(r[0], "store"), { x: x0, y, w: NAMEW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    if (o.mode === "bar") {
      const zx = px + PLOTW / 2, w = Math.abs(g) / gmax * PLOTW / 2, bh = rowH * 0.55;
      s.addShape(pptx.ShapeType.rect, { x: g >= 0 ? zx : zx - w, y: cy - bh / 2, w: Math.max(w, 0.01), h: bh, fill: { color: ok ? C.good : C.bad } });
    } else {
      const X = v => px + (Math.min(Math.max(v, lo), hi) - lo) / (hi - lo) * PLOTW, d = 0.15;
      const xa = X(r[1]), xb = X(r[2]);
      s.addShape(pptx.ShapeType.rect, { x: Math.min(xa, xb), y: cy - 0.015, w: Math.max(Math.abs(xb - xa), 0.01), h: 0.03, fill: { color: C.rule } });
      s.addShape(pptx.ShapeType.ellipse, { x: xa - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: C.navy } });
      s.addShape(pptx.ShapeType.ellipse, { x: xb - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: ok ? C.good : C.bad } });
    }
    const vx = px + PLOTW + 0.05;
    s.addText(FV(r[1]), { x: vx, y, w: VW, h: rowH, fontFace: F, fontSize: 13, color: C.navy, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(FV(r[2]), { x: vx + VW, y, w: VW, h: rowH, fontFace: F, fontSize: 13, color: C.ink, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(FG(g), { x: vx + 2 * VW, y, w: GW, h: rowH, fontFace: F, fontSize: 13, bold: true, color: ok ? C.goodInk : C.badInk,
      align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });

  const TOP = 1.92, BOT = bodyBot(o.takeaway) - 0.07;
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
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

  const perCol = Math.ceil(rows.length / 2), TOP = 2.0, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);
  const COLX = [0.72, 6.79], NAMEW = 2.1, VALW = 0.82, BARW = 5.7 - NAMEW - VALW - EXW - 0.14;
  const sides = o.sides.map(sd => {
    const r = sd.rows.map(([n, v]) => ({ n, v })).sort((a, b) => up ? b.v - a.v : a.v - b.v);
    if (sd.avg != null) { let gi = r.findIndex(x => up ? x.v < sd.avg : x.v > sd.avg); if (gi < 0) gi = r.length;
      r.splice(gi, 0, { n: sd.avgLabel ?? "AVERAGE", v: sd.avg, isGroup: true }); }
    return r;
  });
  const perCol = Math.max(...sides.map(r => r.length)), TOP = 1.92, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
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
  s.addShape(pptx.ShapeType.rect, { x: 6.6, y: 1.5, w: 0.012, h: BOT - 1.5, fill: { color: C.rule } });
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const n = o.tiles.length, gap = 0.22, X0 = 0.78, W = (11.77 - gap * (n - 1)) / n;
  o.tiles.forEach((t, i) => {
    const x = X0 + i * (W + gap), ink = t.tone === "b" ? C.badInk : t.tone === "g" ? C.goodInk : C.navy;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.62, w: W, h: 3.3, fill: { color: C.zebra } });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.62, w: W, h: 0.07, fill: { color: t.tone === "b" ? C.bad : t.tone === "g" ? C.good : C.blue } });
    s.addText(words(t.label, "tile label").toUpperCase(), { x: x + 0.2, y: 1.85, w: W - 0.4, h: 0.62, fontFace: F, fontSize: sz(14, "tile label"), bold: true,
      color: C.ink600, charSpacing: 0.8, valign: "top", isTextBox: true, margin: 0 });
    s.addText(words(t.big, "tile number"), { x: x + 0.2, y: 2.5, w: W - 0.4, h: 0.95, fontFace: F, fontSize: sz(o.bigSize ?? 40, "tile number"), bold: true,   // o.bigSize ADDED 2026-10-19
     
      color: ink, valign: "middle", fit: "shrink", isTextBox: true, margin: 0 });
    s.addText(words(t.sub, "tile sub"), { x: x + 0.2, y: 3.5, w: W - 0.4, h: 1.3, fontFace: F, fontSize: sz(o.subSize ?? 15, "tile sub"),   // o.subSize ADDED 2026-10-24

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
  // Oct 24 (Noah): rank numbers and the thin gray line between the two store columns, as on the other rankers.
  // o.signed: values can be negative (variance from a benchmark) - bars run right of a zero line in green for
  // above, left in red for below; o.fmt should carry the sign.
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const rows = [...o.rows].sort((a, b) => b[1] - a[1]);
  const EX = o.extra ?? [], EXW = EX.reduce((t, e) => t + e.w, 0);
  const RANKW = 0.34, COLX = [0.72, 6.79], NAMEW = EXW > 1 ? 1.95 : 2.2, VALW = o.valW ?? 0.86, BARW = 5.7 - RANKW - NAMEW - VALW - EXW - 0.14;   // o.valW ADDED 2026-10-19 for 7-digit dollars
  const perCol = Math.ceil(rows.length / 2), TOP = 1.86, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
  const lo = Math.min(0, ...rows.map(r => r[1])), hi = Math.max(0, ...rows.map(r => r[1])), span = hi - lo || 1;
  const X = (bx, v) => bx + (v - lo) / span * BARW;
  [0, 1].forEach(c => {
    let hx = COLX[c] + RANKW + 0.05 + NAMEW + 0.08 + BARW + 0.06;
    [{ label: o.valLabel ?? "", w: VALW }, ...EX].forEach(e => {
      s.addText(e.label, { x: hx, y: 1.44, w: e.w, h: 0.36, fontFace: F, fontSize: sz(14, "column header"), bold: true,
        color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      hx += e.w;
    });
    if (lo < 0) {   // zero line, through the rows of this column
      const n = c === 0 ? perCol : rows.length - perCol, zx = X(COLX[c] + RANKW + 0.05 + NAMEW + 0.08, 0);
      s.addShape(pptx.ShapeType.rect, { x: zx - 0.006, y: TOP, w: 0.012, h: n * rowH, fill: { color: C.ink300 } });
    }
  });
  rows.forEach((r, i) => {
    const c = i < perCol ? 0 : 1, k = i - c * perCol, x0 = COLX[c], y = TOP + k * rowH, bx = x0 + RANKW + 0.05 + NAMEW + 0.08;
    const neg = r[1] < 0, fill = o.signed ? (neg ? C.bad : C.good) : (o.color ?? C.blue), ink = o.signed ? (neg ? C.badInk : C.goodInk) : (o.ink ?? C.ink);
    s.addText(String(i + 1), { x: x0, y, w: RANKW, h: rowH, fontFace: F, fontSize: 12, color: C.ink300, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    // 12pt when two extra columns narrow the name column, so CHEVY GMC BULLHEAD stays on one line (Noah, Oct 25)
    s.addText(r[0], { x: x0 + RANKW + 0.05, y, w: NAMEW, h: rowH, fontFace: F, fontSize: EXW > 1 ? 12 : 13, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
    const bh = rowH * 0.55, zx = X(bx, 0), vx = X(bx, r[1]);
    s.addShape(pptx.ShapeType.rect, { x: Math.min(zx, vx), y: y + (rowH - bh) / 2, w: Math.max(Math.abs(vx - zx), 0.02), h: bh, fill: { color: fill } });
    s.addText(o.fmt(r[1]), { x: bx + BARW + 0.06, y, w: VALW, h: rowH, fontFace: F, fontSize: 13, bold: true,
      color: ink, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    let ex = bx + BARW + 0.06 + VALW;
    EX.forEach(e => { s.addText(e.vals[r[0]] ?? "", { x: ex, y, w: e.w, h: rowH, fontFace: F, fontSize: 13,
      color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
      // e.marks = { NAME: "g" | "r" } (Noah, Oct 26): the same top / bottom three bar as the change rankers
      const mk = e.marks?.[r[0]];
      if (mk) s.addShape(pptx.ShapeType.rect, { x: ex + e.w + 0.06, y: y + 0.05, w: 0.06, h: rowH - 0.1, fill: { color: mk === "g" ? C.good : mk === "n" ? C.ink300 : C.bad } });
      ex += e.w; });
  });
  s.addShape(pptx.ShapeType.rect, { x: 6.655, y: TOP, w: 0.012, h: BOT - TOP, fill: { color: C.rule } });
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
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const rows = [...o.rows].sort((p, q) => q[1] - p[1]);
  const COLX = [0.72, 6.79], NAMEW = 2.12, VALW = 1.45, PLOTW = 5.7 - NAMEW - VALW - 0.3;
  const perCol = Math.ceil(rows.length / 2), TOP = 1.9, BOT = bodyBot(o.takeaway) - 0.07, rowH = (BOT - TOP) / perCol;
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
  const TPB = bodyBot("The ask: " + o.ask), TPH = TPB - BODY.top;
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 7.0, h: TPH, label: null });
  panel(pptx, s, { x: 7.68, y: BODY.top, w: 5.15, h: TPH, label: null });
  const head = (x, w, t) => s.addText(words(t, "tp head").toUpperCase(), { x, y: 1.5, w, h: 0.34, fontFace: F, fontSize: sz(14, "tp head"),
    bold: true, color: C.navy, charSpacing: 1.0, valign: "middle", isTextBox: true, margin: 0 });
  // Day 2 (Noah, Oct 21): no "What we said · <prior meeting>" header unless o.saidFrom is passed; bullets start at the top
  if (o.saidFrom) {
    head(0.78, 6.5, "What we said · " + o.saidFrom);
    s.addShape(pptx.ShapeType.rect, { x: 0.78, y: 1.9, w: 6.45, h: 0.012, fill: { color: C.rule } });
  }
  head(7.95, 4.7, o.nowLabel ?? "Where we are · Jan–Sep 2026");
  s.addShape(pptx.ShapeType.rect, { x: 7.95, y: 1.9, w: 4.62, h: 0.012, fill: { color: C.rule } });
  const fs = o.said.length > 6 ? 16 : 18, SY = o.saidFrom ? 2.05 : 1.55;
  s.addText(o.said.map(t => ({ text: words(t, "tp bullet"), options: { bullet: { indent: 18 }, breakLine: true, paraSpaceAfter: 8 } })), {
    x: 0.78, y: SY, w: 6.45, h: TPB - SY - 0.15, fontFace: F, fontSize: sz(fs, "tp bullet"), color: C.ink, valign: "top", isTextBox: true, margin: 0 });
  const n = o.now.length, rowH = (TPB - 2.15) / n;
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
  const H2 = bodyBot(o.ours) - Y;
  panel(pptx, s, { x: 0.5, y: Y, w: 4.75, h: H2, label: null });
  panel(pptx, s, { x: 5.4, y: Y, w: 7.43, h: H2, label: null });
  const head = (x, w, t) => s.addText(t, { x, y: Y + 0.14, w, h: 0.32, fontFace: F, fontSize: sz(14, "explainer head"), bold: true, color: C.navy, charSpacing: 1.0, valign: "middle", isTextBox: true, margin: 0 });
  head(0.72, 4.3, o.whyHead ?? "WHY IT MATTERS"); head(5.62, 7.0, o.leverHead ?? "WHAT MOVES IT");
  s.addText(o.why.map(t => ({ text: words(t, "why"), options: { bullet: { indent: 16 }, breakLine: true, paraSpaceAfter: 6 } })),
    { x: 0.72, y: Y + 0.55, w: 4.35, h: H2 - 0.7, fontFace: F, fontSize: sz(o.why.length > 3 || H2 < 3.4 ? 15 : 16, "why"), color: C.ink, valign: "top", isTextBox: true, margin: 0 });
  const L = o.levers, COLS = L.length > 4 && H2 < 3.6 ? 3 : 2, rows = Math.ceil(L.length / COLS), cw = (7.0 - 0.12 * (COLS - 1)) / COLS;
  // o.leverRows: optional row weights, e.g. [1.6, 1] when one row's lines run longer (Oct 23)
  const wt = o.leverRows ?? Array(rows).fill(1), wsum = wt.reduce((a, b) => a + b, 0);
  const rh = wt.map(v => (H2 - 0.65) * v / wsum), ry = rh.map((_, r) => rh.slice(0, r).reduce((a, b) => a + b, 0));
  L.forEach(([name, line], i) => {
    const r = Math.floor(i / COLS), ch = rh[r], x = 5.62 + (i % COLS) * (cw + 0.12), y = Y + 0.55 + ry[r];
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.05, h: ch - 0.12, fill: { color: C.blue } });
    s.addText([{ text: words(name, "lever"), options: { bold: true, color: C.ink, breakLine: true } }, { text: words(line, "lever"), options: { color: C.ink600 } }],
      { x: x + 0.14, y, w: cw - 0.2, h: ch - 0.12, fontFace: F, fontSize: sz(14, "lever"), valign: "top", isTextBox: true, margin: 0 });
  });
  takeaway(pptx, s, o.ours, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  if (o.notes) s.addNotes(words(o.notes, "notes"));
  return s;
}

/* ---------- MOVERS - top five and bottom five on one change -----------------
   ADDED 2026-10-22 (Noah): one per metric, after its change ranker - a conversation starter on what each store did.
   Laid out like Who's Moving Now. o.rows = [{ n, v0, v1, d }] already formatted v0/v1 (text), d = signed change
   used to sort and colour, o.fmtD formats d. o.sub = the one line under the title. The 2026 value, its arrow and the
   change read green when better than last year, red when worse (navy for a neutral metric).
   --------------------------------------------------------------------- */
export function movers(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  const dir = dirOf(o.metric), neutral = dir === "neutral", up = dir !== "down";
  const rows = [...o.rows].sort((a, b) => up ? b.d - a.d : a.d - b.d), K = o.k ?? 5;
  const top = rows.slice(0, K), bot = rows.slice(-K).reverse();   // bottom list reads worst first
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  s.addText(words(o.sub, "movers sub"), { x: 0.78, y: 1.42, w: 11.77, h: 0.34, fontFace: F, fontSize: sz(14, "movers sub"),
    color: C.ink600, isTextBox: true, margin: 0 });
  const TOP = 2.32, BOT = bodyBot(o.takeaway) - 0.12, rh = (BOT - TOP) / K;
  [[top, 0.72, "TOP " + K, C.good, C.goodInk, i => i + 1], [bot, 6.79, "BOTTOM " + K, C.bad, C.badInk, i => rows.length - i]].forEach(([list, x0, head, bar, ink, rank]) => {
    const W = 5.82, X = { rk: x0 + 0.1, n: x0 + 0.5, v: x0 + 2.45, c: x0 + W - 1.22 };
    s.addShape(pptx.ShapeType.rect, { x: x0, y: 1.9, w: W, h: 0.05, fill: { color: neutral ? C.blue : bar } });
    s.addText(head, { x: X.rk, y: 1.98, w: 2.0, h: 0.3, fontFace: F, fontSize: sz(14, "movers head"), bold: true, color: neutral ? C.navy : ink, charSpacing: 1.0, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(o.valHead ?? "2025 → 2026", { x: X.v, y: 1.98, w: 2.0, h: 0.3, fontFace: F, fontSize: sz(14, "column header"), bold: true, color: C.ink600, valign: "middle", isTextBox: true, margin: 0 });
    s.addText("Change", { x: X.c, y: 1.98, w: 1.15, h: 0.3, fontFace: F, fontSize: sz(14, "column header"), bold: true, color: C.ink600, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    list.forEach((r, j) => {
      const y = TOP + j * rh, better = up ? r.d >= 0 : r.d <= 0, col = neutral ? C.navy : (better ? C.goodInk : C.badInk);
      if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: x0, y, w: W, h: rh, fill: { color: C.zebra } });
      s.addText(String(rank(j)), { x: X.rk, y, w: 0.35, h: rh, fontFace: F, fontSize: sz(15, "movers rank"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
      s.addText(r.n, { x: X.n, y, w: 1.9, h: rh, fontFace: F, fontSize: sz(14, "movers store"), bold: true, color: C.ink, valign: "middle", isTextBox: true, margin: 0 });
      s.addText([{ text: r.v0 + "  ", options: { color: C.ink600 } }, { text: "→  ", options: { color: col, bold: true } }, { text: r.v1, options: { color: col, bold: true } }],
        { x: X.v, y, w: 2.1, h: rh, fontFace: F, fontSize: sz(14, "movers values"), valign: "middle", isTextBox: true, margin: 0 });
      s.addText(o.fmtD(r.d), { x: X.c, y, w: 1.15, h: rh, fontFace: F, fontSize: sz(16, "movers change"), bold: true, color: col, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  if (o.notes) s.addNotes(words(o.notes, "notes"));
  return s;
}

/* Three lists side by side (Oct 23, every top/bottom slide in the deck): top K and bottom K year over year, then top K
   moving now (within 2026). No sub line: each column head says what it ranks. o.cols = [{ head, sub, rows: [{ n, v0,
   v1, d }], tone: "good"|"bad"|"blue", rank }]; o.up false for lower-is-better metrics, o.neutral for no colour. */
export function movers3(pptx, deck, o) {
  const s = deck.addSlide();
  page(pptx, s, { eyebrow: o.eyebrow, title: o.title });
  panel(pptx, s, { x: 0.5, y: BODY.top, w: 12.33, h: bodyH(o.takeaway), label: null });
  const K = o.k, gap = 0.22, W = (12.33 - 0.44 - gap * 2) / 3, HEAD = 1.5, TOP = 2.3, BOT = bodyBot(o.takeaway) - 0.12, rh = (BOT - TOP) / K;
  const big = K <= 3;
  // Oct 24 (Noah): a store on both the first list (top vs last year) and the last (moving now) gets a green bar down
  // the right edge of its row on both lists
  const both = new Set(o.cols[0].rows.map(r => r.n).filter(n => o.cols[o.cols.length - 1].rows.some(r => r.n === n)));
  o.cols.forEach((col, ci) => {
    const x0 = 0.72 + ci * (W + gap);
    const bar = { good: C.good, bad: C.bad, blue: C.blue }[col.tone], ink = { good: C.goodInk, bad: C.badInk, blue: C.navy }[col.tone];
    s.addShape(pptx.ShapeType.rect, { x: x0, y: HEAD, w: W, h: 0.05, fill: { color: bar } });
    s.addText(col.head, { x: x0 + 0.1, y: HEAD + 0.08, w: W - 0.2, h: 0.3, fontFace: F, fontSize: sz(15, "movers head"), bold: true, color: ink, charSpacing: 1.0, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(col.sub, { x: x0 + 0.1, y: HEAD + 0.4, w: W - 0.2, h: 0.3, fontFace: F, fontSize: sz(12, "movers col sub"), color: C.ink600, valign: "middle", isTextBox: true, margin: 0 });
    col.rows.forEach((r, j) => {
      const y = TOP + j * rh, better = o.up === false ? r.d <= 0 : r.d >= 0, c = o.neutral ? C.navy : (better ? C.goodInk : C.badInk);
      if (j % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: x0, y, w: W, h: rh, fill: { color: C.zebra } });
      if ((ci === 0 || ci === o.cols.length - 1) && both.has(r.n)) s.addShape(pptx.ShapeType.rect, { x: x0 + W - 0.07, y: y + 0.04, w: 0.07, h: rh - 0.08, fill: { color: C.good } });
      s.addText(String(col.rank ? col.rank(j) : j + 1), { x: x0 + 0.1, y, w: 0.4, h: rh, fontFace: F, fontSize: sz(big ? 20 : 15, "movers rank"), bold: true, color: C.navy, valign: "middle", isTextBox: true, margin: 0 });
      s.addText([{ text: r.n, options: { bold: true, color: C.ink, fontSize: sz(big ? 17 : 14, "movers store"), breakLine: true } },
                 { text: r.v0 + "  →  ", options: { color: C.ink600, fontSize: sz(big ? 14 : 12, "movers values") } },
                 { text: r.v1, options: { color: c, bold: true, fontSize: sz(big ? 14 : 12, "movers values") } }],
        { x: x0 + (big ? 0.55 : 0.48), y, w: W - (big ? 1.72 : 1.82), h: rh, fontFace: F, valign: "middle", isTextBox: true, margin: 0, paraSpaceAfter: 2 });
      // r.d2 (Oct 24): a smaller second change under the main one, e.g. +0.30 hrs under +17.4%
      if (r.d2 != null) {
        s.addText(o.fmtD(r.d), { x: x0 + W - 1.32, y: y + rh * 0.12, w: 1.22, h: rh * 0.5, fontFace: F, fontSize: sz(16, "movers change"), bold: true, color: c, align: "right", valign: "bottom", isTextBox: true, margin: 0 });
        s.addText(r.d2, { x: x0 + W - 1.32, y: y + rh * 0.58, w: 1.22, h: rh * 0.32, fontFace: F, fontSize: sz(13, "movers change sub"), color: C.ink600, align: "right", valign: "top", isTextBox: true, margin: 0 });
      } else
      s.addText(o.fmtD(r.d), { x: x0 + W - (big ? 1.12 : 1.32), y, w: big ? 1.02 : 1.22, h: rh, fontFace: F, fontSize: sz(big ? 18 : 16, "movers change"), bold: true, color: c, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    });
  });
  takeaway(pptx, s, o.takeaway, o.tone ?? "n");
  foot(pptx, s, o.source, o.page);
  if (o.notes) s.addNotes(words(o.notes, "notes"));
  return s;
}
