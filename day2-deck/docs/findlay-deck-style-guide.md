# Findlay GM Meeting Deck — Style Guide

Locked August 2026. Implemented in the `findlay-meeting-deck` skill; this doc records the *decisions*. Change this doc and the skill together, never one alone.

---

## Format — PowerPoint is primary

A deck is written as a spec and generated:

    node pptx/build.mjs decks/<meeting>.mjs "Meeting.pptx"   # primary
    node build.mjs      decks/<meeting>.mjs "Meeting.html"   # screen-first

**PowerPoint is the deliverable** (Tray, Aug 2026). Charts generate as **native chart objects** — a bar can be clicked and changed. The generator keeps 100+ slide decks consistent and enforces the metric registry at build time; editing afterwards in PowerPoint is expected.

**HTML is parked for large meetings** (Tray, Aug 16 2026) — he wants it for smaller presentations later, so keep it working, but do not build or deliver it for a GM meeting unless he asks.

`Findlay Deck Template.potx` — 12 named layouts — is for building by hand without the generator.

16:9 throughout. HTML canvas 1600×900; PowerPoint 13.333×7.5in. **1px = 0.6pt.**

---

## 🔒 Naming — the name of a slide stays the name of the slide

Locked by Tray, Aug 16 2026. This was the second thing he had to correct, so it is enforced in code, not by discipline.

1. **The segment name is the fine print on every slide in that segment** — the eyebrow above the title *and* the footer, identical on all of them. "Voice AI" on every Voice AI slide. "2026 1st Half YOY" on every Segment 01 slide. It does not become "FLAI · JUL 16 – AUG 15, 2026" on one slide and something else on the next.
2. **Both slides of a data pair carry the SAME title.** A group card and the store view that follows it are one idea shown twice, not two slides needing two names.
3. **The title is the source line name and nothing else.** `New Vehicle Units`. `Total Gross Profit - Service`. `Net Profit`. No "— 2026 vs 2025", no "change by store", no invented descriptive wording.
4. **The period never goes in the title or the eyebrow.** The segment eyebrow states the frame once. Where a different window applies — Flai's rolling 30 days — annotate it on the slide; "30 day" is sufficient, the full date range is not wanted.

Implementation: `seg(name)` returns `{section, eyebrow}` from a single string, so the two can never drift. **Never set `eyebrow` by hand in a deck spec.**

---

## Segment rhythm (Tray, Aug 2026)

**separator → key talking points → data (1–2 slides) → call to action / takeaway**

Four to five slides per segment beats one dense slide. Use `section` → `content` → data types → `statement`.

**Keep slide count down** (Tray, Aug 16). Consolidate lines rather than adding pairs — he replaced ten Segment 01 metrics with eight for exactly this reason.

---

## Colour

**Findlay blue ramp:** `050 #F1F7FC · 100 #DCEBF8 · 200 #B9D7F0 · 300 #8FBEE6 · 400 #5EA0D8 · 500 #2E7FC4 · 600 #1B6BAF · 700 #14548C · 800 #103F69 · 900 #0B2C4A`

**Cool neutral ink:** `050 #F7F9FC` → `900 #0C1420`.

**Status — green is good, red is bad, always.**

| role | fill | ink | background |
|---|---|---|---|
| good | `#12A566` | `#0B7A4A` | `#E6F6EE` |
| bad | `#C2352F` | `#A82B26` | `#FBEBEA` |
| warning | `#C88A16` | — | `#FBF3E2` |

Validated on white: CVD ΔE 9.3, normal-vision ΔE 31.0, both clear 3:1. Position and a signed label always accompany colour.

Blue is **identity**, not judgment. Green and red appear only where a direction is established.

**Store names always render in ink `#0C1420` — never grey** (Tray, Aug 2026, and he had to say it twice). Applies to ranker labels, diverging labels, chart category axes and table store cells. ⚠ The PowerPoint chart default is grey (`catAxisLabelColor: C.n500`) — `ranker` and `diverging` must override it explicitly on every chart. Check it on every render.

**PowerPoint theme slots:** accent1 blue · accent2 light blue · accent3 navy (GROUP) · **accent4 GOOD** · **accent5 BAD** · accent6 warning.

---

## 🔒 Lead with the good news

Locked by Tray, Aug 16 2026. Store charts sort **best-first**, so the top of the left-hand column is the strongest performer — that is the first thing the room reads.

"Best" is **direction-aware**, read from the metric registry: largest gain leads on a production line, largest *cut* leads on an expense line. `deltaRows()` does this and throws if the metric is not registered rather than guessing a sort order. Never hand-sort a chart.

---

## Direction semantics

Every metric carries a direction in `metrics.mjs`; every delta, variance bar, table cell, scored ranker, YOY card and chart sort reads from it. Nothing is hand-coloured.

- **up is good** — units, gross, net, all PVR, penetration, products per deal, ELR, HPRO, absorption, traffic, leads, closing %, trade ratio, DSR
- **down is good** — every expense line, headcount, chargebacks, F&I zeros, days to sell / frontline / photos / recon, aged units, inventory age, days supply, trade bumps, price-to-sale gap, wholesale loss
- **closer to zero is good** — recon estimation, appraisal variance, forecast variance
- **neutral** — inventory units, UA/OA, wholesale units

**Rulings (Tray, Aug 2026):** Days Supply lower is better · Finance % and Lease % up good, Cash % up bad · Recon estimation closer to zero is better · **Voice AI transfer rate UP IS GOOD** — *"we interpret that as giving the customer what they want."* A transfer is the correct outcome; the failure mode is `no pickup %` after the handoff.

**Unknown metrics are not guessed.** The build throws and stops. Ask, then add with a dated comment.

---

## Data sources (Tray, Aug 2026)

- **Sales Log → PACE only.** New + Used + Fleet.
- **Profit & Loss → all historical units, gross, expense.** Authoritative for everything reported.
- **Flai dashboard → Voice AI only.** 25 locations, not the 34 rooftops.

Never mix two sources on one slide. Always state the period and the source; stamp the pull date on anything from a live app.

⚠ **Qlik null trap.** `Sum({<set>} (Jan_Amt+Feb_Amt+…))` silently drops any row with a null in any month — null + number = null. It cost $3.7M of group net profit and up to $699K on a single store, with no error. Use `RangeSum(...)`, or sum seven separate `Sum()`s. Per-month measures are unaffected. **Always reconcile store rows to the group total before shipping a pair.**

---

## Rooftops

**34 operating rooftops**, canonical list in `rooftops.mjs`, rebuilt Aug 2026 from the authoritative `Division` field of the P&L app. Title Case in decks; aggregate row is always `GROUP`, never `AVERAGE`. Lamborghini is a social page, not a rooftop. `normalizeRows()` fails the build on an unrecognised name.

**Vendor location lists are not the rooftop list.** Flai reports 25 locations whose names do not map 1:1 (Post Falls appears twice; "Lexus Spokane" has no P&L twin). Vendor rows must **not** be run through `canonical()` — keep them in their own namespace so a vendor number can never wander onto a P&L slide.

**Partial-year stores.** INEOS opened 2026. It stays in dollar and unit YOY charts, where zero is a real base, with a footnote — and is **excluded from per-unit charts**, where a $0 baseline is not a baseline (`deltaRows(..., {skipPartial:true})`). Counts in per-unit takeaways read "of the 33 rooftops with a prior year."

---

## Typography

Inter — embedded in HTML, **must be installed** for PowerPoint.

| role | HTML px | PPT pt |
|---|---|---|
| cover title | 76 | 46 |
| section break | 80 | 48 |
| statement | 56 | 34 |
| slide title | 42 | 28 |
| takeaway | 21 | 14 |
| bullets | 21 | 14 |
| table body | 18 | 11–13 |
| store names in charts | 17 | 9 |
| data labels | 16.5 | 9 |
| footer | 15.5 | 9.5 |
| eyebrow | 14 | 9 caps, tracked |

**Built for a room, not a laptop.** When legibility and density conflict, density loses. Sentence case titles, except where the title is a source line name — then it is reproduced exactly. ALL CAPS only in the eyebrow.

**Takeaways are two lines, ~190 characters.** Past that they spill a third line onto the legend, which is only visible after rendering — so the PowerPoint build warns at 190+ and names the slide. Trim rather than ignore.

---

## Charts

Native in both outputs. No screenshots where the numbers exist.

- Every data point labelled when labels fit; HTML auto-sizes 17px→12px before falling back to ends and extremes
- Gridlines hairline, solid, never dashed
- Text wears ink tokens; the coloured mark carries identity
- **34-rooftop rankers split into two columns** on a shared axis, `GROUP` marked
- **Diverging axes round out past the longest bar, padded by label width** — otherwise the end label lands on the opposite column's store names. `$2,323` needs more runway than `274`.

**`yoyCard`** — the group YOY format: a card showing 2025 · 2026 · Variance · Variance % for the period, a month-by-month table beneath, and a tinted YTD footer row. Optional methodology `note` sits **between card and table** — a PowerPoint table row is a *minimum* height, so anything placed under it can collide. Card values auto-size. `a`/`b` override the YTD totals and are **required for any ratio metric**, where YTD is YTD gross ÷ YTD units, not the mean of the months.

Store views are `diverging` (YOY change) rather than level rankers — the point is decomposing the group variance.

---

## Rules that are not negotiable

1. **Every chart slide carries a one-line takeaway** — direct, numeric, names the store. If you can't write it, the slide doesn't belong.
2. **Fill the frame.** `table` has `summary`, `content` has `strip`, `kpis` has `insight`.
3. **One string per rooftop.** Aggregate is `GROUP`.
4. **Screenshots get redrawn** unless the visual itself is the point.
5. **Show the whole group, good and bad** — on screen, not called out by name from the podium.
6. **Verify before sending.** PowerPoint: `validate.py`, convert to images, look at 5–6 slides.
7. **Never fabricate a number to fill a slide.** An unbuilt slide is an explicit TO BUILD slide carrying what it needs and which decisions are outstanding — never a placeholder figure. Vendor estimates are labelled as estimates.

## Open items

- Real logo **resolved** — vector-traced SVG, `assets/findlay-logo.svg` is the master.
- Undecided: per-store commitment scorecards one-per-slide vs a scored table with exceptions broken out.
