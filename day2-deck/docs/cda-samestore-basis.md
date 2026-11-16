# Same-store basis — CDA Aug 2026 deck

**Decision (Tray, Aug 16 2026):** INEOS opened in 2026, so a year-over-year comparison against it is a comparison against zero. It comes out of **every YOY figure in the deck — the group cards and month tables as well as the store charts.** 33 rooftops.

Not adjusted, because they are not prior-year comparisons: **Inventory** (on-hand snapshot, Segment 04) and **Super Gross per Employee** (three-month level, Segment 03). Both keep all 34.

Implementation: `data/cda-ineos.mjs` (the adjustment) → `data/cda-derive.mjs` (`ss()` + INEOS filtered out of `BY_STORE` once, so every chart, count and ranker follows).

---

## Why 2025 needed no adjustment

Verified in Qlik: All In Sales Gross for 2025 is **176,163,907.60 whether INEOS is excluded or not** — identical to the cent. INEOS contributed nothing in 2025, so same-store 2026 = group 2026 − INEOS 2026 and 2025 passes through untouched. That is what makes this adjustment cheap and auditable.

Exclusion set syntax that works: `Division={"*"}-{'INEOS'}`.

## What changed on the headline numbers

| | As reported (34) | Same-store (33) |
|---|---|---|
| New units | −485, −2.3% | **−527, −2.5%** |
| Used units | +998, +6.1% | **+927, +5.7%** |
| Total units delivered | +513 | **+400** |
| All in sales gross | −$13,244,486, −7.5% | **−$13,779,775, −7.8%** |
| All in gross per unit | −$414, −8.8% | **−$415, −8.8%** |
| Service gross | +$4,937,350, +5.2% | **+$4,735,091, +4.9%** |
| Parts gross | +$203,842, +0.4% | **+$132,110, +0.3%** |
| Fixed ops gross | +$4,392,625, +3.1% | **+$4,120,801, +2.9%** |
| Total expense | +$2,441,531, +1.0% | **+$1,248,323, +0.5%** |
| Expense per unit | $6,279 → $6,258 | **$6,279 → $6,245** |
| Net profit | −$10,481,590, −13.5% | **−$10,090,384, −13.0%** |
| Personnel | −$2,568,032, −1.6% | **−$3,057,454, −2.0%** |
| Personnel % of super gross | 48.8% → 49.4% | **48.8% → 49.4%** |
| Advertising | **+$71,722, +0.3%** | **−$127,769, −0.5%** |
| Advertising per unit | $706 → $699 | **$706 → $696** |
| Floorplan net credit | +$1,362,232, +50.6% | **+$1,437,233, +53.4%** |
| Floorplan % of all-in | 1.53% → 2.49% | **1.53% → 2.54%** |
| Headcount Aug vs Aug | +214, +8.3% | **+206, +7.9%** |
| Affordability — new units | +238, +2.6% | **+218, +2.4%** |
| Affordability — used units | +894, +13.4% | **+851, +12.8%** |

**Three of these change what Tray says out loud:**

1. **Total expense.** INEOS alone was $1,193,208 of the $2,441,531 increase — half of it. Same-store, expense is up only 0.5% on 400 more units, and expense per unit is down $34 rather than $21. The "expense is not the problem" argument gets materially stronger.
2. **Advertising flips sign.** As reported it was up $71,722. Same-store it is **down $127,769**, and the store split moves from 17/17 to 17 spending less, 16 more. Do not say "essentially flat."
3. **Personnel.** The cut grows from $2,568,032 to $3,057,454 (−2.0%), which sharpens the rate story: personnel fell 2.0% and super gross fell 3.0% — gross still outran the cuts.

Counts that moved: new units 9 rooftops grew (was 10) · used 19 (was 20) · all-in gross 9 (was 10) · service 21 (was 22) · parts 18 up (was 19) · fixed ops 20 (was 21) · net profit 28 of 33 profitable · the two level rankers re-split 13/20 and 19/14.

## Reconciliation

All 33 store rows sum to the same-store group card **to the dollar**, both years, on all ten dollar and unit lines. Run that check before any further change.

INEOS's own monthly 2026 figures reconcile to its row in the store table to the cent on every one of the twelve lines pulled.

## The trade — say it before someone else does

**Same-store group figures no longer match the consolidated P&L.** Net profit reads $67,657,105 here against $67,265,899 consolidated; all-in gross $162,384,133 against $162,919,421. Chad presents accounting immediately before this segment.

Mitigation in the deck: **every group card label carries "· SAME-STORE, 33 ROOFTOPS"** and every store chart note carries "Same-store — INEOS excluded, opened 2026." A GM checking a number against the P&L sees the basis without having to ask. If Tray wants the as-reported figure alongside, the card note is the place for it.

## Where INEOS still appears, deliberately

- **Inventory rankers** (Segment 04) — snapshot, all 34, "of 34" language intact.
- **Super Gross per Employee** (Segment 03) — three-month level, INEOS at $19,052, above the $15,000 line.
- **Total expense store chart** — INEOS's $1,193,208 of first-year cost used to be called out in the takeaway; that sentence is gone. Replaced with "Fifteen added cost, led by Honda Henderson at +$1,470,590."

## Guardrails added

- `ss(key, series)` **throws** if a series has no INEOS counterpart, so a new group series cannot quietly ship on the wrong basis.
- `levelRows()` now requires a registry metric and sorts direction-aware, same as `deltaRows()`.
- The headcount variance strip is **derived** from `H.AUG`, not typed. It had been hard-coded at "+214 / +8.3%" and went stale the moment the basis changed — that class of bug is now impossible on that slide.
