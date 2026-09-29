---
name: findlay-technician-proficiency-report
description: "Rebuild Findlay Automotive Group's 39-tab Technician Proficiency workbook (Proficiency, Proficiency Ranker, Tech Ranker, Capacity, Capacity Ranker, and 34 individual store sheets) for a new reporting period. Use this whenever the user asks for an updated, new, or refreshed Technician Proficiency report, mentions 'tech proficiency workbook', asks to rebuild it from a new ADP employee hours export, or references this report by name -- even if they don't attach files yet (ask for the two required inputs, the ADP Tech Employee List / Tech Efficiency Report Data export for the new period, and the prior period's finished workbook). Do not attempt to hand-build this report from scratch or from memory of its format -- always run the bundled scripts in this skill, which encode every formatting, formula, and parsing rule discovered and fixed through hands-on iteration."
---

# Findlay Technician Proficiency Report

Rebuilds the full 39-tab workbook for a new period from two inputs, reusing tested Python
(openpyxl) scripts rather than reconstructing formatting/formulas from description. This
guarantees byte-for-byte consistent output every time -- do not skip the scripts and write
new cell-by-cell logic; every rule below was learned by fixing a real mismatch against the
user's original example, and skipping the scripts risks reintroducing those same bugs.

## Required inputs (ask the user for both if not already provided)

1. **ADP roster/hours export** for the new period -- filename like
   `Tech_Employee_List__000____Tech_Efficiency_Report_Data_<start>_-_<end>.xlsx`. Single sheet
   named `1`, one row per employee: Location Description, Position ID, Legal Last Name,
   Legal First Name, Job Title Description, Position Status, Regular Hours Total,
   Overtime Hours Total. Trailer rows literally named "Grand Totals" / "Count" are ignored
   automatically.
2. **The prior period's finished workbook** (the last one this skill produced, or the
   original example the user gave the first time). Used as a style/reference source for
   Capacity's Total Stalls (see step 4) -- not touched otherwise.

You'll also need **live Qlik Cloud MCP access** to the "Closed Repair Orders" app (Fixed
space) to pull this period's technician-level production data -- see step 2.

## Workflow

### Step 1: Determine the new period

Ask the user (or infer from the roster filename) the exact new period's date range, in both
forms:
- Human label, e.g. `"Aug 25 to Sep 9, 2026"`
- Qlik-format start/end, e.g. `"8/25/2026"` / `"9/9/2026"` (M/D/YYYY, no leading zeros --
  this must match how `RO_Header.closedate` values are formatted in Qlik; confirm by
  spot-checking `qlik_get_field_values` on that field if unsure)

### Step 2: Pull this period's technician-level data from Qlik

Query the Closed Repair Orders app (Fixed space; find its appId via `qlik_search` if not
already known) with dimensions `Division` and `RO_Detail.techno`, for **all divisions at
once** (one query, not per-store), with these measures:

```
Sold Hours:  Sum({<RO_Header.closedate={">=<start><=<end>"}>} RO_Detail.soldhours)
RO Count:    Count(DISTINCT {<RO_Header.closedate={">=<start><=<end>"}, RO_Detail.hasLabor={1}>} %RO)
Labor Sale:  Sum({<RO_Header.closedate={">=<start><=<end>"}>} RO_Detail.laborsale)
Labor Gross: Sum({<RO_Header.closedate={">=<start><=<end>"}>} RO_Detail.laborgross)
```

Use ALL pay types (no PayType filter) -- this report measures total technician
productivity, not customer-pay-only. Paginate `qlik_get_chart_data` (the result set runs
several hundred rows; ~756 last time) until `offset + count >= totalRows`.

Save the full result as a JSON file (`tech_data.json`) of rows shaped exactly like:
```json
[["ACURA", "438", 40.6, 14, 8709.84, 6535.25], ["ACURA", "459", 67.5, 38, ...], ...]
```
`[Division, TechNo, SoldHours, ROCount, LaborSale, LaborGross]` -- Division is Qlik's
upper-snake-case value (e.g. `"CHEVY GMC BULLHEAD"`), TechNo is a **string**.

### Step 3: Write config.json and stage the working folder

Copy `scripts/*.py` into your working directory, place the roster `.xlsx`, the prior
workbook, and `tech_data.json` alongside them, and write `config.json` (see
`references/config.example.json`) with `roster_path`, `qlik_data_path`, `template_path`,
`output_path`, `new_period_label`, `qlik_date_start`, `qlik_date_end`.

### Step 4: Capacity is pay-period based; Total Stalls is carried forward

Capacity measures the **pay period only**, not a month. `assemble_workbook.py` writes the
period start/end from `qlik_date_start`/`qlik_date_end` into Capacity `B3:C3`, and:
- **Days Open/Period (F)** = `NETWORKDAYS.INTL($B$3,$C$3,11)` -- Mon-Sat days in the period
  (replaces the old 26 days/month).
- **Days Worked/ Period (M)** = `NETWORKDAYS.INTL($B$3,$C$3,1)` -- Mon-Fri days in the period
  (replaces the old 22 days/month).
- **Actual Period Labor Gross (P)** = a live link to the store sheet: TOTAL / AVERAGE row
  Labor Gross (col K) + the EXCLUDED / POOLED CODES Labor Gross (col E). Same dates and
  population as ELR and GP Retention %, so utilization compares like with like. (It
  replaced the old carried-forward "Actual Monthly Gross".) With that population, Tech
  Utilization works out to sold hours ÷ (techs × 10 × tech days), and Stall Utilization to
  sold hours ÷ (stalls × 12 × stall days) -- a handy sanity check.

**Total Stalls** (bay count per store) is the one input not derivable from the two input
files; it is carried forward unchanged from the prior workbook's Capacity sheet (matched by
store name). Tell the user, and refresh it from the stall count reference if it changed.

### Step 5: Run the build

```bash
cd <working directory containing config.json and the copied scripts>
python3 assemble_workbook.py
python3 /mnt/skills/public/xlsx/scripts/recalc.py <output_path> 120
```
Confirm `"status": "success"` and `"total_errors": 0`. If `total_errors` is nonzero, do not
ship the file -- something in this period's data is triggering a formula edge case; find
and fix the root cause (e.g. a division-by-zero not wrapped in IFERROR) rather than
suppressing the error.

### Step 6: Sanity-check before presenting

- Sheet count is 39, sheet order is the 5 summary tabs then 34 stores alphabetically,
  active sheet is "Proficiency".
- Pick 1-2 stores and confirm technicians within each category are sorted by
  **Proficiency % (Sold Hours ÷ Actual Hours) descending, not Sold Hours** -- this was a
  real bug once; see `references/lessons-learned.md`.
- Spot-check that Tech No parsing didn't silently misclassify real employees as
  "unmapped/pooled" -- see the parsing rule in `references/lessons-learned.md` before
  trusting the unmapped counts.
- Copy the output to wherever the user can access it and present it via the file-sharing
  tool.

## Known caveats to mention to the user

- Capacity's Total Stalls is carried forward, not refreshed (see Step 4). Capacity and
  utilization cover the pay period only.
- A handful of Qlik tech codes with no roster match (e.g. `MULT`, `999`, `9999`, `77777`)
  land in each store's "EXCLUDED / POOLED CODES" block, exactly as in the original design --
  this is expected, not a bug.

## Reference files

- `references/lessons-learned.md` -- every formatting/parsing bug found and fixed while
  building this skill (tech-number parsing, sort order, merged-cell fills, border color,
  Capacity number formats/headers). Read this before making ANY manual edit to the
  generated workbook, since it documents exactly what "correct" looks like and why.
- `references/config.example.json` -- template for config.json.
- `references/workbook-structure.md` -- full structural reference (sheet list, store-sheet
  row layout, category order, formula reference) for anyone needing to modify the scripts
  themselves.
