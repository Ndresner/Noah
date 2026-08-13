# Findlay Technician Proficiency Workbook — Build Specification

This is the full rebuild spec for the **Technician Proficiency workbook** (39 tabs: 5 group-level
summary/ranker/capacity tabs + 34 store tabs). It is written so Claude Code can rebuild this report
from scratch for any new period without re-deriving conventions from the sample file.

Reference sample output: `Findlay_Technician_Proficiency_WhatIf_Jul25-Aug9_2026.xlsx`

---

## 1. Data sources

### 1.1 Qlik Cloud — Closed Repair Orders app
- **appId:** `c3efc739-b063-47dd-a4d3-d1c4cac07ad0` (space: "Fixed", managed)
- Always `qlik_clear_selections` first; prefer set analysis over app-level selections for one-off pulls.
- Store filter field: **`Division`** (not `Store`, which carries company-code suffixes).
- Date filter field: **`RO_Header.closedate`**, format `M/D/YYYY` (no leading zeros), e.g.
  `RO_Header.closedate={">=7/25/2026<=8/9/2026"}` — use double quotes (search-mode) for the range operator.
- Per-store, per-technician pull — one `qlik_create_data_object` call per store:
  - **Dimension:** `RO_Detail.techno` (label "TechNo")
  - **Measures** (all wrapped in the same set-analysis modifier, `Division={'STORE'}` + date range):
    - `Sum(RO_Detail.soldhours)` → SoldHours
    - `Count(DISTINCT RO_Detail.ronumber)` → ROCount
    - `Sum(RO_Detail.laborsale)` → LaborSale
    - `Sum(RO_Detail.laborgross)` → LaborGross
  - Sort by SoldHours descending. Set `limit.count` generously (150–200) since larger stores
    (Toyota Henderson, Chevy LV) can return 40–75 distinct tech codes; paginate with
    `qlik_get_chart_data` + `offset` if `totalRows` exceeds the returned `count`.
- Run this for **all 34 stores** (see Division value list in §7). Save all 34 result sets before
  building anything — this is the single most expensive part of the rebuild.

### 1.2 ADP roster upload (per-period Excel file, provided by the user)
Columns: `Location Description | Position ID | Legal Last Name | Legal First Name | Job Title
Description | Position Status | Regular Hours Total | Overtime Hours Total`

- **Actual Hours (Reg+OT)** = `Regular Hours Total + Overtime Hours Total`, rounded to 2 decimals.
- **Location Description → Division mapping is NOT literal.** Confirm every store name maps
  correctly before trusting it — company names in ADP frequently differ from the Qlik `Division`
  value (e.g. `MOTOR COMPANY` → **Chevy GMC Bullhead**, `AUDI RENO TAHOE` → **Audi Reno**,
  `CHRYSLER POST FALLS` → **Cdjr Post Falls**, `JAGUAR LAND ROVER HENDERSON` → **Lr Henderson**,
  `Findlay Honda Spokane` → **Honda Spokane**, `VW ST. GEORGE` → **Vw St George**). Validate any
  ambiguous mapping (like MOTOR COMPANY) by cross-checking that the resulting Tech Nos actually
  appear in that store's Qlik pull, not just by string similarity.
- **Tech No from Position ID:** Position IDs observed so far are always **9 characters**: a
  company-code prefix of **4 characters** (letters and/or digits) followed by a **5-digit,
  zero-padded** technician number. Strip the first 4 characters and parse the remaining 5 digits
  as an integer (drops leading zeros) to get the Tech No that matches Qlik's `RO_Detail.techno`.
  - `YD6091249` → strip `YD60` → `91249`
  - `Z7Q005020` → strip `Z7Q0` → `05020` → `5020`
  - Validate this against at least one store's known-good Tech No list before trusting it broadly
    for a new roster export, since ADP ID formats have changed before.

---

## 2. Technician categorization

Map each `Job Title Description` (upper-cased) to a **group** and a display **label**:

| Job Title (ADP)         | Group                              | Label                  |
|---|---|---|
| SERVICE TECHNICIAN      | Service Technician                 | Service Technician     |
| EXPRESS TECHNICIAN      | Express Technician                 | Express Technician     |
| SERVICE TEAM LEADER     | Service Team Leader/Shop Foreman   | Service Team Leader    |
| SHOP FOREMAN            | Service Team Leader/Shop Foreman   | Shop Foreman           |
| BODY SHOP HELPER        | Body Shop                          | Body Shop Helper       |
| BODY SHOP PAINTER       | Body Shop                          | Body Shop Painter      |
| BODY SHOP TECH          | Body Shop                          | Body Shop Tech         |
| DETAILER                | Detailer                           | Detailer               |
| DETAIL MANAGER          | Detailer                           | Detail Manager         |
| TINTER                  | Tinter                             | Tinter                 |

**Group display order on every store sheet** (skip any group with zero techs that period):
`Service Technician → Express Technician → Service Team Leader/Shop Foreman → Body Shop →
Detailer → Tinter → Unmapped`

**Qualifying groups** for group-level and cross-store rankings (Proficiency, Proficiency Ranker,
Tech Ranker): `Service Technician`, `Express Technician`, `Service Team Leader/Shop Foreman` only.
Body Shop, Detailer, Tinter, and Unmapped are shown on the store sheet and rolled into the store's
`TOTAL / AVERAGE` row, but excluded from those three cross-store views.

### 2.1 Pool / excluded codes
A Qlik `techno` value is treated as a **generic/pooled code** — excluded from the main table and
from every total — if (case-insensitive) it's one of:
`MULT, DS, 999, 9999, 997, 998, 1000, 77777`
These go in a separate **"EXCLUDED / POOLED CODES"** block at the bottom of the store sheet
(Tech No, Sold Hours, RO Count, Labor Sale, Labor Gross only — no name/position/proficiency).

### 2.2 Unmapped technicians
A `techno` that is **not** a pool code and **not found** in that store's ADP roster is still shown
in the main table, in an **"Unmapped"** group at the end, with:
- Name = `(unmapped)`, Position = blank
- Actual Hours (Reg+OT) = the literal text `"N/A"` (not zero — this must flow through the
  Proficiency %/Status formulas as N/A, not divide-by-zero)

### 2.3 Outlier exclusion (summary/ranking calculations only — never removes rows from a store sheet)
When computing store-level qualifying totals (Proficiency, Proficiency Ranker, Tech Ranker), drop:
- Service Team Leader / Shop Foreman techs at or below **20%** proficiency
- Service / Express Technicians at or below **2%** proficiency
- Anyone with Actual Hours = "N/A" or 0 (no hours data)

There is **no upper-bound exclusion** — a technician at 1000%+ proficiency (very low actual hours,
some sold hours) is legitimately included and can rank #1 on Tech Ranker. This is intentional.

### 2.4 Sort order within each group (store sheet)
Descending by Proficiency % (`Sold Hours / Actual Hours`). Rows with Actual Hours = "N/A" sort to
the bottom of their group, secondarily sorted by Sold Hours descending.

---

## 3. Workbook-level conventions

- **Font:** Arial throughout, size 11 for body/headers, size 13 bold for titles.
- **Colors:**
  - Navy `#1A2744` — title backgrounds, section headers, TOTAL/AVERAGE row, separator rows
  - Gold `#C9A04B` — GROUP SUBTOTAL rows
  - Black `#000000` — EXCLUDED/POOLED CODES sub-header fill
  - Yellow `#FFFF00` — input/editable cells (What-If target proficiency)
  - Green `#C6EFCE` fill / `#006100` font — "good" conditional formatting state
  - Yellow `#FFEB9C` fill / `#9C6500` font — "warning" state
  - Red `#FFC7CE` fill / `#9C0006` font — "bad" state
- **Borders:** thin black borders (`THIN_BORDER`, all 4 sides) on every data cell and header cell
  in tables. Title/back-link rows use a lighter partial border (left/top/bottom only).
- **Number formats:**
  - Hours: `#,##0.0`
  - Currency: `\$#,##0` (whole dollars) or `\$#,##0.00` for ELR/per-unit values
  - Percent: `0.0%`
- Every sheet: `showGridLines = True`.
- **Sheet order:** Proficiency, Proficiency Ranker, Tech Ranker, Capacity, Capacity Ranker, then
  all 34 store sheets (order below). Workbook opens active on **Proficiency**.
- Title text pattern: `"{Store or Group} -- {Report Name} -- {Period Label}"`, e.g.
  `"Cadillac -- Technician Proficiency -- Jul 25 to Aug 9, 2026"`. Period label is human-readable
  (`Mon D to Mon D, YYYY` or `Mon D to Mon D, YYYY` across months), derived from the actual Qlik
  date-range filter used for that build — never hardcode a stale label.

---

## 4. Store sheet layout (repeat for all 34 stores)

Columns A–S. Freeze panes at `A5`. Column widths:
`A=17, B=24, C=21, D=12, E=13, F=12, G=23, H=15, I=16, J=12, K=13, L=12, M=13, N=3 (spacer),
O=14, P=13, Q=13, R=13, S=15`

| Row | Content |
|---|---|
| 1 | Merged `A1:L1`, title, Arial 13 bold navy text, no fill |
| 2 | Merged `A2:L2`, `=HYPERLINK("#'Proficiency'!A1","<< Back to Proficiency")`, blue hyperlink font |
| 3 | Merged `O3:S3`, "Technician Proficiency What-If Scenario", Arial 13 bold navy, centered. Row height 15.75 |
| 4 | Header row (navy fill, white bold, centered+wrapped): see columns below |
| 5…N | One row per technician, grouped and sorted per §2 |
| — | `GROUP SUBTOTAL` row after each group (gold fill, bold) |
| — | Navy separator row (merged `A:M` and `O:S`, no text) between consecutive groups — **not** after the last group |
| — | One unstyled blank row, then `TOTAL / AVERAGE` row (navy fill, white bold) |
| — | Blank row, then `STORE-WIDE WHAT-IF SCENARIO` block |
| — | Blank row, then `EXCLUDED / POOLED CODES` block |

### 4.1 Column headers (row 4) and formulas (row `r`)

| Col | Header | Cell content |
|---|---|---|
| A | Tech No | value |
| B | Name | value |
| C | Position | value |
| D | Sold Hours | value, `#,##0.0` |
| E | RO Count | value |
| F | Avg Hrs/RO | `=IFERROR(D{r}/E{r},0)`, `#,##0.0` |
| G | Actual Hours (Reg+OT) | value or text `"N/A"`, `#,##0.0` |
| H | Proficiency % | `=IF(G{r}="N/A","N/A",IFERROR(D{r}/G{r},0))`, `0.0%` |
| I | Status | `=IF(H{r}="N/A","No Hours Data",IF(H{r}>=1,"Above Target",IF(H{r}>=0.85,"On Target","Under Target")))` |
| J | Labor Sale | value, `$#,##0` |
| K | Labor Gross | value, `$#,##0` |
| L | GP% | `=IFERROR(K{r}/J{r},0)`, `0.0%` |
| M | ELR | `=IFERROR(J{r}/D{r},0)`, `$#,##0.00` |
| N | *(spacer, width 3, no header)* |
| O | Current Proficiency % | `=H{r}`, `0.0%` |
| P | Target Proficiency % | **blank, yellow fill** — user input, `0.0%` |
| Q | Proficiency Increase | `=IF(P{r}="","",IF(G{r}="N/A","N/A",P{r}-O{r}))`, `0.0%` |
| R | Additional Sold Hours | `=IF(P{r}="","",IF(G{r}="N/A","N/A",G{r}*P{r}-D{r}))`, `#,##0.0` |
| S | Additional Labor Gross | `=IF(R{r}="","",IF(R{r}="N/A","N/A",IFERROR(R{r}*M{r}*L{r},0)))`, `$#,##0` |

### 4.2 GROUP SUBTOTAL row (row `sr`, group spans `grp_start:grp_end`)

| Col | Formula |
|---|---|
| A | `"GROUP SUBTOTAL"` (gold fill, bold, applied across the whole row) |
| D | `=SUM(D{grp_start}:D{grp_end})` |
| E | `=SUM(E{grp_start}:E{grp_end})` |
| F | `=IFERROR(D{sr}/E{sr},0)` |
| G | `=SUMIF(G{grp_start}:G{grp_end},"<>N/A")` |
| H | `=IFERROR(D{sr}/G{sr},0)` |
| J | `=SUM(J{grp_start}:J{grp_end})` |
| K | `=SUM(K{grp_start}:K{grp_end})` |
| L | `=IFERROR(K{sr}/J{sr},0)` |
| M | `=IFERROR(J{sr}/D{sr},0)` |
| O | `=H{sr}` |
| P | `=IFERROR((D{sr}+R{sr})/G{sr},0)` |
| Q | `=IFERROR(P{sr}-O{sr},0)` |
| R | `=SUM(R{grp_start}:R{grp_end})` |
| S | `=SUM(S{grp_start}:S{grp_end})` |

### 4.3 TOTAL / AVERAGE row (row `tr`; sums each GROUP SUBTOTAL row, not raw ranges)
Same formula shapes as §4.2 but each term is a literal sum of the group-subtotal cells, e.g.
`D{tr} = D{subtotal_1}+D{subtotal_2}+...`. Navy fill, white bold. No Status column value.

### 4.4 STORE-WIDE WHAT-IF SCENARIO block
```
Row: "STORE-WIDE WHAT-IF SCENARIO"                (merged A:E, bold)
Row: "Metric" (merged A:D, navy/white) | "Value" (navy/white, col E)
Row: "Current Store Proficiency %"      (merged A:D) | =H{total_row}
Row: "Target Store Proficiency %"       (merged A:D) | blank, yellow fill — input
Row: "Proficiency Increase"             (merged A:D) | =IF(E{tgt}="","",E{tgt}-E{cur})
Row: "Additional Sold Hours (Store-Wide)"(merged A:D)| =IF(E{tgt}="","",IFERROR(G{tr}*E{tgt}-D{tr},0))
Row: "Additional Labor Gross Profit (Store-Wide)"(A:D)| =IF(E{hrs}="","",IFERROR(E{hrs}*M{tr}*L{tr},0))
```

### 4.5 EXCLUDED / POOLED CODES block
```
Row: "EXCLUDED / POOLED CODES (not individually named technicians)"   (bold, partial border)
Row: Tech No | Sold Hours | RO Count | Labor Sale | Labor Gross       (black fill, white bold)
Rows: one per pool code found in that store's Qlik pull (values only, no formulas)
```

### 4.6 Conditional formatting (apply per data range, one rule set per group's row range, plus store-wide cells)
- **Status column (I):** `="Under Target"` → red; `="On Target"` → yellow; `="Above Target"` → green
- **GP% column (L):** `>=0.8` → green; `>=0.75 AND <0.8` → yellow; `<0.75` → red
- **Current Proficiency % (O):** `>=1` → green; `>=0.85 AND <1` → yellow; `<0.85` → red
- **Proficiency Increase (Q):** numeric `>0` → green; numeric `<0` → red
- **Additional Labor Gross (S):** numeric `>0` → green; numeric `<0` → red
- **Additional Sold Hours (R):** numeric `>0` → bold black on green; numeric `<0` → bold black on red
- **Name column (B):** highlight yellow whenever `P{row}<>""` (a what-if target has been entered)
- Same three-tier rule (green/yellow/red) repeated on the store-wide block's E58 (Current
  Proficiency), E60 (Increase), E61/E62 (Additional Hours/Gross) — increase/gross use the bold
  black-on-green/red variant, not the colored-font variant.

---

## 5. Proficiency sheet
- Row 1 title, Row 2 explanatory subtitle (see exact wording in the sample — states which
  categories are included/excluded and the outlier rule), Row 4 headers, data from Row 5.
- **One row per store, sorted alphabetically by store name.**
- Columns: `Store (hyperlink to store sheet) | Technicians | Total Sold Hours | Total Actual Hours |
  Avg Proficiency % | Status | Total Labor Sale | Total Labor Gross | Avg Labor GP%`
- All values are **hardcoded numbers computed from the qualifying, non-outlier tech list** per
  store (not live formulas into the store sheets) — this mirrors the original build.
- Status formula: `=IF(B{r}=0,"No Data",IF(E{r}>=1,"Above Target",IF(E{r}>=0.85,"On Target","Under Target")))`
- Column widths: `A=20, B=13, C=18, D=20, E=21, F=16, G=18, H=19, I=20`. Freeze `A5`.

## 6. Proficiency Ranker sheet
Identical columns/logic to Proficiency, but **sorted by Avg Proficiency % descending** instead of
alphabetically. Subtitle wording differs slightly to describe the ranking instead of the
alphabetical sort.

## 7. Tech Ranker sheet
- Ranks **every individual qualifying, non-outlier technician across all 34 stores** by
  Proficiency %, descending.
- Columns: `Rank | Name | Tech No | Store | Position | Sold Hours | RO Count | Actual Hours
  (Reg+OT) | Proficiency % | Status | Labor Sale | Labor Gross | GP% | ELR`
- Every row **except Rank and Store** is a live formula pointing back at the technician's row on
  their own store sheet, e.g. `='Chevy Lv'!B5`, `='Chevy Lv'!A5`, ... — this is why the store
  sheet's row numbers for qualifying techs must be tracked while building each store sheet.
- Column widths: `A=10.71, B=22, C=14.14, D=20, E=22, F=16.84, G=15.57, H=16, I=13, J=13,
  K=16.43, L=18.29, M=10, N=12`. Freeze `A5`.
- Status column (J) gets the same red/yellow/green conditional formatting as store sheets.

## 8. Capacity sheet
Uses a **different, fixed-assumption methodology** — not derived from the What-If proficiency
targets. All 34 stores, no Main/Express split (stalls and techs treated as one combined pool).

**Fixed assumptions for every store (do not vary by period):**
- Hours Open/Day = 12
- Days Open/Month (stall calc) = 26
- Desired Flag Hrs/Day (tech calc) = 10
- Days Worked/Month (tech calc) = 22
- **Total Stalls per store is a physical constant — carry forward from the prior build, do not
  recompute.** Values (by store sheet name):
  `Acura=20, Audi Henderson=32, Audi Reno=16, Cadillac=45, Cdjr Post Falls=31,
  Chevy Gmc Bullhead=35, Chevy Lv=82, Gmc Prescott=15, Honda Flagstaff=9, Honda Henderson=38,
  Honda North=35, Honda Spokane=25, Hyundai Prescott=16, Hyundai St George=15, Ineos=4,
  Kia Lv=27, Kia St George=14, Lexus=13, Lincoln=26, Lr Henderson=12, Lr Lv=28, Lr Reno=10,
  Mazda=13, Nissan=26, Subaru Lv=43, Subaru Prescott=10, Subaru St George=12,
  Toyota Flagstaff=21, Toyota Henderson=93, Toyota Prescott=36, Toyota Spokane=29, Volvo=11,
  Vw Henderson=22, Vw St George=8`
  If a new store opens or a store's bay count changes, this table needs a manual update from the
  facilities/ops side — it is never derived from Qlik or ADP data.

**Per-store computed values (Tech Count uses every category, not just qualifying groups):**

| Col | Header | Formula/source |
|---|---|---|
| A | Store | hyperlink |
| B | Total Stalls | constant (table above) |
| C | Hours Open/Day | 12 |
| D | ELR (Blended) | `total store Labor Sale / total store Sold Hours` (all techs, hardcoded) |
| E | Stall Labor $/Day | `=B*C*D` |
| F | Days Open/Month | 26 |
| G | Stall Labor Sales/Month | `=E*F` |
| H | GP Retention % | `total store Labor Gross / total store Labor Sale` (all techs, hardcoded) |
| I | Stall-Based Potential Gross/Month | `=G*H` |
| J | Tech Count | count of every tech row on the store sheet (all groups + unmapped, excludes pool codes) |
| K | Desired Flag Hrs/Day | 10 |
| L | Tech Labor $/Day | `=J*K*D` |
| M | Days Worked/Month | 22 |
| N | Tech Labor Sales/Month | `=L*M` |
| O | Tech-Based Potential Gross/Month | `=N*H` |
| P | Actual Monthly Gross | total store Labor Gross, hardcoded |
| Q | Stall Capacity Utilization % | `=IFERROR(P/I,0)` |
| R | Stall Status | `=IF(Q>=0.75,"At/Above Target","Below Target")` |
| S | Tech Capacity Utilization % | `=IFERROR(P/O,0)` |
| T | Tech Status | `=IF(S<0.85,"Below Target",IF(S<=0.95,"In Target","Above Target"))` |

Column widths: `A=21, B=10, C=13, D=11, E=12, F=13, G=13, H=11, I=16, J=9, K=11, L=13, M=13,
N=13, O=16, P=14, Q=13, R=15, S=14, T=15`. Freeze `A6` (one row lower than other sheets, because
row 4/5 are a two-row merged header for this sheet).

> **Caveat to flag to the user whenever the pull period is shorter than a full month:**
> Actual Monthly Gross reflects only the days actually pulled, but Potential Gross assumes a full
> 22–26 day month. A 2-week pull will show every store "Below Target" on both utilization metrics
> even in a strong period — this is expected, not a data error. Consider prorating F/M (Days
> Open/Worked) to the actual period length if Capacity needs to be meaningful for a partial period.

## 9. Capacity Ranker sheet
All 34 stores ranked by **Tech Capacity Utilization % descending**. Every data column is a live
formula pointing at the corresponding `Capacity` sheet row (`=Capacity!J{r}`, etc.) — only Rank
and the Store hyperlink are static.
Columns: `Rank | Store | Tech Count | Total Stalls | Tech Capacity Utilization % | Tech Status |
Stall Capacity Utilization % | Stall Status | Actual Monthly Gross | Tech-Based Potential
Gross/Month | Stall-Based Potential Gross/Month`
Column widths: `A=7, B=22, C=11, D=13, E=14, F=15, G=14, H=15, I=15.51, J=13, K=13`. Freeze `A5`.

---

## 10. Store sheet name / Division value list (fixed order)

| Sheet name | Qlik `Division` value |
|---|---|
| Ineos | INEOS |
| Vw St George | VW ST GEORGE |
| Hyundai Prescott | HYUNDAI PRESCOTT |
| Honda Flagstaff | HONDA FLAGSTAFF |
| Hyundai St George | HYUNDAI ST GEORGE |
| Kia St George | KIA ST GEORGE |
| Lr Reno | LR RENO |
| Lincoln | LINCOLN |
| Mazda | MAZDA |
| Acura | ACURA |
| Subaru Prescott | SUBARU PRESCOTT |
| Lr Henderson | LR HENDERSON |
| Gmc Prescott | GMC PRESCOTT |
| Subaru St George | SUBARU ST GEORGE |
| Cdjr Post Falls | CDJR POST FALLS |
| Volvo | VOLVO |
| Lexus | LEXUS |
| Chevy Gmc Bullhead | CHEVY GMC BULLHEAD |
| Honda Spokane | HONDA SPOKANE |
| Audi Reno | AUDI RENO |
| Vw Henderson | VW HENDERSON |
| Toyota Flagstaff | TOYOTA FLAGSTAFF |
| Toyota Prescott | TOYOTA PRESCOTT |
| Nissan | NISSAN |
| Kia Lv | KIA LV |
| Lr Lv | LR LV |
| Toyota Spokane | TOYOTA SPOKANE |
| Honda North | HONDA NORTH |
| Subaru Lv | SUBARU LV |
| Cadillac | CADILLAC |
| Audi Henderson | AUDI HENDERSON |
| Honda Henderson | HONDA HENDERSON |
| Chevy Lv | CHEVY LV |
| Toyota Henderson | TOYOTA HENDERSON |

This is a fixed layout order carried over from the original build — do not re-sort it
alphabetically; only Proficiency (alphabetical) and Proficiency/Capacity Ranker (ranked) sheets
reorder stores.

---

## 11. Build sequence (recommended)

1. Pull Qlik data for all 34 stores (§1.1) and the ADP roster (§1.2) before writing any Excel code.
2. Merge: validate every ADP `Location Description` maps to the correct `Division`; parse Tech No
   from Position ID; categorize by job title (§2); flag pool codes (§2.1) and unmapped techs (§2.2).
3. Build each of the 34 store sheets first (§4), and **while building, record the row number of
   every qualifying, non-outlier technician** — Tech Ranker needs these as cross-sheet references.
4. Build Proficiency (§5) and Proficiency Ranker (§6) from the same per-store qualifying-tech
   aggregates (hardcoded values, not formulas).
5. Build Tech Ranker (§7) from the row numbers recorded in step 3, globally sorted by Proficiency %.
6. Build Capacity (§8) and Capacity Ranker (§9).
7. Set sheet order and active tab (§3).
8. **Recalculate with `recalc.py` and confirm `total_errors: 0`.** Never ship with unresolved
   formula errors. Spot-check at least one small store and one large store (e.g. Toyota Henderson,
   ~75 techs) after recalculation to confirm row-shifting logic held up at scale.
9. Note the pull-period length to the user, and flag the Capacity caveat (§8) if it's not a full month.
