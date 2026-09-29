# Workbook structure reference

## Sheets (39 total, in this order)

1. Proficiency (tab color gold `C9A04B`)
2. Proficiency Ranker (no tab color)
3. Tech Ranker (no tab color)
4. Capacity (tab color gold)
5. Capacity Ranker (no tab color)
6-39. The 34 store sheets, alphabetical by sheet name, tab color navy `1A2744`

Active sheet on open: Proficiency. Freeze panes: `A5` on Proficiency/Proficiency
Ranker/Tech Ranker, `A6` on Capacity/Capacity Ranker, `A5` on every store sheet.

## Colors

- Navy `1A2744`, Gold `C9A04B`, White `FFFFFF`
- Status green: font `006100`, fill `C6EFCE`
- Status yellow: font `9C6500`, fill `FFEB9C`
- Status red: font `9C0006`, fill `FFC7CE`
- What-if target-entered flag: fill `FFFF00` (yellow), no font change
- Capacity sheet borders: thin, color `CCCCCC` (light gray)
- Font throughout: Arial

## Store -> Division -> Sheet-name mapping

Qlik's `Division` field values are the store name in upper-snake-case with spaces (e.g.
`"CHEVY GMC BULLHEAD"`); the sheet name is simply `Division.title()` (e.g.
`"Chevy Gmc Bullhead"`). The ADP roster's `Location Description` values are a *different,
looser* naming convention and need the explicit `LOC_TO_DIVISION` mapping in
`build_data.py` -- notably `"MOTOR COMPANY"` in the roster is Chevy Gmc Bullhead's ADP
location name, and doesn't resemble the store name at all. If a new store is added, add it
to both `LOC_TO_DIVISION` and `DIVISION_TO_SHEET`.

## Store sheet row layout (dynamic per store -- exact rows vary)

1. Title: `"{Store} -- Technician Proficiency -- {period}"`, merged A1:L1
2. Back link: `=HYPERLINK("#'Proficiency'!A1","<< Back to Proficiency")`, merged A2:L2
3. (blank in A:M) + "Technician Proficiency What-If Scenario" header at O3, merged O3:S3
4. Column headers (bold white on navy): Tech No | Name | Position | Sold Hours | RO Count |
   Avg Hrs/RO | Actual Hours (Reg+OT) | Proficiency % | Status | Labor Sale | Labor Gross |
   GP% | ELR | (N blank) | Current Proficiency % | Target Proficiency % | Proficiency
   Increase | Additional Sold Hours | Additional Labor Gross
5+. For each category present, in this order -- Service Technician, Express Technician,
   Service Team Leader, Shop Foreman, Body Shop Tech, Body Shop Painter, Body Shop Helper,
   Detailer, Tinter (skip any category with zero technicians at this store):
   - One row per technician, sorted by Proficiency % descending (see lessons-learned #2)
   - GROUP SUBTOTAL row (gold fill)
   - Divider row: navy-filled merged `A:M` + `O:S`, UNLESS this is the last category (then
     plain blank row instead)
   TOTAL / AVERAGE row (navy fill, white bold) -- sums each category's GROUP SUBTOTAL row
   Blank row
   STORE-WIDE WHAT-IF SCENARIO block (see lessons-learned #7 for the proficiency cell)
   Blank row
   EXCLUDED / POOLED CODES block (only if any unmapped codes exist for this store)

### Per-technician-row formulas (row `r`)
```
F{r}  Avg Hrs/RO            =IFERROR(D{r}/E{r},0)
H{r}  Proficiency %         =IF(G{r}="N/A","N/A",IFERROR(D{r}/G{r},0))
I{r}  Status                =IF(H{r}="N/A","No Hours Data",IF(H{r}>=1,"Above Target",IF(H{r}>=0.85,"On Target","Under Target")))
L{r}  GP%                   =IFERROR(K{r}/J{r},0)
M{r}  ELR                   =IFERROR(J{r}/D{r},0)
O{r}  Current Proficiency % =H{r}
Q{r}  Proficiency Increase  =IF(P{r}="","",IF(G{r}="N/A","N/A",P{r}-O{r}))
R{r}  Additional Sold Hours =IF(P{r}="","",IF(G{r}="N/A","N/A",G{r}*P{r}-D{r}))
S{r}  Additional Labor Gross=IF(R{r}="","",IF(R{r}="N/A","N/A",IFERROR(R{r}*M{r}*L{r},0)))
```
(D/E/G/J/K are hardcoded values: Sold Hours, RO Count, Actual Hours, Labor Sale, Labor
Gross. P is a blank yellow-filled input cell for the user's what-if target.)

### GROUP SUBTOTAL row formulas (category rows `cs`-`ce`, this row `sub_r`)
```
D  =SUM(D{cs}:D{ce})          G  =SUMIF(G{cs}:G{ce},"<>N/A")
E  =SUM(E{cs}:E{ce})          H  =IFERROR(D{sub_r}/G{sub_r},0)
F  =IFERROR(D{sub_r}/E{sub_r},0)
J  =SUM(J{cs}:J{ce})          K  =SUM(K{cs}:K{ce})
L  =IFERROR(K{sub_r}/J{sub_r},0)   M  =IFERROR(J{sub_r}/D{sub_r},0)
O  =H{sub_r}                  P  =IFERROR((D{sub_r}+R{sub_r})/G{sub_r},0)
Q  =IFERROR(P{sub_r}-O{sub_r},0)
R  =SUM(R{cs}:R{ce})          S  =SUM(S{cs}:S{ce})
```

### TOTAL / AVERAGE row (this row `tot_r`, subtotal rows list `subtotal_rows`)
Same shape as GROUP SUBTOTAL but each SUM/SUMIF becomes a plain `+`-chain across every
category's subtotal row, e.g. `D{tot_r} = D{sr1}+D{sr2}+...`.

### Store-Wide What-If Scenario block
```
Current Store Proficiency %   -- see lessons-learned #7 (computed value, not a formula)
Target Store Proficiency %    -- blank yellow input cell
Proficiency Increase          =IF(E{tgt}="","",E{tgt}-E{cur})
Additional Sold Hours          =IF(E{tgt}="","",IFERROR(G{tot_r}*E{tgt}-D{tot_r},0))
Additional Labor Gross Profit  =IF(E{addhrs}="","",IFERROR(E{addhrs}*M{tot_r}*L{tot_r},0))
```

## Proficiency / Proficiency Ranker sheets

Hardcoded computed values (not formulas), one row per store: Technicians, Total Sold
Hours, Total Actual Hours, Avg Proficiency %, Status, Total Labor Sale, Total Labor Gross,
Avg Labor GP% -- aggregated only from Service Technician/Express Technician/Service Team
Leader/Shop Foreman rows that pass `qualifies_for_prof_subset()` (excludes outlier records:
STL/Shop Foreman <=20% proficiency, Service/Express Tech <=2%). Proficiency is sorted
alphabetically by store; Proficiency Ranker is the same data sorted by Avg Proficiency %
descending.

## Tech Ranker

One row per technician in Service Technician/Express Technician/Service Team
Leader/Shop Foreman (no outlier exclusion, unlike Proficiency), ranked by Proficiency %
descending across all 34 stores, with live formulas referencing each technician's actual
row on their store sheet (e.g. `='Chevy Lv'!B5`).

## Capacity / Capacity Ranker

See `references/lessons-learned.md` items 3, 5, 6 for merges/borders/number-formats.
Capacity is pay-period based: period dates in `B3:C3`, Days Open (F) = Mon-Sat and Days
Worked (M) = Mon-Fri via `NETWORKDAYS.INTL`, Actual Period Labor Gross (P) linked to each
store sheet (TOTAL row K + EXCLUDED / POOLED CODES col E). Total Stalls is carried forward
from the prior workbook (see SKILL.md Step 4); everything else (ELR, GP Retention %, Tech
Count, and all derived Stall-Based/Tech-Based potential-gross formulas) is recomputed from
this period's data.
Tech Count combines every technician category at the store, including unmapped/pooled
codes -- broader than Tech Ranker's population.
