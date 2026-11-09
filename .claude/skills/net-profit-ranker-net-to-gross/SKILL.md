---
name: net-profit-ranker-net-to-gross
description: Add the "Net to Gross" ranker tab to Findlay's monthly Fixed Ops Net Profit Ranker reports (Fixed/Combined, Service, Parts) and draft the distribution email. Use whenever the user asks to refresh, update or rebuild the Net Profit Ranker, the net to gross ranker, net-to-gross %, or "do the net profit ranker for <month>", even if they only attach the three Net Profit Ranker workbooks. Pulls department gross from the Qlik P&L app, reconciles it to the workbook's net, and runs the bundled script. Do not hand-build the tab.
---

# Net Profit Ranker: Net to Gross tab

The user supplies the three finished **Net Profit Ranker** workbooks every month. Each one has a single `Total` tab: 34 stores, ranked by YTD net, with the months listed across the top. This skill adds a second tab, **Net to Gross**, to each workbook. The new tab has the same look and ranks the stores by net ÷ gross.

| File | Scope | `--` flag |
|---|---|---|
| Net Profit Ranker - Fixed | Service + Parts combined | `--fixed` |
| Net Profit Ranker - Service | Service dept | `--service` |
| Net Profit Ranker - Parts | Parts dept | `--parts` |

"Fixed" is **Service + Parts only**. It is not the whole store, and it leaves out Body Shop. This was confirmed against the P&L: Toyota Henderson Sep-26 Service net −605,568 plus Parts net −322,950 gives $928,519, which matches the Fixed workbook to the penny.

## Inputs

1. The three Net Profit Ranker workbooks for the month (any subset works).
2. Nothing else. Gross comes from Qlik. Net comes from the workbook and is checked against Qlik.

Put each upload in its own folder in the scratchpad (for example `<scratch>/in/`). The output folder must be a different folder, because the script refuses to overwrite its inputs.

## Step 1: Qlik pull (P&L app `bc30d66f-4564-4df0-89a7-368f2c9a931e`)

Follow CLAUDE.md's Qlik rules:
- Call `qlik_clear_selections` first.
- `Statement={'GM'}` stays selected and can't be cleared. That's harmless.
- Use no more than 3 measures per call.
- Division is the store key. Its values match column B of the workbook exactly.

**Other users' selections can show up mid-session** (seen Nov-2026: Division, Month and Dept were selected). Always put `Division=,Month=` in the set so selections are ignored, and check that each result has 34 rows.

Use `Dept` (the standardized field). Do **not** use `departmentname`, which is spelled differently per store and misses about 9 stores. Use `Line Desc` P&L lines and flip the sign (the app stores credits as negative):

```
gross (month m)  = -Sum({<Division=,Month=,Year={'YYYY'},Dept={'<D>'},[Line Desc]={'Total Departmental Gross Profit'}>} <Mon>_MTD_Amt)
gross YTD        = same with <LastMon>_YTD_Amt
gross prior year = same with Year={'YYYY-1'} and Dec_YTD_Amt
net YTD          = -Sum({<Division=,Month=,Year={'YYYY'},Dept={'<D>'},[Line Desc]={'Net Profit Or Loss'}>} <LastMon>_YTD_Amt)
net prior year   = same with Year={'YYYY-1'} and Dec_YTD_Amt
```

- `<D>` is `Service` or `Parts`. Pull each department separately, with `Division` as the only dimension and 3 measures per call. Through month M, that is M + 4 measures per department, so 10 calls in October.
- Month field prefixes: `Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec`.
- What doesn't work:
  - `accounttype`/`StdAccountType` sums under `Statement=GM` double-count, because accounts are mapped to several statement lines.
  - The `Line Desc` values `'Net Income'` and `'Total Gross Profit'` return 0 under GM.

Write the results to `<scratch>/qlik_ntg_<YYYY-MM>.json`. Use the shape in `references/qlik_ntg.example.json`: for each store, `service` and `parts`, each holding `gross_months[M]`, `gross_ytd`, `gross_py`, `net_ytd`, `net_py`. **Never commit this JSON**, because it holds store financials. The script builds Fixed as Service + Parts.

## Step 2: Build, then recalc (mandatory)

```bash
S=.claude/skills/net-profit-ranker-net-to-gross/scripts
python3 -I $S/build_ntg.py --data <scratch>/qlik_ntg_YYYY-MM.json --outdir <scratch>/out \
  --fixed "<scratch>/in/Net Profit Ranker - Fixed.xlsx" \
  --service "<scratch>/in/Net Profit Ranker - Service.xlsx" \
  --parts "<scratch>/in/Net Profit Ranker - Parts.xlsx"
for f in <scratch>/out/*.xlsx; do python3 <xlsx-skill>/scripts/recalc.py "$f"; done   # each: status success, 0 errors
```

Never re-save with openpyxl after recalc.

**The script stops and builds nothing** if any of these fail:
- The stores in the JSON don't match the stores in the workbook.
- A store's monthly gross doesn't sum to its YTD gross. This catches typing errors in the JSON.
- The workbook's YTD net or prior-year net is off from Qlik by more than $1.

A net miss usually means the workbook was built from different data (for example, before month-end close), or you pulled the wrong `Dept`. Find the cause; don't loosen the tolerance.

## What the tab contains

- Columns: Store | each month | `YYYY YTD` | `YYYY-1 FINAL` | `YTD +/- PY`. AVG, PACE and PACE % CHG are dropped on purpose, because a pace of a ratio would just repeat YTD.
- Month % = that month's net ÷ that month's gross.
- YTD % = YTD net ÷ YTD gross. It is **not** an average of the monthly %.
- `YTD +/- PY` is the change in percentage points.
- Total row = group net ÷ group gross.
- Stores are ranked by YTD %, highest first. A store with 0 gross gets a blank and is ranked last.
- Styling is copied cell by cell from the `Total` tab:
  - Fonts, widths, hidden columns and rows, frozen panes at C6, gridlines off, and the same conditional formatting rules. The right-edge border rule is re-pointed to `YTD +/- PY`.
  - Negative values use red font `D90613`, set per cell by sign. The source sets red per cell too; it is not in the number format.
- The `Total` tab's values are not changed. It stays first and is the tab that opens.

Things that were learned the hard way and are already handled in the script:
- **CF priority ties.** The source has two rules at priority 9. LibreOffice (recalc.py) drops one of a tied pair on save, which removed the right border on PACE % CHG in the Total tab. The script gives every rule a unique priority on every sheet, in the same order.
- **What the LibreOffice resave changes.** It writes the number formats in an equivalent form (`\$#,##0;"-$"#,##0`) and rounds widths to 2 decimals. Both look the same. No patch is needed.

## Step 3: Check, then deliver

- Render each `Net to Gross` tab and look at it. Use LibreOffice → PDF → `pdftoppm` on a throwaway copy that has landscape fit-to-page; never touch the deliverable. Look for: red only on negatives, a border on the right edge, shading on YTD/FINAL, and a bold Total row.
- Give the user the group YTD vs prior year for each file, the top 3 and bottom 3 stores, and anything notable. Notable examples: a store that loses money in one department but not the other, or a big drop. Flag tiny denominators (for example INEOS Parts 2025 gross of about $16K) as noise.
- Send the three workbooks with SendUserFile. Do not commit them.
- Draft the email from `references/email_template.md`. Fill in the month and year. Add a short "what stands out" note only if the user asks for one.
