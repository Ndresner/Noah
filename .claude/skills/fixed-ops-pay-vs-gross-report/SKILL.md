---
name: fixed-ops-pay-vs-gross-report
description: Rebuild or refresh Findlay's Parts & Service Employee Pay-vs-Group workbook with the Qlik gross and workload comparison layer (pay as % of department gross, gross per employee, ROs / sold hours per employee, Service Advisor gross per RO and ROs per day, color-coded outlier flags). Use whenever the user asks to update, refresh, re-run or rebuild the "Earnings - Parts and Service Employees" report, the pay vs gross / comp vs gross report, the gross-adjusted pay comparison, or the service advisor pay-vs-price analysis — even if they only attach a new ADP earnings workbook. Do not hand-build it; run the bundled script.
---

# Parts & Service Pay vs Gross Report

Answers one question: **is a store paying a role more than the group because the store produces more gross or more work, or only because its prices are higher?** It does this by comparing each store's pay for a position to the group on three measures: department gross per employee, pay as a % of department gross, and workload per employee.

## Inputs to collect

1. **Base workbook**: the pay-vs-group workbook with these tabs: Master Summary, Store Ranker, Raw Data, Store & Position Summary, Group Position Summary, and one tab per store. A store tab's A1 is "← Back to Master Summary". It can be either:
   - the plain base build (only the pay-vs-group tabs), or
   - last period's finished output. The script strips the old gross layer and rebuilds it.
2. **Pay period facts** from the user:
   - the pay check date (e.g. `8/15`), and
   - the **last day worked** in that pay period (e.g. `2026-08-09`). This is the hire cutoff and the end date for months employed.
   - Commission employees on draw are trued up through the prior month-end. Treat that as close enough; the user chose not to split pay types.
3. **Gross period**: the full closed months before the pay cutoff (e.g. Jan–Jul 2026 = 7 months).

If the user only has a raw ADP "Earnings - Parts and Service Employees" export, the base workbook has to be built first. That base build is **not** in this script. Tell the user, and ask for last period's finished workbook to use as the template.

## Step 0: Environment

`recalc.py` needs LibreOffice **Calc**. If a two-cell test file times out or LibreOffice says "source file could not be loaded", install Calc:

```bash
apt-get install -y --no-install-recommends libreoffice-calc
```

## Step 1: Qlik pulls (34 stores, one row per store)

Follow CLAUDE.md's Qlik rules. Run `qlik_clear_selections` first on each app. Use `Divison_ADP_PayrollCompanyCode` as the dimension; it matches ADP `Company Code` directly. Year is text. In the examples below, M is the last closed month (7 for July).

**P&L app `bc30d66f-4564-4df0-89a7-368f2c9a931e`**
- Take the YTD field for month M (`Jul_YTD_Amt` for July).
- Use no more than 3 measures per call.
- `Statement={'GM'}` stays selected and can't be cleared. That's expected and harmless.

```
service_gross  = -Sum({<Year={'2026'},Dept={'Service'},StdAccountType={'S','C'}>} Jul_YTD_Amt)
parts_gross    = -Sum({<Year={'2026'},Dept={'Parts'},StdAccountType={'S','C'}>} Jul_YTD_Amt)
bodyshop_gross = -Sum({<Year={'2026'},Dept={'Body Shop'},StdAccountType={'S','C'}>} Jul_YTD_Amt)
```

Sanity check: the YTD field should match the sum of Jan..M `*_MTD_Amt` to within about 0.01%.

**Closed ROs app `c3efc739-b063-47dd-a4d3-d1c4cac07ad0`** (filter `Month={"<=7"}`)

```
cpw_ros      = Count({<Year={'2026'},Month={"<=7"}>*(<RO_Header.totalsalecustomerpay={">0"}>+<RO_Header.totalsalewarranty={">0"}>)} DISTINCT %RO)
sold_hours   = Sum({<Year={'2026'},Month={"<=7"}>} RO_Header.soldhours)
cpw_ro_gross = Sum({<Year={'2026'},Month={"<=7"}>} RO_Header.laborgrosscustomerpay+RO_Header.partsgrosscustomerpay+RO_Header.laborgrosswarranty+RO_Header.partsgrosswarranty)
```

Customs (J01) has no Qlik division. The script shows N/A / "No Gross Data" for it.

## Step 2: Write the Qlik JSON

Write it to the scratchpad. **Do not commit it**; it holds store financials. The template is `qlik_template.json`.

```json
{"period_label":"Jan-Jul 2026","year":"2026","months":7,"pull_date":"9/25/26",
 "stores":{"HUZ":{"division":"AUDI HENDERSON","service_gross":0,"parts_gross":0,"bodyshop_gross":0,
                  "cpw_ros":0,"sold_hours":0,"cpw_ro_gross":0}, "...":{}}}
```

Keys are ADP company codes. Values are period totals, not monthly amounts; the script divides by `months`.

## Step 3: Build, then recalc (mandatory)

```bash
python3 .claude/skills/fixed-ops-pay-vs-gross-report/scripts/add_gross_workload.py \
  --base "<base or last output>.xlsx" --qlik <scratchpad>/qlik.json \
  --out "<scratchpad>/Earnings - Parts and Service Employees - Gross Adjusted.xlsx" \
  --cutoff 2026-08-09 --pay-check 8/15 [--band 0.20]
python3 <xlsx-skill>/scripts/recalc.py "<out>.xlsx" 300     # must report status success, 0 errors
```

Never re-save with openpyxl after the recalc; it strips the cached values.

## Step 4: Verify before delivering

- `recalc.py` reports `total_errors: 0`. The formula count should be around 18.4K.
- Recompute one store's Service Advisor row separately in pandas from Raw Data and the JSON, and check it against the workbook:
  - avg monthly pay = gross pay ÷ ((cutoff − max(hire, 1/1) + 1) / 30.4)
  - gross per advisor = service gross per month ÷ advisor count
  - pay % = pay ÷ gross per advisor
  - ROs per advisor = CP+W ROs per month ÷ advisor count

  The Jan–Jul 2026 reference for Audi Henderson advisors was $16,300 pay, $124,793 gross per advisor, 13.1% (group 16.2%), 124 ROs (group 168), and $1,087 gross per RO (group $392).
- The workbook opens on Master Summary, and Customs shows N/A.

## Step 5: Deliver

Send the file with SendUserFile. It contains employee names and pay, so **never commit or publish the workbook or the JSON**.

## Format spec (what the script produces)

**Raw Data**
- Months employed = (cutoff − max(hire date, 1/1) + 1) / 30.4.
- "Included in Averages" = hired on or before the cutoff. Anyone hired after the last day worked has no legitimate pay in the period; if they do have pay, it's from before a rehire.
- The Master Summary and Store Ranker notes are updated with the cutoff date.

**Qlik Store Data tab** (new, placed after Group Position Summary)
- Inputs (blue): B4 months, B5 outlier band (default ±20%), B6 working days per month = 52×5/12 ≈ 21.7.
- One row per store: Qlik figures in blue, then monthly amounts, tech headcount, and gross per CP+W RO.
- A Group Total row at the bottom.

**Group Position Summary**, columns E–J:
- E Gross Basis
- F Workload Unit
- G Dept Monthly Gross (stores with the role)
- H Group Pay as % of Dept Gross
- I Group Dept Gross per Employee
- J Group Workload per Employee (whole numbers)

Group ratios are weighted by headcount and include only stores with Qlik data. A note sits two rows below Grand Total.

**Store & Position Summary**, columns F–M: helper columns for gross basis, store dept gross, pay %, unit, allocated workload, workload per employee, pay and headcount at stores with gross.

**Master Summary**, columns H–M:
- Service Pay % of Service Gross, **excluding techs** (Service Tech, Express Tech, Shop Foreman, Team Leader). Their pay is already in cost of labor sales, so service gross is net of it.
- Group %, and a High / In Line / Low status.
- The same three columns for Parts.

**Each store tab**
- The original table (A–H) and its totals block are unchanged.
- Starting 5 rows below the totals: "<Store> - Gross & Workload Comparison", columns A–K:
  - A Job Title (formula pointing to the table above)
  - B Gross Basis
  - C/D Store vs Group Dept Gross per Employee
  - E/F Store vs Group Pay as % of Dept Gross
  - G Workload Unit
  - H/I Store vs Group Workload per Employee (whole numbers)
  - J Gross-Adjusted Status
  - K Workload Status
- Then a FLAG COUNT row.
- Then the **Service Advisor Detail** block: header, then "Gross per RO (CP+W labor + parts)" and "ROs per Advisor per Day (5-day week)". Columns: Store | Group | Status (Above Group / In Line / Below Group) | Variance vs Group (%) in `+0.0%` format.
- Then the color key, the band cell, and a method note.
- The Service Advisor row's Workload Unit reads like `ROs | 5.7/day vs 7.7 grp`, and the workload cells stay numeric. **Never** put text into numeric cells, because that breaks the coloring and status logic.

**Definitions**
- Gross basis: Parts roles use parts department gross, Body Shop roles use body shop gross, and everyone else uses service department gross.
- Workload:
  - **Sold Hrs**: Service Tech, Express Tech, Shop Foreman and Service Team Leader. Store sold hours are split evenly across that combined headcount.
  - **ROs**: all other roles. Distinct CP+W ROs (internal excluded) ÷ employees in that role.

**Colors**, applied to C, E and H, the two status columns, and the advisor block (house palette):
- Green `C6EFCE/006100` = better than the group for store economics: more gross or work per employee, or lower pay %.
- Red `FFC7CE/9C0006` = worse than the group.
- Yellow `FFEB9C/9C6500` = in line, within the band.
- Pay % is inverted: higher shows red.

**Layout**: Calibri to match the base. Data is centered on every sheet, except merged titles and notes, text over 45 characters, and the "←" back-links. The workbook opens on Master Summary.

## Known limitations (tell the user when relevant)

- Pay covers about 1/1 to the cutoff; gross and ROs cover full closed months. They're close, not exact to the penny. The user accepted this.
- Rehires with pay from earlier in the year can inflate averages. ADP has only a Hire/Rehire date; the original hire date would fix it.
- The export includes Active and Leave employees only. Terminated employees' pay isn't in the per-position figures.
- Gross per RO and ROs per day are store-level averages across all advisors, not per advisor. Per-advisor figures need a mapping from ADP names to Closed ROs `ServiceAdvisor.Name` ("LAST,FIRST").
- Store Ranker's row order is set by the base build and isn't re-sorted.
