---
name: fixed-ops-gross-records-report
description: Build Findlay's monthly Fixed Ops Gross Profit Records report ("<Month>_<Year>_Fixed_Ops_Gross_Records" .xlsx + 3-page PDF) and the recognition email that goes with it. Each store's month of Service, Parts and Combined gross is ranked against its all-time monthly history from Qlik Fixed Ops Accting TB. Use whenever the user asks for the fixed ops records report, gross records, records review, all-time records, "who set records last month", or the records recognition email, even when they don't attach last month's file. Do not hand-build it; run the bundled script.
---

# Fixed Ops Gross Profit Records Report

Answers one question: **which stores set all-time monthly gross records last month, and where did the group land in company history?** It ranks each store's month against the store's full monthly history (Jan 2021 onward) for Service, Parts and Service + Parts. It also ranks the group total.

Script: `scripts/build_records.py`. It reads the prior month's finished workbook. That workbook's `Data` tab holds the full monthly history. The script appends the new month, applies post-close restatements, writes both files and runs `recalc.py` last.

The Sep 2026 build was approved as the reference. The script was checked by rebuilding Aug 2026 from July history, and every formula and value matched the delivered August workbook.

## Inputs

1. **Last month's finished workbook**, e.g. `August_2026_Fixed_Ops_Gross_Records.xlsx`. Ask for it if it isn't attached. Its `Data` tab is the only copy of the 2021+ history; nothing in this repo holds it.
2. **Qlik pull** (Step 1) for the new month, plus a re-pull of the current year's earlier months to catch restatements.

## Step 1: Qlik pull

App **Fixed Ops - Accting TB** `e077f027-fc9f-4f03-adf3-14ce059d8b3c`. Run `qlik_clear_selections` first.

- Dimension: **`Store`**. This is the one exception to the "filter by Division" rule: the report keys on Store names like `ACURA FA-5`, which match the `Data` tab exactly. It returns exactly the 34 stores. Don't filter; just dimension.
- `Year` is numeric in this app: `Year={2026}` works.
- Amounts sit in per-month columns (`Jan_MTD_Amt` … `Dec_MTD_Amt`) on rows for each Year. Gross is the sign-flipped sum (app master-measure convention):
  - Service: `Sum({<Year={2026},Family={'FIXED GROSS'},Class={'SERVICE'}>} Sep_MTD_Amt*-1)`
  - Parts: `Sum({<Year={2026},Family={'FIXED GROSS'},Class={'PARTS'}>} Sep_MTD_Amt*-1)`
- Use `qlik_create_data_object` with chartType `table`, Store sorted alphabetically, **≤3 measures per call**. 34 rows fit in one call.
- Pull the new month (Service + Parts), then every earlier month of the current year (Service + Parts). Five or six calls in parallel cover Jan–Aug.

Save it as two CSVs in the scratchpad:
- `new.csv`: `Store,Service,Parts` for the new month, rounded to 2 dp.
- `restate.csv`: `Store,Year,MonthNum,Service,Parts`, holding **only** the rows where the re-pull differs from the prior workbook's `Data` tab by more than $0.005. Compare in Python against the prior workbook. Sep 2026 found 7 July rows; August found a June reclass. Tell the user what changed and the size of the largest moves.

The month is complete only once it has closed. Pull after the first week of the following month, and say in the report that the following month is excluded as partial.

## Step 2: Build

```bash
python3 .claude/skills/fixed-ops-gross-records-report/scripts/build_records.py \
  --prior <last month workbook>.xlsx --new new.csv --restate restate.csv \
  --notes notes.json --year 2026 --month 9 --pull-date 2026-10-08 --outdir <scratchpad>/out
```

Outputs `<Month>_<Year>_Fixed_Ops_Gross_Records.xlsx`, `.pdf` and `_summary.json` (group ranks, record lists, generated text). Recalc must report `total_errors: 0`.

`notes.json` overrides generated text. Always write these two:
- `exec_summary`: a written paragraph, in the August/September style. First run once without it, read `_summary.json`, and pull these facts with a short Python check against the built Data:
  - the group's rank and the months above it
  - the change vs the prior month and vs the same month last year (and whether it's the best such month on record)
  - who set Combined / Service / Parts records, with amounts vs prior record
  - repeat record-setters from last month
  - the biggest month-over-month movers
- `restatement_note`: which months were re-pulled, which changed, and the largest adjustments. Use plain `-` for negatives, not `−`.

Optional: `near_records_p1` and `near_records_p2`. By default these list every non-record #2/#3 (p1: #2 any category or #3 combined; p2: #3 Service/Parts). Also optional: `store_openings` and `cdk_note`.

Then render the PDF pages (`pdftoppm -r 80 -png`) and look at them before sending.

## Layout (matches the approved template; the script encodes it)

**Workbook**, 4 tabs, all live formulas against `Data`:
- `<Month> <Year> vs Records`: Store, then Service / Parts / Service + Parts blocks of Month $, Rank (COUNTIFS strictly-greater + 1), Record $ (SUMPRODUCT(MAX)). Sorted by the month's combined gross, with a GROUP TOTAL row. Navy `1B2A41` headers. Record cells are green `C6EFCE`/`006100` bold; this highlight is static, a snapshot at pull. Freeze `B5`.
- `Data`: Store, Year, MonthNum, Month, Service, Parts, Total `=ROUND(E+F,2)`. Prior history in its original order, with the new month appended at the bottom. Pre-opening $0 months stay; they never rank above a real month.
- `All-Time Records`: record $ and the month it was set (CHOOSE/SUMPRODUCT) per category, banded `F2F2F2`.
- `Methodology`: source, measure convention, window, partial-month exclusion, restatement note, rank rule, CDK and store-opening notes.

**PDF** (reportlab, letter, Helvetica, 3 pages):
- Page 1: navy header band, 4 KPI cards (group Service, Parts, Combined with all-time rank; new combined records count), Executive Summary, "Records Set in <Month>" table (Category / Store / value / prior record / prior record month), near-records note.
- Page 2: all 34 stores, value / RK / record per category, REC in green, GROUP TOTAL with group ranks.
- Page 3: Methodology & Notes bullets.

Store names in the PDF tables are `store.title()` (e.g. "Lr Reno Lrr-31"), as in the original template. Narrative text uses short names.

## Step 3: Recognition email

After the report, draft the group recognition email in the format of `references/email_format.md`. Use full brand/city names: Land Rover Reno, Kia Las Vegas, Kia St. George, Chevy Las Vegas, CDJR Post Falls. Rules:
- The opening line states where the month landed in Findlay history. Be accurate: a #7 month is not "record-setting". Lead with the strongest true framing, such as best September ever or up X% on last year.
- "Additional Service/Parts Record Setters" lists only stores **not** already in the combined list. Note stores that set records in all three categories inside the combined section.
- The outstanding-performances bullets cover every #2 and #3, matching the report's near-records notes.
- Close with the next month's name.
- Deliver it as ready-to-send text. Never send it.

## Data notes
- June 2024 is depressed group-wide (CDK outage). It doesn't affect maxima.
- Store openings (history starts here): LR Henderson Jul 2021, Hyundai Prescott Mar 2022, Nissan Henderson Nov 2023, Honda Spokane Aug 2024, Ineos Nov 2025. A brand-new store isn't in the prior `Data` tab, and the script stops and asks you to backfill it.
- The outputs hold store-level financials. Keep them in the scratchpad and never commit them.
