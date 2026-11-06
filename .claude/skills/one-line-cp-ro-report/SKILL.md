---
name: one-line-cp-ro-report
description: Build the One-Line Customer Pay RO report for a Findlay store - share of customer-pay visits with exactly one CP labor line, by advisor, technician and month, what the single line was, maintenance sold on repair visits, customer-pay work sold on warranty/recall/factory-maintenance visits, and the gross opportunity - from Qlik Closed Repair Orders, with original + "C" split tickets merged into one visit. Use when a GM or manager asks for one-line ROs, one-line customer pay, single-line repair orders, upsell / MPI / menu recommendation follow-through by advisor or technician, maintenance attach on repair visits, or customer pay sold on warranty or recall visits. Do not hand-build it; run the bundled scripts.
---

# One-Line Customer Pay RO Report

Answers: **are advisors and technicians recommending work from the multi-point inspection and the time-and-mileage menu, or just doing what the customer asked for?**

A **one-line CP visit** is a closed visit with exactly one customer-pay line that has labor charged. For example, a customer asks for an oil change and that's all that's sold.

Reference run: Subaru St George, Aug–Oct 2026 (requested by GM Curtis Whitehead):
- 1,900 CP visits, 1,107 one-line (58.3%), 82 "C" split tickets merged.
- A re-run of that period should reproduce these numbers.

## Inputs

1. **Store.** The Qlik `Division` value, e.g. `SUBARU ST GEORGE`. Look it up with `qlik_search_field_values`.
2. **Period.** Default to the last 3 **full closed months**. Never include the current partial month; the user corrected this in the reference run.
3. **Technician names.** Closed ROs only carry `RO_Detail.techno`. Names come from ADP:
   - the ADP "Tech Employee List / Tech Efficiency Report Data" export, or the `adp-pull` skill's `tech-hours`/`roster`;
   - tech number = Position ID characters 4–9 with leading zeros dropped (same rule as the proficiency report).
4. **Technician groups** (line vs. express/lube). ADP usually titles everyone "Service Technician", so infer the group from work mix and confirm with the user. Express techs have many ROs, few hours and mostly oil-change lines.

## Scripts

```bash
S=/home/user/Noah/.claude/skills/one-line-cp-ro-report/scripts
W=<scratchpad>/oneline_<store>
Q="python3 -I $S/qlik_queries.py --division '<DIV>' --start 2026-08-01 --end 2026-10-31"
```

- `qlik_queries.py` writes the Qlik arguments for each bundle. Pass `dimensions`/`measures`/`sort`/`limit` exactly as written. Each bundle prints a NOTE explaining how its output maps into results.json.
- `build_oneline_report.py --results $W/results.json --out "<Store> - One-Line Customer Pay ROs <Period>.xlsx"` builds the workbook and merges split tickets.
- `references/results_template.json` is the results.json schema.

## Step 1: Profile and confirm (RO app `c3efc739-b063-47dd-a4d3-d1c4cac07ad0`)

1. Run `qlik_clear_selections`, then run `profile`. Use `MaxClose` to confirm the data covers the period.
2. Confirm with the user:
   - **Customer-pay labor types.** All `PayType.PayType = 'C'` types count by default. If the user wants retail only, they may exclude CEXT (extended warranty), CSSC (service contract), CPPM (prepaid maintenance) and CSEMP (employee).
   - **Warranty groups:**
     - `--ws`: warranty repairs and recalls (Subaru WS, WR);
     - `--wsc`: service contract (WSC);
     - `--wsm`: factory-paid scheduled maintenance (Subaru WSM; Hyundai Complimentary Care is CCCE, a **C** pay type, so check);
     - `--pdi`: PDI types to exclude.
   - **Op-code groups:**
     - `--oil`: oil-change op codes;
     - `--repair`: repair/diagnostic op codes (INFO is often the catch-all for customer concerns);
     - `--maint`: maintenance op codes.

     The defaults are Subaru St George's. Other stores' op codes differ, so propose lists from `profile` and `ops`, then confirm.

## Step 2: Pull the bundles

```bash
for b in advisor tech ops oil_addon repair_adv repair_tech warranty_adv warranty_tech warranty_median split_list; do
  eval $Q --bundle $b --out $W/$b.json
done
```

Rows with a null `Type` are parts-only CP lines, so ignore them.

## Step 3: Split tickets ("C" ROs). Do not skip this.

Many stores move part of a visit onto a companion RO ending in "C" (e.g. 290486 + 290486C). At Subaru St George it was mostly the customer-pay work sold on a warranty visit. Counted per ticket:
- the warranty visit looks like nothing was sold;
- the C ticket looks like a one-line RO.

Merging the tickets moved warranty-repair conversion from 38% to 46%.

1. `split_list`: list the C tickets. Also check for other suffixes with a dimension `=Right(RO_Detail.ronumber,1)`.
2. Save the C ticket numbers to `$W/c_list.json`, then run:
   ```bash
   eval $Q --bundle split_pairs --c-list $W/c_list.json --out $W/split_pairs.json
   ```
   - Original ROs are stored as **numbers**, so `qlik_select_values` can't select them as strings, and a set-analysis literal list would bloat every measure. That's why the bundle puts the list once into a calculated dimension.
   - Page through it with `qlik_get_chart_data`.
3. Put every row into `results.split_rows`, replacing NaN techs with "". The builder merges each pair into one visit:
   - lines from both tickets are combined;
   - the visit is credited to the original ticket's advisor and close month;
   - the (visit − ticket) difference is applied to every table.

   The mileage counts and the single-line op table stay per ticket.

## Step 4: results.json → workbook → recalc

- Fill in `results.json` per the template. Check that advisor names match across sections.
- Write plain-language op descriptions for `single_line_ops`.
- Optional: `extra_findings` and `definitions_extra`.

```bash
python3 -I $S/build_oneline_report.py --results $W/results.json --out "$W/<Store> - One-Line Customer Pay ROs <Period>.xlsx"
python3 <xlsx-skill>/scripts/recalc.py "$W/<Store> - One-Line Customer Pay ROs <Period>.xlsx"   # status success, 0 errors
```

Never re-save with openpyxl after recalc.

**Tabs:**

| Tab | Shows |
|---|---|
| Summary | KPIs, auto findings, advisor chart |
| By Advisor | One-line rate, oil-change-only rate, average sale, mileage counts |
| Monthly Trend | One-line rate by month, with a chart |
| By Technician | One-line and oil-change-only rates |
| What Was the One Line | The single job on one-line visits |
| Repair Visits | Maintenance attach on repair visits, by advisor and technician |
| Warranty Visits | CP sold on warranty/recall/service-contract/factory-maintenance visits, plus an opportunity block |
| Opportunity | Oil-change add-on value, with editable inputs |
| Definitions | How each number was built, and the limitations |

**Don't commit** results.json or the workbook. They name advisors and technicians.

## Step 5: Deliver

- Clear selections.
- Send the workbook.
- Lead the reply with:
  1. the one-line rate and its trend;
  2. how many one-liners were oil changes;
  3. the advisor spread;
  4. repair-visit maintenance attach;
  5. warranty-visit CP rate;
  6. the gross opportunity.
- Offer the GM email. Include a "What's in the workbook" section and ask the user to confirm labor-type meanings before sending.

## Gotchas (from the reference run)

- **Close-date set analysis needs M/D/YYYY:** `RO_Header.closedate={">=8/1/2026<=10/31/2026"}`. ISO dates return 0 silently.
- **Default selections can come back after an engine reset.** The RO app came back with `Year=2026` / `MonthName=Oct 2026`, and `MonthName` is not overridden by a Year/Month set. Clear selections and check the result's `selections` list is empty.
- **No line order in Qlik.** CDK line letters (A–Z) aren't in Qlik, and `RO_Detail.bookedtime` is the RO open time, not when the line was added. The report can say whether CP work was sold on a visit, but not whether it was added after the inspection.
- **Declined recommendations aren't in the RO data.** A one-line visit may have had recommendations that were declined. Suggest pairing with the store's MPI/menu tool export if the GM wants "not offered" vs. "declined."
- **Technician one-line rates follow dispatch.** Express techs do the oil changes, so compare technicians within their group.
- **Opportunity uses the median add-on, not the average.** The average roughly doubled the result at Subaru St George.
