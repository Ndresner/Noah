# CLAUDE.md

This file is auto-loaded by Claude Code at the start of every session in this repo. It gives context on Findlay Automotive Group Fixed Ops reporting work — read this before starting any task in this repo.

## Who this is for

Noah, Fixed Operations analytics/reporting at Findlay Automotive Group (34-store dealership group: NV, AZ, ID, WA). Builds recurring and ad hoc Excel/PDF reports pulling from Qlik Cloud.

## Data sources

- Qlik Cloud: `findlayauto.us.qlikcloud.com` (via Qlik MCP connector)
- Primary Qlik apps:
  - Closed Repair Orders — `c3efc739-b063-47dd-a4d3-d1c4cac07ad0`
  - Profit & Loss — `bc30d66f-4564-4df0-89a7-368f2c9a931e`
  - Fixed Ops Accounting TB — `e077f027`
  - Productivity — `948ab4f4`
- Some reports use non-Qlik sources: ADP exports, RO Detail exports, Tech Pay Master file, vendor pricing spreadsheets.

## Qlik MCP rules — always follow

- Run `qlik_clear_selections` first, every time.
- Filter stores by `Division`, never `Store` (Store carries company-code suffixes).
- Year/Month need quoted text in set analysis: `Year={'2026'}`, `Month={'Jun'}`.
- `qlik_create_data_object` caps at 50 rows — use `qlik_get_chart_data` with offset pagination for larger pulls.
- Auth failures on first attempt are normal — retry once immediately, don't stop.
- P&L app: `Statement={'GM'}` persists as a default selection and can't be cleared — this is harmless, ignore it.
- Keep P&L calls to ≤3 measures per request or expression errors follow.
- P&L sign convention: Sales (S) negative, Cost (C)/Expense (E) positive. Gross = −(S+C). Net = Gross − E.
- VIN retention/gain-loss logic: P()-based set analysis needs Division hardcoded per store — 34 separate calls, no exceptions (P() doesn't inherit dimensional row context).

## Excel build rules — always follow

- Run `recalc.py` as the final step on every `.xlsx` build. No exceptions.
- Never reload-and-resave with openpyxl after recalc — it strips cached formula values. If you need to patch something post-recalc (e.g. column widths), edit `sheetN.xml` directly via zipfile/ElementTree.
- Column width audits: skip merged multi-column rows, add +3 padding, enforce min-width floors on short-text columns (e.g. Status = 16).

## PDF build rules

- Reportlab renders Δ, ×, → as blank glyphs. Use plain ASCII substitutes instead.
- Ledger build scripts live in `/home/claude/ledger/` and reset each session — rerun the full pipeline from scratch each month, don't assume state persists.

## Active reports — key specs

**Technician Proficiency Workbook** (36 sheets, all 34 stores)
Structure: Proficiency → Proficiency Ranker → Tech Ranker → Capacity → Capacity Ranker → 34 store sheets. Proficiency is the default active sheet.
Tab colors: Proficiency/Capacity = gold `FFC9A04B`; store sheets = navy.
Proficiency/Proficiency Ranker filter: Service Technician, Express Technician, Service Team Leader, Shop Foreman only. Exclude outliers: STL/Shop Foreman ≤20%, Service/Express Tech ≤2% proficiency.
Tech Ranker excludes Body Shop, Detailer.
Proficiency thresholds: ≥100% Above Target (green) · 85–99.99% On Target (yellow) · <85% Under Target (red). Colors: green `C6EFCE/006100`, yellow `FFEB9C/9C6500`, red `FFC7CE/9C0006`.
Capacity Status uses same colors but inverted logic: Above Target = yellow (overcapacity warning, not an achievement).
GP% thresholds: ≥80% green, 75–79.99% yellow, <75% red.
Store sheet category order: Service Tech → Express Tech → STL/Shop Foreman → Body Shop → Detailer/Detail Manager → Tinter → Unmapped, with grey band rows (`D9D9D9`) between groups and gold "GROUP SUBTOTAL" rows.
Tech Ranker column widths (A–N): 10.71, 22, 14.14, 20, 22, 16.86, 15.57, 16, 13, 13, 16.43, 18.29, 10, 12.
Capacity methodology (combined pool, no Main/Express split, **pay period only** -- dates in Capacity B3:C3):
- Stall-Based = Stalls × 12 hrs/day × ELR × Mon–Sat days in period × GP Retention %
- Tech-Based = Techs × 10 flag hrs/day × ELR × Mon–Fri days in period × GP Retention %
- Actual (col P) = store-wide period Labor Gross linked from each store sheet (TOTAL row K + pooled codes E), not a monthly P&L figure.
Capacity cell M5 header text: "Days Worked/ Period" (space before Period is intentional, don't "fix" it).
Build it with the repo skill `.claude/skills/findlay-technician-proficiency-report/` (has the pay-period capacity logic).

**Service Customer Gain/Loss Report**
TTM vs prior TTM. Source: Closed ROs, PayType C+W, totalsale>0, VIN-level. 3 sheets: Summary, Monthly Trend, Pay Type Split.
Known caveat: CP + Warranty net changes won't sum exactly to combined net (VIN dedup).

**Fixed Ops Findlay Ledger** (5-page PDF)
Pages 1–2: P&L ledger, 34-store MoM/YoY/YTD with Gross vs Expense lever attribution.
Pages 3–5: Driver Analysis (Pay Type Gross, GP% decomposition w/ Volume/Rate effect, headcount).

**Technician Pay Rate Audit**
Source: RO Detail export + Tech Pay Master (not Qlik).
Rules: SC=A rows → Sold Hours × Hourly Rate. Blank SC → Skill-A % × Sold Amount.
Flat-fee exceptions: PPM, SMOG/SMOGU. Negative ISP credit lines flagged separately.
4 sheets: Technician Summary, Technician Rates, Original Data, Variance Detail (Python-filtered, not FILTER spill).

**Monthly Alignment Benchmark Report**
3 tabs: Monthly, Yearly (YTD), Raw Data. Store column width 32, Arial. Benchmark = 15%. CF: green ≥15%, yellow 10–15%, red ≤10%.

**GL Account 74074 Audit**
Shop supplies credit-only account; flags incorrect debit postings. Store Summary + 11 per-store detail tabs, sorted by $ mischarged.

**Parts & Service Pay vs Gross Report** (Earnings - Parts and Service Employees - Gross Adjusted)
Skill: `.claude/skills/fixed-ops-pay-vs-gross-report/` — always use it (script + Qlik pull spec). ADP pay vs group, plus Qlik gross/workload layer: pay % of dept gross, gross per employee, ROs or sold hrs per employee, Service Advisor gross per RO and ROs/day, ±20% outlier colors. Hire cutoff = last day worked in the pay period. Contains employee pay — never commit the workbook or Qlik JSON.

**Other builds on file:** SWE (Service Website Effectiveness) Analysis, Used Tire Disposal Cost Analysis, Used Oil Reimbursement Comparison, Voice AI Deployment Tracker, Opcode Legend Reformatter, Pay Plan Calculator, Work Type Mix Report, Honda Henderson/Honda North Productivity Reports, BOB Stats Summary formatting.

## Full narrative reference

See `Findlay_FixedOps_Reports_Reference.md` in this repo for complete methodology notes, caveats, and historical detail on every report above — this file is the condensed/instruction version for quick session loading.
