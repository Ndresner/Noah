# Findlay Automotive — Fixed Ops Reporting Reference

This document consolidates the methodology, structure, and key parameters for all Fixed Operations reports and projects built to date. Intended as Claude Project knowledge so context carries forward without re-explaining from scratch each session. Replace the prior version of this file with this one — it includes everything through early August 2026.

**Data sources:** Qlik Cloud at findlayauto.us.qlikcloud.com (via Qlik MCP connector), plus ad hoc Excel/ADP/vendor file uploads. Primary Qlik apps: Closed Repair Orders (appId c3efc739-b063-47dd-a4d3-d1c4cac07ad0), Profit & Loss (appId bc30d66f-4564-4df0-89a7-368f2c9a931e), Fixed Ops Accounting TB (appId e077f027), Productivity (appId 948ab4f4).

---

## Cross-Report Qlik MCP Patterns

- Always run `qlik_clear_selections` first.
- Filter stores using the **Division** field, not Store (which carries company-code suffixes).
- Year/Month values require quoted text in set analysis: `Year={'2026'}`, `Month={'Jun'}`.
- `qlik_create_data_object` caps at 50 rows — use `qlik_get_chart_data` with offset pagination for larger pulls.
- Auth failures on first attempt are normal — one immediate retry resolves it.
- P&L app: `Statement={'GM'}` persists as a default selection and can't be cleared — harmless, ignore it.
- Keep P&L calls to ≤3 measures per request to avoid expression errors.
- P&L sign convention: Sales (S) booked negative, Cost (C) and Expense (E) positive; Gross = −(S+C); Net = Gross − E.
- P()-based set-analysis intersection (VIN retention logic) requires **Division hardcoded per store** — 34 separate calls — since P() doesn't inherit dimensional row context.
- Reportlab gotcha: Δ, ×, → render as blank glyphs — use plain ASCII substitutes.
- After any `.xlsx` build: run `recalc.py` as the **final** step, always. Never reload-and-resave with openpyxl after recalc — it strips cached formula values.

---

## 1. Technician Proficiency Workbook (all 34 stores)

**Cadence:** Rebuilt periodically for rolling date windows (most recent: 6/25/26–7/24/26).
**Source:** Qlik Closed ROs pulls + ADP roster matching for actual Reg+OT hours (not assumed schedule hours).

**Structure (36 sheets):** Proficiency → Proficiency Ranker → Tech Ranker → Capacity → Capacity Ranker → 34 individual store sheets. Proficiency sheet is the default active sheet on open.

**Tab colors:** Proficiency & Capacity = gold (FFC9A04B). Store sheets = navy.

**Filtering:**
- Proficiency / Proficiency Ranker: Service Technician, Express Technician, Service Team Leader, Shop Foreman only. Outlier exclusions: STL/Shop Foreman ≤20%, Service/Express Technician ≤2% proficiency.
- Tech Ranker: excludes Body Shop, Detailer positions.

**Conditional formatting:**
- Proficiency Status: Under Target = red (FFC7CE/9C0006), On Target = yellow (FFEB9C/9C6500), Above Target = green (C6EFCE/006100).
- Capacity Status (utilization-based): Below Target = red, In Target = green, Above Target = yellow (overcapacity signal, not an achievement).
- Tech Ranker Status CF must match Proficiency Ranker exactly.

**Proficiency thresholds:** ≥100% Above Target · 85–99.99% On Target · <85% Under Target.
**GP% thresholds:** ≥80% green · 75–79.99% yellow · <75% red.

**Store sheet layout:** Category order fixed — Service Tech → Express Tech → Service Team Leader/Shop Foreman → Body Shop → Detailer/Detail Manager → Tinter → Unmapped. Grey band rows (D9D9D9, unlabeled) separate groups. Subtotals labeled "GROUP SUBTOTAL" (gold). Grand total row navy.

**Tech Ranker fixed column widths:** A=10.71, B=22, C=14.14, D=20, E=22, F=16.86, G=15.57, H=16, I=13, J=13, K=16.43, L=18.29, M=10, N=12.

**Capacity methodology (combined pool, no Main/Express split):**
- Stall-Based Potential Gross = Total Stalls × 12 hrs/day × ELR × 26 days/month × GP Retention %
- Tech-Based Potential Gross = Tech Count × 10 flag hrs/day × ELR × 22 days/month × GP Retention %
- Corrected lift counts: Hyundai Prescott=16, Audi Henderson=32, Kia LV=27, Kia St George=14, Toyota Henderson=93, VW St George=8.
- Capacity M5 header: "Days Worked/ Month" (space before "Month" is deliberate).

**Column-width audit rule:** skip merged multi-column rows; add +3 padding; enforce min-width floors for short-text columns like Status (16).

Full technician rosters (techno → Name(Position) per store, all 34 stores) are on file from prior builds.

---

## 2. Service Customer Gain/Loss Report

**Window:** TTM vs. prior TTM (most recent: 7/11/25–7/10/26).
**Source:** Closed Repair Orders app, PayType C+W, totalsale > 0, VIN-level.

**Methodology:** P()-based set analysis for VIN intersection (Retained VINs); Division hardcoded per store — 34 separate calls.

**Structure (3 sheets):** Summary (34 stores, worst net change first), Monthly Trend (13-point rolling series + LineChart), Pay Type Split.

**Caveat:** CP + Warranty net changes don't sum exactly to the combined net figure, due to VIN dedup across pay types.

**July 2026 baseline (for trend comparison):** Group net −9,255 VINs (−3.1%), Retention 56.3%. Warranty attrition (−8.0% YoY) far outpacing CP (−1.1%). Worst decliners: Honda Henderson (warranty-driven), Honda Spokane, CDJR Post Falls, Nissan (CP-driven). Weakest retention: Kia LV 42.4%, CDJR Post Falls 40.7%.

---

## 3. Fixed Ops Findlay Ledger (monthly PDF, reportlab)

**Structure:** 5 pages.
- Pages 1–2: P&L-based ledger (34-store, MoM/YoY/YTD, top-5/bottom-5 movers with Gross vs. Expense lever attribution).
- Pages 3–5: Driver Analysis — Pay Type Gross breakdown (from Closed ROs), GP% decomposition by Service/Parts/Body Shop with Volume/Rate effect (from Fixed Ops Accounting TB app), headcount analysis (from Productivity app).

**Build note:** Build scripts live in `/home/claude/ledger/` and reset each session — the full pipeline is re-run from scratch every month.

---

## 4. Op Code Penetration / Daily Service Doc / Advisor Analysis

Multi-tab workbooks analyzing alignment, balance, brake flush, fuel service, and major service penetration by store vs. benchmarks, plus an advisor scorecard (closing % vs. store average, monthly impact estimates). Low-volume advisors (<20 CP ROs) excluded.

**Definitions:** PL Pace = net profit pace on trailing 2-month expense baseline. PL% = net profit as % of labor gross.

---

## 5. Service Website Effectiveness (SWE) Analysis

**Scope:** 34-store analysis joining SWE scores to scheduler vendor (CDK, Auto.Live, hybrids).

**Key finding:** Platform-locked features showed within-vendor disagreement pointing to evaluator subjectivity rather than platform capability; "easy to navigate" and "appointment availability" metrics showed ceiling effects.

**Structure (6 tabs):** Key Findings, Vendor Summary, Metric Matrix, Timing Metrics by Vendor, Store Detail, Raw Data.

---

## 6. Technician Pay Rate Audit

**Source:** RO Detail export + Tech Pay Master file (not Qlik).

**Pay rules:**
- SC=A rows → Sold Hours × Hourly Rate.
- Blank SC rows → Skill-A % × Sold Amount.
- Known flat-fee exceptions: PPM prepaid maintenance, SMOG/SMOGU emissions.
- Negative ISP credit lines flagged separately.

**Structure (4 sheets):** Technician Summary, Technician Rates, Original Data (4 helper columns), Variance Detail (live formula links back to source rows — Python-filtered, not a FILTER spill formula).

**Build rule:** recalc.py must be the true final step.

---

## 7. GL Account 74074 Audit

Shop supplies credit-only account with stores incorrectly posting debits. Structure: Store Summary tab + 11 per-store detail tabs, sorted by total dollar amount mischarged.

---

## 8. Used Tire Disposal Cost Analysis

**Scope:** Group-wide cost analysis across all 34 stores to support potential vendor consolidation.

**Build:** Data-collection spreadsheet `Findlay_Tire_Pickup_Pricing_by_Store.xlsx` distributed to parts and service managers at each store to fill in current vendor pricing data. Analysis pending full data collection.

---

## 9. Used Oil Reimbursement Comparison (Cleantech vs. SafetyClean)

One-time spreadsheet comparing reimbursement rates across gallon increments. Not a recurring report.

---

## 10. Voice AI Deployment Tracker

34-store spreadsheet mapping store abbreviations to full names and assigned Voice AI agent (Flai or Pam). As of last build, ~11 stores were unassigned pending full rollout.

---

## 11. Opcode Legend Reformatter

Flat long-format table (one row per code): Store #, Code, Service Type (TRANS SERVICE / COOLANT SERVICE / DIFF SERVICE). 192 data rows across 34 stores.

---

## 12. Pay Plan Calculator (FMG_GROSS_PAY_CALCULATOR.xlsx)

New tab added to an existing workbook. Inputs: base salary, two flat bonuses (Y/N toggles), manual GP scenario box (commission rate × GP → commission/monthly/annual pay). Large comparison table: $600K–$880K GP in $20K steps × 5 commission rate tiers, with three parallel blocks (Commission, Monthly Pay, Annual Pay) and conditional formatting highlighting the cell matching the manually entered scenario.

---

## 13. Monthly Alignment Benchmark Report (recurring)

**Structure (3 tabs):** Monthly, Yearly (YTD), Raw Data.

**Formatting:** Store column fixed width 32; Arial throughout. 3-tier CF on Actual % — green ≥15%, yellow 10–15%, red ≤10%. Benchmark = 15%. SUMIF/COUNTIF summary block below totals (Alignment Variance in col G, Gross Variance in col I, same two rows).

**Build rule:** Always run recalc.py before delivery.

---

## 14. Work Type Mix Report (ad hoc)

Keyword-based classification (Maintenance vs. Repair) on `RO_Detail.opcodedescription` via WildMatch set analysis. Hours-weighted (Sold Hours), not line count. Maintenance checked first to resolve overlaps. Not yet formalized as a recurring report.

---

## 15. Technician Productivity Report — Honda Henderson & Honda North

Full technician name/position rosters on file. Honda North techno format strips the "YD900" prefix. Unmapped pooled codes (2000, 3000, 7352 at Honda North) are likely quick-lane/pooled, pending confirmation. Available Hours = working days × 8 hrs/day flat assumption (not Qlik WorkingDays). Productivity % routinely exceeds 100% in flat-rate shops — expected, not a data error.

---

## 16. BOB Stats Summary Formatting

Multi-year (2020–2025) store grades/scores workbook. Color-coded conditional formatting by grade tier, title case with automotive abbreviations preserved, formatting stripped from rows 40+.

---

## Superseded / Foundational Builds

Earlier single-store technician productivity work and initial report drafts that were later folded into the 34-store versions above are not detailed separately here — ask if historical detail on a specific superseded version is needed.
