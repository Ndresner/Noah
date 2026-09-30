---
name: fixed-ops-71034-policy-audit
description: Build or refresh Findlay's monthly Service Policy (GL 71034) Employee & Advisor-Owned RO audit workbook — closed ROs on each store's policy labor type matched against the ADP Active/Leave employee list to find advisors writing policy ROs to themselves, ROs on employee customer numbers, and employee-name matches (Confirmed / Likely / Review / Low). Use whenever the user asks to run, update, refresh or rebuild the 71034 audit, the service policy audit, the employee/advisor policy RO audit, or "advisor = customer" checks — even if they only attach a new ADP Employee List export. Do not hand-build it; run the bundled scripts.
---

# Service Policy (71034) Employee & Advisor RO Audit

Answers one question: **which policy (GL 71034) ROs this month went to our own employees, and did any advisor write one to themselves?** Output is a 4-tab workbook: Summary, Flagged ROs, Not in 71034 Scope, Method & Limits.

Two scripts:
1. `scripts/qlik_pull_expr.py` prints the arguments for the one Qlik call that pulls the whole month.
2. `scripts/build_audit.py` takes the Qlik result file plus the ADP export and writes the workbook and a QA JSON.

The store crosswalk (Qlik logon ↔ ADP Store ID ↔ display name) is the shared repo file `reference/stores.json`.

Config lives in `config/`. It contains no employee or financial data.
- `policy_labor_types.json`: which labor types post to 71034 at each store, from the Aug-2026 71034 Labor Type Mapping.
- `common_surnames.txt`: about 200 frequent US surnames. It decides Low vs. Review on other-store name matches.

## Inputs to collect

1. **ADP "Employee List – Active/Leave" export (.xlsx)**, pulled as close to month-end as possible. It needs these columns: Location Description, Store ID, CDK Employee ID #, Legal Last Name, Legal First Name, Position Status.
2. **The audit month** (e.g. `2026-09`).
3. **Whether the labor type mapping changed.** The default is the Aug-2026 mapping. If the user has a newer `71034_Policy_Labor_Type_Mapping_*.xlsx`, update `policy_labor_types.json` from its Store Mapping column K and its Labor Type Index "(unconfirmed)" rows before running. Stores missing from `primary` are treated as having no 71034 activity that month.

## Step 0: Environment

`recalc.py` needs LibreOffice Calc. If recalc times out, install it:
`apt-get update && apt-get install -y --no-install-recommends libreoffice-calc`
Run `apt-get update` first; the stale package index returns 404s.

## Step 1: Qlik pull (one call)

Run `qlik_clear_selections` on Closed ROs `c3efc739-b063-47dd-a4d3-d1c4cac07ad0` first. Then generate the call:

```bash
S=.claude/skills/fixed-ops-71034-policy-audit/scripts
python3 $S/qlik_pull_expr.py --month 2026-09        # prints appId / dimensions / measures JSON
```

Pass the printed JSON to `qlik_create_data_object` exactly as printed.
- The result comes back as 34 rows, one per store: `logon`, `n`, and `rows`, where `rows` holds all RO × labor-type lines joined with `~`, with fields separated by `|`.
- The result is about 400K characters. The client saves it to a `tool-results/…txt` file instead of showing it inline. Copy that file into the scratchpad and **don't read it into context**.
- Why this shape: page-based pulls cap at 100 rows per call (about 48 calls a month). The Concat shape gets everything in one call and nothing has to be retyped.
- If the result does come back inline (a very small month), split the call into two by `%Logon` and have a subagent save each JSON verbatim, or write it as CSV (see the `build_audit.py` docstring for the columns).
- Auth failures on the first attempt are normal. Retry once.

The labor type filter covers every mapped type in every store. `build_audit.py` then decides scope for each store. The `n` measure is a row-count check: the script stops if any store's Concat came back truncated.

## Step 2: Build, then recalc (mandatory)

```bash
python3 $S/build_audit.py --qlik <scratchpad>/qlik_<mon>.json \
  --adp "<ADP Employee List>.xlsx" --month 2026-09 \
  --out "<scratchpad>/71034_Employee_Advisor_Policy_Audit_Sep2026.xlsx"   # [--no-secondary]
python3 <xlsx-skill>/scripts/recalc.py "<out>.xlsx" 240   # must report success, 0 errors
```

Never re-save with openpyxl after recalc; it strips cached values. The script prints the population, a count per flag, and any warnings. The QA JSON sits next to the output file.

## Step 3: Verify before delivering

- recalc shows `total_errors: 0`. Summary rows C13, D13 and E13 are all 0 (counts, sale and cost tie to Flagged ROs).
- Check the QA JSON:
  - `adp_unmapped_locations` should hold only `J01 CUSTOMS`. Any other code means `reference/stores.json` needs a new store.
  - `mapped_stores_without_rows` should be empty. A mapped store with no policy ROs usually means the labor type changed.
  - `unmapped_policy_types` lists labor types in the pull that aren't mapped at that store, e.g. `15:ISP` (Lincoln posts IOA). Tell the user if a new one shows up.
  - `id_hits_name_mismatch` lists customer # = employee ID hits whose names disagree. They are dropped, not flagged. List them for the user if there are any.
- **Reference run (Aug 2026):**
  - Population 4,595 ROs across 32 stores, ADP 2,813 rows and 35 locations, 30 employee-ID hits.
  - 61 flagged: 6 / 2 / 20 / 15 / 2 / 6 / 10 by flag. Policy sale $4,572.18, policy cost $5,157.84.
  - 4 ROs on the out-of-scope tab (stores 3 and 42).
  - This rebuild matched the original hand-built Aug workbook row for row. The one extra row (VW Henderson RO 482937, "RANDS JR,RICHARD") is caught because suffixes are now stripped. It carries a "Suffix differs" note.

## Step 4: Deliver

Send the workbook. Lead with:
- Confirmed and Likely ROs, with policy cost.
- Advisors who wrote policy ROs to themselves (red rows).
- The stores with the most flags.
- Any new QA warnings.

**Never commit the workbook, the Qlik file or the QA JSON.** They hold employee names and IDs. Only this skill's scripts and config belong in the repo.

## How the tests work (build_audit.py)

**Scope**
- An RO counts only on the policy labor type(s) mapped for its store. Secondaries are included unless `--no-secondary` is passed.
- Stores not in the mapping are checked only on `out_of_scope_types` (ISP and ISPJ). Their Confirmed and Likely hits go to "Not in 71034 Scope".

**One row per RO.** The highest-priority hit wins, in the `Sort` order:

| Sort | Flag | Confidence | Rule |
|---|---|---|---|
| 1 | Advisor = Customer (number) | Confirmed | cust # = numeric part of `%serviceadvisor` (4th `\|` field) |
| 2 | Advisor = Customer (name; RO on regular cust #) | Confirmed | cust name key = advisor name key |
| 3 | Employee # = Customer # | Confirmed | cust # = CDK ID of a **same-store** employee (IDs repeat across stores) AND last names agree |
| 4 | Employee name, same store (regular cust #) | Likely | cust name key = a same-store employee |
| 5 | Same last name as advisor | Review | possible relative |
| 6 | Employee name, other store (distinctive name) | Review | name key matches an employee at another store |
| 7 | Employee name, other store (common name) | Low | same as 6, but the surname is in `common_surnames.txt` |

**Name matching**
- The name key is (LAST, first given name). Apostrophes and periods are dropped, hyphens become spaces, and JR/SR/II/III/IV are ignored.
- Names are compared as whole words. That rules out prefix false positives like DANIELLE vs. DANIEL.
- Business names (no comma) never match.
- ADP legal names are "First Middle" / "Last Jr". The key uses the first token of the first name.

**Notes column**
- Filled automatically for: possible relatives; an advisor number that equals the advisor's CDK ID; multiple other-store ADP matches; suffix mismatches.
- Add store feedback by hand after delivery if the user asks.

**Dollars.** Policy Sale = labor + parts sale on the policy lines. Policy Cost = labor + parts cost on the same lines. MLS and tax are excluded because they carry no labor type.

## Layout (matches the original Aug-2026 build)

- **Fonts and colors:** Arial 10. Headers are navy `1A2744` with white bold text. Total rows are gold `C9A04B`. Advisor-self rows (Sort 1–2) get red fill `F8CBAD`. Hardcoded Qlik/ADP values are blue `0000FF`. Formulas are black. Thin `BFBFBF` borders.
- **Summary:**
  - Flag table (rows 4–13) with COUNTIFS/SUMIFS against Flagged ROs, plus tie-out checks.
  - Store table from row 15. It lists only stores with flags, in logon order, with Confirmed + Likely ROs, sale and cost, and All Flagged.
  - The workbook opens on Summary.
- **Flagged ROs:**
  - Columns A–T, sorted by Sort, then logon, close date, RO.
  - O = `=M+N`. Gold TOTAL row with COUNTA/SUM.
  - Freeze at G2 and autofilter.
- **Widths:**
  - Summary: A46 B12 C8 D13 E13 F70.
  - Flagged ROs: A6 B34 C11 D8 E20 F10 G11 H13 I30 J10 K24 L10 M12 N13 O13 P11 Q24 R26 S12 T30.
  - Method: A24 B110.

## Known limitations (tell the user when relevant)

- ADP holds current employees only. Anyone who termed between the 1st and the export date is missed.
- Name matching misses nicknames, maiden names and typos.
- "Same last name" rows are possible relatives, not proven.
- Qlik RO amounts haven't been tied to the GL for these specific ROs.
- The labor type mapping comes from a separate GL tie-out (the 71034 Labor Type Mapping workbook). It isn't rebuilt monthly. Re-confirm it when `unmapped_policy_types` or `mapped_stores_without_rows` changes, or at least quarterly.
- IXP (store 12), ISPB (34) and IMGT (46) are still unconfirmed secondaries. They are included by default, as in August.
