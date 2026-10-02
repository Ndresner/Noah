---
name: adp-pull
description: Pull Findlay ADP Workforce Now data (roster/headcount, Parts & Service YTD earnings, technician regular + overtime hours) straight from the ADP API with the bundled CLI at tools/adp/adp_cli.py, written in the same .xlsx layout as the manual ADP report exports. Use whenever a report needs ADP data — the Pay vs Gross report's "Earnings - Parts and Service Employees" export, the Technician Proficiency "Tech Employee List / Tech Efficiency Report Data" export, roster matching, or headcount — or when the user asks to pull, refresh or check anything from ADP. Fall back to the user's manual export only if the API pull fails.
---

# ADP pull (CLI)

`tools/adp/adp_cli.py` replaces the manual ADP report exports. Its output files have the same sheet name (`1`) and the same columns in the same order, so the downstream report scripts run unchanged.

| Command | Replaces | Feeds |
|---|---|---|
| `earnings --pay-date YYYY-MM-DD` | "Earnings - Parts and Service Employees" | pay-vs-gross `build_base.py --adp` |
| `tech-hours --start --end` | "Tech Employee List / Tech Efficiency Report Data" | technician proficiency `build_data.py` roster |
| `roster` | ad hoc roster pulls | headcount, roster matching (Ledger driver pages, proficiency) |
| `probe` | — | credential and API-access check |
| `seed-config` / `compare` | — | setup and validation (offline) |

## Before the first run

1. **Credentials.** Set these as environment secrets in the cloud environment settings, never in the repo or chat: `ADP_CLIENT_ID`, `ADP_CLIENT_SECRET`, `ADP_CERT`, `ADP_KEY`. ADP requires a mutual-TLS certificate, so the cert and key are either file paths or the PEM text.
2. **Network.** Allow `accounts.adp.com` and `api.adp.com` in the environment's network settings. If they aren't allowed, the proxy rejects every call.
3. `pip install requests openpyxl`
4. `python3 tools/adp/adp_cli.py probe` must show OK for the token, the Workers API, pay statements and pay statement detail. A FAILED line usually means the ADP API Central subscription doesn't include that API; take that to payroll/IT.
5. **Filters.** `cp tools/adp/adp_config.example.json tools/adp/adp_config.json`, then
   `python3 tools/adp/adp_cli.py seed-config --earnings-export <last manual earnings export> --tech-export <last manual tech export>`.
   This copies the exact job-title lists from the manual reports. `earnings` and `tech-hours` refuse to run with an empty title list, because an empty filter would pull the whole group's payroll.
6. **Validate once against a manual export for the same period** before trusting the API file:
   `python3 tools/adp/adp_cli.py compare --api <api file> --export <manual export>`.
   It reports rows missing or extra on either side, per-column mismatches, and the Gross Pay / hours totals. Fix any differences before switching a report over (see "Known gaps").

## Running

Write every output to the scratchpad. The files contain employee pay and PII. `tools/adp/.gitignore` blocks .xlsx, .json, .csv and cert files, but don't rely on that alone.

```bash
S=<scratchpad>
python3 tools/adp/adp_cli.py earnings  --pay-date 2026-08-15 --out $S/adp_earnings_2026-08-15.xlsx
python3 tools/adp/adp_cli.py tech-hours --start 2026-09-01 --end 2026-09-30 --out $S/adp_tech_hours_2026-09.xlsx
python3 tools/adp/adp_cli.py roster --out $S/adp_roster.xlsx
```

Options: `--dump-raw $S/raw` saves the raw ADP JSON for debugging. `--pause 0.2` slows the calls if ADP rate-limits. `--include-zero` keeps matched employees who had no pay in the range.

## Logic

- **Earnings Gross Pay is YTD.** It sums every pay statement dated Jan 1 through `--pay-date`, because `build_base.py` divides Gross Pay by months employed since Jan 1. `--pay-date-from` changes the start date.
- **Tech hours** add up the pay statement earning lines that carry hours. A line counts as Regular or Overtime only on a whole-word match to `hours.regular_codes` / `hours.overtime_codes`. Holiday, PTO and anything else are left out and listed in the run log. Add codes to the config if the manual report counts them.
- `--basis period_end` (default) includes a pay statement when its pay-period **end date** falls in start..end. `--basis pay_date` uses the check date instead. Use whichever matches the manual report; `compare` settles it.
- **Worker fields** come from the primary work assignment:
  - Company Code = `payrollGroupCode`
  - Position ID = `positionID` (company code + 6-digit file number; the proficiency script parses Tech No from it)
  - Location Description = `homeWorkLocation` name
  - Hire/Rehire Date = the rehire date when there is one, otherwise the original hire date
- Terminated employees are kept in `earnings` and `tech-hours` when they could have pay in the range. `roster` drops them unless `--include-terminated` is set.
- If more than 5% of employees' pay-statement calls fail, the run stops and writes nothing.

## Known gaps — verify with `compare` on the first live run

The CLI was built and tested against a mock of ADP's documented API shapes. It has not yet run against Findlay's ADP tenant. Check these on the first run:

- **Location names:** `homeWorkLocation` must match the strings in the proficiency `LOC_TO_DIVISION` map (e.g. "VW ST. GEORGE", "LEXUS SPOKANE"). Any unmapped location stops that build loudly.
- **Earning codes:** confirm what Findlay's Regular and Overtime codes are actually called.
- **Job-title filter vs department filter:** the manual reports may filter by ADP department rather than by title, so new titles could be missed. `compare` will show missing rows.
- **Pay dates per call:** if one call doesn't reach back to Jan 1, raise it with `--lookback N`.
