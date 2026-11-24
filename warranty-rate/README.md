# In-house warranty labor rate / parts markup submissions

Goal: find out whether Findlay can build its own warranty rate submissions (manufacturer procedure or state statute) instead of paying a vendor (DealerUplift, an Armatus company), and how close our numbers land to theirs.

Status as of Nov 24, 2026: method built and tested for Toyota Flagstaff (AZ) and Honda Henderson / Honda North (NV). Historical comparison against DealerUplift submissions is in progress. Utah, Idaho and Washington statutes not yet pulled (sites blocked by the cloud network policy until their domains are allowed).

## Folder
- `statutes/` – state statute text (verbatim, with source URL and retrieval date) and notes. NV and AZ done; UT, ID, WA to add.
- `oem_policies/` – summaries of manufacturer labor-rate procedures (Toyota 7.1–7.4, Honda SOM 6.15). Originals are dealer-confidential and are not stored here.
- `scripts/` – Qlik pull expression generator, Concat parser, and the three builds:
  - `toyota_74_30day.py` – Toyota Policy 7.4 30-RO calculation (single window).
  - `toyota_vs_az_statute.py` – Toyota 7.4 vs. Arizona ARS 28-4451, 180-day look-back.
  - `honda_vs_nv_statute.py` – Honda SOM 6.15 vs. Nevada NRS 482.36385, two stores.
- `config/store_labor_types.json` – which labor types count as customer pay, employee, third-party and policy per store (confirmed by Noah).

## Data rules
- Source: Qlik Closed Repair Orders (`c3efc739-…`), line grain `RO_Detail`, filtered by `Division` and `RO_Header.openeddate`. Aftermarket parts: Closed Parts Tickets (`d3d25914-…`), `Parts_Detail.source = 200`.
- Pull pattern: select `Division` + an `RO_Header.openeddate` match range (`>=M/D/YYYY<=M/D/YYYY`) with `qlik_select_values`, then one `qlik_create_data_object` with a `Concat(Aggr(...), %RO_Detail)` measure per pay type (see `scripts/qlik_line_pull_expr.py` for the row layout). 60-day chunks; results are saved to file by the MCP client – copy to the scratchpad and parse with `parse_qlik_concat.py`. Clear selections afterward.
- Qlik sale amounts are already net of discounts (no separate discount lines). Advertised/manufacturer specials can't be identified until the CDK discount feed lands.
- Hours = `RO_Detail.soldhours` (billed to customer).
- Outputs contain customer numbers and RO detail: **never commit the workbooks, CSVs or Qlik JSON**. Scratchpad only.

## Findings so far (estimates, not filings)
| Store | Current warranty labor (Qlik W lines) | Manufacturer method | Statute method (warranty-like) | DealerUplift prelim |
|---|---|---|---|---|
| Toyota Flagstaff (AZ) | $157.84 | Toyota 7.4 median $153.61 (best $172.99) | AZ median $169.89, best $187.35; parts markup median 77.8% | n/a (Toyota not with DealerUplift) |
| Honda Henderson (NV) | $186.12 | Honda 6.15 median $196.60, best $231.18 | NV best window $208.10 | $207.28 (Nov 20, 2026) |
| Honda North (NV) | $178.74 | Honda 6.15 median $186.58, best $225.01 | NV best window $198.74 | $199.28 (Nov 20, 2026) |

- Our implied current rates match DealerUplift's within $0.50; our best-window statute rates land within about $1 of their preliminary results at both Honda stores.
- Toyota Flagstaff: Option B (manufacturer) at the median start date pays less than today; the AZ statutory route is worth ~$127K/yr at the median window (labor + parts). Alignments ($111/hr) are the biggest drag under Toyota's rules.
- Nevada has no statutory method; the literal reading (all customer-pay work) is below current pay. The statute's value is leverage (NRS 482.36423 treble damages).

## Open questions
- Approved warranty labor rates and parts reimbursement on file for each store (we use Qlik W-line proxies).
- Toyota Flagstaff CPI agreement renewal/opt-out terms (last year of 3-year term).
- Honda 6.15 eligibility (iExam warranty mechanical labor GP% ≤ 80%).
- Counsel: "warranty-like" scope (alignments only with a repair; brake pads/rotors as wear items); separate labor and parts windows under ARS 28-4451.
- Whether Findlay policy on vendor use (Armatus/DealerUplift) also covers in-house statutory filings.
