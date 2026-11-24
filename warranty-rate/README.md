# In-house warranty labor rate / parts markup submissions

Goal: find out whether Findlay can build its own warranty rate submissions (manufacturer procedure or state statute) instead of paying a vendor (DealerUplift, an Armatus company), and how close our numbers land to theirs.

Status as of Nov 24, 2026: method built and tested for Toyota Flagstaff (AZ) and Honda Henderson / Honda North (NV). All five state statutes stored. DealerUplift submission history inventoried (57 cycles, 30 stores); historical comparison waiting on labor-type confirmation per store (`config/labor_types_proposed.csv`).

## Folder
- `statutes/` – state statute text (verbatim, with source URL and retrieval date) and notes: NV (NRS 482.36385), AZ (ARS 28-4451), UT (13-14-204), ID (49-1626), WA (RCW 46.96.105).
- `history/` – DealerUplift/Armatus submission inventory from the mailbox (`dealeruplift_submissions.csv`, one row per store x labor/parts x cycle, Sep 2024 – Nov 2026) and its summary (outcomes, gaps, methodology statements).
- `config/labor_types_proposed.csv` – proposed customer-pay / employee / third-party labor types for every store (from Qlik, May–Nov 2026), with open questions. Confirmed rows move into `store_labor_types.json`.
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

## State statute comparison (quick reference)
| | NV | AZ | UT | ID | WA |
|---|---|---|---|---|---|
| Standard | ≥ retail charged | fair & reasonable; dealer may declare | ≥ retail for like parts/service | ≥ retail (if not unreasonable vs. area) | ≥ retail; ≥ 2010 schedule |
| Method in statute | none | 100 seq. or 90 days, ≤180 days | 100 seq. qualified or 90 days, ≤180 days | 100 seq. or 90 days, ≤180 days | parts: 100 seq. or 90 days; labor: retail rate |
| "Warranty-like" | – | "warranty-like repairs" (undefined) | defined: would be warranty but for time/miles | – | – |
| Alignments | – | not named | **excluded** | not named | not named |
| Brake pads/rotors | – | not named | **excluded** | not named | not named |
| Batteries / bulbs | – | not named | non-EV batteries, bulbs excluded | not named | **excluded** |
| Menu pricing | – | – | – | **excluded** | – |
| Aftermarket parts | – | – | excluded | – | – |
| Manufacturer response | court action | rebut in 30 days → director/ALJ | object or request ±30-day ROs within 30 days | info request within 30 days; effective ≤90 days | refuse only if rate unreasonably exceeds all area dealers |
| Frequency | – | 12 months | 12 months (each) | 12 months | once per calendar year |

## Historical check vs. DealerUplift (Nov 24, 2026)
23 DealerUplift cycles (11 labor, 12 parts; see `config/history_cycles.json`) rebuilt from Qlik with `scripts/history_vs_dealeruplift.py`: 180-day look-back ending at DealerUplift's results date, 100 sequential qualifying ROs (or 90 days), state rules per `statutes/`, description-based line classification. Workbook is scratchpad-only.

Labor ($/hr): DealerUplift vs. our best / median window
| Store | DU | Ours best | Ours median |
|---|---|---|---|
| Honda Henderson (10/25) | 185.83 | 193.11 | 175.86 |
| Honda North (11/25) | 179.22 | 186.56 | 173.17 |
| Nissan Henderson (5/26) | 216.39 | 236.83 | 210.80 |
| Land Rover Henderson (2/26) | 278.57 (approved 270) | 273.67 | 269.33 |
| Jaguar Henderson (11/25, 30 ROs) | 290.27 (approved 250) | 306.66 | 294.67 |
| Land Rover Las Vegas (8/26) | 308.58 | 294.78 | 281.36 |
| Jaguar Las Vegas (11/25) | 274.61 (approved 250) | 276.46 | 266.94 |
| Subaru Prescott (8/26) | 164.82 | 180.40 | 172.96 |
| BGMC Prescott (11/26) | 200.42 | 211.57 | 207.37 |
| Chrysler Post Falls (12/25) | 183.36 | 184.61 | 178.24 |
| Lexus Spokane (7/26) | 224.51 | 221.89 | 218.21 |

Parts markup: our best window is at or above DealerUplift's in 10 of 11 cycles, within ~3 points at Kia St George, Subaru Prescott, Chrysler Post Falls and Lexus Spokane; far below at Subaru St George (77.6% vs 122.9%, Utah rules - needs review). Our parts windows swing widely (median well below best), so parts needs a tighter line definition before it is relied on.
