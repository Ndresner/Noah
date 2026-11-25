# In-house warranty labor rate / parts markup submissions

Goal: find out whether Findlay can build its own warranty rate submissions (manufacturer procedure or state statute) instead of paying a vendor (DealerUplift, an Armatus company), and how close our numbers land to theirs.

Status as of Nov 25, 2026: method built and tested for Toyota Flagstaff (AZ) and Honda Henderson / Honda North (NV). All five state statutes stored. DealerUplift submission history inventoried (57 cycles, 30 stores) and 23 cycles rebuilt (see "Historical check" below). Shadow run of live DealerUplift cycles under way (`shadow/`), with a scheduled mailbox check for new results.

## Folder
- `statutes/` – state statute text (verbatim, with source URL and retrieval date) and notes: NV (NRS 482.36385), AZ (ARS 28-4451), UT (13-14-204), ID (49-1626), WA (RCW 46.96.105).
- `history/` – DealerUplift/Armatus submission inventory from the mailbox (`dealeruplift_submissions.csv`, one row per store x labor/parts x cycle, Sep 2024 – Nov 2026) and its summary (outcomes, gaps, methodology statements).
- `config/labor_types_proposed.csv` – proposed customer-pay / employee / third-party labor types for every store (from Qlik, May–Nov 2026), with open questions. Confirmed rows move into `store_labor_types.json`.
- `oem_policies/` – summaries of manufacturer warranty policies and labor-rate procedures: Toyota 7.1–7.4, Lexus, Honda SOM 6.15, Acura, Hyundai, Volkswagen, Ford/Lincoln, INEOS, Jaguar Land Rover (partial), Kia (forms + Section 7). Originals are dealer-confidential and are not stored here. Still to come (requested from service managers Nov 2026): Nissan, Subaru, Mazda, Audi/Lamborghini, GM, Stellantis, Volvo; JLR NAS local market bulletin.
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
- On-brand vehicles only (`RO_Header.make` = the franchise make) – submissions count only the manufacturer's own line-make.
- Lines with ≥15 sold hours at under $30/hr are dropped as clerical errors (e.g. Chrysler Post Falls RO 198882: alignment booked at 150.0 hours).
- Outputs contain customer numbers and RO detail: **never commit the workbooks, CSVs or Qlik JSON**. Scratchpad only.

## Findings so far (estimates, not filings)
All statutory figures count **on-brand vehicles only** (Toyota/Scion at Toyota Flagstaff, Honda at the Honda stores). Corrected Nov 25, 2026: the first versions of the Toyota-vs-AZ and Honda-vs-NV builds counted every make in the statutory method; the effect was small (see below).

| Store | Current warranty labor (Qlik W lines) | Manufacturer method | Statute method (warranty-like) | DealerUplift prelim |
|---|---|---|---|---|
| Toyota Flagstaff (AZ) | $157.84 | Toyota 7.4 median $153.61 (best $172.99) | AZ median $169.64, best $187.73; parts markup median 79.8% | n/a (Toyota not with DealerUplift) |
| Honda Henderson (NV) | $186.12 | Honda 6.15 median $196.60, best $231.18 | NV median $193.96, best $207.21 (Honda op-code rules) / $204.82 (description rules, shadow run) | $207.28 (Nov 20, 2026) |
| Honda North (NV) | $178.74 | Honda 6.15 median $186.58, best $225.01 | NV median $186.72, best $198.74 (Honda op-code rules) / $195.60 (description rules, shadow run) | $199.28 (Nov 20, 2026) |

- On-brand correction: Toyota Flagstaff AZ best +$0.38, median −$0.25 (was $187.35 / $169.89); Honda Henderson NV best −$0.89 (was $208.10); Honda North unchanged.
- Our implied current rates match DealerUplift's within $0.50. At the Honda stores, the store-specific op-code rules land within about $0.50 of DealerUplift's prelims; the brand-agnostic description rules used for the shadow run land $2.46 (Henderson) and $3.68 (North) short. The RO detail requested from DealerUplift will show which line treatment matches theirs.
- Toyota Flagstaff: Option B (manufacturer) at the median start date pays less than today; the AZ statutory route is worth ~$134K/yr at the median window (labor $43K + parts $91K). Alignments ($111/hr) are the biggest drag under Toyota's rules.
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
23 DealerUplift cycles (11 labor, 12 parts; see `config/history_cycles.json`) rebuilt from Qlik with `scripts/history_vs_dealeruplift.py`: 180-day look-back ending at DealerUplift's results date, 100 sequential qualifying ROs (or 90 days), state rules per `statutes/`, description-based line classification, on-brand vehicles only, clerical-hours guard. Workbook is scratchpad-only.

Labor ($/hr): DealerUplift vs. our best / median window
| Store | DU | Ours best | Ours median |
|---|---|---|---|
| Honda Henderson (10/25) | 185.83 | 193.11 | 175.86 |
| Honda North (11/25) | 179.22 | 186.56 | 173.17 |
| Nissan Henderson (5/26) | 216.39 | 236.83 | 213.17 |
| Land Rover Henderson (2/26) | 278.57 (approved 270) | 273.67 | 269.33 |
| Jaguar Henderson (11/25, 30 ROs) | 290.27 (approved 250) | 306.66 | 294.67 |
| Land Rover Las Vegas (8/26) | 308.58 | 294.78 | 281.36 |
| Jaguar Las Vegas (11/25) | 274.61 (approved 250) | 276.46 | 266.94 |
| Subaru Prescott (8/26) | 164.82 | 180.40 | 172.96 |
| BGMC Prescott (11/26) | 200.42 | 211.57 | 207.37 |
| Chrysler Post Falls (12/25) | 183.36 | 184.61 | 178.24 |
| Lexus Spokane (7/26) | 224.51 | 221.89 | 218.21 |

Parts markup: our best window is at or above DealerUplift's in 11 of 12 cycles, within ~3 points at Kia St George, Subaru Prescott, Chrysler Post Falls and Lexus Spokane; far below at Subaru St George (77.6% vs 122.9%, Utah rules - needs review). Our parts windows swing widely (median well below best), so parts needs a tighter line definition before it is relied on.

First OEM-audited parts result (VW St George, Utah, Nov 25, 2026): DealerUplift asked 96.48%; VW's line-by-line review landed at **89.69%** (after agreeing batteries and bulbs are excluded under Utah law). Our rebuild for that cycle: median 96.3%, best 118.9% – i.e. our median matched DealerUplift's ask, and both were about 7 points above what VW accepted. Use 89.69% as the calibration target when tightening the parts line rules (details in `oem_policies/Volkswagen_Warranty_Policy_2026_summary.md`).

## Manufacturer procedures at a glance (from `oem_policies/`)
| Brand | Own labor-rate route | Sample | Parts reimbursed | Own parts-markup route | Defers to state statute? |
|---|---|---|---|---|---|
| Toyota | Option A CPI or Option B ELR (Region sets start date) | first 30 qualified CP ROs, ≤120 days | (Toyota rate) | none | "unless prohibited by state law" on limits |
| Lexus | Same as Toyota: Option A CPI or Option B ELR (Area Office sets start date); alignments and timing/serpentine belts excluded; survey = 3 imports + 1 domestic (ELR worksheet Rev. 10-26); increase ≥ $250; not on WAS Tier II/III | first 30 qualified Lexus CP ROs, ≤120 days | MSRP (OTC cost + 30%) | none | "unless prohibited by state law" on limits only |
| Honda / Acura | 6.15 (iExam GP% ≤ 80%) | 20 consecutive qualifying CP ROs from 50+, dealer picks; parts required | dealer net + 40% | none | yes – 6.15 is "an alternative" to the statute |
| Hyundai | 5.10, capped at 5% below sample rate | 25 most recent sequential CP ROs | dealer net + 40% | none | yes – statutory requests "governed by state law"; over-cap requests sent to the statute |
| Volkswagen | Option 1 CPI (3-yr, CPI-U, cap 4%) or Option 2 retail rate | 30 sequential warranty-like CP ROs, ≤12 months, dealer picks | MSRP (MY2003+) | none (contact Warranty Labor Rate Specialist) | yes – "unless contradicted by relevant law"; statute-only retail-rate form |
| Ford / Lincoln | Index Rate (3-yr, CPI + 0.5%) or Retail Rate option, both via FordWarrantyRetailReimbursement.com | on the platform instructions/template (not yet obtained) | dealer price + 40% | yes – same online platform (since Feb 2021) | generic only |
| INEOS | Labor Rate Review Form (not in manual), 12-month dealers only | in the form | per state law | none – "in accordance with applicable state law" | "to extent allowed by applicable law" |
| Jaguar Land Rover | Annual market-level review by JLRNA (no dealer-filed request); rate kept within a 75–90% band; US details in the NAS local market bulletin (not yet obtained) | none at dealer level | net + 40% (10% batteries, some Jaguar components) | none | "if market legislation exists ... evidence must be provided" |
| Kia | "Effective Retail Labor Rate Method" (written procedure not in Section 7 – forms only); survey of Ford, GM, Honda, Hyundai, Mazda, Nissan, Toyota; Kia checks District Average | up to 100 CP repair lines (window/sequence not stated) | warranty cost incl. parts handling (rate not stated) | none found | not stated |

Check before any statutory filing: whether the store is enrolled in a CPI/index program (Toyota Option A, VW Option 1, Ford Index Rate) and its exit terms.
