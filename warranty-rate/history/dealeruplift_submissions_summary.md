# DealerUplift (Armatus) submission inventory - summary

Built 2026-11-24 from the Fixed Ops Director's Outlook mailbox (read-only). File: dealeruplift_submissions.csv (message links removed) (57 rows, one per store x type x cycle).

## Coverage
- Mailbox range searched: earliest DealerUplift mail 2024-09-17 (Honda Henderson labor) through 2026-11-24. Sparse gap between Dec 2024 and Oct 2025 (almost no Armatus mail in that window - likely retention/archiving, not inactivity).
- Searches: sender dealeruplift.com (all 230 hits paged), "preliminary results", "Ready to Ship", "final results", "Billing Explanation", "Approval Reminder", "Shipped", store-name queries. Many Armatus originals exist only as quotes inside Noah's replies/forwards.
- Rows: 57 cycles across 30 stores. Outcomes: approved 38 (many inferred, see below), pending 5, rebutted/reduced 3 (Honda Henderson 2024, Chevrolet labor 2025-26, VW St George parts 2026), not submitted/fell out 2 (BGMC Prescott labor Jul 2026, Audi Henderson labor May 2026), unknown 9.
- Stores: Honda Henderson, Honda North, Honda Flagstaff, Honda Spokane, Hyundai St George, Hyundai Prescott, Buick GMC Prescott, Chevrolet, Cadillac, Subaru St George, Subaru Prescott, Subaru Las Vegas, VW St George, VW Henderson, Lincoln, Jaguar Reno, Jaguar Las Vegas, Jaguar Henderson Service, Land Rover Reno / Las Vegas / Henderson, Acura, Lexus Spokane, Mazda, Chrysler Post Falls, Audi Reno Tahoe, Audi Henderson, Kia St George, Nissan Henderson, Lamborghini Las Vegas.

## How "approved" was determined
- Explicit: "Final Results & Billing Explanation" emails, "approved as submitted" statements, OEM correspondence (Chevrolet, VW, Honda 2024).
- Billing evidence: Armatus invoices that open "Congratulations on your recent Warranty Reimbursement Rate Increase" (first $2,250 invoice per project).
- Armatus "YTD RWR Financials - Findlay Automotive Group 2026 (002).xlsx" (Mike McGuire, 2026-10-21): lists 2026 YTD revenue secured by store, Labor $1,661,368 + Parts $2,126,832 = $3,788,100. Some rows rely only on this.
- OEM-approved rates are rarely stated in email text; approval letters are mostly scanned PDFs or images. oem_approved_rate is blank unless a number was stated.
- implied_current_rate = supported rate minus stated increase (arithmetic only). Final-stage numbers are in the extra columns final_results_date / final_supported_rate / final_increase / final_annual_value, added after the requested columns.

## Upcoming (Noah 2026-11-13, Armatus start dates)
Chevrolet Labor R7 11/16; Chrysler Post Falls Labor R7 11/16; Honda Flagstaff Labor R5 11/16; Honda Henderson Labor R5 11/16; Honda North Labor R6 11/16; Findlay Kia Labor R5 11/19; Subaru Las Vegas Labor R6 11/20; Jaguar Las Vegas Labor R6 11/25; Jaguar Henderson Service Labor R5 11/25; Honda Spokane Labor R4 12/02; Findlay Motor Company (GM) Labor R6 12/07; Hyundai Prescott Labor R4 12/10; VW St George Labor R7 12/14. ("R#" appears to be the project round number for that store.)

## Relevant threads not interpreted (or not fully read)
- Jaguar Reno Parts & Labor prelim (2026-09-02 / 09-10): rates not captured.
- Land Rover Reno Labor 2026 prelim figures (same thread as parts).
- Subaru Prescott "Preliminary Numbers" (2026-02-23) and Audi Reno "Approval Needed" (2026-06-05): figures are in table images.
- Cadillac labor prelim (2025-11-07), VW Henderson "Preliminary Numbers" (2026-05-06): figures not captured.
- Subaru St George Parts 2026 "Parts opportunity" (2026-11-03): figures not captured.
- Honda Henderson 2024 AHM response letter (2024-12-02) and 2025 "HONDA WARR RATE UPDATE 12.17.25.pdf": scanned, text not extractable.
- "PARTS REVIEW STORE LIST" (2026-02-18): Armatus says 12 stores have Parts agreements (attachment not opened).
- Lamborghini Las Vegas parts-at-retail audit (2025-12-09): not read.
- Lincoln 2026: Armatus's YTD list counts labor and parts as secured, but on 10/19 Noah had not seen a Ford letter.
- Mismatches between quoted annual value and the YTD list (Audi Reno parts $36.1K vs $146.8K; Nissan Henderson parts $16.9K vs $119.4K; Honda North parts $179K vs $320K) are not reconciled.
- The BGMC Prescott GM range table (Nov 2026) shows current parts 79.38%, but the Mar 2026 parts final was 73.78%. Not reconciled.

## Methodology statements found in emails
- Jeff White (2025-12-01): approval timing is set by state statute, not brand. NV has no statutory deadline; ID, AZ and WA are 30 days.
- Prelim emails ask OK to finalize "if we achieve a WLR of at least $X" (floor about $3-$10 below the prelim rate). Final ("Ready to Ship") emails often say how much was lost to discounting (Honda Henderson 2025: "lost over $9.00 to discounting"; BGMC parts "lost over 3...").
- Cadillac prelim: "best possible audit result from the most recent 6-months of store customer pay data". This is the only stated look-back.
- GM (Chevrolet 2026): 100 qualified ROs, consecutive RO range 1028440-1030249; GM's exclusions/inclusions debated: discount allocation, internal and extended-warranty pay types, warranty jobs, non-warranty-like repairs, batteries/TPMS.
- Honda (2024): AHM excluded ROs showing $0.00 customer payment; proof of customer payment needed.
- JLR (2026): proof of payment required for every qualified RO.
- VW (2026): letter of intent, then RO review. VW counted batteries; the store cites Utah 13-14-204(3)(a)(ii)(A), which excludes non-EV batteries as routine maintenance. Non-OEM parts (brake cleaner) and non-warranty-like repairs are excluded.
- Kia/Utah: Kia has never asked for supplemental ROs in Utah (Jeff).
- AZ statutes say "reasonable", so an OEM may reduce a request it deems unreasonable (Jeff, BGMC 2026).
- Nissan typically responds within 24-48 hours on labor (Mike McGuire). Hyundai now requires a portal submission with store credentials and 2FA.
- Labor projects need a store Labor Worksheet (LWS); at Nissan Henderson an updated LWS changed the value by about $76K/yr.
- No email stated a sample size or RO count for Honda Henderson or Honda North 2026. Noah requested this detail from Jeff on 2026-11-24, with no reply yet.
