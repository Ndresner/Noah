# Day 2 Deck - Data Audit (Oct 28, 2026)

Scope: the 194 meeting-flow slides of the 237-slide test deck (Open Decisions not audited). Every figure was recomputed from the committed data files, independently of the build scripts. CP group totals were re-pulled live from Qlik Closed ROs on Oct 28 (CP ROs 279,300 / 279,756; CP hours 504,018.97 / 523,141.86; CP labor sale $80,067,347 / $84,966,724 = ELR $158.86 / $162.42; CP & W ROs 379,921 / 371,180) and match the deck exactly. Store rows sum to the group on every input, both years.

Result: no wrong source data. The issues are display rounding, a few labels/periods, and stale notes.

## Fix (visible on a slide)
Status Oct 28: 1-4 fixed (round 21). GMC Prescott body shop out of absorption and Lexus (not gray) done. 5-9 and the judgment calls: done Oct 28 round 22 per Noah (143 keeps 20% and rewords; 158 tie named on the band); option 8 (section dividers) in review.

| # | Slide(s) | Issue | Fix |
|---|---|---|---|
| 1 | 142 | Band: "6 stores at or above 85% still leave $537K a month". 20 stores are at or above 85% and leave $817,537 a month; the 6 are those leaving $50K+ each. | "20 stores at or above 85% still leave $818K a month" |
| 2 | 42, 56, 57 | "Every 0.1 hour per RO is worth $4,543,691 of labor" is Jan-Sep labor SALE (0.1 x 279,756 x $162.42), not a year; slide 60 and Where the Money Is say $6,602,564 a year gross. | Use $6.6M a year gross, or label "Jan-Sep labor sale" |
| 3 | 94, 98, 102, 106; movers 101, 105, 109, 125; Store Rx | Opcode rates truncated but changes/gaps rounded (11.9% -> 13.2% shown as +1.4 pts; "1.7 pts under 15%" off a shown 13.2%). Store Rx rounds the same rates (47 values 0.1 higher; Lexus balance 0.1% vs 0.0%). | One rule: compute change/gap from the shown (truncated) values; truncate on Store Rx |
| 4 | 156-157 vs 47, 49, 51, 61, 70 | Close truncates, earlier slides round: CP RO count +0.1% vs +0.2%, Kia LV +13.6% vs +13.7%, etc. | Same rule everywhere |
| 5 | 159 | Opcode row shows Mazda alignment, but the footer says each row is the top mover; Audi Henderson tops alignment. | Reword "Mazda's biggest opcode move" or show Mazda balance (top mover) |
| 6 | 143 | Share column rounds 19.97% and 19.57% to "20%", so 7 rows show 20%+ while the band says 5 stores miss a fifth. | Truncate the column |
| 7 | 158 | Left list cuts Nissan (2 top-5 finishes, tied with the shown stores) by insertion order. | Show ties or cut at a clean break |
| 8 | 146, 155 | Section chips stale: Warranty omits Warranty Gross Profit; Close still lists "Before you leave". | Update the section list |
| 9 | 45 | "6 stores at 100%" callback, 5 listed. | List all 6 or say "busiest five" |

## Notes-only cleanups
27 (breakdown "next slide"), 147 ("Warranty Rate Increases"), 13 and 32 (Closed RO / Roundtable boilerplate on P&L explainers), 43 (old Flai pull note), 69 ("OPTION A"), 92 (photo "not a meeting slide"), 99 and 107 ("CHECK OP-CODE MAPPING"), 95/99/103/107 (rounded rates vs truncated on slide), 85 ("-0.0 pts").

## Judgment calls for Noah
- GMC Prescott absorption includes the body shop (75.6%, rank 17); without it about 78.6% (above the group, green mark). Group unchanged at 78.1%.
- Lexus tire balance 0.07%: op-code mapping or manufacturer doesn't call for it (gray like the Audi stores)?
- Slide 65: LR Reno and Hyundai Prescott show identical work-type shares (56.8 / 23.9 / 19.3), likely a copy error in the recovered v1 data; no raw pull to check.
- Slide 12: "$20.59M a month over the last twelve months" sits under a Jan-Sep 2026 header (Jan-Sep average $20.98M).
- Slide 3: 27 cents tile labelled all stores; it's same-store (still 27 cents with INEOS).
- "Best September at +$1.61M" (5, 14, 18, 22) can read as best September ever; "largest gain in September".
- Slides 127/129: headcount -76 (Sep to Sep) and -83 (average) both shown.
- Slide 44: "last 90 days" covers 93.
- Slide 110: "7 of 33 ahead across the four services" means ahead on net dollars; only Honda Henderson is above all four.
- Near-zero changes (+0.0 / -0.0 pts) are counted in raised / declined bands (30, 35, 39).

## Still open (before the meeting)
ELR benchmark 70% vs 80%; Store Prescriptions handout or not, and a "vs Jan-Sep 2025" label; Most Improved extra metric; Flai pickup-rate definition; Tommy-Gary transcript; warranty audit wording vs OEM policy; delete Open Decisions (43 slides); Inter on the presenting PC.
