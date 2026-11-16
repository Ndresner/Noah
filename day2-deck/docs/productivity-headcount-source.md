# Productivity headcount source: ADP, not Qlik (decided Oct 20, 2026)

**Decision (Noah):** Productivity slides use headcount from the ADP Employee Master extract (Employee Count by Position workbook, pulled 10/19/2026), not the Qlik Productivity app. Gross and RO counts still come from Qlik.

**Why:** The Qlik headcount history undercounts 2025. In Sep-25, 1,451 fixed ops people match between the two sources. ADP has 226 more:
- 169 have no other store record, yet Qlik doesn't show them until Oct-25 to Oct-26. These are a Qlik error.
- 34 are transfers that ADP back-dates to the original hire date. These are an ADP quirk, and Qlik is right on them.
- 18 are at Findlay Customs, which isn't a store.

By Sep-26 the two sources agree: 1,537 vs 1,537, with 1,549 of 1,550 people matched.

**Basis used:**
- Month-end snapshot.
- Qlik's Fixed Employee title list, applied to ADP titles.
- Same-store 33 (INEOS out). GMC Prescott body shop titles, gross and body-only ROs out.
- A person open at two stores is counted at the old store until that record closes, then at the new one. Each person counts once, so store rows add to the group.
- Caveat: ADP job titles are as of the extract date.

**What changed on the slides (Jan–Sep monthly average):**

| Metric | Qlik basis (old) | ADP basis (new) |
|---|---|---|
| Fixed employees | 1,433 → 1,484 (+3.6%) | 1,624 → 1,541 (−5.1%) |
| Sep-to-Sep headcount | +97 | −76, to 1,540 |
| Fixed gross per fixed employee | $14,062 → $14,137 | $12,410 → $13,617 (+$1,207, +9.7%) |
| ROs per advisor | 189 → 172 | 173 → 166 |
| ROs per fixed employee | 29.5 → 27.8 | 26.0 → 26.8 |
| ROs per fixed manager | 440 → 403 | 377 → 390 |

**Rebuild:** run engine `hc-build_prod.py`, then `prod-deck.mjs`, `fge-deck.mjs`, `fge-gap-deck.mjs` and `tp-deck.mjs`. Swap the result into the master with `swap_prod.py`, then run `merge_test.py`.

**Open item:** the Qlik headcount load needs fixing at the source. The file "Headcount - In ADP but not Qlik, Sep 2025.csv" lists every person affected.
