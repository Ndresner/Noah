---
name: parts-matrix-peer-comparison
description: Compare one Findlay store's customer-pay parts pricing against sister stores of the same brand - its repair matrix vs. theirs, what a matrix change would cost (break-even), repair penetration, and total job tickets - using the saved CDK matrices in reference/parts_matrices plus Qlik Closed Parts Tickets and Closed Repair Orders. Use when a GM or manager asks whether a store's parts are "too expensive", "too aggressive", "more than our other stores", whether pricing is costing business, or what lowering the matrix would cost. Not for the single-store "are we hitting our own matrix" check - that is parts-matrix-audit. Do not hand-build it; run the bundled scripts.
---

# Parts Pricing Peer Comparison

Answers: **is this store's matrix-priced parts pricing out of line with our other stores of the same brand, is it costing business, and what would changing it cost?**

This is optional. Most matrix work is the single-store audit (`.claude/skills/parts-matrix-audit/`). Run this skill only when the user asks for a comparison against other stores.

Reference run: Toyota Prescott vs. Flagstaff, Henderson and Spokane, Aug–Oct 2026.
- Prescott's repair matrix was 1.71x list; the peer matrices were 1.22x–1.33x.
- Prescott's actual price on those lines was 1.42x list (409 lines, 229 ROs, $78,186.41 sale, $32,552.83 cost).
- Midpoint scenario: $6,258.58 of gross given up per quarter. Matching Flagstaff: $12,506.89 per quarter.
- A re-run of Prescott for the same period should reproduce these.

## Inputs

1. **Target store and peer stores.** All of them must be in `reference/parts_matrices/manifest.json`, with a current parsed matrix.
   - If a store is missing, add it first following `reference/parts_matrices/README.md`.
   - Peers should be the same brand. A different brand has a different list/cost structure, so the comparison doesn't hold.
2. **Each store's manifest entry needs two fields:**
   - `matrix_labor_types`: the customer-pay labor type(s) that get matrix pricing.
   - `main_repair_source`: the source whose cost-based rule is the store's main OEM repair matrix.
   - If either is missing, build the labor-type profile with the audit skill's `profile` bundle (or a Division × `RO_Detail.labortype` table), propose values, and **confirm them with the user before running**. Then write them into the manifest.
3. **Period:** default to the last 3 full closed months. The price-vs-list and same-part sections can use 1 month to keep the Qlik work down.

## What the scripts do (don't re-derive)

- `scripts/peer_expr.py` generates the Qlik arguments for each bundle.
  - It imports the matrix-pricing logic from `../parts-matrix-audit/scripts/qlik_expr.py`.
  - The target's **repair basket** is every target source whose rule is identical to its `main_repair_source` (Prescott: 100–105, 110, 200). The script prints the list; check that it looks right.
  - Each peer is priced with its own `main_repair_source` rule. Use `--sources-override key=source` to try a different source.
- `scripts/build_peer_workbook.py` builds the workbook from `results.json` (schema: `references/results_template.json`).

```bash
S=/home/user/Noah/.claude/skills/parts-matrix-peer-comparison/scripts
A=/home/user/Noah/.claude/skills/parts-matrix-audit/scripts
W=<scratchpad>/peer_<store>
ST=toyota_prescott,toyota_flagstaff,toyota_henderson,toyota_spokane   # target first
```

## Step 1: Target RO list (parts app state)

Part lines have no labor type, so the basket is restricted to ROs whose part-bearing lines are all the target's matrix labor type. This is the same as audit Step 3:

```bash
python3 -I $A/qlik_expr.py --division "<TARGET DIV>" --year 2026 --months Aug,Sep,Oct --bundle rolist --matrix-types CPT --out $W/rolist.json
```

1. Run it on the RO app (`c3efc739-b063-47dd-a4d3-d1c4cac07ad0`).
2. Load the RO list into parts-app (`d3d25914-349d-4a55-857c-645fccefbecc`) state `tgt`, field `Parts_Detail.referno`.
3. **Check** that `stateCounts.selected` = `n`.
4. Re-check `Count({tgt} distinct Parts_Detail.referno)` before each batch. States drop silently.

## Step 2: Repair basket and break-even (parts app)

```bash
for b in basket scenarios; do python3 -I $S/peer_expr.py --stores $ST --year 2026 --months Aug,Sep,Oct --bundle $b --state tgt --out $W/$b.json; done
```

- **`basket`**: by unit-cost band. Returns Lines, ROs, Actual, Cost, List, the price under the target's own matrix, and the price under each peer matrix.
  - This goes into `results.basket.rows` as `[band, Lines, ROs, Actual, Cost, List, target, peer1, ...]` in `--stores` order.
- **`scenarios`**: totals for the whole basket.
  - `ROs`.
  - `Mid_<peer>`: halfway between the target and the first peer (`--mid-peer` to change).
  - `Cap_<peer>`: each line capped at RangeMin(actual, peer matrix).
  - `LinesAboveOwn` / `LinesBelowOwn`.
- **Reconcile:**
  - The basket's band ROs sum to more than `scenarios.ROs`, because an RO can appear in several bands. Use `scenarios.ROs`.
  - Actual ≥ every Cap value.

## Step 3: Repair penetration and job tickets (RO app, no state needed)

```bash
for b in jobs cpros broad; do python3 -I $S/peer_expr.py --stores $ST --year 2026 --months Aug,Sep,Oct --bundle $b --out $W/$b.json; done
python3 -I $S/peer_expr.py --stores $ST --year 2025 --months Aug,Sep,Oct --bundle cpros --out $W/cpros_py.json   # prior year
python3 -I $S/peer_expr.py --stores $ST --year 2025 --months Aug,Sep,Oct --bundle broad --out $W/broad_py.json
```

- **`jobs`**: Division × job category on matrix labor-type lines with parts.
  - Categories come from op-code description keywords (`JOB_CATS` in `peer_expr.py`).
  - Results go straight into `results.jobs.rows`.
  - Sanity-check the categories with a top-op-codes table per store if a store's naming looks unusual. At Prescott, BG fluid services dominate lines with $50+ of parts.
- **`cpros`**: CP ROs (any C* line) per Division. This is the penetration denominator.
- **`broad`** (optional): ROs with a matrix-type line carrying $50+ of parts cost, for this year and last.
  - It includes fluid services, so use it only as a year-over-year cross-check.

## Step 4: Optional sections

- **Price vs. list mix (`pricepos`):** needs one RO-list state per store, e.g. `--states tgt,p1,p2,p3`.
  - Henderson's month was about 2,000 ROs. Use 1 month for big stores.
- **Same-part prices:** for part numbers sold at both the target and a peer:
  - index = Σ(target qty × target avg price) / Σ(target qty × peer avg price), built with `Aggr` over `Parts_Detail.partnumber` per state.
  - Show a maintenance table and a repair table. Few part numbers overlap in one month, and oil and filters dominate the index.
- **Traffic:** run `cpros` for the same months in both years.

## Step 5: results.json → workbook → recalc

- Write `$W/results.json` following `references/results_template.json`.
- Write the `findings` in plain language and lead with the answer.
- **Don't commit `results.json` or the workbook.** They hold store cost data.

```bash
python3 -I $S/build_peer_workbook.py --results $W/results.json --out "$W/<Target> vs Peers - CP Parts Pricing <Period>.xlsx"
python3 <xlsx-skill>/scripts/recalc.py "$W/<Target> vs Peers - CP Parts Pricing <Period>.xlsx"   # status success, 0 errors
```

Never re-save with openpyxl after recalc.

**Tabs:**

| Tab | Shows |
|---|---|
| Summary | Findings and key numbers |
| Break-Even | Basket by band under every matrix; scenarios with gross given up and added ROs to break even. The labor-gross assumption is in an editable blue cell. |
| Repair Penetration | Job-category ROs per 100 CP ROs, repair parts $ and GP; optional broad YoY |
| Job Tickets | Parts, labor and total per job, vs. the peer average |
| Matrix Curves | Each main repair matrix as a multiple of list by part cost (computed from the matrix library) |
| Price vs List, Same Part Prices, CP Traffic Trend | Optional |
| Methodology | How each number was built, and limitations |

## Step 6: Clean up and deliver

- Clear every state you loaded (`qlik_clear_selections` with each stateName).
- Send the workbook with SendUserFile.
- **Lead the reply with the answer:**
  1. Is the target the expensive store, and where?
  2. What would a change cost per year, and how many added ROs to break even?
  3. Is there evidence of lost repair business?
- Offer to draft the GM email. Include a "What's in the workbook" section listing the tabs. Ask the user whether peer-store prices can be shared with the target's GM.

## Interpreting results (lessons from Prescott)

- **A steep matrix on paper ≠ an expensive store.** Check what customers actually pay: maintenance at list, job tickets, and how far actual pricing sits under the store's own matrix.
- **Don't use a parts-cost threshold as the repair definition.** Use job categories. A threshold picks up BG and fluid services and menus, and it made Prescott look low on repairs when it wasn't.
- **Labor can matter as much as parts.** Prescott's brake jobs had the lowest parts $ and the highest labor $ of the four stores.
- **Watch for short-sample estimates.** One month badly understated the break-even ($2–3K/month from October vs. about $4.2K/month from 3 months). Use at least 3 months for the dollars.
- **Traffic differences are correlation only.** Say so.
