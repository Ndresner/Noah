---
name: parts-matrix-audit
description: Audit whether a Findlay store's customer-pay parts on service ROs are priced per its parts pricing matrix (CDK "price breaks by source" export), using Qlik Closed Parts Tickets + Closed Repair Orders, and build the Parts Pricing vs. Matrix workbook (summary, labor-type review, results by source, line exceptions, matrix reference, methodology). Use whenever the user asks to check parts pricing against the matrix, run a parts matrix compliance / parts pricing audit for a store, see if parts are "hitting the matrix", or attaches a parts pricing matrix export for any store or brand. Do not hand-build it; run the bundled scripts.
---

# Parts Pricing vs. Matrix Audit

Answers one question: **are parts sold on customer-pay ROs priced the way the store's parts matrix says they should be?** Each part line is tested against the matrix price for its source. The store's own matrix export is the rule book.

Reference run: Toyota Prescott, Aug–Oct 2026.
- CPT hit the matrix on 96.7% of lines (7,964 of 8,235 lines on 1,939 ROs).
- Actual $396,262.51 vs. matrix $411,345.16, a net of −$15,083.
- Use these numbers to check that a re-run of Prescott reproduces them.

> **Store asking whether its maintenance parts (filters, pads, rotors) are priced too aggressively?** That's an optional one-off review, not part of the standard audit. Follow `references/maintenance_pricing_review.md` (reference run: Hyundai Prescott, Nov 2026).
>
> **Comparing against other stores?** If the user asks whether this store is more expensive than sister stores, or what lowering the matrix would cost, use `.claude/skills/parts-matrix-peer-comparison/` instead (or after this audit). It reuses this skill's scripts.

## Inputs

1. **Parts pricing matrix export** (.xlsx, required). This is the CDK "Price Breaks by Source" report. Columns: Source | Price Field | Price Break Field | Price Break Range | Formula. The layout is the same for every brand.
2. **Parts source description export** (optional). Qlik's `Source.description` holds the same names per store. It was identical to the export for Prescott. The export only adds sources that had no sales in the period, which shows up in the "sources with no matrix row" note.
3. **Store**: the Qlik `Division` value, for example `TOYOTA PRESCOTT`. Look it up with `qlik_search_field_values`. Filter by Division, never by Store (CLAUDE.md rule).
4. **Period**: default to the last 3 full closed months, so results reflect the current matrix. For very large stores, 1–2 months keeps the RO list manageable (see Step 3).

## How the matrix reads

These rules are encoded in `parse_matrix.py`. Don't re-derive them.
- `COST + 275.00` = cost × 3.75 (cost plus 275%).
- `LIST + 30.00` = list × 1.30.
- `LIST` = list.
- `COST - 10.00` = cost × 0.90.
- The band is chosen by testing the Price Break Field (COST = unit cost, LIST = list) against the range.
- Bands are not in order in the export. For example, `0.00 - 0.00` sits mid-list.
- The CDK print repeats the source number at page breaks. That is the same source, not a new one.
- Sources with no matrix row get the system default price. They show as "No matrix" and are not tested.
- Any formula the parser doesn't recognize is reported as a WARNING, and that source is left untested. If a new brand's export shows a new formula shape, extend `FORMULA_RE` rather than hand-pricing.
- **Several price fields.** A source can carry a rule that sets the **LIST** field (how list is filled in, e.g. "$0 list → cost +67%") and a separate rule that sets the customer-pay field (e.g. `PRICE10`). This happens at Chevrolet LV.
  - The parser keys blocks by source + price field and tests the price field that most sources use. Override it with `--price-field`.
  - Other-field rules are kept under `other_price_fields`. They are not customer-pay pricing.

The math was proven on Prescott, matching to the cent:
- source 100, cost $648.89 → COST+170% = $1,752.00
- source 100, cost $1,562.84 → List $2,049 +5% = $2,151.45

## Step 0: Check the matrix library first

Matrices on file live in `reference/parts_matrices/`; `manifest.json` lists each store's current file.
- If the store is there, use its current parsed JSON and its noted CP price field, and skip asking for the export. Do confirm with the user that the file is still current.
- When the user sends a new or updated export, add it to the library following that folder's README:
  - dated file, never overwrite;
  - parse it, diff it against the prior version, and update the manifest;
  - commit it.

## Step 1: Parse the matrix

```bash
S=.claude/skills/parts-matrix-audit/scripts; W=<scratchpad>/<store>
mkdir -p $W
python3 -I $S/parse_matrix.py --matrix "<matrix export>.xlsx" [--sources "<source export>.xlsx"] --out $W/matrix.json
```

Read any WARNINGs before going on. One to watch for: if more than one Price Field appears, ask which one is the customer-pay price.

## Step 2: Profile labor types (Closed RO app) and confirm the matrix types with the user

Run `qlik_clear_selections` on both apps first.

Generate the call arguments and pass them to `qlik_create_data_object` on the Closed RO app `c3efc739-b063-47dd-a4d3-d1c4cac07ad0`:

```bash
python3 -I $S/qlik_expr.py --division "<DIV>" --year 2026 --months Aug,Sep,Oct --bundle profile --out $W/a_profile.json
python3 -I $S/qlik_expr.py --division "<DIV>" --year 2026 --months Aug,Sep,Oct --bundle opcodes --matrix-types <all C types, comma list> --out $W/a_opcodes.json
```

- `profile` returns one row per labor-type combination on part-bearing lines, with ROs, parts sale and parts cost.
- `opcodes` returns the top op-code descriptions per labor type, which tell you what each type is.
- Customer-pay types start with **C**.

Classify each C type into one role, then **confirm the classification with the user before testing**:

| Role | Meaning | Prescott example |
|---|---|---|
| `matrix` | Should follow the matrix. These are the types tested line by line. | CPT |
| `list` | Sells at list / MSRP | CTX, CTXX, CPTA |
| `reimbursed` | Factory-paid or reimbursed rate | CPTC (ToyotaCare, ~20% margin) |
| `package` | Package or menu pricing | CPTE (Green/Red service, cost +14%) |
| `discount` | Fleet, employee or prepaid discount | Chevy LV CPF/CPF1 (fleet), CEMP (employee) |
| `unclear` | Doesn't fit a role; ask the user | CTXO |
| `mixed` | ROs with 2+ C types on part lines; not testable line by line | 444 ROs at Prescott |

Gross margin % is a useful hint, not proof:
- Matrix types run high (CPT 45%).
- Reimbursed and package types run low (12–20%).

Labor-type codes differ by brand and store, so never assume Toyota codes elsewhere.

## Step 3: Build the RO list and load it into the parts app

Part lines in Qlik carry **no labor type**. The link is the RO: test only ROs whose part-bearing lines are **all** matrix types.

```bash
python3 -I $S/qlik_expr.py --division "<DIV>" --year 2026 --months Aug,Sep,Oct --bundle rolist --matrix-types CPT --out $W/a_rolist.json
```

1. Run `a_rolist.json` on the RO app. It returns `n`, `PartsSale` and `ROs`, a comma-separated list.
   - **Check:** `n` must equal the matrix-type row's RO count from Step 2, and `PartsSale` must match.
2. Load the list into an **alternate state** in the parts app. Call `qlik_select_values` with appId `d3d25914-349d-4a55-857c-645fccefbecc`, stateName `cpt` (any short name), field `Parts_Detail.referno`, and values = the RO list.
   - **Check:** `stateCounts.selected` must equal `n`.
   - A "conflict" warning with the full count selected is harmless.
3. This is the expensive step. The RO list is copied out of one tool result into the next call. Prescott's 1,939 ROs are about 15K characters each way. For a store with more than ~4,000 matrix ROs, run 1–2 months instead of 3.
4. If there's a big gap between the RO app's matrix parts sale and the parts app's `Actual`, stop and investigate before going on.
5. **Alternate states can silently drop** when the Qlik engine session resets, which happened at Chevy LV. When that happens, every `state::field` restriction quietly means "all ROs", and totals balloon.
   - Before each batch of queries, and after any long pause, check `Count({cp} distinct Parts_Detail.referno)`. It must equal `n`.
   - If it doesn't, reload the list.
6. **After a session reset the RO app can come back with default selections** (seen Nov 2026: `Year=2026`, `MonthName=Oct 2026`). Set analysis on `Year`/`Month` does **not** override a `MonthName` selection, so totals quietly shrink to one month.
   - Run `qlik_clear_selections` on both apps whenever a state has dropped, and check that the result's `selections` list is empty.

## Step 4: Matrix test (parts app)

Generate each bundle with `--state cpt`, then run it on the parts app:

```bash
for b in totals summary classify under over; do
  python3 -I $S/qlik_expr.py --matrix $W/matrix.json --division "<DIV>" --year 2026 --months Aug,Sep,Oct --bundle $b --state cpt --out $W/a_$b.json
done
```

| Bundle | Gives | Size |
|---|---|---|
| `totals` | Lines, ROs, Actual, Matrix, List, Cost, MatchLn (whole group) | ~7K chars |
| `summary` | The same by source, plus `Source.description` | ~7K |
| `classify` | Under/Over lines and $, under lines sold at list, below-cost lines and $, $0-list lines, by source | ~25K. It can be split into two calls (counts, then $) |
| `under` / `over` | Top 45 line exceptions each way (`--top N` to change) | ~7K |
| `all_exc` | Every exception line, for when the user wants the full list | Page it with `qlik_get_chart_data` and drop rows with null/0 `VarExt` |

Pass the file's `dimensions`, `measures`, `sort` and `limit` exactly as written. Never hand-edit the matrix expression.

**Reconcile:**
- `totals.Actual` = the sum of `summary.Actual`.
- Under$ + Over$ ≈ Actual − Matrix, within cents of rounding.
- Spot-check 2–3 exception lines by hand with `parse_matrix.unit_price`.

## Step 5: Results JSON → workbook → recalc

Write `$W/results.json` following `references/results_template.json`.
- `labor_types` comes from Step 2, with the confirmed roles plus short notes.
- `mixed` is the RO count and parts sale of the multi-C-type ROs.
- Add store-specific `findings` and `questions` you discovered, for example "$0 list on brake pads" or "RO 403504 repricing". The script already writes the standard findings automatically.

**Don't commit `results.json` or the workbook.** They hold store cost and pricing data; keep them in the scratchpad.

```bash
python3 -I $S/build_workbook.py --results $W/results.json --matrix $W/matrix.json --out "$W/<Store> CP Parts Pricing Audit <Period>.xlsx"
python3 <xlsx-skill>/scripts/recalc.py "$W/<Store> CP Parts Pricing Audit <Period>.xlsx"   # must be status success, 0 errors
```

Never re-save with openpyxl after recalc.

Tabs:
1. **Summary**: KPIs and auto findings.
2. **Labor Type Review**
3. **Matrix Test by Source**: Match % shading is green ≥95%, yellow 85–95%, red <85%.
4. **Top Exceptions**: each line carries a likely cause:
   - sold below cost
   - $0 list in the parts master
   - sold at list instead of matrix
   - priced far above list
   - manual override
5. **Matrix Reference**
6. **Methodology**

### Recommended extras before sending (all done for Chevy LV)

- **GP by pricing outcome.** Use sales lines only (`Parts_Detail.quantitysold={">0"}`). Split each source's lines into hit-matrix / sold-at-list / other, with sale, cost and GP% for each.
  - For sources whose matrix *is* list, the at-list lines are matrix hits; don't double count them.
  - This answers "what GP do parts priced per the matrix actually earn?"
- **Sales lines only.** Core and return credits (negative quantity) distort over/under. Report compliance and under/over on sales lines.
- **Traffic vs. price.** CP ROs with parts and parts sale/gross per RO, year over year (RO app). This keeps a car-count decline from being blamed on pricing.
- **Menu check** for list-priced maintenance sources:
  1. Build a second RO list of matrix-type ROs with **no** menu or maintenance op code on any part-bearing line. Use `P()` minus `RO_Detail.opcode={...menu codes...}`.
  2. Compare the at-list share and dollars on menu ROs vs. repair-only ROs.
  3. Only the repair-only part is a likely miss.

### Optional: "when did the matrix change, and what did it do to gross?"

This was done for Chevrolet LV, Nov 2026, where the change went live 2/2/2026. The method:

1. **Find the date.** No RO list is needed. Score every RO part line (`invsaletype='RO'`, all pay types) against the **current** matrix:
   - by month, with `Year`/`Month` as dimensions, then
   - by source × month around the suspected window, then
   - by `Parts_Detail.opendate` (the line's open date) across about 3 weeks.

   A source whose bands changed goes from roughly 0% to most lines matching. Warranty and internal lines dilute the absolute rate, so look for the step change, not the level. Qlik history for this app starts about May 2025.
2. **Realized gross trend.** Use the Closed RO app, `RO_Detail.labortype` = the matrix type, parts sale and cost by `Year`/`Month`. Compare the same months year over year, and the 3 months before vs. 3 months after the change. Split the result into a rate effect and a volume effect.
3. **Before/after on the matrix sources.** Build a second CP-only RO list for the 3 months before the change and load it as state `cppre`. Then compare GP% **within unit-cost bands**, pre vs. post. Post gross at pre-change GP% gives the mix-adjusted dollar effect. Raw GP% swings are mostly mix, especially from parts over $1,000.
4. **Old matrix.** Infer it as the most common sale/cost multiple per band on pre-change CP lines. Present it as an estimate and report how many pre-change lines it matches (only about 25% at Chevy LV).

Put the timing, trend and before/after tables in `results.json` → `extra_sheets`. Each entry has `name`, `title`, `subtitle`, `widths` and `tables`; each table has `columns`, `rows`, and optional `formats`, `highlight_rows` and `note`. The builder adds one tab per entry.

## Step 6: Clean up and deliver

- Clear the alternate state: `qlik_clear_selections` with stateName `cpt`.
- Send the workbook with SendUserFile.
- Lead the reply with the compliance %, the net $ and $/month, and the top 2–3 sources driving the variance.
- List the labor-type roles and what was not tested (mixed ROs).
- Offer the full exception list (`all_exc`).

## Patterns seen at Prescott

Look for these at every store:
- **Aftermarket or non-OEM source sold at vendor list** instead of cost-plus. At Prescott this was 64% of the under-matrix dollars. Often the matrix itself is unrealistic: oil at $7.17 cost marked up to $32 a quart against a $13.96 list. Frame it as "fix the matrix or fix the behavior."
- **OEM stock overridden to list.** Here the matrix priced above list on most parts.
- **$0 list price in the parts master** on sources whose matrix is LIST. The system prices them at $0, and the counter keys a price by hand. This is a setup fix.
- **Prices shuffled within one RO**: some lines below cost, others many times list. This is a quote-to-total behavior, so name the RO.
- Fractional-quantity fluids (oil and fluid sources) show tiny under-matrix amounts. That's rounding noise, so don't headline it.

## Patterns seen at Chevrolet LV (Nov 2026)

- **Maintenance sources sold at flat list.** Bulk oil, oil and air filters, chemicals and "no touch" sold at list on 90–99% of CP lines while the matrix says list +X%. That was −$88K per quarter vs. matrix. It's likely menu pricing, so ask before calling it a miss.
- **Core and return credits** (negative quantity, credited at cost) show as large "over-matrix" lines on OEM sources. The builder labels them and leaves them out of the repricing finding.
- **A "more aggressive" matrix can actually be lower.** Check direction by band against pre-change CP prices before assuming the change raised prices.
