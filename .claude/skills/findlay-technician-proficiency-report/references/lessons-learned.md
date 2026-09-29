# Lessons learned (read before editing the workbook or the scripts)

Every one of these was a real mismatch against the user's original example, found by
diffing against it directly -- not guessed. If you're modifying `scripts/`, preserve all
of these; if you're tempted to hand-edit the output workbook instead of the scripts, check
here first since the issue is probably already solved in the scripts.

## 1. ADP Position ID -> Tech No parsing

Every Position ID in the roster export is exactly 9 characters: a **3-character store
prefix** + a **6-digit tech number** (leading zeros dropped, matching how Qlik's
`RO_Detail.techno` field stores them). The correct parse is:
```python
tech_no = str(int(position_id[3:]))
```
**Do not assume a 4-character prefix.** An earlier version of this script stripped 4
characters, which happened to work for a few stores (Cadillac, Acura, Chevy Gmc Bullhead)
purely because their true tech numbers start with a redundant leading zero after the
3-char prefix -- but it silently truncated a real leading digit for other stores (Lexus
Spokane's `680024`, Toyota Spokane's `600020`, etc.), misclassifying real named
technicians as "unmapped/pooled" codes.

If you ever suspect this is wrong again, verify empirically per store rather than assuming:
for each division, try stripping 2/3/4/5 characters from each roster Position ID, convert
the remainder to `int()`, and see which strip length produces the most matches against
that division's actual Qlik `techno` set. Strip-3 was correct for all 34 stores when this
was checked exhaustively.

## 2. Sort order within each category

Technicians within a category (Service Technician, Express Technician, etc.) are sorted by
**Proficiency % (Sold Hours / Actual Hours) descending**, not by Sold Hours. This is easy to
get wrong because a sold-hours sort often *looks* plausible until you check the actual
ratio. The "EXCLUDED / POOLED CODES" block at the bottom of each store sheet is the one
exception -- that one *is* sorted by Sold Hours descending, since proficiency isn't a
meaningful concept for pooled/non-named codes (no actual_hours to divide by).

Tech Ranker's row-reference formulas depend on this same sort order lining up with the
actual row positions on each store sheet -- if you change the sort key in
`store_sheet_builder.py`, you must change it identically in `assemble_workbook.py`'s Tech
Ranker construction, or the two will point at mismatched rows.

## 3. Merged cells carry the fill, not every member cell

Several template rows only need `fill` set on the merge's anchor (top-left) cell --
Excel renders the whole merged range in that color once the range is actually merged via
`ws.merge_cells(...)`. Setting fill on only the anchor without merging leaves the rest of
the "band" white with no color; the reverse (filling every individual cell without merging)
works visually but doesn't match the template's actual XML structure. Places this matters:

- Title row (`A1:L1` on store sheets, `A1:T1` on Capacity) and subtitle row (`A2:...`)
- Store-sheet divider row between each category's GROUP SUBTOTAL and the next category:
  merge `A{row}:M{row}` and separately `O{row}:S{row}`, both filled navy -- **except** the
  very last category before TOTAL/AVERAGE, which gets a plain unmerged/unfilled blank row
  instead.
- Store-Wide What-If Scenario block: header merged `A:E` (no fill), "Metric"/"Value" row
  merged `A:D` (navy fill) + `E` separately (navy fill), each metric label row merged `A:D`.
- Capacity sheet: `A4:A5` ("Store", spans both header rows), `B4:I4` ("Stall-Based Capacity
  Calculation"), `J4:O4` ("Tech-Based Capacity Calculation"), `P4:T4` ("Dashboard
  Comparison"), plus `A1:T1` / `A2:T2` for title/subtitle.

## 4. Column N (the spacer between the two tables) must never be filled

On every store sheet, column N sits between the main proficiency table (A:M) and the
What-If table (O:S). It's meant to be permanently blank/unfilled so the two tables read as
visually separate. An earlier version filled N gold on GROUP SUBTOTAL rows and navy on the
TOTAL/AVERAGE row (to match B and C, which *should* carry those fills) -- this visually
bridged the two tables together. Never include `"N"` in the subtotal/total row fill-column
lists.

## 5. Border color on the Capacity sheet is light gray, not default black

Every header and data cell on the Capacity sheet uses `Side(style="thin", color="CCCCCC")`
-- a light gray. Using openpyxl's default thin border (which renders black) looks
noticeably "harder"/darker than the original and was flagged as looking wrong even after
the borders were structurally correct otherwise.

## 6. Capacity sheet number formats

Every dollar column needs explicit `\$#,##0` (or `\$#,##0.00` for ELR specifically), and
percent columns need `0.0%` -- `set_cell()` defaults to `General` if you forget to pass
`number_format`, which is easy to miss across ~20 columns. Cross-check the full column list
against `references/workbook-structure.md` if adding new columns.

## 7. "Current Store Proficiency %" on each store's What-If block

This must match the exact same population and rules as row 2 of the Proficiency sheet --
i.e. `qualifies_for_prof_subset()` in `style_helpers.py`: only Service Technician, Express
Technician, Service Team Leader, Shop Foreman, **and** excluding outlier records (Shop
Foreman/Service Team Leader at or below 20% proficiency, Service/Express Technician at or
below 2%). This is a computed Python value baked into the cell, not a live formula --
because outlier exclusion is per-technician, it can't be expressed as a simple
range-SUM formula the way the rest of the sheet is. If the underlying data changes, this
cell needs a full rebuild, same as the Proficiency sheet itself.

Earlier, simpler versions of this cell (that got superseded by the above) were: (a) a
straight reference to the TOTAL/AVERAGE row (wrong -- includes Detailer/Body Shop/Tinter),
then (b) a SUM of just the Tech-Ranker-eligible category subtotal rows with no outlier
exclusion (closer, but still not matching the Proficiency sheet's methodology). Only (c),
the current per-technician filtered calculation, is correct.
