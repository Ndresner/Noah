# Parts Pricing Matrices

CDK **"Price Breaks by Source"** exports for each store. These are the inputs to the `parts-matrix-audit` skill (`.claude/skills/parts-matrix-audit/`).

`manifest.json` lists, for each store:
- the current file,
- its Qlik `Division`,
- the customer-pay price field,
- confirmed matrix labor types,
- any notes, and
- the version history.

## Layout

```
reference/parts_matrices/
  manifest.json
  <store_slug>/
    YYYY-MM-DD_price_breaks_by_source.xlsx   # raw export as received (date = date received)
    YYYY-MM-DD_matrix.json                   # parsed by parse_matrix.py
    YYYY-MM-DD_source_descriptions.xlsx      # optional
```

## Current files (Nov 2026)

| Store | Qlik Division | CP price field | Received | Notes |
|---|---|---|---|---|
| Toyota Prescott | TOYOTA PRESCOTT | PRICE6 | 2026-11-04 | Matrix labor type CPT |
| Toyota Flagstaff | TOYOTA FLAGSTAFF | PRICE6 | 2026-11-05 | |
| Toyota Henderson | TOYOTA HENDERSON | **PRICE10** | 2026-11-05 | Source 200 has a LIST rule only |
| Toyota Spokane | TOYOTA SPOKANE | PRICE6 | 2026-11-05 | Sources 140–145 price off PRICE7 (not in Qlik), so they are untested |
| Chevrolet Las Vegas | CHEVY LV | **PRICE10** | 2026-11-05 | Matrix labor type CP. Change went live 2026-02-02 |

## Adding a newer version

1. **Save the new export** as `<store_slug>/<date received>_price_breaks_by_source.xlsx`. Never overwrite an older file; old versions are the only record of what the matrix used to be. That record mattered for Chevy LV, where there was no copy of the pre-change matrix.
2. **Parse it:**
   ```bash
   python3 -I .claude/skills/parts-matrix-audit/scripts/parse_matrix.py --matrix <file> [--sources <descriptions>] --out <store_slug>/<date>_matrix.json
   ```
   Read any warnings.
3. **Diff it against the previous parsed JSON.** Note which sources and bands changed and record that in the store's `notes`.
4. **Update `manifest.json`:**
   - set `current_*` and `received` to the new file;
   - append the new file to `history`.

These are internal pricing policies. Keep this repo private. Never store cost or sales output (workbooks, results JSON) here; that stays in the scratchpad.
