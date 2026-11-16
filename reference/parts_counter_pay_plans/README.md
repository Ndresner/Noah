# Parts Counter Pay Plans (reference)

Plan terms from 11 signed parts counter pay plans on file (10 stores, Chevrolet LV has two), collected 2026-11-16 for the
Subaru Las Vegas fixed ops director, who asked for example pay plans.

**Names, signatures and employee numbers are removed on purpose. Never add them.** The source PDFs are not in the repo.

## What's here

- `plans.json` holds, for each plan:
  - the plan terms: base, commission rate and basis, bonus tiers, other pay, draw
  - the compensation schedule, worded as on the original
  - reviewer notes where the original contradicts itself
  - the Jan–Oct 2026 average monthly parts gross snapshot used for estimates (Qlik Productivity, `PL_Line = 'Total Gross Profit - Parts'`)
  - `qlik_division`, so you can refresh that snapshot
- `pay_structures` in the same file summarizes the six ways counter staff are paid across the group.
- `retired_plans` lists plans that are no longer used and are kept out of the examples. One so far: the July 2023 Chevrolet LV objective plan.

## Rebuild the workbook

```bash
python3 tools/pay_plans/build_parts_counter_pay_plans.py --out <scratchpad>/Parts_Counter_Pay_Plan_Examples.xlsx
python3 <path to recalc.py> <scratchpad>/Parts_Counter_Pay_Plan_Examples.xlsx 60
```

Use `--benchmark-gross` and `--benchmark-label` to estimate pay at a different store's volume. Write the output to the
scratchpad only and never commit it.

The workbook has three tabs:
- Plan Comparison: plans side by side, with estimated pay at each plan's own store and at the benchmark store
- Pay Structures
- Plan Documents: each compensation schedule laid out like the originals, one plan per printed page

## Keep in mind

- **Estimates are illustrations, not payroll.** They leave out individual production, CSI, shop hours, YoY bonuses,
  car allowances and spiffs. Use ADP for what people were actually paid.
- **Plans can change.** Chevrolet LV moved to hourly on 10/25/2026, for example. Before reusing these as examples, confirm
  a plan is still current with the store. When a plan is replaced, move the old one to `retired_plans` instead of deleting it.
- **Problems in the originals:**
  - Honda Centennial Hills: $1,000 vs $1,250 salary, and the copy on file is unsigned
  - Audi Henderson: CSI tier wording doesn't match ("at or above 4.60" vs a first tier of 3.60–4.69)
  - Toyota Spokane: lists both a salary and a draw
  - Honda Henderson: still on the 2021 template
