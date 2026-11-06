# Optional playbook: "Are our maintenance parts priced too aggressively?"

Use this only when a store asks this specific question. It is usually after a matrix change, and often comes with "brake jobs look expensive." It's not part of the standard audit.

Reference run: Findlay Hyundai Prescott, Nov 2026 (May–Jul vs. Aug–Oct 2026). The matrix is on file at `reference/parts_matrices/hyundai_prescott/`.

## The question behind the question

A cost-based matrix applied to **every** source prices cheap, fast-moving maintenance parts far above list. With Hyundai list at cost × 1.61:
- cost + 475% on parts under $3 is **3.6x list**;
- cost + 250% on $7–10 parts is **2.2x list**.

There are three separate things to check:
1. What the matrix **says**.
2. What customers **actually paid**. The counter often overrides the matrix.
3. Whether anything **outside** the matrix drives the complaint. At Hyundai Prescott, pads sold at a flat $140.95.

## Steps

1. **Parse the matrix** (`parse_matrix.py`) and save it to the matrix library.
   - Look for one rule on all sources, and for sources with no customer-pay rule.
   - **Compute the brand's list/cost ratio** from a few parts (Hyundai 1.6125). Divide each band's cost multiplier by that ratio to get "x list." Without this step, a matrix of "cost + 200%" doesn't sound alarming to a GM.
2. **Find the change date** (audit SKILL.md, "when did the matrix change").
   - Score all RO lines on the store's main parts sale account against the current matrix by month, then by `Week(Parts_Detail.opendate)`.
   - The step change dates it. The level stays low because warranty, internal and complimentary lines are priced by other rules.
3. **Profile labor types** and confirm with the user.
   - Watch for factory-paid maintenance programs: Hyundai Complimentary Care = `CCCE`; Toyota Care and similar programs elsewhere.
   - They are **not** customer-priced. They usually carry most of the store's maintenance volume, which is why the customer-pay dollars end up small.
4. **RO list.** Use ROs with customer-pay matrix-type parts and **no** factory-maintenance, extended-warranty, employee or internal parts lines:
   `P(CP type) - P({"CCCE","CEX","CEMP","I*",...})`
   - Strict CP-only lists can be very small at low-CP stores. Hyundai Prescott had 271 ROs in 6 months, and allowing warranty lines on the RO gave 391.
   - Maintenance categories on those ROs are effectively customer pay.
   - Load the list as a state, and **check the count in the same object** as the results: `Count(TOTAL {state} distinct Parts_Detail.referno)`.
5. **Price by part category, before vs. after**, using one table:
   - Dimension: the category expression below.
   - Measures: lines, sale, list, cost for each period, plus the full-matrix price for the after period.
   - Report sale/list and GP% per category, the maintenance total, and full-matrix/list (the exposure if overrides stop).
6. **Brake line detail.** List every pad and rotor line: RO, part, cost, list, unit sale and matrix price. Flat prices show up immediately here.
7. **Fast movers.** List the top customer-pay part numbers on matrix sources with current price vs. list vs. matrix. This becomes the candidate list for a maintenance source. Add an editable proposed-price column.
8. **Options** on the after-period maintenance basket:
   - today;
   - at list;
   - list + editable markup;
   - full matrix.

   Show gross per quarter and per year. Volume response isn't modeled.
9. **Traffic context** (RO app): customer-pay ROs, brake ROs and CP parts GP%, for the same months in prior years. Check whether any decline started **before** the change.

## Category dimension (Parts_Detail.description keywords; adjust per brand)

```
=If(Match(Num#(Trim(Parts_Detail.source)),<list-priced fluid source>),'Bulk oil (list)',
 If(Match(Num#(Trim(Parts_Detail.source)),<sources with no CP rule>),'No CP rule',
 If(WildMatch(Upper(Parts_Detail.description),'*PAD KIT*','*BRAKE PAD*','*PAD SET*','*PADS*','PAD-*','*PAD ASSY*'),'Brake pads',
 If(WildMatch(Upper(Parts_Detail.description),'*DISC-FR*','*DISC-RR*','*ROTOR*','DISC ASSY*','DISC-*','*BRAKE DISC*'),'Rotors',
 If(WildMatch(Upper(Parts_Detail.description),'*OIL FILTER*','*FILTER ASSY-ENGINE OIL*','*SERVICE KIT-OIL*','*FILTER-OIL*','*OIL PLUG*'),'Oil filter / kit / drain gasket',
 If(WildMatch(Upper(Parts_Detail.description),'*AIR CLEANER*','*AIR FILTER*','FILTER-AIR','FILTER-AIR *'),'Engine air filter',
 If(WildMatch(Upper(Parts_Detail.description),'*AIR CONDITIONER*','*CABIN*'),'Cabin filter',
 If(WildMatch(Upper(Parts_Detail.description),'*WIPER*'),'Wipers',
 If(WildMatch(Upper(Parts_Detail.description),'*SPARK*'),'Spark plugs',
 If(WildMatch(Upper(Parts_Detail.description),'*BATTERY*'),'Battery',
 If(WildMatch(Upper(Parts_Detail.description),'*BELT-V*','*RIBBED BELT*','*SERPENTINE*'),'Belts','Other parts')))))))))))
```

Pads must come before rotors, because "PAD KIT-FRONT DISC BRAKE" contains DISC.

## Recommendation framework (what the user agreed with)

- **Maintenance source for fast movers**, priced at list. Use list + 10% for pads, rotors, belts and batteries. This is the standard fix and what most stores do.
- **Replace flat prices** such as a one-price pad kit with per-part pricing or a vehicle-specific brake menu.
- **Add a low-cost floor to the repair matrix.** Bands under $10 cost at 3.5–5.75x hit small repair parts too.
- **Frame the dollars.**
  - Gross given up at list is usually small, because factory-paid maintenance carries the volume.
  - The bigger risk is full enforcement of the matrix on maintenance, which would be about 1.8x list at Hyundai Prescott.
- **Other angles to offer:**
  - total job ticket vs. a sister store and local competitors;
  - follow-on CP sales from complimentary-care visits;
  - override rates by advisor or counter person;
  - declined brake recommendations, if available.

## Gotchas from the reference run

- States dropped mid-run, and the RO app came back with `Year=2026` / `MonthName=Oct 2026` selected. See audit Step 3 checks 5–6. Clear both apps and re-verify before trusting any number.
- Bulk oil and some fluid/aftermarket sources were list-priced or had no CP rule. Show them separately so they don't dilute the maintenance ratio.
- Don't commit the workbook. Build it in the scratchpad with an ad hoc openpyxl script, then run recalc. Hyundai Prescott tabs: Summary, Maintenance Pricing, Brake Lines, Fast Movers, Options, Change & Trend, Current Matrix, Methodology.
