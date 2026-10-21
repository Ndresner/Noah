# Findlay Day 2 Fixed Ops Deck: project brief for Claude Code

Read this first. It holds what the code can't tell you: the meeting, the data rules, the house style, the decisions already made, and what's still open.

## The meeting
- **Event:** Findlay Automotive Group GM & Service Manager meeting. Day 1 is Oct 28, 2026. **Day 2, the fixed ops day, is Oct 29, 2026, 8am–4pm.**
- **Audience:** GMs, service managers and fixed ops managers. There are no parts-only managers; some managers run parts and service together. This is the first meeting dedicated to fixed ops for the GMs.
- **Owner:** Noah (Fixed Ops analytics, ndresner@findlayauto.com). He reviews with **Tommy** (also written Tray or T-Rail; call him Tommy), who presents and makes the style rulings.
- **Direction (Oct 20):** shorten the slide count significantly. The priority is **dialogue in the room, not data volume**: the room doesn't know the numbers as deeply as we do. Keep every slide in the test deck for now, so Noah and Tommy can choose what to cut. Don't change the format.

## Folder map
| Path | What it is |
|---|---|
| `build/house.mjs` | **The engine.** Day 1's engine (Tommy's vault `house.mjs`) with the Day 2 renderers merged in: `cover`, `agenda`, `section` (with tracker), `groupYoy`, `levelRanker`, `ranker`, `totalRanker`, `segmentRanker`, `statTiles`, `amountRanker`, `dumbbell`, `talkingPoints`, `explainer`, `recordBoard`, `gapRanker`, … The look lives here. The pre-merge engine is `build/v1/house.v1.mjs`. |
| `build/blocks.mjs` | A metric block in the Day 1 sequence: group picture (`groupYoy`: card over the monthly chart), level ranker (or segment / total ranker), change ranker. `load()` reads `{m25, m26, S}` JSON. |
| `build/*-deck.mjs` | One script per deck (see the build table below). Each writes a `.pptx` into `build/`. |
| `build/opening-deck.mjs` | Cover, agenda, the eight section transitions and the Open Decisions divider. **The section list lives here once**; agenda, kickers and tracker derive from it. |
| `build/merge_v2.py` | Assembles the **v2 test deck** `decks/Fixed Ops - Day 2 Example Test Deck.pptx` (206 slides) from the rebuilt decks. The flow order is defined here. |
| `build/merge_test.py` | The v1 assembler (159 slides). Kept for reference; the v1 output is archived. |
| `decks/v1 - house style/` | **The v1 archive:** all 28 v1 decks byte-for-byte, including the 159-slide v1 test deck (snapshot commit `205030e`). Don't edit. |
| `build/swap_prod.py` | Replaces master slides 74–88 (Productivity) with the rebuilt Productivity deck. |
| `build/prof/` | Technician proficiency data and calculation (`compute.py` plus the helpers from the `findlay-technician-proficiency-report` skill). |
| `build/imp/` | Most Improved and Who's Moving Now data (`compute.py`, `momentum.py`, raw Qlik pulls). |
| `build/hc/` | ADP vs Qlik headcount investigation and `build_prod.py` (productivity on ADP headcount). |
| `build/prod/` | Productivity data built by `hc/build_prod.py`. |
| `inputs/` | The ADP Employee Count workbook Noah provided, plus the saved Qlik pulls `build_prod.py` reads. |
| `decks/` | Every current `.pptx`, plus the CSV of people missing from Qlik. |
| `docs/` | The house style guide and source notes from the claude.ai project. |
| `reference/` | Text of the prior decks used for the talking points: Nov 2024 Roundtable, Jan 2026 and May 2026 GM meetings. |

## Not in git (employee-level data)
`.gitignore` keeps these out of the repo: `build/prof/roster.xlsx`, `role_adj.json`, `tech_data.json`, `tech_gross.json`, `sold.json`, `sold_raw.json`, `bad.json`, `build/hc/diff.json`, the ADP Employee Count workbook and the Qlik headcount-names pull in `inputs/`, and the ADP-not-Qlik CSV in `decks/`. Ask Noah to re-upload them, or re-pull them from ADP and Qlik, before rerunning `prof/compute.py` or `hc/build_prod.py`. Store- and group-level JSON is committed, so every deck script still builds without them.

## Setup
```
npm install            # pptxgenjs 4.0.1
pip install -r requirements.txt --break-system-packages   # python-pptx 1.0.2, openpyxl
```
- **Inter must be installed** on any machine that builds, opens or presents the decks.
- For thumbnails and checks: LibreOffice (`soffice --headless --convert-to pdf`) and `pdftoppm`.

**Run everything from `build/`.** Node scripts write their `.pptx` into `build/`. Copy them to `../decks/` before running `merge_test.py`, which reads only `decks/`.

## Build table
| Script | Output in `decks/` | Data |
|---|---|---|
| `prod-deck.mjs` | Fixed Ops Productivity - House Style.pptx (15) | `prod/*.json` (run `python3 hc/build_prod.py` first) |
| `fge-deck.mjs` / `fge-gap-deck.mjs` | Fixed Gross per Employee 5 Pct / Gap - Review | `prod/fge26.txt` |
| `prof-deck.mjs` | Technician Proficiency - House Style (8) | `prof/prof.json`, `prof/opp.json` (from `prof/compute.py`) |
| `best-options.mjs` | Best Techs - Design Options (3) | prof data |
| `tp-deck.mjs` | Talking Points - Day 2 (14) | typed from verified figures (see below) |
| `explain-deck.mjs` | Metric Explainers - Review (12) | typed group figures |
| `minute-deck.mjs` | Tech Minute - Review (2) | prof data |
| `hpro-deck.mjs` | Hours per RO Value - Review (1) | `prof/hpro.json` |
| `opmiss-deck.mjs` | Op Code Missed Opportunity - Review (6) | `prof/op_raw.txt` |
| `hrs-ro-deck.mjs` | CP Hours & RO a Day - Review (3) | `prof/cphrs.json` |
| `money-deck.mjs` | Where the Money Is - Review (1) | figures from the decks above |
| `improved-deck.mjs` | Most Improved - Review (v1, superseded) | `imp/imp.json` |
| `moving-deck.mjs` | **Most Improved v2 - Review (10)** | `imp/imp.json`, `imp/mom.json` |
| `whos-moving.mjs` | Who's Moving Now - Review (10, top 10 per measure) | `imp/mom.json` |
| `gross-trend-review.mjs` | **Gross Trend - Review (11)**: Jan–Sep 2026 vs 2025 two-line options | `gross/monthly.json` (Qlik P&L, pulled Oct 20, reconciles to verified totals), `prod/trend.json` |
| `records-deck.mjs` | **Store Records - Review (6)**: Record Year, Record Board, two Spotlights, records ranker, ranker by category | `records/records.json` from `records/compute.py` (reads `inputs/September_2026_Fixed_Ops_Gross_Records.xlsx`) |
| `internal-gp-deck.mjs` | **Internal Gross Profit - Review (6)**: internal labor and parts GP%, YOY card, level ranker, change ranker | `internal/internal.json` (Qlik Closed RO, internal pay type, pulled Oct 21) |
| `cp-vs-internal-deck.mjs` | **CP vs Internal GP - Review (4)**: CP vs internal labor and parts GP% with the gap, option A (gap bars) and B (dumbbell) | `internal/cp_vs_internal.json` (Qlik Closed RO, Jan–Sep 2026) |

**Rebuilt Oct 21 (v2).** The Gross, Profitability, CP Engine, Labor Mix, Opcode Benchmarks and Warranty & Internal decks now build from re-pulled Qlik data (`gross-deck.mjs`, `pl-deck.mjs`, `cp-deck.mjs`, `labormix-deck.mjs`, `opcode-deck.mjs`, `wi-deck.mjs`, `door-slides.mjs`):
- Closed RO (set identifier `1`; every pull reconciles store rows to the group): `m5.json` (CP), `wi.json` (warranty/internal), `elr_q3.json` (Jul–Sep ELR), `all.json` (opcodes). Parser: `pulls/to_block.py`.
- P&L: `abs.json` (absorption), `ntgs.json` / `ntgp.json` (service / parts net to gross).
- Recovered from the v1 speaker notes because the raw input was lost: `door.json`, `doorinfo.json` (door rates, grid/no-grid), `wm_shares.json` (work-type shares).

**v2 build:** run each `*-deck.mjs`, then `node opening-deck.mjs`, then `python3 merge_v2.py`.

## Data sources (Qlik Cloud MCP)
| App | ID |
|---|---|
| Closed Repair Orders | `c3efc739-b063-47dd-a4d3-d1c4cac07ad0` |
| Productivity (Accounting) | `948ab4f4-bf9d-47d5-b5ba-6565499db705` |
| Profit & Loss | `bc30d66f-4564-4df0-89a7-368f2c9a931e` |

- **Clear selections first** on every app.
- **Big pulls:** `Concat(Aggr(...), '|')` gets one row per store. If a result comes back inline instead of saved to a file, add a padding measure such as `Repeat('x',3000)` to force a file save, then parse the file. Never retype numbers by hand.

### Basis rules: apply to every slide
1. **Same-store, 33 stores.** INEOS is excluded (it opened in 2026).
2. **GMC Prescott body shop is excluded everywhere.**
   - Labor types: `CPB, CPP, WB, INTB, IUTB, IUCB, ISPB, IBP, ISBS, IBSHP`.
   - In P&L terms: subtract `Gross Profit - Body Shop`. For headcount: drop the `BODY SHOP*` titles. For RO counts: remove body-only ROs.
   - Set pattern (Closed RO app):
```
{<BASE, Division-={'INEOS','GMC PRESCOTT'}> + <BASE, Division={'GMC PRESCOTT'}, RO_Detail.labortype-={'CPB','CPP','WB','INTB','IUTB','IUCB','ISPB','IBP','ISBS','IBSHP'}>}
```
3. **CP RO count = all customer-pay ROs.** That includes the ~0.9% with $0 labor; Noah wants it to match what stores see. Never write "with labor charged". `RO_Detail.hasLabor = 1` on every line, so it doesn't filter anything.
4. **Period: Jan–Sep by close date.** 2026 vs 2025. YTD ratios are YTD numerator ÷ YTD denominator, never the mean of months.
5. **Field quirks.**
   - Closed RO app: `Month={"<=9"}` works.
   - Productivity app: Month is text, so use `Month={"Sep"}`. `Month={9}` silently returns 0.
   - Single quotes for exact values, double quotes for searches and ranges.
6. **Reconcile before shipping.** Store rows must sum to the group on every input, both years. These group figures are verified:

| Measure | 2025 | 2026 |
|---|---|---|
| CP ROs | 279,300 | 279,756 |
| CP hours | 504,018.97 | 523,141.86 |
| CP ELR | $158.86 | $162.42 |
| Fixed gross (Jan–Sep) | $181,331,157 | $188,799,724 |
| CP & W ROs | 379,921 | 371,180 |

Opcode closing % for 2026: alignment 13.2%, balance 8.3%, brake fluid 9.2%, fuel 5.8%.

### Productivity headcount: ADP, not Qlik (decided Oct 20)
- **The problem:** the Qlik Productivity app's headcount history undercounts 2025 by about 170 real people a month. Sep-25: 1,451 matched, 226 in ADP only. Of those, 169 are Qlik lag, 34 are ADP transfer back-dating and 18 are Findlay Customs. By Sep-26 the sources agree.
- **The fix:** headcount now comes from the **ADP Employee Master extract** in `inputs/`. Gross and RO counts still come from Qlik.
- **Rules (`hc/build_prod.py`):**
  - Month-end snapshot: hire ≤ snapshot and (no termination or termination ≥ snapshot).
  - Qlik's Fixed Employee title list (managers included).
  - Technicians = service tech, express tech, shop foreman and team leader. Managers = parts manager, service manager, service director, service drive manager, assistant parts manager and express service manager.
  - A person open at two stores is moved **sequentially**: they stay at the old store until that record closes, then start at the new one the next day.
- **Caveat:** ADP job titles are as of the extract date.
- `decks/Headcount - In ADP but not Qlik, Sep 2025.csv` is for whoever fixes the Qlik load.
- Full write-up: `docs/productivity-headcount-source.md`.

### Technician proficiency (`prof/compute.py`)
- **Formula:** proficiency = Qlik sold hours ÷ ADP regular + OT hours. No PTO or holiday.
- **Scope:** matched on tech number (Position ID characters 3 onward), all pay types, Dec 25 2025 – Oct 9 2026. Label it **Jan–Sep**.
- **Who counts:**
  - Service and express techs only above 10%; team leaders only above 20%; shop foremen only above 30%.
- **Role-change rule:** paid hours over 125% of full time get scaled. It applies only to techs under 85% and can't lift anyone above 85%.
- **Exclusions:**
  - 45 keying-error lines removed (20+ hrs at under $25/hr).
  - Pooled codes (MULT, 999, 77777, 222 …) are credited to no one. Chevy GMC Bullhead's code 222 stays out.
- **Valuing the gap:** gap hours are valued at each store's own counted-tech labor gross, parts gross and labor cost per sold hour.
- **Results:** group 94.7%; 20 of 33 stores at 85% or better; 262 of 583 techs under 85%; gap $1,646,932 gross a month.
- **Wording:** flat-rate techs aren't paid for unflagged hours, so never say the hours are "already paid for".

### Store records (`records/compute.py`, decided Oct 21)
- **Source: Qlik Fixed Ops Accounting TB** (`e077f027`), Family FIXED GROSS, Class SERVICE / PARTS, history Jan 2021 – Sep 2026, from Noah's monthly records workbook (pulled Oct 8; September was closed, so no re-pull).
- TB matches the P&L to the dollar for every store **except GMC PRESCOTT service** (TB higher by $153,746 Jan–Sep 2026, $286,430 Jan–Sep 2025; likely body shop). GMC PRESCOTT's combined record holds on both bases. TB group Jan–Sep 2026 = $188.95M vs P&L $188.80M.
- **Stores with shorter histories count the same as everyone else** (Noah). Months before a store opened are ignored.
- Record = Jan–Sep 2026 above every prior full Jan–Sep for that store; monthly records = each new all-time monthly high in 2026.

### Other fixed facts
- **Days open:** every service department is open 6 days a week, never Sunday, and closed 6 holidays (New Year's, Memorial Day, Independence Day, Labor Day, Thanksgiving, Christmas). **307 days a year** in 2026. Never infer days open from RO close dates.
- **Luxury segment:** Lexus, Lincoln, Cadillac, Volvo, Audi Reno, Audi Henderson, LR Reno, LR LV, LR Henderson. Acura is non-luxury. GMC Prescott is the only body shop.
- **Opcode benchmarks:** alignment 15%, balance 15%, brake fluid 10%, fuel 7.5%. Closing % = ROs with the service ÷ CP ROs on the op-code report basis (`RO_Header.flagExcludeOpCodeReport = 0`).
- **ELR % of door rate target:** 70%.

## House style: enforced in `house.mjs`, keep it that way
- **Font:** Inter, 16:9 (LAYOUT_WIDE). The build stops below the **14pt floor**; store names may go to 13pt.
- **Naming:** always "stores", never "rooftops"; `words()` stops the build on the old term. Store names in caps on the rankers, Title Case in takeaway prose (`proseNames()` converts them).
- **Colour:**
  - **Green is good, red is bad, always.** Below group, segment average or benchmark is red, with no exceptions.
  - **Blue** is for neutral facts and opportunity sizes.
  - Headcount change is neutral.
- **Titles and eyebrows:**
  - Eyebrow = the segment name.
  - The title is the metric's source line name.
  - Both slides of a data pair share the title.
  - No period in the title.
- **Takeaways:**
  - **Derived from the data, never typed:** counts, names and direction ("grew"/"fell") are computed.
  - One to two lines.
  - Every chart slide has one.
- **Number formats:** percentages **truncated**, not rounded, on the improvement slides.
- **Speaker notes:** source and footnotes go in the notes, not on the slide (Tommy, Oct 2026).
- **Opportunity slides:** frame them as small, achievable increments (0.1 hr per RO, one more RO a day, a few %), not out-of-reach targets.
- **Talking-point slides:** what we said (prior meetings), then where we are now (numbers), then the ask.
- Full rules: `docs/findlay-deck-style-guide.md`. It predates this engine; where they differ, `house.mjs` and this file are current.

## Current test deck: v2, 206 slides (155 in the meeting flow + 51 in Open Decisions) (order in `merge_v2.py`)
Rebuilt Oct 21 in the Day 1 style. Cover title "Fixed Operations", audience "General Managers & Service Managers". The v1 deck (159 slides) is in `decks/v1 - house style/`.
1. Cover, Agenda, Why We Are Here.
2. **01 Gross**, with Service Is Sales 2.0.
3. **02 Profitability**, with the Absorption and Net to Gross explainers.
4. **03 The CP Engine**: talking points, ROs (store total ranker added), hours sold (total ranker added), hours per RO, Pricing Strategy, ELR, door rate, $ per RO, Goldilocks Margin, GP%, Who's Moving Now after each measure.
5. **04 Labor Mix**, with Express Service and Tracking & Accountability.
6. **05 Opcode Benchmarks**: Road to the Sale, explainer, four benchmarks each with Who's Moving Now, then Missed Opportunity.
7. **06 Warranty & Internal**: Warranty Compliance, warranty ROs (total ranker added), warranty ELR, internal ELR.
8. **07 Productivity**: People, Pay Plans, gross per employee, ROs per advisor/employee/manager, headcount, then Technician Proficiency and Tech Minute.
9. **08 Close**: Better Than Last Year, Who's Moving Now, Two Ways to Read Improvement, What Changed?, Where the Money Is, Before You Leave.
10. **Open Decisions** (slides 156–206; was the appendix; not presented, for review with Tommy): every slide built but not placed in the deck, grouped in deck order: Gross Trend options A / B / B+Q4 (10), Store Records (6), Door Rate options A–E (5), Internal GP % (6), CP vs Internal GP % options A–B (4), Headcount trend (1), Best Techs options A–C (3), Who's Moving Now alternate versions from Most Improved v2 (6), Most Improved v1 (9). Each eyebrow reads `OPEN DECISIONS · <TOPIC> · <OPTION>` (set in `merge_v2.py`). Nothing here moves into the deck by default (Noah, Oct 21); delete the section before the meeting.

Not in the test deck at all: the door rate options deck (`Door Rate - Options (Jul-Sep).pptx`, duplicates of Open Decisions A–E), the v1 Master Review Deck, and the per-deck title slides (replaced by the section transitions).

## Open decisions (Noah / Tommy)
- **What to cut.** Priority is dialogue.
- **Trend slides (Oct 20):** replace the trailing-12 line with Jan–Sep 2026 vs 2025, two lines. All options are saved in `decks/Gross Trend - Review.pptx`; nothing is in the test deck yet. Pick one:
  - **A** (slides 1–6): two-line chart with "What it says" bullets, table slide kept.
  - **B** (slides 7–9): chart merged with the YOY card, table slide dropped (3 fewer slides). Claude's recommendation.
  - **B + Q4** (slide 10): B with the 2025 line through December. In or out?
  - **Headcount** (slide 11): same two-line treatment for Fixed Ops Headcount. Keep or leave the trailing-12 version?
  - Once chosen: swap into the Gross and Productivity sections the way `swap_prod.py` does.
- **Store records (Oct 21):** all options saved in `decks/Store Records - Review.pptx`; nothing in the test deck yet. Pick which to keep: (1) A Record Year, group bars 2021–2026; (2) Record Board, all 25 stores with a Jan–Sep record; (3) Spotlight by %; (4) Spotlight by $; (5) ranker of monthly records set in 2026; (6) the same ranker split into service, parts and combined counts, with the month of each record in the speaker notes. Claude's recommendation: 1 near the opening, 2, and one spotlight in the Close ahead of "What Changed?".
- **Internal GP% (Oct 21):** six slides in `decks/Internal Gross Profit - Review.pptx`, coloured green at or above the group, red below (Noah). Not in the test deck yet; they belong in Warranty & Internal after Internal ELR.
- **CP vs Internal GP% (Oct 21):** Noah: internal GP% should closely mirror customer pay (internal detail and similar processes explain some gaps). `decks/CP vs Internal GP - Review.pptx` shows CP, internal and the gap by store; green = gap at or better than the group gap. Pick option A (gap bars) or B (dumbbell); it sits beside the internal GP% slides as a talking point. **Open (Noah, Oct 21):** colour by gap vs group (current) or by distance from zero (internal far above CP also red). Decide before finalizing.
- **Door rate design:** pick one of the five options in Open Decisions (A–E). The CP Engine section also still carries the door rate slides (52–57); trim those to the pick.
- **Best Techs Prove It:** dumbbell (current), or option A, B or C. Noah leans C.
- **Labor GP%:** how to colour it (the 78–80% band) and whether it belongs on the "most improved" slides. It's left off for now.
- **Smaller rulings:** Acura's segment; whether the GMC body shop counts in absorption; provisional directions in the metric registry.
- **Logistics:** venue and logo; Inter installed on the presenting PC.
- **Possible slides:** MPI, retention, discounting and capacity. Parked dollar slides: ELR to 70%, $1 of ELR, absorption, net to gross, parts margin.
- **Fix the Qlik headcount load** at the source (CSV in `decks/`).

## How Noah works
- **Respect his format.** Lead with the answer. Build review files first (`* - Review.pptx`) and add to the test deck only when he says so.
- **Send the .pptx only** after an update; no PDF unless he asks (Oct 21). Still render internally to check.
- **Verify before sending:** render to images and look. Reconcile totals. Never fabricate a number.
- **When a fact is uncertain** (data quirks, store anomalies), flag it and ask. Example: he confirmed Mazda's August opcode jump is a real change.
