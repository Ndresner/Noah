import re
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from style_helpers import *
from store_sheet_builder import build_store_sheet, CATEGORY_ORDER
from build_data import build, DIVISION_TO_SHEET
from config import load_config

_cfg = load_config()
TEMPLATE_PATH = _cfg["template_path"]
OUT_PATH = _cfg["output_path"]
# PROF_CATEGORIES and qualifies_for_prof_subset now come from style_helpers (imported above via *)
PERIOD_START = datetime.strptime(_cfg["qlik_date_start"], "%m/%d/%Y").date()
PERIOD_END = datetime.strptime(_cfg["qlik_date_end"], "%m/%d/%Y").date()


def period_day_counts(start, end):
    """(stall days, tech days) in the pay period: Mon-Sat open days for the stall calc,
    Mon-Fri worked days for the tech calc. Must match the NETWORKDAYS.INTL formulas in
    Capacity columns F (weekend code 11) and M (weekend code 1)."""
    days = [start + timedelta(i) for i in range((end - start).days + 1)]
    return sum(d.weekday() < 6 for d in days), sum(d.weekday() < 5 for d in days)


def period_labor_gross_formula(sheet_name, meta):
    """Store-wide labor gross for the period: named techs (TOTAL row, col K) plus the
    EXCLUDED / POOLED CODES block (col E) -- same population as ELR and GP Retention %."""
    q = f"'{sheet_name}'"
    formula = f"={q}!K{meta['total_row']}"
    if meta["excluded_range"]:
        a, b = meta["excluded_range"]
        formula += f"+SUM({q}!E{a}:E{b})"
    return formula


def main():
    template = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    stores_data = build()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for name in ["Proficiency", "Proficiency Ranker", "Tech Ranker", "Capacity", "Capacity Ranker"]:
        wb.create_sheet(name)

    store_meta = {}
    prof_subset = {}
    tech_ranker_rows = []

    sheet_order = sorted(stores_data.keys())

    for sheet_name in sheet_order:
        data = stores_data[sheet_name]
        meta = build_store_sheet(wb, sheet_name, data, tab_color=NAVY)
        store_meta[sheet_name] = meta

        techs = 0
        sold_sum = 0.0
        actual_sum = 0.0
        sale_sum = 0.0
        gross_sum = 0.0
        for cat, cs, ce, sub_r in meta["category_ranges"]:
            if cat not in PROF_CATEGORIES:
                continue
            for row in data["named"][cat]:
                if qualifies_for_prof_subset(row, cat):
                    techs += 1
                    sold_sum += row["sold_hours"]
                    actual_sum += row["actual_hours"]
                    sale_sum += row["labor_sale"]
                    gross_sum += row["labor_gross"]
        prof_subset[sheet_name] = {
            "technicians": techs, "sold_hours": sold_sum, "actual_hours": actual_sum,
            "avg_prof": (sold_sum / actual_sum) if actual_sum else 0.0,
            "labor_sale": sale_sum, "labor_gross": gross_sum,
            "avg_gp": (gross_sum / sale_sum) if sale_sum else 0.0,
        }

        for info in meta["all_named_rows"]:
            if info["category"] in PROF_CATEGORIES:
                tech_ranker_rows.append(info)

    # ---------------- Proficiency sheet ----------------
    ws = wb["Proficiency"]
    ws.sheet_properties.tabColor = GOLD
    set_cell(ws, 1, 1, f"Findlay Automotive Group -- Technician Proficiency -- {NEW_PERIOD_LABEL}",
             font=f(size=14, bold=True, color=NAVY))
    set_cell(ws, 2, 1,
             "Proficiency uses actual hours worked (Regular + Overtime) from the ADP employee hours export, not an "
             "assumed schedule. Store totals include only (SERVICE TECHNICIAN), (EXPRESS TECHNICIAN), and (SERVICE "
             "TEAM LEADER; SHOP FOREMAN) -- excludes (BODY SHOP HELPER; BODY SHOP PAINTER; BODY SHOP TECH), "
             "(DETAILER; DETAIL MANAGER), (TINTER), and unmapped/pooled tech codes. Also excludes outlier records: "
             "Service Team Leaders/Shop Foremen at or below 20% proficiency, and Service/Express Technicians at or "
             "below 2% proficiency. Click a store name to jump to its sheet. Sorted alphabetically -- see "
             "Proficiency Ranker for proficiency ranking (note: Proficiency Ranker still reflects all categories, "
             "unfiltered).", font=f(size=10, italic=True, color="555555"))
    headers = ["Store", "Technicians", "Total Sold Hours", "Total Actual Hours", "Avg Proficiency %", "Status",
               "Total Labor Sale", "Total Labor Gross", "Avg Labor GP%"]
    for i, h in enumerate(headers, start=1):
        set_cell(ws, 4, i, h, font=f(bold=True, color=WHITE), fill_color=NAVY, align=wrap_center())
    r = 5
    prof_start = r
    for sheet_name in sorted(prof_subset.keys()):
        d = prof_subset[sheet_name]
        set_cell(ws, r, 1, f'=HYPERLINK("#\'{sheet_name}\'!A1","{sheet_name}")', font=f(color="0563C1", underline="single"))
        set_cell(ws, r, 2, d["technicians"], font=f())
        set_cell(ws, r, 3, round(d["sold_hours"], 2), number_format="#,##0.0", font=f())
        set_cell(ws, r, 4, round(d["actual_hours"], 2), number_format="#,##0.0", font=f())
        set_cell(ws, r, 5, d["avg_prof"], number_format="0.0%", font=f())
        set_cell(ws, r, 6, f'=IF(B{r}=0,"No Data",IF(E{r}>=1,"Above Target",IF(E{r}>=0.85,"On Target","Under Target")))', font=f())
        set_cell(ws, r, 7, round(d["labor_sale"], 2), number_format="\\$#,##0", font=f())
        set_cell(ws, r, 8, round(d["labor_gross"], 2), number_format="\\$#,##0", font=f())
        set_cell(ws, r, 9, d["avg_gp"], number_format="0.0%", font=f())
        r += 1
    prof_end = r - 1
    add_status_cf(ws, "F", prof_start, prof_end, kind="status_text")
    add_status_cf(ws, "I", prof_start, prof_end, kind="gp")
    widths = [24, 13, 16, 16, 15, 14, 15, 15, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # ---------------- Proficiency Ranker sheet ----------------
    ws = wb["Proficiency Ranker"]
    set_cell(ws, 1, 1, f"Findlay Automotive Group -- Proficiency Ranker -- {NEW_PERIOD_LABEL}",
             font=f(size=14, bold=True, color=NAVY))
    set_cell(ws, 2, 1,
             "Same data as Proficiency, ranked by Avg Proficiency %, highest to lowest -- includes only (SERVICE "
             "TECHNICIAN), (EXPRESS TECHNICIAN), and (SERVICE TEAM LEADER; SHOP FOREMAN), excludes (BODY SHOP "
             "HELPER; BODY SHOP PAINTER; BODY SHOP TECH), (DETAILER; DETAIL MANAGER), (TINTER), and unmapped/pooled "
             "tech codes, and also excludes outlier records: Service Team Leaders/Shop Foremen at or below 20% "
             "proficiency, and Service/Express Technicians at or below 2% proficiency. Click a store name to jump "
             "to its sheet.", font=f(size=10, italic=True, color="555555"))
    for i, h in enumerate(headers, start=1):
        set_cell(ws, 4, i, h, font=f(bold=True, color=WHITE), fill_color=NAVY, align=wrap_center())
    r = 5
    rank_start = r
    ranked = sorted(prof_subset.items(), key=lambda kv: -kv[1]["avg_prof"])
    for sheet_name, d in ranked:
        set_cell(ws, r, 1, f'=HYPERLINK("#\'{sheet_name}\'!A1","{sheet_name}")', font=f(color="0563C1", underline="single"))
        set_cell(ws, r, 2, d["technicians"], font=f())
        set_cell(ws, r, 3, round(d["sold_hours"], 2), number_format="#,##0.0", font=f())
        set_cell(ws, r, 4, round(d["actual_hours"], 2), number_format="#,##0.0", font=f())
        set_cell(ws, r, 5, d["avg_prof"], number_format="0.0%", font=f())
        set_cell(ws, r, 6, f'=IF(B{r}=0,"No Data",IF(E{r}>=1,"Above Target",IF(E{r}>=0.85,"On Target","Under Target")))', font=f())
        set_cell(ws, r, 7, round(d["labor_sale"], 2), number_format="\\$#,##0", font=f())
        set_cell(ws, r, 8, round(d["labor_gross"], 2), number_format="\\$#,##0", font=f())
        set_cell(ws, r, 9, d["avg_gp"], number_format="0.0%", font=f())
        r += 1
    rank_end = r - 1
    add_status_cf(ws, "F", rank_start, rank_end, kind="status_text")
    add_status_cf(ws, "I", rank_start, rank_end, kind="gp")
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # ---------------- Tech Ranker sheet ----------------
    ws = wb["Tech Ranker"]
    set_cell(ws, 1, 1, f"Findlay Automotive Group -- Tech Ranker -- {NEW_PERIOD_LABEL}",
             font=f(size=14, bold=True, color=NAVY))
    set_cell(ws, 2, 1,
             "Ranks (SERVICE TECHNICIAN), (EXPRESS TECHNICIAN), and (SERVICE TEAM LEADER; SHOP FOREMAN) technicians "
             "across 34 stores by Proficiency %, highest to lowest, linking directly to each technician's row "
             "within that category on their store's sheet. Excludes (BODY SHOP HELPER; BODY SHOP PAINTER; BODY "
             "SHOP TECH), (DETAILER; DETAIL MANAGER), (TINTER), and unmapped/pooled tech codes -- see the "
             "Proficiency sheet for store totals that include every category. Click a store name to jump to its "
             "sheet.", font=f(size=10, italic=True, color="555555"))
    tr_headers = ["Rank", "Name", "Tech No", "Store", "Position", "Sold Hours", "RO Count",
                  "Actual Hours (Reg+OT)", "Proficiency %", "Status", "Labor Sale", "Labor Gross", "GP%", "ELR"]
    for i, h in enumerate(tr_headers, start=1):
        set_cell(ws, 4, i, h, font=f(bold=True, color=WHITE), fill_color=NAVY, align=wrap_center())

    ranker_entries = []
    for sheet_name in sheet_order:
        data = stores_data[sheet_name]
        meta = store_meta[sheet_name]
        for cat, cs, ce, sub_r in meta["category_ranges"]:
            if cat not in PROF_CATEGORIES:
                continue
            rows_sorted = sorted(data["named"][cat], key=sort_key_proficiency)
            for idx, row in enumerate(rows_sorted):
                store_row = cs + idx
                actual = row["actual_hours"]
                prof = (row["sold_hours"] / actual) if actual else 0.0
                ranker_entries.append({"sheet": sheet_name, "row": store_row, "prof": prof})

    ranker_entries.sort(key=lambda x: -x["prof"])
    r = 5
    for rank, entry in enumerate(ranker_entries, start=1):
        sn, sr = entry["sheet"], entry["row"]
        quoted = f"'{sn}'" if " " in sn else sn
        set_cell(ws, r, 1, rank, font=f())
        set_cell(ws, r, 2, f"={quoted}!B{sr}", font=f())
        set_cell(ws, r, 3, f"={quoted}!A{sr}", font=f())
        set_cell(ws, r, 4, sn, font=f())
        set_cell(ws, r, 5, f"={quoted}!C{sr}", font=f())
        set_cell(ws, r, 6, f"={quoted}!D{sr}", number_format="#,##0.0", font=f())
        set_cell(ws, r, 7, f"={quoted}!E{sr}", font=f())
        set_cell(ws, r, 8, f"={quoted}!G{sr}", number_format="#,##0.0", font=f())
        set_cell(ws, r, 9, f"={quoted}!H{sr}", number_format="0.0%", font=f())
        set_cell(ws, r, 10, f"={quoted}!I{sr}", font=f())
        set_cell(ws, r, 11, f"={quoted}!J{sr}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 12, f"={quoted}!K{sr}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 13, f"={quoted}!L{sr}", number_format="0.0%", font=f())
        set_cell(ws, r, 14, f"={quoted}!M{sr}", number_format="\\$#,##0.00", font=f())
        r += 1
    tr_end = r - 1
    add_status_cf(ws, "J", 5, tr_end, kind="status_text")
    add_status_cf(ws, "M", 5, tr_end, kind="gp")
    tr_widths = {"A": 10.71, "B": 22, "C": 14.14, "D": 20, "E": 22, "F": 16.86, "G": 15.57, "H": 16, "I": 13,
                 "J": 13, "K": 16.43, "L": 18.29, "M": 10, "N": 12}
    for col, w in tr_widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ---------------- Capacity + Capacity Ranker: pay-period capacity, carry forward Stalls ----------------
    tmpl_cap = template["Capacity"]
    prior_by_store = {}
    for row in tmpl_cap.iter_rows(min_row=6, max_row=39, values_only=False):
        name_cell = row[0]
        if not name_cell.value:
            continue
        val = name_cell.value
        m = re.search(r',"([^"]+)"\)', val) if isinstance(val, str) else None
        store = m.group(1) if m else None
        if store:
            prior_by_store[store] = {
                "stalls": row[1].value, "hours_open": row[2].value,
                "desired_flag_hrs": row[10].value,
            }
    tech_days = period_day_counts(PERIOD_START, PERIOD_END)[1]  # Python-side ranking only

    ws = wb["Capacity"]
    ws.sheet_properties.tabColor = GOLD
    set_cell(ws, 1, 1, f"Findlay Automotive Group -- Capacity -- {NEW_PERIOD_LABEL}", font=f(size=14, bold=True, color=NAVY))
    ws.merge_cells("A1:T1")
    set_cell(ws, 2, 1,
             "Rebuilt using the Findlay Capacity Calculator methodology (STALL Capacity Calculator + Forecasting "
             "Tool - TECH Count + Dashboard). No Main/Express split applied per user direction -- stalls and "
             "technicians are each treated as one combined pool per store. Tech Count (column J) combines every "
             "technician category on the store sheet -- not just the categories ranked in Tech Ranker. Assumptions "
             "used for all stores: 12 hrs/day open and a combined 10 flag hrs/day/tech goal. Capacity covers the pay "
             "period only (B3:C3): Days Open = Mon-Sat days in the period (stall calc), Days Worked = Mon-Fri days "
             "in the period (tech calc). ELR and GP Retention % are each store's own blended actuals for the "
             "period, recomputed from this period's closed-RO data. Actual Period Labor Gross is the store-wide "
             "labor gross for the same dates and population (named techs + pooled codes), linked to each store "
             "sheet. Total Stalls is carried forward unchanged from the prior period; refresh it from its source "
             "when available. Goals: Stall Capacity Utilization = 75%+, Tech Capacity Utilization = 85-95%.",
             font=f(size=9, italic=True, color="555555"))
    ws.merge_cells("A2:T2")
    set_cell(ws, 3, 1, "Pay Period Start / End", font=f(bold=True))
    set_cell(ws, 3, 2, PERIOD_START, number_format="m/d/yyyy", font=f(bold=True))
    set_cell(ws, 3, 3, PERIOD_END, number_format="m/d/yyyy", font=f(bold=True))
    set_cell(ws, 4, 1, "Store", font=f(bold=True, color=WHITE), fill_color=NAVY, align=center())
    ws.merge_cells(start_row=4, start_column=1, end_row=5, end_column=1)
    cap_headers1 = [(2, 9, "Stall-Based Capacity Calculation"), (10, 15, "Tech-Based Capacity Calculation"),
                    (16, 20, "Dashboard Comparison")]
    for start_col, end_col, label in cap_headers1:
        set_cell(ws, 4, start_col, label, font=f(bold=True, color=WHITE), fill_color=NAVY, align=center())
        ws.merge_cells(start_row=4, start_column=start_col, end_row=4, end_column=end_col)
    cap_headers2 = ["Store", "Total Stalls", "Hours Open/Day", "ELR (Blended)", "Stall Labor $/Day",
                    "Days Open/Period", "Stall Labor Sales/Period", "GP Retention %", "Stall-Based Potential Gross/Period",
                    "Tech Count", "Desired Flag Hrs/Day", "Tech Labor $/Day", "Days Worked/ Period",
                    "Tech Labor Sales/Period", "Tech-Based Potential Gross/Period", "Actual Period Labor Gross",
                    "Stall Capacity Utilization %", "Stall Status", "Tech Capacity Utilization %", "Tech Status"]
    for i, h in enumerate(cap_headers2, start=1):
        if i == 1:
            continue  # "Store" already placed at A4, merged into A5
        set_cell(ws, 5, i, h, font=f(bold=True, color=WHITE), fill_color=NAVY, align=wrap_center())

    r = 6
    cap_start = r
    for sheet_name in sheet_order:
        sw = store_meta[sheet_name]["store_wide"]
        prior = prior_by_store.get(sheet_name, {})
        stalls = prior.get("stalls", 1)
        hours_open = prior.get("hours_open", 12)
        desired_flag = prior.get("desired_flag_hrs", 10)
        elr = (sw["labor_sale"] / sw["sold_hours"]) if sw["sold_hours"] else 0.0
        gp_pct = (sw["labor_gross"] / sw["labor_sale"]) if sw["labor_sale"] else 0.0
        set_cell(ws, r, 1, f'=HYPERLINK("#\'{sheet_name}\'!A1","{sheet_name}")', font=f(color="0563C1", underline="single"))
        set_cell(ws, r, 2, stalls, font=f())
        set_cell(ws, r, 3, hours_open, font=f())
        set_cell(ws, r, 4, elr, number_format="\\$#,##0.00", font=f())
        set_cell(ws, r, 5, f"=B{r}*C{r}*D{r}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 6, "=NETWORKDAYS.INTL($B$3,$C$3,11)", font=f())
        set_cell(ws, r, 7, f"=E{r}*F{r}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 8, gp_pct, number_format="0.0%", font=f())
        set_cell(ws, r, 9, f"=G{r}*H{r}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 10, sw["tech_count"], font=f())
        set_cell(ws, r, 11, desired_flag, font=f())
        set_cell(ws, r, 12, f"=J{r}*K{r}*D{r}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 13, "=NETWORKDAYS.INTL($B$3,$C$3,1)", font=f())
        set_cell(ws, r, 14, f"=L{r}*M{r}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 15, f"=N{r}*H{r}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 16, period_labor_gross_formula(sheet_name, store_meta[sheet_name]), number_format="\\$#,##0", font=f())
        set_cell(ws, r, 17, f"=IFERROR(P{r}/I{r},0)", number_format="0.0%", font=f())
        set_cell(ws, r, 18, f'=IF(Q{r}>=0.75,"At/Above Target","Below Target")', font=f())
        set_cell(ws, r, 19, f"=IFERROR(P{r}/O{r},0)", number_format="0.0%", font=f())
        set_cell(ws, r, 20, f'=IF(S{r}<0.85,"Below Target",IF(S{r}<=0.95,"In Target","Above Target"))', font=f())
        r += 1
    cap_end = r - 1

    thin = thin_border()
    for rr in range(4, cap_end + 1):
        for cc in range(1, 21):
            ws.cell(row=rr, column=cc).border = thin

    rng = f"Q{cap_start}:Q{cap_end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"Q{cap_start}<0.75"], font=Font(color=RED_FONT), fill=fill(RED_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"Q{cap_start}>=0.75"], font=Font(color=GREEN_FONT), fill=fill(GREEN_FILL)))
    rng = f"R{cap_start}:R{cap_end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'R{cap_start}="Below Target"'], font=Font(color=RED_FONT), fill=fill(RED_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'R{cap_start}="At/Above Target"'], font=Font(color=GREEN_FONT), fill=fill(GREEN_FILL)))
    add_status_cf(ws, "S", cap_start, cap_end, kind="pct")
    rng = f"T{cap_start}:T{cap_end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'T{cap_start}="Below Target"'], font=Font(color=RED_FONT), fill=fill(RED_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'T{cap_start}="In Target"'], font=Font(color=YELLOW_FONT), fill=fill(YELLOW_FILL_CF)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'T{cap_start}="Above Target"'], font=Font(color=GREEN_FONT), fill=fill(GREEN_FILL)))

    cap_widths = {"A": 21, "B": 12, "C": 13, "D": 14, "E": 14, "F": 13, "G": 16, "H": 13, "I": 18,
                  "J": 10, "K": 14, "L": 13, "M": 13, "N": 15, "O": 18, "P": 15, "Q": 13, "R": 15, "S": 13, "T": 15}
    for col, w in cap_widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"

    # ---------------- Capacity Ranker ----------------
    ws = wb["Capacity Ranker"]
    set_cell(ws, 1, 1, f"Findlay Automotive Group -- Capacity Ranker -- {NEW_PERIOD_LABEL}", font=f(size=14, bold=True, color=NAVY))
    set_cell(ws, 2, 1, "All 34 stores ranked by Tech Capacity Utilization %, highest to lowest. Same data as "
                       "Capacity, combined-pool methodology. Click a store name to jump to its sheet.",
             font=f(size=10, italic=True, color="555555"))
    cr_headers = ["Rank", "Store", "Tech Count", "Total Stalls", "Tech Capacity Utilization %", "Tech Status",
                  "Stall Capacity Utilization %", "Stall Status", "Actual Period Labor Gross",
                  "Tech-Based Potential Gross/Period", "Stall-Based Potential Gross/Period"]
    for i, h in enumerate(cr_headers, start=1):
        set_cell(ws, 4, i, h, font=f(bold=True, color=WHITE), fill_color=NAVY, align=wrap_center())

    cap_rows_by_store = {}
    rr = cap_start
    for sheet_name in sheet_order:
        cap_rows_by_store[sheet_name] = rr
        rr += 1

    util_calc = {}
    for sheet_name in sheet_order:
        sw = store_meta[sheet_name]["store_wide"]
        prior = prior_by_store.get(sheet_name, {})
        elr = (sw["labor_sale"] / sw["sold_hours"]) if sw["sold_hours"] else 0.0
        desired_flag = prior.get("desired_flag_hrs", 10)
        tech_labor_dollar_day = sw["tech_count"] * desired_flag * elr
        tech_labor_sales_period = tech_labor_dollar_day * tech_days
        gp_pct = (sw["labor_gross"] / sw["labor_sale"]) if sw["labor_sale"] else 0.0
        tech_potential_gross_period = tech_labor_sales_period * gp_pct
        util = (sw["labor_gross"] / tech_potential_gross_period) if tech_potential_gross_period else 0.0
        util_calc[sheet_name] = util
    ranked_cap = sorted(sheet_order, key=lambda s: -util_calc[s])
    r = 5
    for rank, sheet_name in enumerate(ranked_cap, start=1):
        cap_row = cap_rows_by_store[sheet_name]
        set_cell(ws, r, 1, rank, font=f())
        set_cell(ws, r, 2, f'=HYPERLINK("#\'{sheet_name}\'!A1","{sheet_name}")', font=f(color="0563C1", underline="single"))
        set_cell(ws, r, 3, f"=Capacity!J{cap_row}", font=f())
        set_cell(ws, r, 4, f"=Capacity!B{cap_row}", font=f())
        set_cell(ws, r, 5, f"=Capacity!S{cap_row}", number_format="0.0%", font=f())
        set_cell(ws, r, 6, f"=Capacity!T{cap_row}", font=f())
        set_cell(ws, r, 7, f"=Capacity!Q{cap_row}", number_format="0.0%", font=f())
        set_cell(ws, r, 8, f"=Capacity!R{cap_row}", font=f())
        set_cell(ws, r, 9, f"=Capacity!P{cap_row}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 10, f"=Capacity!O{cap_row}", number_format="\\$#,##0", font=f())
        set_cell(ws, r, 11, f"=Capacity!I{cap_row}", number_format="\\$#,##0", font=f())
        r += 1
    crk_end = r - 1
    add_status_cf(ws, "E", 5, crk_end, kind="pct")
    rng = f"F5:F{crk_end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['F5="Below Target"'], font=Font(color=RED_FONT), fill=fill(RED_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['F5="In Target"'], font=Font(color=YELLOW_FONT), fill=fill(YELLOW_FILL_CF)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['F5="Above Target"'], font=Font(color=GREEN_FONT), fill=fill(GREEN_FILL)))
    rng = f"G5:G{crk_end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=["G5<0.75"], font=Font(color=RED_FONT), fill=fill(RED_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=["G5>=0.75"], font=Font(color=GREEN_FONT), fill=fill(GREEN_FILL)))
    rng = f"H5:H{crk_end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['H5="Below Target"'], font=Font(color=RED_FONT), fill=fill(RED_FILL)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['H5="At/Above Target"'], font=Font(color=GREEN_FONT), fill=fill(GREEN_FILL)))
    cr_widths = {"A": 8, "B": 21, "C": 12, "D": 12, "E": 15, "F": 15, "G": 15.5, "H": 15.5, "I": 15.5, "J": 18, "K": 18}
    for col, w in cr_widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    wb.active = wb.sheetnames.index("Proficiency")
    wb.save(OUT_PATH)
    print("saved", OUT_PATH)


if __name__ == "__main__":
    main()
