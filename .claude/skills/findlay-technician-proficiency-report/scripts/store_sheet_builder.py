from style_helpers import *


def sort_key_position(pos):
    return CATEGORY_ORDER.index(pos) if pos in CATEGORY_ORDER else len(CATEGORY_ORDER)


def build_store_sheet(wb, sheet_name, store_data, tab_color=NAVY):
    """store_data: {"named": {category: [row dicts]}, "unmapped": [row dicts]}
    row dict keys: tech_no, name, position, sold_hours, ro_count, actual_hours, labor_sale, labor_gross
    Returns metadata: dict with keys:
      total_row, category_ranges (list of (category, start, end, subtotal_row)),
      all_named_rows (list of dict with row_num added, for Tech Ranker),
      store_wide (dict of totals: sold_hours, ro_count, actual_hours, labor_sale, labor_gross, tech_count)
      excluded_range ((first_row, last_row) of the EXCLUDED / POOLED CODES data rows, or None)
    """
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = tab_color

    # Row 1: title
    set_cell(ws, 1, 1, f"{sheet_name} -- Technician Proficiency -- {NEW_PERIOD_LABEL}",
             font=f(size=14, bold=True, color=NAVY), align=center())
    ws.merge_cells("A1:L1")
    # Row 2: back link
    c2 = set_cell(ws, 2, 1, "=HYPERLINK(\"#'Proficiency'!A1\",\"<< Back to Proficiency\")",
                  font=f(color="0563C1", underline="single"))
    ws.merge_cells("A2:L2")
    # Row 3: whatif header
    set_cell(ws, 3, 15, "Technician Proficiency What-If Scenario",
             font=f(size=12, bold=True, color=NAVY), align=center())
    ws.merge_cells("O3:S3")
    # Row 4: column headers
    for i, h in enumerate(DATA_HEADERS, start=1):
        if h is None:
            continue
        set_cell(ws, 4, i, h, font=f(bold=True, color=WHITE), fill_color=NAVY, align=wrap_center())
    ws.row_dimensions[4].height = 40

    r = 5
    category_ranges = []
    all_named_rows = []
    subtotal_rows = []

    present_categories = [c for c in CATEGORY_ORDER if store_data["named"].get(c)]

    for cat in present_categories:
        rows = store_data["named"][cat]
        # sort by sold hours desc within category (matches template convention)
        rows_sorted = sorted(rows, key=sort_key_proficiency)
        cs = r
        for row in rows_sorted:
            set_cell(ws, r, COL["A"], row["tech_no"], font=f())
            set_cell(ws, r, COL["B"], row["name"], font=f())
            set_cell(ws, r, COL["C"], cat, font=f())
            set_cell(ws, r, COL["D"], round(row["sold_hours"], 2), number_format="#,##0.0", font=f())
            set_cell(ws, r, COL["E"], row["ro_count"], font=f())
            set_cell(ws, r, COL["F"], f"=IFERROR(D{r}/E{r},0)", number_format="0.00", font=f())
            set_cell(ws, r, COL["G"], round(row["actual_hours"], 2), font=f())
            set_cell(ws, r, COL["H"], f'=IF(G{r}="N/A","N/A",IFERROR(D{r}/G{r},0))', number_format="0.0%", font=f())
            set_cell(ws, r, COL["I"],
                      f'=IF(H{r}="N/A","No Hours Data",IF(H{r}>=1,"Above Target",IF(H{r}>=0.85,"On Target","Under Target")))',
                      font=f())
            set_cell(ws, r, COL["J"], round(row["labor_sale"], 2), number_format="\\$#,##0", font=f())
            set_cell(ws, r, COL["K"], round(row["labor_gross"], 2), number_format="\\$#,##0", font=f())
            set_cell(ws, r, COL["L"], f"=IFERROR(K{r}/J{r},0)", number_format="0.0%", font=f())
            set_cell(ws, r, COL["M"], f"=IFERROR(J{r}/D{r},0)", number_format="\\$#,##0.00", font=f())
            set_cell(ws, r, COL["O"], f"=H{r}", number_format="0.0%", font=f())
            set_cell(ws, r, COL["P"], None, number_format="0.0%", fill_color=YELLOW_FILL, font=f(bold=True))
            set_cell(ws, r, COL["Q"], f'=IF(P{r}="","",IF(G{r}="N/A","N/A",P{r}-O{r}))', number_format="0.0%", font=f())
            set_cell(ws, r, COL["R"], f'=IF(P{r}="","",IF(G{r}="N/A","N/A",G{r}*P{r}-D{r}))', number_format="#,##0.0",
                      font=f(bold=True))
            set_cell(ws, r, COL["S"], f'=IF(R{r}="","",IF(R{r}="N/A","N/A",IFERROR(R{r}*M{r}*L{r},0)))',
                      number_format="\\$#,##0", font=f())
            all_named_rows.append({"row": r, "sheet": sheet_name, "category": cat,
                                    "sold_hours": row["sold_hours"]})
            r += 1
        ce = r - 1
        add_category_block_cf(ws, cs, ce)
        # GROUP SUBTOTAL row
        sub_r = r
        set_cell(ws, sub_r, COL["A"], "GROUP SUBTOTAL", font=f(bold=True), fill_color=GOLD)
        for col in ["B", "C"]:
            set_cell(ws, sub_r, COL[col], None, fill_color=GOLD)
        set_cell(ws, sub_r, COL["D"], f"=SUM(D{cs}:D{ce})", number_format="#,##0.0", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["E"], f"=SUM(E{cs}:E{ce})", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["F"], f"=IFERROR(D{sub_r}/E{sub_r},0)", number_format="0.00", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["G"], f'=SUMIF(G{cs}:G{ce},"<>N/A")', number_format="#,##0.0", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["H"], f"=IFERROR(D{sub_r}/G{sub_r},0)", number_format="0.0%", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["I"], None, fill_color=GOLD)
        set_cell(ws, sub_r, COL["J"], f"=SUM(J{cs}:J{ce})", number_format="\\$#,##0", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["K"], f"=SUM(K{cs}:K{ce})", number_format="\\$#,##0", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["L"], f"=IFERROR(K{sub_r}/J{sub_r},0)", number_format="0.0%", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["M"], f"=IFERROR(J{sub_r}/D{sub_r},0)", number_format="\\$#,##0.00", font=f(bold=True), fill_color=GOLD)
        set_cell(ws, sub_r, COL["O"], f"=H{sub_r}", number_format="0.0%", font=f(bold=True, color=BLACK), fill_color=GOLD)
        set_cell(ws, sub_r, COL["P"], f"=IFERROR((D{sub_r}+R{sub_r})/G{sub_r},0)", number_format="0.0%", font=f(bold=True, color=BLACK), fill_color=GOLD)
        set_cell(ws, sub_r, COL["Q"], f"=IFERROR(P{sub_r}-O{sub_r},0)", number_format="0.0%", font=f(bold=True, color=BLACK), fill_color=GOLD)
        set_cell(ws, sub_r, COL["R"], f"=SUM(R{cs}:R{ce})", number_format="#,##0.0", font=f(bold=True, color=BLACK), fill_color=GOLD)
        set_cell(ws, sub_r, COL["S"], f"=SUM(S{cs}:S{ce})", number_format="\\$#,##0", font=f(bold=True, color=BLACK), fill_color=GOLD)
        category_ranges.append((cat, cs, ce, sub_r))
        subtotal_rows.append(sub_r)
        r = sub_r + 1
        is_last_category = (cat == present_categories[-1])
        if not is_last_category:
            # solid navy divider bar spanning the full row width (merged, like the template)
            set_cell(ws, r, COL["A"], None, fill_color=NAVY)
            set_cell(ws, r, COL["O"], None, fill_color=NAVY)
            ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["M"])
            ws.merge_cells(start_row=r, start_column=COL["O"], end_row=r, end_column=COL["S"])
            r += 1
        else:
            r += 1  # plain blank separator before TOTAL / AVERAGE

    # TOTAL / AVERAGE row
    tot_r = r
    sum_terms = lambda col: "+".join(f"{col}{sr}" for sr in subtotal_rows)
    set_cell(ws, tot_r, COL["A"], "TOTAL / AVERAGE", font=f(bold=True, color=WHITE), fill_color=NAVY)
    for col in ["B", "C"]:
        set_cell(ws, tot_r, COL[col], None, fill_color=NAVY)
    set_cell(ws, tot_r, COL["D"], f"={sum_terms('D')}", number_format="#,##0.0", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["E"], f"={sum_terms('E')}", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["F"], f"=IFERROR(D{tot_r}/E{tot_r},0)", number_format="0.00", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["G"], f"={sum_terms('G')}", number_format="#,##0.0", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["H"], f"=IFERROR(D{tot_r}/G{tot_r},0)", number_format="0.0%", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["I"], None, fill_color=NAVY)
    set_cell(ws, tot_r, COL["J"], f"={sum_terms('J')}", number_format="\\$#,##0", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["K"], f"={sum_terms('K')}", number_format="\\$#,##0", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["L"], f"=IFERROR(K{tot_r}/J{tot_r},0)", number_format="0.0%", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["M"], f"=IFERROR(J{tot_r}/D{tot_r},0)", number_format="\\$#,##0.00", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["O"], f"=H{tot_r}", number_format="0.0%", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["P"], f"=IFERROR((D{tot_r}+R{tot_r})/G{tot_r},0)", number_format="0.0%", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["Q"], f"=IFERROR(P{tot_r}-O{tot_r},0)", number_format="0.0%", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["R"], f"={sum_terms('R')}", number_format="#,##0.0", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, tot_r, COL["S"], f"={sum_terms('S')}", number_format="\\$#,##0", font=f(bold=True, color=WHITE), fill_color=NAVY)

    r = tot_r + 2  # blank row then whatif block

    # STORE-WIDE WHAT-IF SCENARIO block
    set_cell(ws, r, COL["A"], "STORE-WIDE WHAT-IF SCENARIO", font=f(bold=True))
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["E"])
    r += 1
    set_cell(ws, r, COL["A"], "Metric", font=f(bold=True, color=WHITE), fill_color=NAVY)
    set_cell(ws, r, COL["E"], "Value", font=f(bold=True, color=WHITE), fill_color=NAVY)
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["D"])
    r += 1
    cur_row = r
    set_cell(ws, r, COL["A"], "Current Store Proficiency %", font=f())
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["D"])
    # Matches the Proficiency sheet's row-2 methodology exactly: only Service Technician,
    # Express Technician, Service Team Leader, and Shop Foreman, excluding outlier records
    # (STL/Shop Foreman <=20% proficiency, Service/Express Tech <=2% proficiency).
    qualifying_sold = 0.0
    qualifying_actual = 0.0
    for cat, cs, ce, sub_r in category_ranges:
        for row in store_data["named"].get(cat, []):
            if qualifies_for_prof_subset(row, cat):
                qualifying_sold += row["sold_hours"]
                qualifying_actual += row["actual_hours"]
    cur_value = (qualifying_sold / qualifying_actual) if qualifying_actual else 0.0
    set_cell(ws, r, COL["E"], cur_value, number_format="0.0%", font=f(bold=True))
    r += 1
    tgt_row = r
    set_cell(ws, r, COL["A"], "Target Store Proficiency %", font=f())
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["D"])
    set_cell(ws, r, COL["E"], None, number_format="0.0%", font=f(bold=True), fill_color=YELLOW_FILL)
    r += 1
    inc_row = r
    set_cell(ws, r, COL["A"], "Proficiency Increase", font=f())
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["D"])
    set_cell(ws, r, COL["E"], f'=IF(E{tgt_row}="","",E{tgt_row}-E{cur_row})', number_format="0.0%", font=f(bold=True))
    r += 1
    addhrs_row = r
    set_cell(ws, r, COL["A"], "Additional Sold Hours (Store-Wide)", font=f())
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["D"])
    set_cell(ws, r, COL["E"], f'=IF(E{tgt_row}="","",IFERROR(G{tot_r}*E{tgt_row}-D{tot_r},0))', number_format="#,##0.0", font=f(bold=True))
    r += 1
    addgp_row = r
    set_cell(ws, r, COL["A"], "Additional Labor Gross Profit (Store-Wide)", font=f())
    ws.merge_cells(start_row=r, start_column=COL["A"], end_row=r, end_column=COL["D"])
    set_cell(ws, r, COL["E"], f'=IF(E{addhrs_row}="","",IFERROR(E{addhrs_row}*M{tot_r}*L{tot_r},0))', number_format="\\$#,##0", font=f(bold=True))
    add_status_cf(ws, "E", cur_row, cur_row, kind="pct")
    add_posneg_cf(ws, "E", inc_row, inc_row, black_font=False)
    add_posneg_cf(ws, "E", addhrs_row, addhrs_row, black_font=False)
    add_posneg_cf(ws, "E", addgp_row, addgp_row, black_font=False)

    r += 2

    # EXCLUDED / POOLED CODES block
    unmapped = store_data["unmapped"]
    excluded_range = None  # (first_row, last_row) of the pooled-code data rows, if any
    if unmapped:
        set_cell(ws, r, COL["A"], "EXCLUDED / POOLED CODES (not individually named technicians)", font=f(bold=True))
        r += 1
        for col, label in zip(["A", "B", "C", "D", "E"], ["Tech No", "Sold Hours", "RO Count", "Labor Sale", "Labor Gross"]):
            set_cell(ws, r, COL[col], label, font=f(bold=True, color=WHITE), fill_color=BLACK)
        r += 1
        ex_start = r
        for row in sorted(unmapped, key=lambda x: -x["sold_hours"]):
            set_cell(ws, r, COL["A"], row["tech_no"], font=f())
            set_cell(ws, r, COL["B"], round(row["sold_hours"], 2), number_format="#,##0.0", font=f())
            set_cell(ws, r, COL["C"], row["ro_count"], font=f())
            set_cell(ws, r, COL["D"], round(row["labor_sale"], 2), number_format="\\$#,##0", font=f())
            set_cell(ws, r, COL["E"], round(row["labor_gross"], 2), number_format="\\$#,##0", font=f())
            r += 1
        excluded_range = (ex_start, r - 1)

    # Column widths (fixed, standard across all store sheets)
    widths = {"A": 12, "B": 22, "C": 20, "D": 12, "E": 10, "F": 12, "G": 20, "H": 13, "I": 15,
              "J": 13, "K": 13, "L": 10, "M": 12, "N": 3, "O": 13, "P": 13, "Q": 13, "R": 13, "S": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = "A5"

    # store-wide totals across ALL techs (named + unmapped) for Capacity tab
    all_rows = [row for cat_rows in store_data["named"].values() for row in cat_rows] + unmapped
    store_wide = {
        "tech_count": len(all_rows),
        "sold_hours": sum(x["sold_hours"] for x in all_rows),
        "labor_sale": sum(x["labor_sale"] for x in all_rows),
        "labor_gross": sum(x["labor_gross"] for x in all_rows),
    }

    return {
        "total_row": tot_r,
        "category_ranges": category_ranges,
        "all_named_rows": all_named_rows,
        "store_wide": store_wide,
        "excluded_range": excluded_range,
    }
