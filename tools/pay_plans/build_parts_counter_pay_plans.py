"""Build the Parts Counter Pay Plan Examples workbook from reference/parts_counter_pay_plans/plans.json.

Usage:
  python3 tools/pay_plans/build_parts_counter_pay_plans.py --out <scratchpad>/Parts_Counter_Pay_Plan_Examples.xlsx
  python3 <recalc.py> <out.xlsx> 60        # always recalc as the final step (CLAUDE.md)

Options:
  --benchmark-gross N   monthly parts gross for the "at benchmark store volume" column (default from plans.json)
  --benchmark-label S   label for that column (default from plans.json)

Tabs: Plan Comparison (side by side + estimates), Pay Structures, Plan Documents (each plan's
compensation schedule laid out like the signed originals, one plan per printed page).
Output contains pay terms: write it to the scratchpad, never commit it.
"""
import argparse, json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.normpath(os.path.join(HERE, "..", "..", "reference", "parts_counter_pay_plans", "plans.json"))

NAVY, GOLD = "1F3864", "C9A04B"
hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill("solid", fgColor=NAVY)
body = Font(name="Arial", size=10)
bold = Font(name="Arial", size=10, bold=True)
input_fill = PatternFill("solid", fgColor="FFF2CC")
band = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")


def header_row(ws, row, heads):
    for c, h in enumerate(heads, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill; cell.border = box
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def pay_formula(p, r, g):
    E, F, kind = f"E{r}", f"F{r}", p["estimate_formula"]
    if kind == "std":
        return f"={E}+{F}*{g}"
    if kind == "hourly":
        return f"={p['hourly_rate']}*$L$4"
    if kind == "draw":
        return f"=MAX({E},{F}*{g})"
    if kind == "lr":
        return (f"=MAX({E},{F}*{g}+{g}*IF({g}>=375000,0.004,IF({g}>=350000,0.003,"
                f"IF({g}>=330000,0.002,IF({g}>=300000,0.001,0)))))")
    if kind == "nissan":
        return (f"={E}+{F}*{g}+{g}*IF({g}>=200000,0.0075,IF({g}>=180000,0.005,"
                f"IF({g}>=160000,0.0025,0)))")
    raise ValueError(f"unknown estimate_formula {kind}")


def comparison(wb, data, bench_gross, bench_label):
    ws = wb.active
    ws.title = "Plan Comparison"
    ws.sheet_properties.tabColor = GOLD
    ws["A1"] = "Parts Counter Pay Plans - Examples from Findlay Stores"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color=NAVY)
    ws["A2"] = ("Current signed plans on file for one parts counter employee per store; top parts-gross stores in the "
                "group (Jan-Oct 2026). Employee names removed.")
    ws["A2"].font = Font(name="Arial", italic=True, size=9)
    ws["A4"] = f"{bench_label} avg monthly parts gross - edit to test other volumes:"
    ws["A4"].font = bold
    ws["H4"] = bench_gross
    ws["I4"] = "Monthly hours assumed for hourly plans:"
    ws["I4"].font = bold
    ws["L4"] = data["hours_per_month_for_hourly"]
    for ref, fmt in (("H4", '$#,##0'), ("L4", '0.00')):
        ws[ref].number_format = fmt; ws[ref].fill = input_fill; ws[ref].font = bold; ws[ref].border = box
    heads = ["Store", "Plan Type", "Effective", "Base Type", "Base $ / Month", "Commission % of Total Parts Gross",
             "Commission Detail", "Gross Bonus", "Other Pay", "Store Avg Monthly Parts Gross (Jan-Oct 2026)",
             "Est. Monthly Pay at Own Store", f"Est. Monthly Pay at {bench_label} Volume",
             "What the Estimate Leaves Out / Notes"]
    HR = 6
    header_row(ws, HR, heads)
    ws.row_dimensions[HR].height = 45
    for i, p in enumerate(data["plans"]):
        r = HR + 1 + i
        vals = [p["comparison_label"], p["plan_type"], p["effective"], p["base_type"], None,
                p["commission_pct_total_parts_gross"], p["commission_detail"], p["gross_bonus"], p["other_pay"],
                p["store_avg_monthly_parts_gross_jan_oct_2026"], None, None, p["notes"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = body; cell.border = box; cell.alignment = wrap
            if i % 2:
                cell.fill = band
        ws[f"E{r}"] = f"={p['hourly_rate']}*$L$4" if p["estimate_formula"] == "hourly" else p["base_monthly"]
        ws[f"K{r}"] = pay_formula(p, r, f"J{r}")
        ws[f"L{r}"] = pay_formula(p, r, "$H$4")
        for col in "EJKL":
            c = ws[f"{col}{r}"]
            c.number_format = '$#,##0'; c.font = body; c.border = box; c.alignment = wrap
            if i % 2:
                c.fill = band
        ws[f"F{r}"].number_format = '0.00%'
        ws[f"A{r}"].font = bold
        ws[f"L{r}"].font = bold
    fr = HR + len(data["plans"]) + 2
    ws[f"A{fr}"] = "How to read the estimates"
    ws[f"A{fr}"].font = bold
    notes = [
        "Estimates apply each plan's base and total-department-gross commission to an average month. They are illustrations, not payroll amounts; ADP shows what was actually paid.",
        "Plans pay on a specific OEM financial statement line; the store gross used here is Qlik Productivity 'Total Gross Profit - Parts', which may not tie exactly to that line.",
        "Pay tied to individual production, CSI, shop hours, year-over-year growth, car allowances and spiffs is NOT included - see the Notes column for each plan.",
        "Recoverable draws: estimate shows the greater of the draw or earned commission for an average month.",
        "Hourly plans: hourly rate x monthly hours (cell L4), no overtime.",
        f"{bench_label} column: the same plan formulas applied to that store's volume (cell H4). Bonus tiers built for smaller or larger stores may not translate directly.",
    ]
    for k, t in enumerate(notes):
        ws.cell(row=fr + 1 + k, column=1, value=f"- {t}").font = Font(name="Arial", size=9)
    for c, w in enumerate([28, 30, 16, 22, 13, 14, 45, 40, 38, 17, 15, 16, 48], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B7"
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def structures(wb, data):
    s2 = wb.create_sheet("Pay Structures")
    s2.sheet_properties.tabColor = NAVY
    s2["A1"] = "Ways Parts Counter Staff Are Paid in the Group"
    s2["A1"].font = Font(name="Arial", bold=True, size=14, color=NAVY)
    header_row(s2, 3, ["Structure", "Stores Using It", "Range in the Group", "How It Works", "Strengths", "Watch-Outs"])
    for i, row in enumerate(data["pay_structures"]):
        vals = [row["structure"], row["stores"], row["range"], row["how_it_works"], row["strengths"], row["watch_outs"]]
        for c, v in enumerate(vals, 1):
            cell = s2.cell(row=4 + i, column=c, value=v)
            cell.font = bold if c == 1 else body; cell.border = box; cell.alignment = wrap
            if i % 2:
                cell.fill = band
    for c, w in enumerate([30, 42, 36, 46, 38, 42], 1):
        s2.column_dimensions[get_column_letter(c)].width = w
    s2.page_setup.orientation = "landscape"; s2.page_setup.fitToWidth = 1; s2.page_setup.fitToHeight = 0
    s2.sheet_properties.pageSetUpPr.fitToPage = True


def documents(wb, data):
    d = wb.create_sheet("Plan Documents")
    d.sheet_properties.tabColor = NAVY
    d.sheet_view.showGridLines = False
    for col, w in (("A", 3), ("B", 5), ("C", 5), ("D", 88)):
        d.column_dimensions[col].width = w

    def put(r, col, text, font=body, indent=0, h=None, align="left"):
        c = d.cell(row=r, column=col, value=text)
        c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=align, indent=indent)
        if col < 4:
            d.merge_cells(start_row=r, start_column=col, end_row=r, end_column=4)
        width = {2: 118, 3: 112, 4: 104}[col]
        lines = max(1, -(-len(text) // width)) if text else 1
        d.row_dimensions[r].height = h or (13.5 * lines + 2)
        return r + 1

    title_font = Font(name="Arial", bold=True, size=14, color=NAVY)
    sec_font = Font(name="Arial", bold=True, size=10, underline="single")
    note_font = Font(name="Arial", italic=True, size=9, color="9C0006")
    line_bottom = Border(bottom=Side(style="medium", color=NAVY))
    r = 1
    plans = data["plans"]
    for n, p in enumerate(plans, 1):
        doc = p["document"]
        put(r, 2, doc["title"], title_font, h=22, align="center"); r += 1
        r = put(r, 2, f"Employee Name: [Removed]        Position: {doc['position']}        Effective Date: {doc['effective']}",
                bold, h=18, align="center")
        for col in range(2, 5):
            d.cell(row=r - 1, column=col).border = line_bottom
        r += 1
        if doc["reviewer_note"]:
            r = put(r, 2, "Reviewer note: " + doc["reviewer_note"], note_font)
        r += 1
        r = put(r, 2, "A.   COMPENSATION SCHEDULE:", sec_font)
        r = put(r, 3, "Monthly compensation will be calculated as follows:")
        for i, item in enumerate(doc["compensation_schedule"], 1):
            d.cell(row=r, column=3, value=f"{i}.").font = body
            d.cell(row=r, column=3).alignment = Alignment(vertical="top", horizontal="right")
            r = put(r, 4, item["text"])
            for t in item["tiers"]:
                r = put(r, 4, t, body, indent=12, h=14)
            if item["footnote"]:
                r = put(r, 4, item["footnote"], bold, indent=12 if item["tiers"] else 0)
        if doc.get("draw_note"):
            r += 1
            r = put(r, 3, doc["draw_note"], bold)
        r += 1
        if n < len(plans):
            d.row_breaks.append(Break(id=r - 1))
    d.page_setup.orientation = "portrait"
    d.page_setup.fitToWidth = 1; d.page_setup.fitToHeight = 0
    d.sheet_properties.pageSetUpPr.fitToPage = True
    d.print_options.horizontalCentered = True
    d.page_margins.left = d.page_margins.right = 0.5
    d.oddFooter.center.text = "Page &P of &N"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--data", default=DATA)
    ap.add_argument("--benchmark-gross", type=float)
    ap.add_argument("--benchmark-label")
    a = ap.parse_args()
    data = json.load(open(a.data))
    bench = data["benchmark"]
    wb = Workbook()
    comparison(wb, data, a.benchmark_gross or bench["avg_monthly_parts_gross_jan_oct_2026"],
               a.benchmark_label or bench["label"])
    structures(wb, data)
    documents(wb, data)
    wb.active = 0
    wb.save(a.out)
    print(a.out)


if __name__ == "__main__":
    main()
