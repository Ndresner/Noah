#!/usr/bin/env python3
"""Fixed Ops Gross Profit Records report: workbook + 3-page PDF.

Takes last month's finished workbook (its Data tab is the full monthly history)
plus the new month's store-level Service / Parts gross from Qlik, appends the new
month to Data, applies any post-close restatements, and writes:
  <Month>_<Year>_Fixed_Ops_Gross_Records.xlsx   (4 tabs, all live formulas)
  <Month>_<Year>_Fixed_Ops_Gross_Records.pdf    (3 pages, reportlab)

Usage:
  build_records.py --prior August_2026_Fixed_Ops_Gross_Records.xlsx \
      --new sep_2026.csv --year 2026 --month 9 --pull-date 2026-10-08 \
      [--restate restate.csv] [--notes notes.json] --outdir OUT

--new      CSV: Store,Service,Parts   (Store = Data-tab store name, e.g. ACURA FA-5)
--restate  CSV: Store,Year,MonthNum,Service,Parts  (prior months re-pulled from Qlik
           that differ from the prior workbook; only changed rows are needed)
--notes    JSON overriding generated text: exec_summary, near_records_p1,
           near_records_p2, restatement_note, store_openings, cdk_note
Gross values are positive (Qlik MTD_Amt * -1), matching the Data tab.
"""
import argparse
import csv
import datetime as dt
import json
import os
import subprocess
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_FULL = ["January", "February", "March", "April", "May", "June", "July",
              "August", "September", "October", "November", "December"]
CATS = [("Service", 4), ("Parts", 5), ("Combined", 6)]  # tuple index in a Data row

DEFAULT_OPENINGS = ("Honda Spokane (Aug 2024), Nissan Henderson (Nov 2023), LR Henderson (Jul 2021), "
                    "Hyundai Prescott (Mar 2022), Ineos (Nov 2025)")
DEFAULT_CDK = "June 2024 is depressed across all stores (CDK outage)."
ACRONYMS = {"LR", "LV", "VW", "GMC", "CDJR"}
RECALC = os.environ.get("RECALC_PY", "/mnt/skills/public/xlsx/scripts/recalc.py")


# ----------------------------------------------------------------------------- data
def load_history(prior_path):
    ws = openpyxl.load_workbook(prior_path)["Data"]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        rows.append([r[0], int(r[1]), int(r[2]), r[3], float(r[4] or 0), float(r[5] or 0)])
    return rows


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def num(v):
    v = str(v).replace("$", "").replace(",", "").strip()
    if v.startswith("(") and v.endswith(")"):
        v = "-" + v[1:-1]
    return round(float(v or 0), 2)


def build_rows(history, new_csv, year, month, restate_csv=None):
    stores = sorted({r[0] for r in history})
    if any(r[1] == year and r[2] == month for r in history):
        sys.exit(f"Data already holds {MON[month-1]} {year}; pass the prior month's workbook.")
    changes = []
    if restate_csv:
        idx = {(r[0], r[1], r[2]): r for r in history}
        for x in read_csv(restate_csv):
            key = (x["Store"].strip(), int(x["Year"]), int(x["MonthNum"]))
            if key not in idx:
                sys.exit(f"Restatement row not in history: {key}")
            r = idx[key]
            new_s, new_p = num(x["Service"]), num(x["Parts"])
            if round(r[4], 2) != new_s or round(r[5], 2) != new_p:
                changes.append((key, r[4], r[5], new_s, new_p))
                r[4], r[5] = new_s, new_p
    new = {x["Store"].strip(): (num(x["Service"]), num(x["Parts"])) for x in read_csv(new_csv)}
    missing = [s for s in stores if s not in new]
    extra = [s for s in new if s not in stores]
    if missing:
        sys.exit(f"New-month file is missing stores: {missing}")
    if extra:
        sys.exit(f"New-month file has stores not in history (add backfill first): {extra}")
    rows = history + [[s, year, month, MON[month-1], *new[s]] for s in stores]
    for r in rows:
        r.append(round(r[4] + r[5], 2))  # index 6 = Combined (mirrors =ROUND(E+F,2))
    return rows, stores, changes


def analyze(rows, stores, year, month):
    cur = (year, month)
    by_store = {s: [r for r in rows if r[0] == s] for s in stores}
    res = {}
    for s in stores:
        hist = by_store[s]
        now = next(r for r in hist if (r[1], r[2]) == cur)
        d = {"row": now}
        for cat, i in CATS:
            val = now[i]
            rank = sum(1 for r in hist if r[i] > val) + 1
            prior = [r for r in hist if (r[1], r[2]) != cur]
            best = max(prior, key=lambda r: r[i]) if prior else None
            rec = max(r[i] for r in hist)
            d[cat] = {"val": val, "rank": rank, "record": rec,
                      "prior_val": best[i] if best else 0,
                      "prior_mon": f"{best[3]} {best[1]}" if best else ""}
        res[s] = d
    # group totals by month
    grp = {}
    for r in rows:
        g = grp.setdefault((r[1], r[2]), [0.0, 0.0, 0.0])
        g[0] += r[4]; g[1] += r[5]; g[2] += r[6]
    group = {}
    for cat, gi in (("Service", 0), ("Parts", 1), ("Combined", 2)):
        val = grp[cur][gi]
        ranked = sorted(grp.items(), key=lambda kv: -kv[1][gi])
        group[cat] = {"val": val, "rank": sum(1 for v in grp.values() if v[gi] > val) + 1,
                      "ranked": [(k, v[gi]) for k, v in ranked]}
    order = sorted(stores, key=lambda s: -res[s]["Combined"]["val"])
    return res, group, order


# ----------------------------------------------------------------------------- text
def short(store):
    """'LR RENO LRR-31' -> 'LR Reno' (drops the company-code token)."""
    words = store.split()[:-1] or store.split()
    return " ".join(w if w in ACRONYMS else w.title() for w in words)


def money_m(v):
    return f"${v/1e6:.2f}M"


def ordinal(n):
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def word_ord(n):
    return {1: "best", 2: "second-best", 3: "third-best", 4: "fourth-best", 5: "fifth-best"}.get(
        n, f"{ordinal(n)}-best")


def join(xs):
    xs = list(xs)
    if len(xs) <= 1:
        return "".join(xs)
    if len(xs) == 2:
        return f"{xs[0]} and {xs[1]}"
    return ", ".join(xs[:-1]) + f", and {xs[-1]}"


def num_word(n):
    return ["No", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"][n] if n <= 10 else str(n)


def records_by_cat(res, order):
    return {cat: [s for s in order if res[s][cat]["rank"] == 1] for cat, _ in CATS}


def exec_summary(res, group, order, year, month, restatement_note):
    mlabel = f"{MONTH_FULL[month-1]} {year}"
    c = group["Combined"]
    higher = [(k, v) for k, v in c["ranked"] if v > c["val"]]
    lab = lambda k: f"{MONTH_FULL[k[1]-1]} {k[0]}"
    if c["rank"] == 1:
        nxt = c["ranked"][1]
        s1 = (f"{mlabel} combined Fixed Ops gross of {money_m(c['val'])} set a new all-time company record, "
              f"topping {lab(nxt[0])} ({money_m(nxt[1])}).")
    else:
        parts = [f"{lab(k)}" + (" all-time record" if i == 0 else "") + f" ({money_m(v)})"
                 for i, (k, v) in enumerate(higher[:3])]
        s1 = (f"{mlabel} combined Fixed Ops gross of {money_m(c['val'])} was the {word_ord(c['rank'])} month "
              f"in company history, trailing {join(parts)}" + ("" if len(higher) <= 3 else
              f" among {len(higher)} higher months") + ".")
    s2 = (f"Group Service gross of {money_m(group['Service']['val'])} ranks #{group['Service']['rank']} "
          f"all-time and Group Parts gross of {money_m(group['Parts']['val'])} ranks "
          f"#{group['Parts']['rank']} all-time.")
    rec = records_by_cat(res, order)
    names = {k: [short(s) for s in v] for k, v in rec.items()}
    out = [s1, s2]
    n = len(rec["Combined"])
    if n:
        out.append(f"{num_word(n)} store{'s' if n != 1 else ''} set new all-time combined-month "
                   f"record{'s' if n != 1 else ''} — {join(names['Combined'])}.")
    else:
        out.append("No store set a new all-time combined-month record.")
    if names["Service"]:
        out.append(f"New Service records: {join(names['Service'])}.")
    if names["Parts"]:
        out.append(f"New Parts records: {join(names['Parts'])}.")
    if restatement_note:
        out.append(restatement_note)
    return " ".join(out)


def near_records(res, order, split):
    """Rank-2/3 store-months that are not records. split=True -> (page1, page2)."""
    p1, p2 = [], []
    for s in order:
        r2 = [cat if cat != "Combined" else "combined" for cat, _ in CATS if res[s][cat]["rank"] == 2]
        r3c = res[s]["Combined"]["rank"] == 3
        r3sp = [cat for cat in ("Service", "Parts") if res[s][cat]["rank"] == 3]
        bits = [f"#2 {c}" for c in r2] + (["#3 combined"] if r3c else [])
        if bits:
            p1.append(f"{short(s)} ({', '.join(bits)})")
        if r3sp:
            p2.append(f"{short(s)} (#3 {' and '.join(r3sp)})")
    t1 = ("Near-records: " + "; ".join(p1) + ".") if p1 else ""
    t2 = ("Near-records not shown above: " + "; ".join(p2) + ".") if p2 else ""
    return t1, t2


# ----------------------------------------------------------------------------- xlsx
ARIAL = "Arial"
NAVY = "FF1B2A41"
MONEY = '\\$#,##0;"($"#,##0\\);\\-'
THIN = Side(style="thin", color="FFBFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FONT = Font(name=ARIAL, size=10, bold=True, color="FFFFFFFF")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
REC_FILL = PatternFill("solid", fgColor="FFC6EFCE")
REC_FONT = Font(name=ARIAL, size=10, bold=True, color="FF006100")
BAND = PatternFill("solid", fgColor="FFF2F2F2")
CENTER = Alignment(horizontal="center")


def f(size=10, bold=False, italic=False, color=None):
    return Font(name=ARIAL, size=size, bold=bold, italic=italic, color=color)


def write_xlsx(path, rows, res, order, year, month, pull, first_year, restatement_note,
               openings, cdk_note):
    mon, MONU = MON[month-1], MON[month-1].upper()
    n = len(rows) + 1
    rng = lambda col: f"Data!${col}$2:${col}${n}"
    pull_s = pull.strftime("%b %-d, %Y")
    next_m = MON[month % 12] + f" {year + (1 if month == 12 else 0)}"
    wb = openpyxl.Workbook()

    # --- Sheet 1: <Month> <Year> vs Records
    ws = wb.active
    ws.title = f"{MONTH_FULL[month-1]} {year} vs Records"
    ws["A1"] = (f"FINDLAY AUTOMOTIVE GROUP — {MONTH_FULL[month-1].upper()} {year} "
                f"FIXED OPS GROSS vs ALL-TIME RECORDS")
    ws["A1"].font = f(12, bold=True)
    ws["A2"] = (f"Source: Qlik Fixed Ops - Accting TB (Family=FIXED GROSS), pulled {pull_s}. "
                f"History: Jan {first_year} – {mon} {year}. Rank = all-time rank of the "
                f"{MONTH_FULL[month-1]} {year} month for that store. Sorted by {MONTH_FULL[month-1]} "
                f"{year} combined gross.")
    ws["A2"].font = f(9, italic=True, color="FF595959")
    for c0, lab in ((2, "SERVICE"), (5, "PARTS"), (8, "SERVICE + PARTS")):
        ws.cell(3, c0, lab)
        ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c0 + 2)
        for c in range(c0, c0 + 3):
            ws.cell(3, c).font, ws.cell(3, c).fill, ws.cell(3, c).alignment = HDR_FONT, HDR_FILL, CENTER
    heads = ["Store"] + [f"{mon} {year} ($)", "Rank", "Record ($)"] * 3
    for c, h in enumerate(heads, 1):
        x = ws.cell(4, c, h)
        x.font, x.fill, x.alignment = HDR_FONT, HDR_FILL, CENTER
    for i, s in enumerate(order):
        r = 5 + i
        ws.cell(r, 1, s)
        for k, (cat, col) in enumerate((("Service", "E"), ("Parts", "F"), ("Combined", "G"))):
            vc = 2 + 3 * k
            ws.cell(r, vc, f"=SUMIFS({rng(col)},{rng('A')},$A{r},{rng('B')},{year},{rng('C')},{month})")
            L = openpyxl.utils.get_column_letter(vc)
            ws.cell(r, vc + 1, f'=COUNTIFS({rng("A")},$A{r},{rng(col)},">"&{L}{r})+1')
            ws.cell(r, vc + 2, f"=SUMPRODUCT(MAX(({rng('A')}=$A{r})*{rng(col)}))")
            for c in (vc, vc + 1, vc + 2):
                ws.cell(r, c).number_format = MONEY if c != vc + 1 else "General"
                if c == vc + 1:
                    ws.cell(r, c).alignment = CENTER
            if res[s][cat]["rank"] == 1:  # static highlight, snapshot as of pull
                for c in (vc, vc + 1):
                    ws.cell(r, c).fill, ws.cell(r, c).font = REC_FILL, REC_FONT
        for c in range(1, 11):
            x = ws.cell(r, c)
            x.border = BOX
            if not x.fill.fill_type:
                x.font = f()
    tr = 5 + len(order)
    ws.cell(tr, 1, "GROUP TOTAL")
    for c in (2, 5, 8):
        L = openpyxl.utils.get_column_letter(c)
        ws.cell(tr, c, f"=SUM({L}5:{L}{tr-1})").number_format = MONEY
    for c in range(1, 11):
        ws.cell(tr, c).border = BOX
        if c in (1, 2, 5, 8):
            ws.cell(tr, c).font = f(bold=True)
    for col, w in zip("ABCDEFGHIJ", (26, 13, 7, 13, 13, 7, 13, 13, 7, 13)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B5"

    # --- Data
    d = wb.create_sheet("Data")
    for c, h in enumerate(["Store", "Year", "MonthNum", "Month", "Service", "Parts", "Total"], 1):
        x = d.cell(1, c, h)
        x.font, x.fill = HDR_FONT, HDR_FILL
    blue = f(color="FF0000FF")
    for i, r in enumerate(rows, 2):
        for c, v in enumerate(r[:6], 1):
            x = d.cell(i, c, v)
            x.font = blue if c in (5, 6) else f()
            if c in (5, 6):
                x.number_format = MONEY
        x = d.cell(i, 7, f"=ROUND(E{i}+F{i},2)")
        x.font, x.number_format = f(), MONEY
    for col, w in zip("ABCDEFG", (26, 7, 10, 7, 12, 12, 12)):
        d.column_dimensions[col].width = w

    # --- All-Time Records
    a = wb.create_sheet("All-Time Records")
    a["A1"] = (f"FINDLAY AUTOMOTIVE GROUP — FIXED OPS GROSS PROFIT RECORDS "
               f"(ALL-TIME MONTHLY, {first_year}–{MONU} {year})")
    a["A1"].font = f(12, bold=True)
    for c, h in enumerate(["Store", "Service Record ($)", "Svc Month", "Parts Record ($)", "Pts Month",
                           "Svc+Pts Record ($)", "Total Month"], 1):
        x = a.cell(3, c, h)
        x.font, x.fill, x.alignment = HDR_FONT, HDR_FILL, CENTER
    months = ",".join(f'"{m}"' for m in MON)
    for i, s in enumerate(order):
        r = 4 + i
        a.cell(r, 1, s)
        for k, col in enumerate("EFG"):
            vc = 2 + 2 * k
            V = openpyxl.utils.get_column_letter(vc)
            a.cell(r, vc, f"=SUMPRODUCT(MAX(({rng('A')}=$A{r})*{rng(col)}))").number_format = MONEY
            m = f"({rng('A')}=$A{r})*({rng(col)}={V}{r})"
            a.cell(r, vc + 1, f"=CHOOSE(SUMPRODUCT({m}*{rng('C')}),{months})&\" \"&SUMPRODUCT({m}*{rng('B')})"
                   ).alignment = CENTER
        for c in range(1, 8):
            x = a.cell(r, c)
            x.font, x.border = f(), BOX
            if i % 2 == 0:
                x.fill = BAND
    for col, w in zip("ABCDEFG", (26, 16, 11, 16, 11, 17, 11)):
        a.column_dimensions[col].width = w
    a.freeze_panes = "A4"

    # --- Methodology
    m = wb.create_sheet("Methodology")
    lines = [
        "FIXED OPS GROSS PROFIT RECORDS — METHODOLOGY",
        None,
        f'Source: Qlik Cloud, app "Fixed Ops - Accting TB" (Fixed space, e077f027-fc9f-4f03-adf3-14ce059d8b3c), '
        f"pulled {pull.isoformat()}.",
        "Measure convention (matches app master measures):",
        "  Service gross = Sum({<Family={'FIXED GROSS'}, Class={'SERVICE'}>} [Mon]_MTD_Amt * -1)",
        "  Parts gross   = Sum({<Family={'FIXED GROSS'}, Class={'PARTS'}>}   [Mon]_MTD_Amt * -1)",
        "  Service + Parts = sum of the two for the same store-month. Body Shop is excluded.",
        None,
        f"History window: Jan {first_year} through {mon} {year} (all complete months in the app).",
        f"{next_m} excluded as partial (in progress at pull time).",
        restatement_note,
        "Rank = 1 + count of store-months with a strictly greater value for that store (rank 1 = all-time record).",
        "All values on the report tabs are live formulas (SUMIFS / COUNTIFS / SUMPRODUCT) against the Data tab.",
        f"Green highlight = new all-time record set in {MONTH_FULL[month-1]} {year} "
        f"(static formatting, snapshot as of pull date).",
        None,
        f"Data note: {cdk_note} Does not affect record maxima.",
        f"New stores enter history at opening: {openings}.",
    ]
    for i, t in enumerate(lines, 1):
        if t is not None:
            m.cell(i, 1, t).font = f(bold=(i == 1))
    m.column_dimensions["A"].width = 120
    wb.active = 0
    wb.save(path)


# ----------------------------------------------------------------------------- pdf
def write_pdf(path, res, group, order, year, month, pull, first_year, notes, xlsx_name):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    NAVY = colors.HexColor("#1A2744")
    GREEN = colors.HexColor("#0E7C4A")
    RECBG = colors.HexColor("#C6EFCE")
    RECFG = colors.HexColor("#006100")
    BANDC = colors.HexColor("#F5F6F8")
    CARDB = colors.HexColor("#D9DCE1")
    GREY = colors.HexColor("#595959")
    W, H = letter
    LM = 42
    mlabel = f"{MONTH_FULL[month-1]} {year}"
    mshort = f"{MON[month-1]} {year}"
    footer = f"Findlay Automotive Group — Fixed Ops Gross Profit Records — {mlabel}"

    def deco(c, doc, first):
        c.saveState()
        if first:
            c.setFillColor(NAVY); c.rect(0, H - 116, W, 116, stroke=0, fill=1)
            c.setFillColor(GREEN); c.rect(0, H - 119, W, 3, stroke=0, fill=1)
            c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 9)
            c.drawString(LM, H - 30, "FINDLAY AUTOMOTIVE GROUP  ·  FIXED OPERATIONS")
            c.setFont("Helvetica-Bold", 21)
            c.drawString(LM, H - 61, f"Gross Profit Records Review — {mlabel}")
            c.setFont("Helvetica", 9)
            c.drawString(LM, H - 82, f"Service · Parts · Combined  |  All-time monthly records, "
                                     f"Jan {first_year} – {mshort}")
            c.setFillColor(colors.HexColor("#B8BFCC")); c.setFont("Helvetica", 7.5)
            c.drawRightString(W - LM, H - 27, f"Prepared {pull.strftime('%B %-d, %Y')}")
            c.drawRightString(W - LM, H - 38, "Source: Qlik — Fixed Ops Accting TB")
        c.setFillColor(colors.HexColor("#666666")); c.setFont("Helvetica", 7.5)
        c.drawString(LM, 30, footer)
        c.drawRightString(W - LM, 30, f"Page {doc.page}")
        c.restoreState()

    doc = SimpleDocTemplate(path, pagesize=letter, leftMargin=LM, rightMargin=LM, topMargin=24,
                            bottomMargin=50, title=f"Fixed Ops Gross Profit Records — {mlabel}",
                            author="Findlay Automotive Group")
    h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, textColor=NAVY, leading=16,
                        spaceAfter=2)
    body = ParagraphStyle("b", fontName="Helvetica", fontSize=8.5, leading=12, alignment=TA_LEFT)
    small = ParagraphStyle("s", fontName="Helvetica", fontSize=7.5, leading=11, textColor=GREY)
    bullet = ParagraphStyle("bl", parent=body, spaceAfter=8)
    cell = lambda t, **kw: Paragraph(t, ParagraphStyle("c", fontName="Helvetica", fontSize=7.5,
                                                        leading=9, **kw))
    usd = lambda v: f"${v:,.0f}"
    story = [Spacer(1, 132)]

    # KPI cards
    rec = records_by_cat(res, order)
    lab = ParagraphStyle("kl", fontName="Helvetica-Bold", fontSize=6.5, textColor=GREY, leading=8)
    big = ParagraphStyle("kb", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY, leading=19)
    sub = ParagraphStyle("ks", fontName="Helvetica-Bold", fontSize=6.5, textColor=GREEN, leading=8)
    nrec = len(rec["Combined"])
    cards = [("GROUP SERVICE GROSS", usd(group["Service"]["val"]), f"#{group['Service']['rank']} all-time month"),
             ("GROUP PARTS GROSS", usd(group["Parts"]["val"]), f"#{group['Parts']['rank']} all-time month"),
             ("COMBINED SVC + PARTS", usd(group["Combined"]["val"]), f"#{group['Combined']['rank']} all-time month"),
             ("NEW COMBINED RECORDS", f"{nrec} store{'s' if nrec != 1 else ''}",
              f"{len(rec['Service'])} Service · {len(rec['Parts'])} Parts")]
    kt = Table([[[Paragraph(a, lab), Spacer(1, 7), Paragraph(b, big), Paragraph(c, sub)]
                 for a, b, c in cards]], colWidths=[126] * 4, rowHeights=[70])
    kt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), BANDC), ("BOX", (0, 0), (-1, -1), 0.75, CARDB),
                            ("INNERGRID", (0, 0), (-1, -1), 0.75, CARDB), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 10)]))
    story += [kt, Spacer(1, 18), Paragraph("Executive Summary", h2),
              Paragraph(notes["exec_summary"], body), Spacer(1, 12),
              Paragraph(f"Records Set in {mlabel}", h2)]

    # Records table
    hdr = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=6.5, textColor=colors.white, leading=8)
    hdrR = ParagraphStyle("hr", parent=hdr, alignment=2)
    data = [[Paragraph("CATEGORY", hdr), Paragraph("STORE", hdr), Paragraph(mlabel.upper(), hdrR),
             Paragraph("PRIOR RECORD", hdrR), Paragraph("PRIOR RECORD MONTH", hdrR)]]
    for cat in ("Combined", "Service", "Parts"):
        for s in rec[cat]:
            x = res[s][cat]
            data.append([cat, s.title(), usd(x["val"]), usd(x["prior_val"]) if x["prior_mon"] else "—",
                         x["prior_mon"] or "First month"])
    if len(data) == 1:
        data.append(["—", f"No new all-time records in {mlabel}", "", "", ""])
    rt = Table(data, colWidths=[68, 170, 92, 92, 106], repeatRows=1)
    ts = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("FONT", (0, 1), (-1, -1), "Helvetica", 7.5), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
          ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
          ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 9),
          ("LINEBELOW", (0, -1), (-1, -1), 0.75, NAVY)]
    for i in range(1, len(data)):
        if i % 2 == 0:
            ts.append(("BACKGROUND", (0, i), (-1, i), BANDC))
        if len(rec["Combined"] + rec["Service"] + rec["Parts"]):
            ts += [("BACKGROUND", (2, i), (2, i), RECBG), ("TEXTCOLOR", (2, i), (2, i), RECFG),
                   ("FONT", (2, i), (2, i), "Helvetica-Bold", 7.5)]
    rt.setStyle(TableStyle(ts))
    story += [rt, Spacer(1, 8)]
    if notes["near_records_p1"]:
        story.append(Paragraph(notes["near_records_p1"], small))
    story.append(PageBreak())

    # Page 2: all stores
    story += [Paragraph(f"All Stores — {mlabel} vs. All-Time Records", h2),
              Paragraph(f"Rank = all-time rank of the {mlabel} month within each store's history "
                        f"(Jan {first_year} – {mshort}). Green = new all-time record. Sorted by combined "
                        f"{mlabel} gross.", small), Spacer(1, 6)]
    cols = ["STORE", "SERVICE", "RK", "SVC REC", "PARTS", "RK", "PTS REC", "SVC+PTS", "RK", "TOT REC"]
    data = [cols]
    for s in order:
        row = [s.title()]
        for cat, _ in CATS:
            x = res[s][cat]
            row += [usd(x["val"]), "REC" if x["rank"] == 1 else f"#{x['rank']}", usd(x["record"])]
        data.append(row)
    data.append(["GROUP TOTAL", usd(group["Service"]["val"]), f"#{group['Service']['rank']}", "",
                 usd(group["Parts"]["val"]), f"#{group['Parts']['rank']}", "",
                 usd(group["Combined"]["val"]), f"#{group['Combined']['rank']}", ""])
    at = Table(data, colWidths=[110, 58, 32, 60, 52, 32, 60, 58, 32, 64], repeatRows=1)
    ts = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
          ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 6.5), ("FONT", (0, 1), (-1, -1), "Helvetica", 7),
          ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 4.3), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.3),
          ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
          ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 7), ("LINEABOVE", (0, -1), (-1, -1), 1, NAVY),
          ("TOPPADDING", (0, -1), (-1, -1), 8)]
    for i, s in enumerate(order, 1):
        if i % 2 == 0:
            ts.append(("BACKGROUND", (0, i), (-1, i), BANDC))
        for k, (cat, _) in enumerate(CATS):
            if res[s][cat]["rank"] == 1:
                c0 = 1 + 3 * k
                ts += [("BACKGROUND", (c0, i), (c0 + 1, i), RECBG), ("TEXTCOLOR", (c0, i), (c0 + 1, i), RECFG),
                       ("FONT", (c0, i), (c0 + 1, i), "Helvetica-Bold", 7)]
    at.setStyle(TableStyle(ts))
    story += [at, Spacer(1, 10)]
    if notes["near_records_p2"]:
        story.append(Paragraph(notes["near_records_p2"], small))
    story.append(PageBreak())

    # Page 3: methodology
    nxt = f"{MONTH_FULL[month % 12]} {year + (1 if month == 12 else 0)}"
    bullets = [
        f"Source: Qlik Cloud app “Fixed Ops - Accting TB” (Fixed space, e077f027-fc9f-4f03-adf3-14ce059d8b3c), "
        f"pulled {pull.strftime('%B %-d, %Y')}. Measures follow the app's master-measure convention: "
        f"Sum({{&lt;Family={{'FIXED GROSS'}}, Class={{'SERVICE'|'PARTS'}}&gt;}} [Month]_MTD_Amt x -1). "
        f"Combined = Service + Parts for the same store-month; Body Shop excluded.",
        f"History window is January {first_year} through {mlabel} — all complete months available in the app. "
        f"{nxt} excluded as partial (in progress at pull time). {notes['restatement_note']}",
        f"Store-months of $0 prior to a store's opening are excluded from ranking ({notes['store_openings_pdf']}).",
        f"{notes['cdk_note'].rstrip('.').replace(' (CDK outage)', ' due to the CDK outage')}; this does not affect record maxima but should be considered in any "
        f"average-based comparison.",
        "Rank = 1 + count of store-months with a strictly greater value for that store (rank 1 = all-time "
        "record, shown as REC).",
        f"Companion workbook: {xlsx_name} — all figures are live formulas against a full monthly history tab "
        f"(SUMIFS / COUNTIFS / SUMPRODUCT), recalculated and verified with zero formula errors.",
    ]
    story.append(Paragraph("Methodology &amp; Notes", h2))
    story += [Paragraph("• " + b, bullet) for b in bullets]
    doc.build(story, onFirstPage=lambda c, d: deco(c, d, True), onLaterPages=lambda c, d: deco(c, d, False))


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--prior", required=True)
    ap.add_argument("--new", required=True)
    ap.add_argument("--restate")
    ap.add_argument("--notes")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--pull-date", required=True, help="YYYY-MM-DD")
    ap.add_argument("--outdir", default=".")
    ap.add_argument("--no-recalc", action="store_true")
    a = ap.parse_args()
    pull = dt.date.fromisoformat(a.pull_date)

    history = load_history(a.prior)
    rows, stores, changes = build_rows(history, a.new, a.year, a.month, a.restate)
    first_year = min(r[1] for r in rows)
    res, group, order = analyze(rows, stores, a.year, a.month)

    notes = json.load(open(a.notes)) if a.notes else {}
    if "restatement_note" not in notes:
        if changes:
            ms = sorted({f"{MON[k[2]-1]} {k[1]}" for k, *_ in changes})
            notes["restatement_note"] = (f"Post-close adjustments applied from this pull to: {join(ms)} "
                                         f"({len(changes)} store-months).")
        else:
            notes["restatement_note"] = ("Prior months reflect the prior report's pull; no post-close "
                                         "adjustments were identified in this pull.")
    openings = notes.get("store_openings", DEFAULT_OPENINGS)
    notes["store_openings_pdf"] = openings.replace(" (", " opened ", 1).replace("), ", ", ").replace(" (", " ").rstrip(")")
    notes.setdefault("cdk_note", DEFAULT_CDK)
    notes.setdefault("exec_summary", exec_summary(res, group, order, a.year, a.month, notes["restatement_note"]))
    p1, p2 = near_records(res, order, True)
    notes.setdefault("near_records_p1", p1)
    notes.setdefault("near_records_p2", p2)

    stem = f"{MONTH_FULL[a.month-1]}_{a.year}_Fixed_Ops_Gross_Records"
    os.makedirs(a.outdir, exist_ok=True)
    xlsx = os.path.join(a.outdir, stem + ".xlsx")
    pdf = os.path.join(a.outdir, stem + ".pdf")
    write_xlsx(xlsx, rows, res, order, a.year, a.month, pull, first_year, notes["restatement_note"],
               openings, notes["cdk_note"])
    write_pdf(pdf, res, group, order, a.year, a.month, pull, first_year, notes, os.path.basename(xlsx))
    if not a.no_recalc:  # final step on every .xlsx build (CLAUDE.md); never re-save after this
        out = subprocess.run([sys.executable, RECALC, xlsx, "180"], capture_output=True, text=True)
        print(out.stdout.strip() or out.stderr.strip())
    json.dump({"group": {k: {"val": v["val"], "rank": v["rank"]} for k, v in group.items()},
               "records": records_by_cat(res, order), "restatements": [list(map(str, c)) for c in changes],
               "notes": notes}, open(os.path.join(a.outdir, stem + "_summary.json"), "w"), indent=1)
    print(xlsx, pdf, sep="\n")


if __name__ == "__main__":
    main()
