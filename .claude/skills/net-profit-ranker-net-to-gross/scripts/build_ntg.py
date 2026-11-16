#!/usr/bin/env python3
"""Add a "Net to Gross" ranker tab to the monthly Net Profit Ranker workbooks.

Reads the three finished Net Profit Ranker files (Fixed = Service + Parts combined,
Service, Parts) and a JSON of Service and Parts gross/net pulled from the Qlik P&L app.
Writes a copy of each workbook with a second tab, "Net to Gross", styled from the first
tab (fonts, widths, frozen panes, conditional formatting, negative-red font). The
Total tab is not touched.

Net to gross % = department net / department total departmental gross profit.
  Months     : monthly net (from the workbook) / monthly gross (from Qlik)
  YYYY YTD   : YTD net / YTD gross (not an average of the monthly %)
  YYYY FINAL : prior-year net / prior-year gross
  YTD +/- PY : YTD % minus prior-year %, in percentage points
  Total row  : group net / group gross
Stores are ranked by YTD %, highest first.

Before writing, it checks the workbook against Qlik and stops on any miss over $1:
  - monthly gross sums to YTD gross (catches transcription errors)
  - workbook YTD net and prior-year net == Qlik net, per store
  - store list in the workbook == store list in the JSON

Usage:
  python3 build_ntg.py --data qlik_ntg.json --outdir <scratchpad> \
      --fixed "Net Profit Ranker - Fixed.xlsx" \
      --service "Net Profit Ranker - Service.xlsx" \
      --parts "Net Profit Ranker - Parts.xlsx"
Any of --fixed/--service/--parts may be omitted. Output file names match the inputs.
Run recalc.py on every output afterwards; never re-save with openpyxl after that.
"""
import argparse, copy, json, os, sys
import openpyxl

SHEET = 'Net to Gross'
TITLE = 'Net to Gross Ranker'
DELTA_HDR = 'YTD +/- PY'
PCT = '0.0%;-0.0%'
DELTA = '+0.0%;-0.0%;0.0%'
RED, BLACK = 'FFD90613', 'FF000000'
TOL = 1.0  # dollars


def dept_series(stores, dept):
    """Return {store: dict(gm=[...], gy, gp, ny, np)} for 'service', 'parts' or 'fixed'."""
    out = {}
    for s, d in stores.items():
        if dept == 'fixed':
            a, b = d['service'], d['parts']
            out[s] = dict(gm=[x + y for x, y in zip(a['gross_months'], b['gross_months'])],
                          gy=a['gross_ytd'] + b['gross_ytd'], gp=a['gross_py'] + b['gross_py'],
                          ny=a['net_ytd'] + b['net_ytd'], np=a['net_py'] + b['net_py'])
        else:
            a = d[dept]
            out[s] = dict(gm=a['gross_months'], gy=a['gross_ytd'], gp=a['gross_py'],
                          ny=a['net_ytd'], np=a['net_py'])
    return out


def ratio(n, g):
    return n / g if g else None


def layout_of(ws):
    """Find header columns on the source tab (row 5)."""
    hdr = {c: ws.cell(5, c).value for c in range(2, ws.max_column + 1)}
    months, ytd, final, last = [], None, None, None
    for c, v in hdr.items():
        if v is None:
            continue
        v = str(v)
        last = c
        if v.endswith('YTD'):
            ytd = c
        elif v.endswith('FINAL'):
            final = c
        elif ytd is None and c > 2:
            months.append(c)
    if not (months and ytd and final):
        raise SystemExit(f'Unexpected header row on {ws.title}: {hdr}')
    total_row = next(r for r in range(6, ws.max_row + 1) if ws.cell(r, 2).value == 'Total')
    return months, ytd, final, last, total_row


def style(dst, src, fmt=None):
    dst.font = copy.copy(src.font); dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border); dst.alignment = copy.copy(src.alignment)
    dst.protection = copy.copy(src.protection)
    dst.number_format = fmt or src.number_format


def build(path, dept, stores, outdir):
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    months, c_ytd, c_final, c_last, r_total = layout_of(ws)
    M = len(months)
    data = dept_series(stores, dept)

    # ---- reconcile ----
    issues = []
    book = {}
    for r in range(6, r_total):
        s = ws.cell(r, 2).value
        if s is None:
            continue
        book[s] = dict(nm=[ws.cell(r, c).value or 0 for c in months],
                       ny=ws.cell(r, c_ytd).value or 0, np=ws.cell(r, c_final).value or 0)
    if set(book) != set(data):
        issues.append(('store list', sorted(set(book) ^ set(data))))
    for s in book.keys() & data.keys():
        d, b = data[s], book[s]
        if len(d['gm']) != M:
            issues.append((s, f'{len(d["gm"])} gross months in JSON, {M} in workbook'))
            continue
        if abs(sum(d['gm']) - d['gy']) > TOL:
            issues.append((s, 'gross months != gross YTD', round(sum(d['gm']) - d['gy'], 2)))
        if abs(b['ny'] - d['ny']) > TOL:
            issues.append((s, 'YTD net: workbook vs Qlik', b['ny'], d['ny']))
        if abs(b['np'] - d['np']) > TOL:
            issues.append((s, 'prior-year net: workbook vs Qlik', b['np'], d['np']))
    if issues:
        print(f'[{dept}] RECONCILE FAILED for {os.path.basename(path)}:')
        for i in issues:
            print('  ', i)
        raise SystemExit(1)

    # ---- compute ----
    rows = []
    for s, b in book.items():
        d = data[s]
        m = [ratio(b['nm'][i], d['gm'][i]) for i in range(M)]
        y, p = ratio(d['ny'], d['gy']), ratio(d['np'], d['gp'])
        rows.append((s, m, y, p, (y - p) if (y is not None and p is not None) else None))
    rows.sort(key=lambda x: (x[2] is None, -(x[2] or 0)))
    tn = [sum(book[s]['nm'][i] for s in book) for i in range(M)]
    tg = [sum(data[s]['gm'][i] for s in book) for i in range(M)]
    ty = ratio(sum(data[s]['ny'] for s in book), sum(data[s]['gy'] for s in book))
    tp = ratio(sum(data[s]['np'] for s in book), sum(data[s]['gp'] for s in book))
    total = ('Total', [ratio(tn[i], tg[i]) for i in range(M)], ty, tp, ty - tp)

    # ---- new sheet ----
    ns = wb.create_sheet(SHEET)
    ns.sheet_format = copy.copy(ws.sheet_format)
    for k, d in ws.column_dimensions.items():
        nd = ns.column_dimensions[k]
        nd.width, nd.hidden, nd.min, nd.max = d.width, d.hidden, d.min, d.max
    for r, d in ws.row_dimensions.items():
        ns.row_dimensions[r].height = d.height
    ns.sheet_view.showGridLines = ws.sheet_view.showGridLines
    ns.freeze_panes = ws.freeze_panes
    ns.page_margins = copy.copy(ws.page_margins)
    ns.sheet_properties.tabColor = ws.sheet_properties.tabColor
    for c in ('B1', 'B2', 'B3'):
        style(ns[c], ws[c]); ns[c].value = ws[c].value
    ns['B1'].value = TITLE

    # (new col, source col for styling, header)
    lay = [(2, 2, ws.cell(5, 2).value)] + [(3 + i, c, ws.cell(5, c).value) for i, c in enumerate(months)]
    n = 3 + M
    lay += [(n, c_ytd, ws.cell(5, c_ytd).value), (n + 1, c_final, ws.cell(5, c_final).value),
            (n + 2, c_last, DELTA_HDR)]
    for nc, sc, h in lay:
        style(ns.cell(5, nc), ws.cell(5, sc)); ns.cell(5, nc).value = h

    def write(r, srow, rec):
        s, m, y, p, dlt = rec
        for (nc, sc, _), v in zip(lay, [s] + m + [y, p, dlt]):
            cell = ns.cell(r, nc)
            style(cell, ws.cell(srow, sc), None if nc == 2 else (DELTA if nc == n + 2 else PCT))
            cell.value = v
            if nc != 2:
                f = copy.copy(cell.font)
                f.color = RED if (v is not None and v < 0) else BLACK
                cell.font = f

    for i, rec in enumerate(rows):
        write(6 + i, 6 + i, rec)
    write(r_total, r_total, total)

    # Same CF rules; the right-edge border rule keys on the new last header.
    last_hdr = str(ws.cell(5, c_last).value)
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            nr = copy.copy(rule)
            if rule.formula == [f'A$5="{last_hdr}"']:
                nr.formula = [f'A$5="{DELTA_HDR}"']
            ns.conditional_formatting.add(str(cf.sqref), nr)

    # The source CF has two rules sharing priority 9 (right edge + FINAL column).
    # LibreOffice (recalc.py) silently drops one of a tied pair on save, which erases
    # the right border of the last column. Renumber every rule to a unique priority,
    # keeping the original order, on every sheet.
    for sh in wb.worksheets:
        rules = sorted((r for cf in sh.conditional_formatting for r in cf.rules),
                       key=lambda r: r.priority)
        for i, r in enumerate(rules, 1):
            r.priority = i

    wb.active = 0
    for s in wb.worksheets:
        s.sheet_view.tabSelected = (s is ws)
    out = os.path.join(outdir, os.path.basename(path))
    wb.save(out)
    print(f'[{dept}] {out}')
    print(f'   {"Total":<20} YTD {ty:7.1%}  PY {tp:7.1%}  chg {ty - tp:+.1%}')
    for r in rows[:3] + rows[-3:]:
        y = f'{r[2]:7.1%}' if r[2] is not None else '    n/a'
        print(f'   {r[0]:<20} YTD {y}')
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True)
    ap.add_argument('--outdir', required=True)
    ap.add_argument('--fixed'); ap.add_argument('--service'); ap.add_argument('--parts')
    a = ap.parse_args()
    stores = json.load(open(a.data))['stores']
    os.makedirs(a.outdir, exist_ok=True)
    jobs = [(p, d) for p, d in ((a.fixed, 'fixed'), (a.service, 'service'), (a.parts, 'parts')) if p]
    if not jobs:
        sys.exit('pass at least one of --fixed / --service / --parts')
    for p, d in jobs:
        if os.path.abspath(os.path.join(a.outdir, os.path.basename(p))) == os.path.abspath(p):
            sys.exit(f'--outdir would overwrite the input {p}; use a different folder')
    for p, d in jobs:
        build(p, d, stores, a.outdir)


if __name__ == '__main__':
    main()
