#!/usr/bin/env python3
"""Build the Parts Pricing Peer Comparison workbook from results.json (see references/results_template.json).

Tabs (optional ones only when their section is in results.json):
  Summary, Break-Even, Repair Penetration, Job Tickets, [Price vs List], Matrix Curves, [Same Part Prices],
  [CP Traffic Trend], Methodology

Run recalc.py on the output afterwards and never re-save it with openpyxl.

Usage:
  python3 build_peer_workbook.py --results results.json --out "<Store> vs Peers - CP Parts Pricing <Period>.xlsx"
"""
import argparse, json, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, '..', '..', 'parts-matrix-audit', 'scripts'))
from parse_matrix import unit_price  # noqa: E402
from peer_expr import load_store, main_rule, OTHER_JOB  # noqa: E402

NAVY, GOLD = '1F3864', 'FFC9A04B'
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor=NAVY)
T = Font(bold=True, size=14, color=NAVY); B = Font(bold=True); SEC = Font(bold=True, size=12, color=NAVY)
NOTE = Font(italic=True, color='595959')
TGT = PatternFill('solid', fgColor='FFF2CC'); INPUT = PatternFill('solid', fgColor='DDEBF7')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
USD, USD2, PCT, X = '$#,##0;[Red]-$#,##0', '$#,##0.00', '0.0%', '0.00"x"'
REPAIR_JOBS = ['02 Front brakes', '03 Rear brakes', '04 Battery', '05 Spark plugs', OTHER_JOB]
PER_RO_JOBS = ('09 4x4/diff', OTHER_JOB)
CURVE_COSTS = [2, 5, 10, 20, 40, 75, 125, 200, 350, 600, 1200]


def hdr(ws, r, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(r, i, h); c.font = H; c.fill = HF; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical='center')


def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = T
    ws['A2'] = sub; ws['A2'].font = NOTE


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def note(ws, r, text, ncols, height=45):
    c = ws.cell(r, 1, text); c.alignment = Alignment(wrap_text=True, vertical='top'); c.font = NOTE
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols); ws.row_dimensions[r].height = height


def short(st):
    return st['store'].replace('Findlay ', '')


def build(res, out):
    keys = [res['target']] + res['peers']
    S = {k: load_store(k) for k in keys}
    name = {k: short(S[k]) for k in keys}
    div = {k: S[k]['qlik_division'] for k in keys}
    tgt, peers = res['target'], res['peers']
    period = res['period']
    wb = Workbook()

    # ---------- Break-Even data ----------
    bk = res['basket']
    rows = bk['rows']
    tot = lambda i: sum(r[i] for r in rows)  # noqa: E731
    lines, actual, cost, lst = tot(1), tot(3), tot(4), tot(5)
    sc = res['scenarios']
    jobs = {(d, j): dict(lines=a, ros=b, ps=c, pc=e, ls=f, lc=g) for d, j, a, b, c, e, f, g in res['jobs']['rows']}
    cp = res['jobs']['cp_ros']
    g = jobs.get((div[tgt], OTHER_JOB))
    lab_gross = res.get('labor_gross_per_ro') or (round((g['ls'] - g['lc']) / g['ros'], 2) if g else 0)

    # ---------- Summary ----------
    ws = wb.active; ws.title = 'Summary'; ws.sheet_properties.tabColor = GOLD
    title(ws, f"{name[tgt]} vs. {', '.join(name[p] for p in peers)} - Customer-Pay Parts Pricing",
          f"{period} | CP part lines on ROs whose part-bearing lines are all the store's matrix labor type | "
          'Source: Qlik Closed Parts Tickets + Closed Repair Orders, CDK matrix exports')
    ws['A4'] = 'Bottom line'; ws['A4'].font = SEC
    findings = res.get('findings', [])
    for i, t in enumerate(findings):
        r = 5 + i
        c = ws.cell(r, 1, f'{i + 1}. {t}'); c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); ws.row_dimensions[r].height = 48
    r = 6 + len(findings)
    ws.cell(r, 1, 'Key numbers').font = SEC; r += 1
    hdr(ws, r, ['Measure', 'Value', 'Note']); r += 1
    kn = [('Repair-matrix lines tested', lines, '#,##0', f"Sources {', '.join(bk['repair_sources'])}; {sc['ros']:,} ROs"),
          ('Actual sale as multiple of list', actual / lst, X, 'What customers paid on those lines'),
          (f"{name[tgt]} matrix as multiple of list", tot(6) / lst, X, 'What the matrix calls for')]
    for i, p in enumerate(peers):
        kn.append((f"{name[p]} matrix as multiple of list", tot(7 + i) / lst, X, 'Same parts, peer matrix'))
    kn.append(('Lines below / above own matrix', f"{bk.get('lines_below_own', '-')} / {bk.get('lines_above_own', '-')}", None, '+/- $0.02'))
    for lab, v, fmt, nt in kn:
        ws.cell(r, 1, lab).border = BOX
        c = ws.cell(r, 2, round(v, 4) if isinstance(v, float) else v); c.border = BOX
        if fmt: c.number_format = fmt
        ws.cell(r, 3, nt).border = BOX
        r += 1
    widths(ws, [40, 14, 50, 12, 12, 12, 12, 12, 12])

    # ---------- Break-Even ----------
    ws = wb.create_sheet('Break-Even')
    title(ws, f"What would it cost {name[tgt]} to bring its repair matrix in line with the other stores?",
          f"{name[tgt]} CP sales lines on repair-matrix sources {', '.join(bk['repair_sources'])} | {period} | {sc['ros']:,} ROs with a repair-matrix part")
    ws.cell(4, 1, 'A. The basket priced through each matrix (straight matrix, every line)').font = SEC
    cols = ['Unit cost band', 'Lines', 'Cost $', 'List $', f'{name[tgt]} actual $', f'{name[tgt]} matrix $'] + [f'{name[p]} matrix $' for p in peers]
    nc = len(cols)
    hdr(ws, 5, cols)
    for i, row in enumerate(rows):
        rr = 6 + i
        vals = [row[0][2:] if row[0][:2] in ('1 ', '2 ', '3 ', '4 ', '5 ', '6 ', '7 ') else row[0], row[1], row[4], row[5], row[3]] + list(row[6:])
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            if j >= 3: c.number_format = USD
    rt = 6 + len(rows)
    ws.cell(rt, 1, 'Total').font = B
    for j in range(2, nc + 1):
        col = get_column_letter(j)
        c = ws.cell(rt, j, f'=SUM({col}6:{col}{rt - 1})'); c.font = B; c.number_format = USD if j >= 3 else '#,##0'
    ws.cell(rt + 1, 1, 'Multiple of list').font = B
    for j in range(5, nc + 1):
        col = get_column_letter(j)
        c = ws.cell(rt + 1, j, f'={col}{rt}/$D${rt}'); c.number_format = X; c.fill = TGT
    if 'lines_below_own' in bk:
        ws.cell(rt + 2, 1, f"{bk['lines_below_own']} of {lines} lines sold below {name[tgt]}'s own matrix and "
                           f"{bk.get('lines_above_own', 0)} above (+/- $0.02).").font = NOTE
    r = rt + 4
    ws.cell(r, 1, 'B. Break-even: cap each line at the lower of its actual price or the new matrix').font = SEC
    ws.cell(r + 1, 1, 'Lines already priced below the new matrix stay where they are. Parts gross only; annual = period x (12 / months in period).').font = NOTE
    r += 2
    hdr(ws, r, ['Scenario', 'Parts sale $', 'Parts cost $', 'Parts gross $', 'GP %', 'Gross given up / period',
                'Gross given up / yr', 'Gross per added RO (parts + labor)', 'Added ROs / period to break even', f"Lift vs. {sc['ros']:,} ROs"])
    mp = sc['midpoint']
    scen = [('Today (actual)', actual), (f"Midpoint: halfway {name[tgt]} -> {name[mp['peer']]}", mp['sale'])]
    scen += [(f'Match {name[p]} matrix', sc['cap'][p]) for p in peers if p in sc['cap']]
    r0 = r + 1
    ra = r0 + len(scen) + 1
    months = res.get('months_in_period', 3)
    for i, (nm, sale) in enumerate(scen):
        rr = r0 + i
        ws.cell(rr, 1, nm); ws.cell(rr, 2, round(sale, 2)); ws.cell(rr, 3, round(cost, 2))
        ws.cell(rr, 4, f'=B{rr}-C{rr}'); ws.cell(rr, 5, f'=D{rr}/B{rr}')
        if i:
            ws.cell(rr, 6, f'=$D${r0}-D{rr}'); ws.cell(rr, 7, f'=F{rr}*12/{months}')
            ws.cell(rr, 8, f"=D{rr}/{sc['ros']}+$B${ra}")
            ws.cell(rr, 9, f'=ROUNDUP(F{rr}/H{rr},0)'); ws.cell(rr, 10, f"=I{rr}/{sc['ros']}")
        for j in range(1, 11):
            c = ws.cell(rr, j); c.border = BOX
            c.number_format = {5: PCT, 9: '#,##0', 10: PCT}.get(j, USD)
        if i == 1:
            for j in range(1, 11): ws.cell(rr, j).fill = TGT
    ws.cell(ra, 1, 'Labor gross per added repair RO (assumption)').font = B
    c = ws.cell(ra, 2, lab_gross); c.number_format = USD2; c.fill = INPUT
    note(ws, ra + 1, (f"Assumption: each added RO carries the average repair-matrix parts gross at the new price plus {name[tgt]}'s "
                      f"average labor gross on a diagnostic/general-repair RO (${lab_gross:,.0f}, {period}). Change the blue cell to test "
                      'other values. Ignores any lift in maintenance traffic or retention from a better price reputation.'), 10)
    widths(ws, [40, 14, 13, 14, 12, 15, 16, 17, 15, 12])

    # ---------- Repair Penetration ----------
    ws = wb.create_sheet('Repair Penetration')
    title(ws, f'Are {name[tgt]} customers buying repairs at the same rate as the other stores?',
          f"{period} | Matrix labor-type lines with parts | Denominator = ROs with any customer-pay (C*) labor line")
    lbl = {'02 Front brakes': 'Front brakes', '03 Rear brakes': 'Rear brakes', '04 Battery': 'Battery',
           '05 Spark plugs': 'Spark plugs', OTHER_JOB: 'Diagnostic / general repair'}
    hdr(ws, 4, ['Store', 'CP ROs'] + [f'{lbl[c]} - ROs per 100 CP ROs' for c in REPAIR_JOBS] +
        ['Repair parts $ per CP RO', 'Repair parts GP %', 'Repair parts gross per CP RO'])
    for i, k in enumerate(keys):
        rr = 5 + i; d = div[k]; n = cp[d]
        ws.cell(rr, 1, name[k]); ws.cell(rr, 2, n).number_format = '#,##0'
        for j, cat in enumerate(REPAIR_JOBS):
            x = jobs.get((d, cat))
            ws.cell(rr, 3 + j, round(x['ros'] / n * 100, 2) if x else 0).number_format = '0.0'
        ps = sum(jobs[(d, c)]['ps'] for c in REPAIR_JOBS if (d, c) in jobs)
        pc = sum(jobs[(d, c)]['pc'] for c in REPAIR_JOBS if (d, c) in jobs)
        ws.cell(rr, 8, round(ps / n, 2)).number_format = USD2
        ws.cell(rr, 9, round((ps - pc) / ps, 4) if ps else 0).number_format = PCT
        ws.cell(rr, 10, round((ps - pc) / n, 2)).number_format = USD2
        for j in range(1, 11):
            ws.cell(rr, j).border = BOX
            if k == tgt: ws.cell(rr, j).fill = TGT
    r = 6 + len(keys)
    if res.get('broad'):
        br = res['broad']
        ws.cell(r, 1, 'Broad check: ROs with any matrix-type line carrying $50+ of parts cost (includes fluid services), per 100 CP ROs').font = SEC
        hdr(ws, r + 1, ['Store', f"{br['prior_label']} CP ROs", 'ROs w/ $50+ parts line', 'Per 100', f"{br['current_label']} CP ROs",
                        'ROs w/ $50+ parts line', 'Per 100', f"Parts GP % {br['prior_label']}", f"Parts GP % {br['current_label']}"])
        for i, k in enumerate(keys):
            rr = r + 2 + i; a = br['prior'][div[k]]; b = br['current'][div[k]]
            vals = [name[k], a[0], a[1], f'=C{rr}/B{rr}*100', b[0], b[1], f'=F{rr}/E{rr}*100',
                    round((a[2] - a[3]) / a[2], 4), round((b[2] - b[3]) / b[2], 4)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(rr, j, v); c.border = BOX
                c.number_format = {4: '0.0', 7: '0.0', 8: PCT, 9: PCT}.get(j, '#,##0')
                if k == tgt: c.fill = TGT
        r += 3 + len(keys)
    note(ws, r, 'Job categories come from op-code description keywords (see Methodology) and are approximate. Store op-code naming and job content differ.', 10)
    widths(ws, [20, 10, 14, 14, 12, 12, 16, 14, 12, 15])

    # ---------- Job Tickets ----------
    ws = wb.create_sheet('Job Tickets')
    title(ws, f'What a customer pays for common jobs - parts + labor, {period}',
          'Matrix labor-type lines with parts. Job content differs by store (e.g. pads only vs. pads + rotors), so compare with care.')
    hdr(ws, 4, ['Job', 'Measure'] + [name[k] for k in keys] + [f'{name[tgt]} vs. peer avg'])
    cats = sorted({j for _, j in jobs if j != '01 Menu/maint' and j != '10 Tires'})
    rr = 5
    pcol, qcols = get_column_letter(3), (get_column_letter(4), get_column_letter(2 + len(keys)))
    for job in cats:
        per_ro = job in PER_RO_JOBS
        unit = 'RO' if per_ro else 'line'
        meas = [(f'Jobs ({unit}s)', lambda x: x['ros'] if per_ro else x['lines'], '#,##0'),
                (f'Parts $ per {unit}', lambda x: x['ps'] / (x['ros'] if per_ro else x['lines']), USD),
                (f'Labor $ per {unit}', lambda x: x['ls'] / (x['ros'] if per_ro else x['lines']), USD),
                (f'Total ticket per {unit}', lambda x: (x['ps'] + x['ls']) / (x['ros'] if per_ro else x['lines']), USD)]
        for kk, (ml, fn, fmt) in enumerate(meas):
            ws.cell(rr, 1, job[3:] if kk == 0 else ''); ws.cell(rr, 2, ml)
            for j, k in enumerate(keys):
                x = jobs.get((div[k], job))
                ws.cell(rr, 3 + j, round(fn(x), 2) if x and (x['ros'] if per_ro else x['lines']) else None).number_format = fmt
            lc = 3 + len(keys)
            if kk:
                ws.cell(rr, lc, f'=IFERROR({pcol}{rr}/AVERAGE({qcols[0]}{rr}:{qcols[1]}{rr})-1,"")').number_format = PCT
            for j in range(1, lc + 1):
                ws.cell(rr, j).border = BOX
                if kk == 3: ws.cell(rr, j).font = B
            ws.cell(rr, 3).fill = TGT
            rr += 1
    note(ws, rr + 1, '4x4/diff and diagnostic/general repair are per RO because one job spans several lines. Blank = no matching op codes at that store.', 3 + len(keys))
    widths(ws, [24, 22] + [14] * len(keys) + [18])

    # ---------- Price vs List (optional) ----------
    if res.get('pricepos'):
        pp = res['pricepos']
        ws = wb.create_sheet('Price vs List')
        title(ws, 'Where each store prices relative to list', f"{pp['period']} | share of CP parts sales dollars by price position (unit sale vs. list, per line)")
        n = len(keys)
        hdr(ws, 4, ['Price position'] + [f'{name[k]} sale $' for k in keys] + [f'{name[k]} % of $' for k in keys])
        buckets = list(pp['rows'])
        tr = 5 + len(buckets)
        for i, b in enumerate(buckets):
            rr = 5 + i
            ws.cell(rr, 1, b[2:] if b[1] == ' ' else b).border = BOX
            for j, k in enumerate(keys):
                c = ws.cell(rr, 2 + j, round(pp['rows'][b][k][0], 2)); c.number_format = USD; c.border = BOX
                col = get_column_letter(2 + j)
                c = ws.cell(rr, 2 + n + j, f'={col}{rr}/{col}${tr}'); c.number_format = PCT; c.border = BOX
        ws.cell(tr, 1, 'Total').font = B
        for j in range(2 * n):
            col = get_column_letter(2 + j)
            c = ws.cell(tr, 2 + j, f'=SUM({col}5:{col}{tr - 1})'); c.font = B; c.number_format = USD if j < n else PCT
        r = tr + 2
        hdr(ws, r, ['Store', 'Sale / list', '% of $ above list', 'Markup on above-list $ (x list)'])
        above = [b for b in buckets if 'List +' in b]
        for i, k in enumerate(keys):
            rr = r + 1 + i
            S_ = sum(pp['rows'][b][k][0] for b in buckets if not b.startswith('0'))
            L_ = sum(pp['rows'][b][k][1] for b in buckets)
            aS = sum(pp['rows'][b][k][0] for b in above); aL = sum(pp['rows'][b][k][1] for b in above)
            tS = sum(pp['rows'][b][k][0] for b in buckets)
            for j, (v, f) in enumerate([(name[k], None), (S_ / L_ if L_ else 0, X), (aS / tS if tS else 0, PCT), (aS / aL if aL else 0, X)], 1):
                c = ws.cell(rr, j, round(v, 4) if isinstance(v, float) else v); c.border = BOX
                if f: c.number_format = f
                if k == tgt: c.fill = TGT
        widths(ws, [22] + [15] * (2 * n))

    # ---------- Matrix Curves ----------
    ws = wb.create_sheet('Matrix Curves')
    ratio = round(lst / cost, 3) if cost else 1.6
    title(ws, 'Main repair matrix - price as a multiple of list, by part cost',
          f"Each store's main repair source. List assumed = cost x {ratio} (list/cost on {name[tgt]}'s repair basket); list-break matrices use that list.")
    hdr(ws, 4, ['Unit cost'] + [f'{name[k]} (x list)' for k in keys] + [f'{name[tgt]} vs. peer avg'])
    rules = {k: main_rule(S[k]) for k in keys}
    for i, c_ in enumerate(CURVE_COSTS):
        rr = 5 + i; l_ = round(c_ * ratio, 2)
        ws.cell(rr, 1, c_).number_format = USD2
        for j, k in enumerate(keys):
            c = ws.cell(rr, 2 + j, round(unit_price(rules[k], c_, l_) / l_, 3)); c.number_format = X; c.border = BOX
        ws.cell(rr, 2).fill = TGT
        lc = 2 + len(keys)
        c = ws.cell(rr, lc, f'=B{rr}/AVERAGE(C{rr}:{get_column_letter(lc - 1)}{rr})-1'); c.number_format = PCT; c.border = BOX
    rr = 6 + len(CURVE_COSTS)
    ws.cell(rr, 1, 'Main repair source').font = B
    for j, k in enumerate(keys):
        ws.cell(rr, 2 + j, str(S[k].get('main_repair_source')))
    widths(ws, [12] + [18] * len(keys) + [20])

    # ---------- Same Part Prices (optional) ----------
    if res.get('same_part'):
        sp = res['same_part']
        ws = wb.create_sheet('Same Part Prices')
        title(ws, 'Same part number, average CP unit price', sp.get('subtitle', ''))
        r = 4
        if sp.get('index'):
            hdr(ws, r, ['Peer', f'{name[tgt]} price index vs. peer', 'Matched part numbers', f'{name[tgt]} $ covered'])
            for i, (p, idx, n_, d_) in enumerate(sp['index']):
                for j, v in enumerate([name.get(p, p), idx, n_, d_], 1):
                    c = ws.cell(r + 1 + i, j, v); c.border = BOX
                ws.cell(r + 1 + i, 2).number_format = '0.000'; ws.cell(r + 1 + i, 4).number_format = USD
            r += 2 + len(sp['index'])
            ws.cell(r, 1, f'Index < 1.00 = {name[tgt]} cheaper on the same parts, weighted by {name[tgt]} quantity.').font = NOTE
            r += 2
        for tb in sp.get('tables', []):
            ws.cell(r, 1, tb['heading']).font = SEC; r += 1
            hdr(ws, r, ['Part', 'Description'] + [name[k] for k in keys] + ['List']); r += 1
            for row in tb['rows']:
                for j, v in enumerate(row, 1):
                    c = ws.cell(r, j, v if v is not None else '-'); c.border = BOX
                    if j >= 3 and v is not None: c.number_format = USD2
                ws.cell(r, 3).fill = TGT
                r += 1
            r += 1
        widths(ws, [20, 26] + [12] * (len(keys) + 1))

    # ---------- CP Traffic Trend (optional) ----------
    if res.get('traffic'):
        tr_ = res['traffic']
        ws = wb.create_sheet('CP Traffic Trend')
        title(ws, 'Customer-pay RO count, current vs. prior year', 'ROs with at least one customer-pay (C*) labor line; Closed Repair Orders app')
        hdr(ws, 4, ['Store', f"{tr_['prior_label']} CP ROs", f"{tr_['current_label']} CP ROs", 'YoY',
                    f"{tr_['prior_label']} CP parts $", f"{tr_['current_label']} CP parts $", 'YoY'])
        for i, k in enumerate(keys):
            rr = 5 + i; a, b, c_, d_ = tr_['rows'][div[k]]
            for j, v in [(1, name[k]), (2, a), (3, b), (5, c_), (6, d_)]:
                ws.cell(rr, j, v)
            ws.cell(rr, 4, f'=C{rr}/B{rr}-1'); ws.cell(rr, 7, f'=F{rr}/E{rr}-1')
            for j in range(1, 8):
                c = ws.cell(rr, j); c.border = BOX
                c.number_format = PCT if j in (4, 7) else (USD if j in (5, 6) else '#,##0')
                if k == tgt: c.fill = TGT
        note(ws, 6 + len(keys), 'Correlation only. Traffic is driven by many factors (UIO, marketing, capacity, staffing, menu pricing).', 7, 30)
        widths(ws, [20, 14, 14, 9, 16, 16, 9])

    # ---------- Methodology ----------
    ws = wb.create_sheet('Methodology')
    title(ws, 'Methodology and limitations', '')
    lt = '; '.join(f"{name[k]} {'/'.join(S[k]['matrix_labor_types'])}" for k in keys)
    rows_ = [
        ('Period', period + (f" (price vs. list: {res['pricepos']['period']})" if res.get('pricepos') else '')),
        ('Matrix labor types', lt + '. Confirmed with the user before the run.'),
        ('Repair basket', f"{name[tgt]} CP sales lines (qty > 0, invoice type RO) on ROs whose part-bearing lines are all its matrix labor type, "
                          f"on sources sharing its main repair rule ({', '.join(bk['repair_sources'])}). Re-priced with each store's main repair "
                          'matrix from the CDK exports in reference/parts_matrices. Same parts, cost and list; only the matrix changes.'),
        ('Break-even', 'Each scenario caps the unit price at the lower of actual or the scenario matrix. Midpoint = average of the target and '
                       'midpoint-peer matrix price per line. Added-RO gross = new average parts gross per affected RO + the labor-gross assumption.'),
        ('Repair penetration / job tickets', 'Closed Repair Orders app, matrix labor-type lines with parts cost > 0. Job categories from op-code '
                                             'description keywords (spark plug; menu/maintenance; front/rear brake; battery; transmission; coolant; '
                                             'induction; 4x4/diff; tire; everything else = diagnostic/general repair). Denominator = ROs with any C* line.'),
        ('Limitations', 'Store op-code naming and job content differ. Peer matrices may route a part to a different source than the target. '
                        'Annual dollars are the period scaled to 12 months. Competitor (non-Findlay) pricing is not included.'),
    ] + [tuple(x) for x in res.get('methodology_extra', [])]
    for i, (k, v) in enumerate(rows_):
        ws.cell(4 + i, 1, k).font = B; ws.cell(4 + i, 1).alignment = Alignment(vertical='top')
        c = ws.cell(4 + i, 2, v); c.alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[4 + i].height = 60
    widths(ws, [24, 110])

    wb.save(out)
    print(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--results', required=True)
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    build(json.load(open(a.results)), a.out)


if __name__ == '__main__':
    main()
