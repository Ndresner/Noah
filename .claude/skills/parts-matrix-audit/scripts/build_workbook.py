#!/usr/bin/env python3
"""Build the Parts Pricing vs. Matrix audit workbook from results.json + matrix.json.

    python3 build_workbook.py --results results.json --matrix matrix.json --out "<Store> CP Parts Pricing Audit.xlsx"

Then ALWAYS run recalc.py on the output (CLAUDE.md rule) and never re-save it with openpyxl afterwards.
results.json schema: see references/results_template.json.
"""
import argparse, json
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

F = 'Arial'
NAVY = PatternFill('solid', fgColor='1F3864'); GOLD = PatternFill('solid', fgColor='FFC9A04B')
GREY = PatternFill('solid', fgColor='D9D9D9')
GRN = PatternFill('solid', fgColor='C6EFCE'); YEL = PatternFill('solid', fgColor='FFEB9C'); RED = PatternFill('solid', fgColor='FFC7CE')
HF = Font(name=F, bold=True, color='FFFFFF', size=10); BF = Font(name=F, size=10); BB = Font(name=F, size=10, bold=True)
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
MONEY = '$#,##0.00;[Red]-$#,##0.00'; MONEY0 = '$#,##0;[Red]-$#,##0'; PCT = '0.0%'; INT = '#,##0'

ROLE_TEXT = {
    'matrix': 'MATRIX - tested line by line',
    'list': 'LIST / MSRP type - should sell at list',
    'reimbursed': 'NOT MATRIX - reimbursed / factory-paid rate',
    'package': 'NOT MATRIX - package pricing',
    'discount': 'NOT MATRIX - discounted (fleet / employee / prepaid)',
    'unclear': 'UNCLEAR - confirm intended rule',
    'mixed': 'NOT TESTED - lines span several labor types',
}


def header(ws, row, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row, i, h); c.font = HF; c.fill = NAVY; c.border = BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = 32


def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = Font(name=F, size=14, bold=True, color='1F3864')
    ws['A2'] = sub; ws['A2'].font = Font(name=F, size=9, italic=True, color='595959')


def body(ws, r1, r2, ncol, fmts):
    for r in range(r1, r2 + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c); cell.font = BF; cell.border = BOX
            if c in fmts: cell.number_format = fmts[c]


def rule_text(matrix, src):
    r = matrix['sources'].get(str(src))
    if not r: return 'No matrix row (system default price)'
    if r.get('unsupported'): return 'Unsupported formula - not tested'
    if len(r['bands']) == 1:
        b = r['bands'][0]
        return b['base'].title() + (f" {'+' if b['pct'] > 0 else '-'}{abs(b['pct']):g}%" if b['pct'] else '')
    lo = min(b['pct'] for b in r['bands']); hi = max(b['pct'] for b in r['bands'])
    bases = sorted({b['base'].title() for b in r['bands']})
    return f"{r['break_field'].title()}-break matrix: {'/'.join(bases)} {lo:+g}% to {hi:+g}% ({len(r['bands'])} bands)"


def split_key(k):
    # %Parts_Detail looks like "3PA62984/404712*10/2026-08-25" (logon/RO*line/open date)
    k = k.split('/', 1)[1] if k.count('/') >= 2 else k
    ro, rest = k.split('*', 1)
    line, dt = rest.split('/', 1)
    return ro, int(line), dt


def cause(e):
    if e['qty'] < 0: return 'Core / return credit (negative qty) - not a pricing miss'
    if e['unit_sale'] < e['unit_cost'] - 0.01: return 'Sold below cost'
    if e['list'] == 0: return '$0 list in parts master - manual price'
    if e['var'] < 0 and abs(e['unit_sale'] - e['list']) <= max(0.02, e['list'] * 0.005): return 'Sold at list instead of matrix'
    if e['var'] > 0 and e['list'] > 0 and e['unit_sale'] > e['list'] * 2: return 'Priced far above list'
    return 'Manual price override'


def usd(x):
    return ('-' if x < 0 else '') + f'${abs(x):,.0f}'


def auto_findings(res):
    t, bs, mt = res['totals'], res['by_source'], '/'.join(res['matrix_types'])
    out = []
    pct = t['match_lines'] / t['lines'] if t['lines'] else 0
    under = sum(s['under_amt'] for s in bs); over = sum(s['over_amt'] for s in bs)
    out.append(f"{mt} parts hit the matrix on {pct:.1%} of lines ({t['match_lines']:,} of {t['lines']:,}). "
               f"Actual {usd(t['actual'])} vs. matrix {usd(t['matrix'])}: net {usd(t['actual'] - t['matrix'])} "
               f"({usd(under)} under, {usd(over)} over).")
    top = sorted([s for s in bs if s['under_amt'] < 0], key=lambda s: s['under_amt'])[:3]
    if top and under:
        parts = [f"source {s['source']} {s.get('description', '')}".strip() + f" {usd(s['under_amt'])}"
                 f" ({s['under_at_list_lines']} of {s['under_lines']} under lines sold at list)" for s in top]
        share = sum(s['under_amt'] for s in top) / under
        out.append(f"Largest under-matrix sources ({share:.0%} of under $): " + '; '.join(parts) + '.')
    atl = sum(s['under_at_list_lines'] for s in bs); ul = sum(s['under_lines'] for s in bs)
    if ul:
        out.append(f"{atl} of {ul} under-matrix lines ({atl / ul:.0%}) were sold at list instead of the matrix price.")
    zl = [s for s in bs if s['zero_list_lines']]
    if zl:
        out.append(f"{sum(s['zero_list_lines'] for s in zl)} lines carry a $0 list price in the parts master (sources "
                   f"{', '.join(str(s['source']) for s in zl)}). Where the matrix keys off list, the system cannot price them and the counter keys a price by hand.")
    bc = sum(s['below_cost_lines'] for s in bs)
    if bc:
        out.append(f"{bc} lines sold below cost ({usd(sum(s['below_cost_amt'] for s in bs))}).")
    ex = res.get('exceptions', {})
    ros = defaultdict(set)
    for d in ('under', 'over'):
        for e in ex.get(d, []):
            if e['qty'] > 0:   # core / return credits are not pricing decisions
                ros[split_key(e['key'])[0]].add(d)
    both = [r for r, s in ros.items() if len(s) == 2]
    if both:
        out.append(f"RO(s) {', '.join(both)} show parts repriced both below and above matrix on the same RO - "
                   f"looks like prices were moved to hit a quoted total. Worth a conversation.")
    for lt in res.get('labor_types', []):
        if lt['role'] in ('list', 'reimbursed', 'package', 'unclear') and lt.get('note'):
            out.append(f"{lt['code']}: {lt['note']}")
    return out


def build(res, matrix, out):
    wb = openpyxl.Workbook()
    mt = '/'.join(res['matrix_types'])
    period = (f"{res['period_detail']}. Source: Qlik Closed Parts Tickets + Closed Repair Orders, "
              f"Division = {res['division']}.")
    bs = sorted(res['by_source'], key=lambda s: s['under_amt'])
    n = len(bs); first, last = 5, 4 + n; tot = last + 1

    # ---------- Summary ----------
    ws = wb.active; ws.title = 'Summary'; ws.sheet_properties.tabColor = 'C9A04B'
    title(ws, f"{res['store']} - Customer Pay Parts Pricing vs. Matrix ({res['period_label']})", period)
    ws['A4'] = f"{mt} part lines tested against the source matrix"; ws['A4'].font = BB
    src = "'Matrix Test by Source'"
    kp = [('ROs tested', f'{src}!E{tot}', INT),
          ('Part lines tested', f'{src}!D{tot}', INT),
          ('Lines priced per matrix (within 0.5% / $0.02)', f'{src}!H{tot}', INT),
          ('Line compliance %', '=B7/B6', PCT),
          ('Actual parts sale', f'{src}!F{tot}', MONEY0),
          ('Matrix parts sale', f'{src}!G{tot}', MONEY0),
          ('Under-matrix $', f'{src}!J{tot}', MONEY0),
          ('Over-matrix $', f'{src}!M{tot}', MONEY0),
          ('Net variance (actual - matrix)', '=B9-B10', MONEY0),
          ('Net variance per month', f"=B13/{res.get('months', 1)}", MONEY0)]
    for i, (k, f, fmt) in enumerate(kp, 5):
        ws.cell(i, 1, k).font = BF; c = ws.cell(i, 2, '=' + f if not f.startswith('=') else f)
        c.number_format = fmt; c.font = BB; ws.cell(i, 1).border = BOX; c.border = BOX
    r = 17
    ws.cell(r, 1, 'Key findings').font = BB
    for t in auto_findings(res) + res.get('findings', []):
        r += 1; c = ws.cell(r, 1, t); c.font = BF; c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6); ws.row_dimensions[r].height = 42
    if res.get('questions'):
        r += 2; ws.cell(r, 1, 'Points to confirm').font = BB
        for q in res['questions']:
            r += 1; c = ws.cell(r, 1, '- ' + q); c.font = BF; c.alignment = Alignment(wrap_text=True, vertical='top')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6); ws.row_dimensions[r].height = 30
    ws.column_dimensions['A'].width = 52; ws.column_dimensions['B'].width = 16
    for col in 'CDEF': ws.column_dimensions[col].width = 16

    # ---------- Labor Type Review ----------
    ws = wb.create_sheet('Labor Type Review'); ws.sheet_properties.tabColor = 'C9A04B'
    title(ws, 'Customer Pay Labor Types - How Each Is Treated', period + ' Parts sale/cost from RO lines with parts.')
    cols = ['Labor Type', 'What it is', 'ROs', 'Parts Sale', 'Parts Cost', 'Gross Margin %', 'Treatment', 'Notes']
    header(ws, 4, cols)
    lts = res.get('labor_types', [])
    for i, lt in enumerate(lts, 5):
        vals = [lt['code'], lt.get('description', ''), lt.get('ros'), lt.get('parts_sale'), lt.get('parts_cost'),
                f'=IF(N(D{i})>0,(D{i}-E{i})/D{i},"")', ROLE_TEXT.get(lt['role'], lt['role']), lt.get('note', '')]
        for j, v in enumerate(vals, 1): ws.cell(i, j, v)
    if lts:
        body(ws, 5, 4 + len(lts), 8, {3: INT, 4: MONEY0, 5: MONEY0, 6: PCT})
        for rr, lt in enumerate(lts, 5):
            ws.cell(rr, 7).fill = GRN if lt['role'] == 'matrix' else (YEL if lt['role'] == 'unclear' else GREY)
            for cc in (2, 7, 8): ws.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical='top')
            ws.row_dimensions[rr].height = 30
    for col, w in zip('ABCDEFGH', [12, 40, 9, 13, 13, 11, 36, 50]): ws.column_dimensions[col].width = w
    ws.freeze_panes = 'B5'

    # ---------- Matrix Test by Source ----------
    ws = wb.create_sheet('Matrix Test by Source'); ws.sheet_properties.tabColor = '1F3864'
    title(ws, f'{mt} - Actual vs. Matrix by Parts Source',
          period + f" ROs whose part-bearing lines are all {mt} ({res['totals']['ros']:,}).")
    cols = ['Source', 'Description', 'Matrix Rule', 'Lines', 'ROs', 'Actual Sale', 'Matrix Sale', 'Lines Match', 'Match %',
            'Under-Matrix $', 'Under Lines', 'Under Lines Sold at List', 'Over-Matrix $', 'Over Lines',
            'Below-Cost Lines', '$0 List Lines', 'Net Var (Actual-Matrix)']
    header(ws, 4, cols)
    for i, s in enumerate(bs, first):
        vals = [int(s['source']) if str(s['source']).isdigit() else s['source'], s.get('description', ''),
                rule_text(matrix, s['source']), s['lines'], s['ros'], s['actual'], s['matrix'], s['match_lines'],
                f'=IF(D{i}>0,H{i}/D{i},"")', s['under_amt'], s['under_lines'], s['under_at_list_lines'],
                s['over_amt'], s['over_lines'], s['below_cost_lines'], s['zero_list_lines'], f'=F{i}-G{i}']
        for j, v in enumerate(vals, 1): ws.cell(i, j, v)
    ws.cell(tot, 1, 'TOTAL'); ws.cell(tot, 2, mt)
    for j in range(4, 18):
        Lc = get_column_letter(j)
        ws.cell(tot, j, {9: f'=H{tot}/D{tot}', 17: f'=F{tot}-G{tot}', 5: res['totals']['ros']}.get(j, f'=SUM({Lc}{first}:{Lc}{last})'))
    body(ws, first, tot, 17, {4: INT, 5: INT, 6: MONEY0, 7: MONEY0, 8: INT, 9: PCT, 10: MONEY0, 11: INT, 12: INT,
                              13: MONEY0, 14: INT, 15: INT, 16: INT, 17: MONEY0})
    for j in range(1, 18):
        ws.cell(tot, j).font = BB; ws.cell(tot, j).fill = GOLD
    rng = f'I{first}:I{last}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['0.95'], fill=GRN, font=Font(color='006100')))
    ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['0.85', '0.9499'], fill=YEL, font=Font(color='9C6500')))
    ws.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['0.85'], fill=RED, font=Font(color='9C0006')))
    ws.cell(tot + 2, 1, 'Under/Over use a tolerance of the greater of $0.02 or 0.5% of the matrix unit price. ROs column '
            'is distinct per source; the total is distinct ROs overall. Net variance includes sub-tolerance rounding.').font = Font(name=F, size=8, italic=True)
    ws.column_dimensions['A'].width = 9; ws.column_dimensions['B'].width = 30; ws.column_dimensions['C'].width = 44
    for j in range(4, 18): ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = 'D5'

    # ---------- Exceptions ----------
    ex = res.get('exceptions', {})
    ws = wb.create_sheet('Top Exceptions' if not res.get('exceptions_complete') else 'All Exceptions')
    ws.sheet_properties.tabColor = '1F3864'
    title(ws, f'{mt} - Line-Level Pricing Exceptions',
          'Every exception line.' if res.get('exceptions_complete') else
          f"Largest {len(ex.get('under', []))} under-matrix and {len(ex.get('over', []))} over-matrix lines by $.")
    header(ws, 4, ['Direction', 'RO #', 'RO Line', 'Open Date', 'Part Number', 'Source', 'Qty', 'Unit Cost', 'List',
                   'Unit Sale', 'Matrix Price', 'Ext Variance $', 'Sale vs List %', 'Likely Cause'])
    r = 5
    for d, label in (('under', 'Under'), ('over', 'Over')):
        for e in ex.get(d, []):
            ro, line, dt = split_key(e['key'])
            vals = [label, int(ro) if ro.isdigit() else ro, line, dt, e['part'], int(e['source']) if str(e['source']).isdigit() else e['source'],
                    e['qty'], e['unit_cost'], e['list'], e['unit_sale'], f'=IF(G{r}<>0,J{r}-L{r}/G{r},"")', e['var'],
                    f'=IF(I{r}>0,J{r}/I{r}-1,"no list")', cause(e)]
            for j, v in enumerate(vals, 1): ws.cell(r, j, v)
            r += 1
    if r > 5:
        body(ws, 5, r - 1, 14, {7: INT, 8: MONEY, 9: MONEY, 10: MONEY, 11: MONEY, 12: MONEY, 13: PCT})
        for rr in range(5, r):
            f = ws.cell(rr, 14).value
            ws.cell(rr, 14).fill = RED if f in ('Sold below cost', 'Priced far above list') else (YEL if f.startswith('$0') else GREY)
    for col, w in zip('ABCDEFGHIJKLMN', [10, 9, 8, 11, 18, 8, 6, 11, 11, 11, 12, 13, 12, 36]): ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A5'

    # ---------- Extra analysis sheets (optional) ----------
    for xs in res.get('extra_sheets', []):
        ws = wb.create_sheet(xs['name'][:31]); ws.sheet_properties.tabColor = 'C9A04B'
        title(ws, xs['title'], xs.get('subtitle', ''))
        r = 4
        for tb in xs['tables']:
            if tb.get('heading'):
                ws.cell(r, 1, tb['heading']).font = BB; r += 1
            header(ws, r, tb['columns']); r0 = r + 1
            for row in tb['rows']:
                r += 1
                for j, v in enumerate(row, 1): ws.cell(r, j, v)
            fm = {int(k): {'money': MONEY0, 'money2': MONEY, 'pct': PCT, 'int': INT, 'x': '0.00"x"'}.get(v, v)
                  for k, v in tb.get('formats', {}).items()}
            if r >= r0: body(ws, r0, r, len(tb['columns']), fm)
            for hl in tb.get('highlight_rows', []):
                for j in range(1, len(tb['columns']) + 1): ws.cell(r0 + hl, j).fill = YEL
            if tb.get('note'):
                r += 1; ws.cell(r, 1, tb['note']).font = Font(name=F, size=8, italic=True)
            r += 2
        for j, w in enumerate(xs.get('widths', []), 1): ws.column_dimensions[get_column_letter(j)].width = w

    # ---------- Matrix Reference ----------
    ws = wb.create_sheet('Matrix Reference'); ws.sheet_properties.tabColor = '808080'
    title(ws, 'Pricing Matrix by Source (as provided)',
          'CDK price code: "COST + 275.00" = cost +275%; "LIST + 30.00" = list +30%. Break field = value tested against the range.')
    header(ws, 4, ['Source', 'Description', 'Price Field', 'Break Field', 'Break Range', 'Formula'])
    descs = {str(s['source']): s.get('description', '') for s in res['by_source']}
    descs.update({k: v['description'] for k, v in matrix.get('source_info', {}).items() if v.get('description')})
    r = 5
    for s, rule in matrix['sources'].items():
        for b in sorted(rule['bands'], key=lambda b: b['lo']):
            vals = [int(s) if s.isdigit() else s, descs.get(s, ''), rule['price_field'], rule['break_field'],
                    f"{b['lo']:,.2f} - {b['hi']:,.2f}", b['raw']]
            for j, v in enumerate(vals, 1): ws.cell(r, j, v)
            r += 1
    body(ws, 5, r - 1, 6, {})
    nm = matrix.get('sources_without_matrix') or []
    if nm:
        ws.cell(r + 1, 1, 'Sources with no matrix row (system default price): ' + ', '.join(nm)).font = Font(name=F, size=9, italic=True)
    for col, w in zip('ABCDEF', [9, 32, 12, 12, 24, 18]): ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A5'

    # ---------- Methodology ----------
    ws = wb.create_sheet('Methodology'); ws.sheet_properties.tabColor = '808080'
    title(ws, 'Methodology, Definitions & Limitations', '')
    mx = res.get('mixed') or {}
    notes = [
        ('Period', res['period_detail'] + ' (Qlik Year/Month = RO close date).'),
        ('Store filter', f"Division = {res['division']} (both Qlik apps)."),
        ('Part lines', 'Qlik "Closed Parts Tickets", invoice sale type = RO. Grain = one part line on an RO.'),
        ('Labor type', f'Part lines carry no labor type in Qlik. Each RO was classified in "Closed Repair Orders" from its labor lines '
                       f'that carry parts cost; only ROs whose part-bearing lines are ALL {mt} were tested.'),
        ('Not tested', f"ROs with parts on any other labor type (warranty, internal, list-priced or reimbursed C types). "
                       + (f"{mx.get('ros', 0):,} ROs mixed {mt} with other C types (${mx.get('parts_sale', 0):,.0f} parts sale)." if mx else '')),
        ('Matrix price', 'Source rule from the matrix export applied to the line\'s unit cost or list (per Break Field), rounded to the cent.'),
        ('Match tolerance', 'Unit sale within the greater of $0.02 or 0.5% of the matrix price = Match. "Sold at list" = within $0.02 / 0.5% of list.'),
        ('Cost / list timing', 'Unit cost and list are as stored on the ticket line at time of sale.'),
    ] + [(k, v) for k, v in res.get('method_notes', [])]
    header(ws, 3, ['Item', 'Detail'])
    for i, (a, b) in enumerate(notes, 4):
        ws.cell(i, 1, a).font = BB; c = ws.cell(i, 2, b); c.font = BF; c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.cell(i, 1).alignment = Alignment(vertical='top'); ws.row_dimensions[i].height = 40
        ws.cell(i, 1).border = BOX; c.border = BOX
    ws.column_dimensions['A'].width = 24; ws.column_dimensions['B'].width = 110

    wb.active = 0
    wb.save(out)
    print('saved', out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--results', required=True); ap.add_argument('--matrix', required=True); ap.add_argument('--out', required=True)
    a = ap.parse_args()
    build(json.load(open(a.results)), json.load(open(a.matrix)), a.out)


if __name__ == '__main__':
    main()
