#!/usr/bin/env python3
"""Build the One-Line Customer Pay RO workbook from results.json (schema: references/results_template.json).

Split tickets: results['split_rows'] holds every original RO + its "C" split ticket (one row per ticket, same
columns as qlik_queries.py bundle `split_pairs`). The builder merges each pair into one visit and applies the
(visit - ticket) difference to every advisor / technician count, so the Qlik pulls themselves stay ticket-level.

Tabs: Summary, By Advisor, Monthly Trend, By Technician, What Was the One Line, Repair Visits, Warranty Visits,
Opportunity, Definitions. Run recalc.py on the output afterwards and never re-save it with openpyxl.

Usage:
  python3 build_oneline_report.py --results results.json --out "<Store> - One-Line Customer Pay ROs <Period>.xlsx"
"""
import argparse, collections, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule

NAVY, GOLD = '1F3864', 'FFC9A04B'
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor=NAVY)
T = Font(bold=True, size=14, color=NAVY); B = Font(bold=True); SEC = Font(bold=True, size=12, color=NAVY)
NOTE = Font(italic=True, color='595959')
RED = PatternFill('solid', fgColor='FFC7CE'); YEL = PatternFill('solid', fgColor='FFEB9C'); GRN = PatternFill('solid', fgColor='C6EFCE')
GREY = PatternFill('solid', fgColor='D9D9D9'); INPUT = PatternFill('solid', fgColor='DDEBF7')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
PCT, USD, N0 = '0.0%', '$#,##0', '#,##0'


def hdr(ws, r, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(r, i, h); c.font = H; c.fill = HF; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')


def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = T
    ws['A2'] = sub; ws['A2'].font = NOTE


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def note(ws, r, text, ncols, h=45):
    c = ws.cell(r, 1, text); c.alignment = Alignment(wrap_text=True, vertical='top'); c.font = NOTE
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols); ws.row_dimensions[r].height = h


def rate_cf(ws, rng, first, good, bad, low_is_good=True):
    if low_is_good:
        rules = [(f'{first}>={bad}', RED), (f'AND({first}>{good},{first}<{bad})', YEL), (f'{first}<={good}', GRN)]
    else:
        rules = [(f'{first}<{bad}', RED), (f'AND({first}>={bad},{first}<{good})', YEL), (f'{first}>={good}', GRN)]
    for f, fill in rules:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),{f})'], fill=fill))


def first_last(n):
    p = n.split()
    return p[0] + ' ' + p[-1] if len(p) > 1 else n


# ----------------------------------------------------------------------------------------------- split merge
def apply_split(res, ADV, TECH, REP_ADV, REP_TECH, W_ADV, W_TECH):
    rows = res.get('split_rows') or []
    start, end, mkeys = res['period']['start'], res['period']['end'], res['period']['month_keys']
    inp = lambda d: start <= d <= end
    techs = lambda s: [t for t in (s or '').split('/') if t]

    def unit(r):
        return dict(adv=r[1], cpn=r[2], oil=r[3], rep=r[4], mnt=r[5], ws=r[6], wsc=r[7], wsm=r[8], rec=r[9], sale=r[10], gross=r[11],
                    cpt=techs(r[12]), rept=techs(r[13]), wt=techs(r[14]), mon=mkeys.index(r[15][:7]) if r[15][:7] in mkeys else 0)

    tickets = [unit(r) for r in rows if inp(r[15])]
    groups = collections.OrderedDict()
    for r in rows:
        groups.setdefault(r[0].rstrip('C'), []).append(r)
    visits = []
    for base, rs in groups.items():
        if not any(inp(r[15]) for r in rs):
            continue
        b = [r for r in rs if r[0] == base]
        anchor = b[0] if b and inp(b[0][15]) else [r for r in rs if inp(r[15])][0]
        u = unit(anchor)
        for k in ('cpn', 'oil', 'rep', 'mnt', 'ws', 'wsc', 'wsm', 'rec', 'sale', 'gross'):
            u[k] = 0
        u['cpt'], u['rept'], u['wt'] = [], [], []
        for r in rs:
            x = unit(r)
            for k in ('cpn', 'oil', 'rep', 'mnt', 'ws', 'wsc', 'wsm', 'rec', 'sale', 'gross'):
                u[k] += x[k]
            for k in ('cpt', 'rept', 'wt'):
                u[k] += [t for t in x[k] if t not in u[k]]
        visits.append(u)

    def metrics(units):
        m = collections.defaultdict(float)
        for u in units:
            a = u['adv']
            if u['cpn'] > 0:
                one = u['cpn'] == 1
                m[(a, 'one' if one else 'multi')] += 1
                m[(a, 'sale1' if one else 'saleM')] += u['sale']
                m[(a, 'mon_one' if one else 'mon_multi', u['mon'])] += 1
                if u['oil'] > 0:
                    m[(a, 'oilonly' if one else 'oilmulti')] += 1
                for t in u['cpt']:
                    m[('T', t, 'one' if one else 'multi')] += 1
                    if u['oil'] > 0 and (one or len(u['cpt']) == 1):
                        m[('T', t, 'oo' if one else 'om')] += 1
            if u['rep'] > 0:
                m[(a, 'rv')] += 1; m[(a, 'rvsale')] += u['sale']
                if u['mnt'] > 0: m[(a, 'rvm')] += 1
                if u['cpn'] == 1: m[(a, 'rv1')] += 1
                for t in u['rept']:
                    m[('T', t, 'rv')] += 1
                    if u['mnt'] > 0: m[('T', t, 'rvm')] += 1
                    if u['cpn'] == u['rep']: m[('T', t, 'rvonly')] += 1
            if u['ws'] + u['wsc'] + u['wsm'] > 0:
                cp = 1 if u['cpn'] > 0 else 0
                m[(a, 'w')] += 1; m[(a, 'wcp')] += cp
                for k in ('ws', 'wsc', 'wsm'):
                    if u[k] > 0:
                        m[(a, k)] += 1; m[(a, k + 'cp')] += cp
                if u['rec'] > 0:
                    m[(a, 'rec')] += 1; m[(a, 'reccp')] += cp
                m[(a, 'wsale')] += u['sale']; m[(a, 'wgross')] += u['gross']
                for t in u['wt']:
                    if u['ws'] + u['wsc'] > 0:
                        m[('T', t, 'wr')] += 1; m[('T', t, 'wrcp')] += cp
                    if u['wsm'] > 0:
                        m[('T', t, 'wm')] += 1; m[('T', t, 'wmcp')] += cp
        return m

    old, new = metrics(tickets), metrics(visits)
    d = lambda *k: new.get(k, 0) - old.get(k, 0)
    ADV = [(n, int(o + d(n, 'one')), int(mu + d(n, 'multi')), s1 + d(n, 'sale1'), sm + d(n, 'saleM'), cpl,
            int(oo + d(n, 'oilonly')), int(om + d(n, 'oilmulti')),
            tuple(int(x + d(n, 'mon_one', i)) for i, x in enumerate(mo)), tuple(int(x + d(n, 'mon_multi', i)) for i, x in enumerate(mm)), k60, o30)
           for (n, o, mu, s1, sm, cpl, oo, om, mo, mm, k60, o30) in ADV]
    TECH = {t: (v[0], v[1], int(v[2] + d('T', t, 'one')), int(v[3] + d('T', t, 'multi')), int(v[4] + d('T', t, 'oo')), int(v[5] + d('T', t, 'om')))
            for t, v in TECH.items()}
    REP_ADV = [(n, int(v + d(n, 'rv')), int(mm + d(n, 'rvm')), int(o + d(n, 'rv1')), s + d(n, 'rvsale'), mi) for (n, v, mm, o, s, mi) in REP_ADV]
    REP_TECH = [(t, int(v + d('T', t, 'rv')), int(mm + d('T', t, 'rvm')), int(o + d('T', t, 'rvonly'))) for (t, v, mm, o) in REP_TECH]
    W_ADV = [(n, (ws[0] + int(d(n, 'ws')), ws[1] + int(d(n, 'wscp'))), (wc[0] + int(d(n, 'wsc')), wc[1] + int(d(n, 'wsccp'))),
              (wm[0] + int(d(n, 'wsm')), wm[1] + int(d(n, 'wsmcp'))), al + int(d(n, 'w')), alc + int(d(n, 'wcp')),
              s + d(n, 'wsale'), g + d(n, 'wgross'), rc + int(d(n, 'rec')), rcc + int(d(n, 'reccp')), mi)
             for (n, ws, wc, wm, al, alc, s, g, rc, rcc, mi) in W_ADV]
    W_TECH = [(t, a + int(d('T', t, 'wr')), b + int(d('T', t, 'wrcp')), c + int(d('T', t, 'wm')), e + int(d('T', t, 'wmcp'))) for (t, a, b, c, e) in W_TECH]
    n_c = sum(1 for r in rows if r[0].endswith('C') and inp(r[15]))
    return ADV, TECH, REP_ADV, REP_TECH, W_ADV, W_TECH, n_c


# ----------------------------------------------------------------------------------------------- build
def build(res, out):
    store, per = res['store'], res['period']
    months = per['month_labels']; nm = len(months)
    lab = res['labels']
    ADV = [(a['name'], a['one'], a['multi'], a['sale_one'], a['sale_multi'], a['cp_lines'], a['oil_only'], a['oil_multi'],
            tuple(a['month_one']), tuple(a['month_multi']), a.get('one_60k', 0), a.get('oil_only_30k', 0)) for a in res['advisors']]
    TECH = {t: (v['name'], v['group'], v['one'], v['multi'], v['oil_only'], v['oil_multi']) for t, v in res['techs'].items()}
    REP_ADV = [(r['name'], r['visits'], r['with_maint'], r['one_line'], r['cp_sale'], r.get('avg_miles', 0)) for r in res['repair_adv']]
    REP_TECH = [(r['tech'], r['visits'], r['with_maint'], r['repair_only']) for r in res['repair_tech']]
    W_ADV = [(w['name'], tuple(w['ws']), tuple(w['wsc']), tuple(w['wsm']), w['visits'], w['with_cp'], w['cp_sale'], w['cp_gross'],
              w['recall'][0], w['recall'][1], w.get('avg_miles', 0)) for w in res['warranty_adv']]
    W_TECH = [(w['tech'], w['repair'][0], w['repair'][1], w['maint'][0], w['maint'][1]) for w in res['warranty_tech']]
    ADV, TECH, REP_ADV, REP_TECH, W_ADV, W_TECH, N_C = apply_split(res, ADV, TECH, REP_ADV, REP_TECH, W_ADV, W_TECH)
    OPS = res['single_line_ops']
    oa = res['oil_addon']; ADDON_MEDIAN, ADDON_AVG, ADDON_GM = oa['median_sale'], oa['avg_sale'], oa['gross_margin']
    W_MEDIAN = res['warranty_addon_median']
    ONE_TOT = sum(a[1] for a in ADV); MULTI_TOT = sum(a[2] for a in ADV)
    OIL_ONLY = sum(a[6] for a in ADV); OIL_MULTI = sum(a[7] for a in ADV)
    MIN_N = res.get('min_ros_for_rating', 10)
    MAIN = [a for a in ADV if a[1] + a[2] > MIN_N]
    nadv = len(ADV)

    wb = Workbook()
    # ---------------- Summary ----------------
    ws = wb.active; ws.title = 'Summary'; ws.sheet_properties.tabColor = GOLD
    title(ws, f'{store} - One-Line Customer Pay Repair Orders',
          f'Closed ROs, {per["label"]} | One-line CP visit = exactly one customer-pay line with labor charged | Source: Qlik Closed Repair Orders')
    kp = [('Customer-pay visits', ONE_TOT + MULTI_TOT, N0), ('One-line CP visits', ONE_TOT, N0), ('One-line rate', ONE_TOT / (ONE_TOT + MULTI_TOT), PCT),
          ('Avg CP sale, one-line visit', sum(a[3] for a in ADV) / ONE_TOT, USD), ('Avg CP sale, multi-line visit', sum(a[4] for a in ADV) / MULTI_TOT, USD),
          ('Oil-change visits that were oil change only', OIL_ONLY / (OIL_ONLY + OIL_MULTI), PCT)]
    for i, (k, v, f) in enumerate(kp):
        c = ws.cell(4, 1 + i * 2, k); c.font = NOTE; c.alignment = Alignment(wrap_text=True)
        c = ws.cell(5, 1 + i * 2, round(v, 4)); c.font = Font(bold=True, size=16, color=NAVY); c.number_format = f
    ws.row_dimensions[4].height = 30
    ws.cell(7, 1, 'What the data says').font = SEC
    r1 = {a[0]: a[1] / (a[1] + a[2]) for a in MAIN}
    oo = {a[0]: a[6] / (a[6] + a[7]) for a in MAIN if a[6] + a[7]}
    mo = [sum(a[8][m] for a in ADV) / max(1, sum(a[8][m] + a[9][m] for a in ADV)) for m in range(nm)]
    BEST_OO = min(oo, key=oo.get)
    tr = {t: v[2] / (v[2] + v[3]) for t, v in TECH.items() if v[2] + v[3] > 20}
    grp_rates = collections.defaultdict(list)
    for t, r in tr.items():
        grp_rates[TECH[t][1]].append(r)
    rep = {r[0]: r[2] / r[1] for r in REP_ADV if r[1]}
    rept = {t: m / v for t, v, m, o in REP_TECH if v >= 20}
    low_rept = sorted(rept, key=rept.get)[:2]
    wsr = {w[0]: w[1][1] / w[1][0] for w in W_ADV if w[1][0]}; wsm = {w[0]: w[3][1] / w[3][0] for w in W_ADV if w[3][0]}
    BEST_WS, BEST_WSM = max(wsr, key=wsr.get), (max(wsm, key=wsm.get) if wsm else None)
    wtot, wcp = sum(w[4] for w in W_ADV), sum(w[5] for w in W_ADV)
    wst, wsc = sum(w[1][0] for w in W_ADV), sum(w[1][1] for w in W_ADV)
    wmt, wmc = sum(w[3][0] for w in W_ADV), sum(w[3][1] for w in W_ADV)
    conv = round((OIL_ONLY + OIL_MULTI) * (OIL_ONLY / (OIL_ONLY + OIL_MULTI) - oo[BEST_OO]))
    trend = ', '.join(f'{x:.0%} in {mlab}' for x, mlab in zip(mo, months))
    FIND = [
        f'{ONE_TOT / (ONE_TOT + MULTI_TOT):.0%} of customer-pay visits ({ONE_TOT:,} of {ONE_TOT + MULTI_TOT:,}) left with exactly one customer-pay line. By month: {trend}.',
        f'{OIL_ONLY / ONE_TOT:.0%} of one-line visits ({OIL_ONLY} of {ONE_TOT:,}) were an oil change and nothing else; about {sum(a[11] for a in ADV)} of those vehicles had 30,000+ miles, where maintenance beyond an oil change is typically coming due.',
        f'By advisor, the one-line rate runs from {min(r1.values()):.0%} ({first_last(min(r1, key=r1.get))}) to {max(r1.values()):.0%} ({first_last(max(r1, key=r1.get))}). On oil-change visits, {first_last(BEST_OO)} is lowest at {oo[BEST_OO]:.0%} oil only; the others run {min(v for k, v in oo.items() if k != BEST_OO):.0%}-{max(oo.values()):.0%}.',
        'By technician (one-line rate of the CP visits each tech worked): ' + '; '.join(f'{g} {min(v):.0%}-{max(v):.0%}' for g, v in grp_rates.items()) +
        '. Technician rates reflect the work they are given as much as the inspections they write up; compare within a group.',
        f'Repair visits ({lab["repair_desc"]}): {sum(r[1] for r in REP_ADV)} visits; {sum(r[2] for r in REP_ADV) / sum(r[1] for r in REP_ADV):.0%} also sold any maintenance. Advisors range from {min(rep.values()):.0%} ({first_last(min(rep, key=rep.get))}) to {max(rep.values()):.0%}'
        + (f'; lowest technicians: {first_last(TECH[low_rept[0]][0])} ({rept[low_rept[0]]:.0%})' + (f' and {first_last(TECH[low_rept[1]][0])} ({rept[low_rept[1]]:.0%})' if len(low_rept) > 1 else '') if low_rept else '') + '. See Repair Visits.',
        f'Warranty visits ({lab["ws"]}, {lab["wsc"]}, {lab["wsm"]}; PDI excluded): {wtot} visits, {wcp / wtot:.0%} also had customer-pay work. {lab["ws"]} visits converted {wsc / wst:.0%} ({first_last(BEST_WS)} {wsr[BEST_WS]:.0%}, {first_last(min(wsr, key=wsr.get))} {min(wsr.values()):.0%})'
        + (f'; {lab["wsm"]} visits {wmc / wmt:.0%} ({first_last(BEST_WSM)} {wsm[BEST_WSM]:.0%}, {first_last(min(wsm, key=wsm.get))} {min(wsm.values()):.0%})' if wmt else '') + '. See Warranty Visits.',
        f'Opportunity: if every advisor matched {first_last(BEST_OO)}\'s oil-change add-on rate, about {conv} more oil-change visits per period would add a service; at the median add-on (${ADDON_MEDIAN:,.0f} sale, {ADDON_GM:.0%} gross) that is roughly ${conv * ADDON_MEDIAN * ADDON_GM * 12 / nm / 1000:,.0f}K of gross a year. See Opportunity.',
    ]
    if N_C:
        FIND.append(f'Split tickets: {N_C} visits were split onto a second RO ending in "C". All counts treat the original and its "C" ticket as one visit.')
    for i, t in enumerate(res.get('extra_findings', []), 0):
        FIND.append(t)
    for i, t in enumerate(FIND):
        r = 8 + i
        c = ws.cell(r, 1, f'{i + 1}. {t}'); c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = 46
    r0 = 9 + len(FIND)
    ws.cell(r0, 1, 'One-line rate by advisor').font = SEC
    hdr(ws, r0 + 1, ['Advisor', 'One-line rate', 'Oil change only rate'])
    for i, a in enumerate(MAIN):
        rr = r0 + 2 + i
        ws.cell(rr, 1, a[0]).border = BOX
        c = ws.cell(rr, 2, round(a[1] / (a[1] + a[2]), 4)); c.number_format = PCT; c.border = BOX
        c = ws.cell(rr, 3, round(a[6] / (a[6] + a[7]), 4) if a[6] + a[7] else None); c.number_format = PCT; c.border = BOX
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'One-line rate and oil-change-only rate by advisor'
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=r0 + 1, max_row=r0 + 1 + len(MAIN)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=r0 + 2, max_row=r0 + 1 + len(MAIN)))
    ch.y_axis.numFmt = '0%'; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.x_axis.scaling.orientation = 'maxMin'
    ch.height = 7.5; ch.width = 16
    ws.add_chart(ch, f'E{r0}')
    widths(ws, [22, 14, 14, 3, 16, 3, 16, 3, 16, 3, 16, 3])

    # ---------------- By Advisor ----------------
    ws = wb.create_sheet('By Advisor')
    title(ws, 'One-line customer-pay visits by service advisor', f'Advisor on the RO header, {per["label"]}')
    hdr(ws, 4, ['Advisor', 'CP visits', 'One-line visits', 'One-line rate', 'Multi-line visits', 'Avg CP sale one-line', 'Avg CP sale multi-line',
                'CP lines per visit', 'Oil-change visits', 'Oil change only', 'Oil change only rate', 'One-line 60K+ miles', 'Oil-only 30K+ miles'])
    for i, a in enumerate(ADV):
        rr = 5 + i
        vals = [a[0], f'=C{rr}+E{rr}', a[1], f'=C{rr}/B{rr}', a[2], f'=IFERROR({round(a[3], 2)}/C{rr},0)', f'=IFERROR({round(a[4], 2)}/E{rr},0)',
                f'={a[5]}/B{rr}', a[6] + a[7], a[6], f'=IFERROR(J{rr}/I{rr},"")', a[10], a[11]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            c.number_format = {4: PCT, 6: USD, 7: USD, 8: '0.00', 11: PCT}.get(j, N0)
    rt = 5 + nadv
    ws.cell(rt, 1, 'Store total').font = B
    for j in (2, 3, 5, 9, 10, 12, 13):
        col = get_column_letter(j); c = ws.cell(rt, j, f'=SUM({col}5:{col}{rt - 1})'); c.font = B; c.number_format = N0
    for j, f, fm in [(4, f'=C{rt}/B{rt}', PCT), (11, f'=J{rt}/I{rt}', PCT), (6, f'={round(sum(a[3] for a in ADV), 2)}/C{rt}', USD),
                     (7, f'={round(sum(a[4] for a in ADV), 2)}/E{rt}', USD), (8, f'={sum(a[5] for a in ADV)}/B{rt}', '0.00')]:
        c = ws.cell(rt, j, f); c.font = B; c.number_format = fm
    for j in range(1, 14): ws.cell(rt, j).fill = GREY
    th = res.get('thresholds', {})
    t1, t2 = th.get('one_line', [0.55, 0.62]), th.get('oil_only', [0.50, 0.60])
    rate_cf(ws, f'D5:D{rt - 1}', 'D5', *t1); rate_cf(ws, f'K5:K{rt - 1}', 'K5', *t2)
    note(ws, rt + 2, f'Shading: one-line rate green <={t1[0]:.0%}, red >={t1[1]:.0%}; oil-change-only rate green <={t2[0]:.0%}, red >={t2[1]:.0%}. Discussion thresholds, not a Findlay standard. '
                     f'Oil-change visits = visits with a customer-pay {lab["oil_ops"]} line.', 13, 45)
    widths(ws, [20, 9, 10, 10, 10, 12, 12, 10, 11, 11, 11, 13, 13])

    # ---------------- Monthly ----------------
    ws = wb.create_sheet('Monthly Trend')
    title(ws, 'One-line rate by month', 'Is it getting better or worse?')
    cols = ['Advisor']
    for mlab in months:
        cols += [f'{mlab} one-line', f'{mlab} CP visits', f'{mlab} rate']
    hdr(ws, 4, cols)
    for i, a in enumerate(ADV):
        rr = 5 + i
        ws.cell(rr, 1, a[0]).border = BOX
        for m in range(nm):
            c0 = 2 + m * 3
            ws.cell(rr, c0, a[8][m]); ws.cell(rr, c0 + 1, a[8][m] + a[9][m])
            ws.cell(rr, c0 + 2, f'=IFERROR({get_column_letter(c0)}{rr}/{get_column_letter(c0 + 1)}{rr},"")').number_format = PCT
            for j in range(c0, c0 + 3): ws.cell(rr, j).border = BOX
    rt = 5 + nadv
    ws.cell(rt, 1, 'Store').font = B
    for m in range(nm):
        c0 = 2 + m * 3
        for j in (c0, c0 + 1):
            col = get_column_letter(j); c = ws.cell(rt, j, f'=SUM({col}5:{col}{rt - 1})'); c.font = B
        c = ws.cell(rt, c0 + 2, f'={get_column_letter(c0)}{rt}/{get_column_letter(c0 + 1)}{rt}'); c.font = B; c.number_format = PCT
        rate_cf(ws, f'{get_column_letter(c0 + 2)}5:{get_column_letter(c0 + 2)}{rt - 1}', f'{get_column_letter(c0 + 2)}5', *t1)
    for j in range(1, 2 + 3 * nm): ws.cell(rt, j).fill = GREY
    ws.cell(rt + 2, 1, 'Month'); ws.cell(rt + 2, 2, 'Store one-line rate')
    for m, mlab in enumerate(months):
        ws.cell(rt + 3 + m, 1, mlab); c = ws.cell(rt + 3 + m, 2, f'={get_column_letter(4 + m * 3)}{rt}'); c.number_format = PCT
    lc = LineChart(); lc.title = 'Store one-line rate by month'; lc.style = 12
    lc.add_data(Reference(ws, min_col=2, min_row=rt + 2, max_row=rt + 2 + nm), titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=1, min_row=rt + 3, max_row=rt + 2 + nm))
    lc.y_axis.numFmt = '0%'; lc.legend = None; lc.height = 7; lc.width = 14
    ws.add_chart(lc, f'D{rt + 2}')
    widths(ws, [20] + [11, 11, 10] * nm)

    # ---------------- By Technician ----------------
    ws = wb.create_sheet('By Technician')
    title(ws, 'One-line customer-pay visits by technician', 'Visits where the technician worked a customer-pay line. A multi-line visit counts for every tech who worked on it.')
    hdr(ws, 4, ['Technician', 'Tech #', 'Group', 'CP visits worked', 'One-line visits', 'One-line rate', 'Oil-change visits', 'Oil change only', 'Oil change only rate'])
    for i, (tn, (nmn, grp, one, multi, o1, om)) in enumerate(TECH.items()):
        rr = 5 + i
        vals = [nmn, tn, grp, one + multi, one, f'=IFERROR(E{rr}/D{rr},"")', o1 + om, o1, f'=IFERROR(H{rr}/G{rr},"")']
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            if j in (6, 9): c.number_format = PCT
    note(ws, 6 + len(TECH), 'Technicians do not control what is sold, but their multi-point inspection findings feed the advisor\'s recommendations. Express/lube techs are mostly given oil changes, so their one-line rates are naturally higher. '
                            'Compare within a group. Catch-all tech codes (e.g. 999, MULT) are omitted.', 9, 45)
    widths(ws, [22, 9, 16, 12, 11, 11, 13, 12, 13])

    # ---------------- What was the one line ----------------
    ws = wb.create_sheet('What Was the One Line')
    title(ws, 'On one-line visits, what was the single customer-pay line?', f'Top op codes, {per["label"]} (per ticket; split tickets not merged here)')
    hdr(ws, 4, ['Op code', 'What it is', 'One-line ROs', '% of one-line visits', 'CP sale', 'Avg CP sale', 'Avg mileage'])
    for i, o in enumerate(OPS):
        rr = 5 + i
        vals = [o['op'], o['desc'], o['ros'], f'=C{rr}/{ONE_TOT}', o['cp_sale'], f'=E{rr}/C{rr}', o.get('avg_miles')]
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            c.number_format = {3: N0, 4: PCT, 5: USD, 6: USD, 7: N0}.get(j, 'General')
    note(ws, 6 + len(OPS), 'Repair/diagnostic one-liners are expected to some degree (the customer came in for one problem) but are still a chance to sell due maintenance. '
                           'A one-line 30K/60K-type service package is less of a concern because it already bundles several items.', 7, 45)
    widths(ws, [10, 40, 12, 12, 12, 11, 11])

    # ---------------- Repair Visits ----------------
    ws = wb.create_sheet('Repair Visits')
    title(ws, 'When a customer comes in for a repair, is maintenance recommended and sold?',
          f'Repair visit = CP visit with a repair/diagnostic line ({lab["repair_ops"]}). Maintenance = any CP maintenance line on the same visit.')
    ws.cell(4, 1, 'By advisor').font = SEC
    hdr(ws, 5, ['Advisor', 'Repair visits', 'Also sold maintenance', 'Maintenance attach rate', 'No maintenance sold', 'One-line repair visits', 'One-line rate', 'Avg CP sale per repair visit', 'Avg mileage'])
    for i, (a, n, m, o, sale, mi) in enumerate(REP_ADV):
        rr = 6 + i
        vals = [a, n, m, f'=IFERROR(C{rr}/B{rr},"")', f'=B{rr}-C{rr}', o, f'=IFERROR(F{rr}/B{rr},"")', f'=IFERROR({round(sale, 2)}/B{rr},"")', mi]
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            c.number_format = {4: PCT, 7: PCT, 8: USD, 9: N0}.get(j, N0 if j > 1 else 'General')
    rt = 6 + len(REP_ADV)
    ws.cell(rt, 1, 'Store').font = B
    for j in (2, 3, 5, 6):
        col = get_column_letter(j); c = ws.cell(rt, j, f'=SUM({col}6:{col}{rt - 1})'); c.font = B
    for j, f, fm in [(4, f'=C{rt}/B{rt}', PCT), (7, f'=F{rt}/B{rt}', PCT), (8, f'={round(sum(r[4] for r in REP_ADV), 2)}/B{rt}', USD)]:
        c = ws.cell(rt, j, f); c.font = B; c.number_format = fm
    for j in range(1, 10): ws.cell(rt, j).fill = GREY
    t3 = th.get('attach', [0.50, 0.45])
    rate_cf(ws, f'D6:D{rt - 1}', 'D6', t3[0], t3[1], low_is_good=False)
    r = rt + 2
    ws.cell(r, 1, 'By technician (the tech who did the repair)').font = SEC
    hdr(ws, r + 1, ['Technician', 'Tech #', 'Group', 'Repair visits', 'Also sold maintenance', 'Maintenance attach rate', 'Repair only (no other CP work)'])
    for i, (tn, n, m, o) in enumerate(REP_TECH):
        rr = r + 2 + i
        vals = [TECH.get(tn, (tn, ''))[0], tn, TECH.get(tn, ('', ''))[1], n, m, f'=IFERROR(E{rr}/D{rr},"")', o]
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            if j == 6: c.number_format = PCT
    f0, f1 = r + 2, r + 1 + len(REP_TECH)
    rate_cf(ws, f'F{f0}:F{f1}', f'F{f0}', t3[0], t3[1], low_is_good=False)
    note(ws, f1 + 2, f'Maintenance attach rate = share of repair visits where at least one customer-pay maintenance line was also sold (any technician). Shading: green >= {t3[0]:.0%}, red < {t3[1]:.0%} (discussion thresholds). '
                     'Technician rates reflect MPI findings and recommendations; the advisor decides what is presented. Small samples (under 20 visits) are not shown in the findings.', 9, 45)
    widths(ws, [20, 10, 16, 12, 12, 12, 14, 14, 11])

    # ---------------- Warranty Visits ----------------
    ws = wb.create_sheet('Warranty Visits')
    title(ws, 'When a customer comes in for free warranty work, do we sell anything?',
          f'Visits with a warranty line ({lab["ws"]}: {lab["ws_types"]}; {lab["wsc"]}: {lab["wsc_types"]}; {lab["wsm"]}: {lab["wsm_types"]}; PDI excluded) that also had a customer-pay line with labor')
    ws.cell(4, 1, 'By advisor').font = SEC
    hdr(ws, 5, ['Advisor', 'Warranty visits', 'With CP work', 'CP rate', f'{lab["ws"]} visits', 'With CP', 'Rate', f'{lab["wsc"]} visits', 'With CP', 'Rate',
                f'{lab["wsm"]} visits', 'With CP', 'Rate', 'Recall visits', 'With CP', 'CP sale on these visits', 'CP gross'])
    for i, w in enumerate(W_ADV):
        rr = 6 + i
        vals = [w[0], w[4], w[5], f'=IFERROR(C{rr}/B{rr},"")', w[1][0], w[1][1], f'=IFERROR(F{rr}/E{rr},"")', w[2][0], w[2][1], f'=IFERROR(I{rr}/H{rr},"")',
                w[3][0], w[3][1], f'=IFERROR(L{rr}/K{rr},"")', w[8], w[9], round(w[6], 2), round(w[7], 2)]
        for j, v in enumerate(vals, 1):
            c = ws.cell(rr, j, v); c.border = BOX
            c.number_format = PCT if j in (4, 7, 10, 13) else (USD if j >= 16 else N0)
    rt = 6 + len(W_ADV)
    ws.cell(rt, 1, 'Store').font = B
    for j in (2, 3, 5, 6, 8, 9, 11, 12, 14, 15, 16, 17):
        col = get_column_letter(j); c = ws.cell(rt, j, f'=SUM({col}6:{col}{rt - 1})'); c.font = B; c.number_format = USD if j >= 16 else N0
    for j, (a, b) in {4: ('C', 'B'), 7: ('F', 'E'), 10: ('I', 'H'), 13: ('L', 'K')}.items():
        c = ws.cell(rt, j, f'=IFERROR({a}{rt}/{b}{rt},"")'); c.font = B; c.number_format = PCT
    for j in range(1, 18): ws.cell(rt, j).fill = GREY
    r = rt + 2
    ws.cell(r, 1, 'By technician (the tech on the warranty line)').font = SEC
    hdr(ws, r + 1, ['Technician', 'Tech #', 'Group', f'{lab["ws"]} / {lab["wsc"]} visits', 'With CP', 'Rate', f'{lab["wsm"]} visits', 'With CP', 'Rate'])
    for i, (tn, a, b, c_, d_) in enumerate(W_TECH):
        rr = r + 2 + i
        vals = [TECH.get(tn, (tn, ''))[0], tn, TECH.get(tn, ('', ''))[1], a, b, f'=IFERROR(E{rr}/D{rr},"")', c_, d_, f'=IFERROR(H{rr}/G{rr},"")']
        for j, v in enumerate(vals, 1):
            cc = ws.cell(rr, j, v); cc.border = BOX
            if j in (6, 9): cc.number_format = PCT
    r = r + 3 + len(W_TECH)
    ws.cell(r, 1, 'Opportunity (edit blue cells)').font = SEC
    rows = [(f'{lab["ws"]} visits per period', wst, N0, False), ('Current CP rate', round(wsc / wst, 4) if wst else 0, PCT, False),
            (f'Target CP rate (default = best advisor, {first_last(BEST_WS)})', round(wsr[BEST_WS], 4), PCT, True),
            (f'{lab["wsm"]} visits per period', wmt, N0, False), ('Current CP rate', round(wmc / wmt, 4) if wmt else 0, PCT, False),
            (f'Target CP rate (default = best advisor{", " + first_last(BEST_WSM) if BEST_WSM else ""})', round(wsm[BEST_WSM], 4) if BEST_WSM else 0, PCT, True),
            ('Additional visits with CP work per period', None, N0, False), ('CP sale per converted visit (default = median)', W_MEDIAN, USD, True),
            ('Gross margin on that work', round(sum(w[7] for w in W_ADV) / max(1, sum(w[6] for w in W_ADV)), 4), PCT, True), ('Added gross per year', None, USD, False)]
    for i, (k, v, f, inp) in enumerate(rows):
        rr = r + 1 + i
        ws.cell(rr, 1, k).border = BOX
        c = ws.cell(rr, 2, v); c.number_format = f; c.border = BOX
        if inp: c.fill = INPUT
    a0 = r + 1
    ws.cell(a0 + 6, 2, f'=MAX(0,ROUND(B{a0}*(B{a0 + 2}-B{a0 + 1}),0))+MAX(0,ROUND(B{a0 + 3}*(B{a0 + 5}-B{a0 + 4}),0))')
    ws.cell(a0 + 9, 2, f'=B{a0 + 6}*B{a0 + 7}*B{a0 + 8}*{12 / nm:g}'); ws.cell(a0 + 9, 1).font = B; ws.cell(a0 + 9, 2).font = B
    wsct, wscc = sum(w[2][0] for w in W_ADV), sum(w[2][1] for w in W_ADV)
    note(ws, a0 + 11, 'Qlik does not carry CDK line letters (A-Z) and each line\'s booked time is the RO open time, so the report shows whether any customer-pay work was sold on the visit, not when it was added. '
                      f'{lab["wsc"]} visits ({wscc / wsct:.0%} with CP) are left out of the opportunity.' if wsct else '', 9, 60)
    widths(ws, [26, 11, 10, 9, 13, 9, 9, 12, 9, 9, 13, 9, 9, 9, 9, 13, 11])

    # ---------------- Opportunity ----------------
    ws = wb.create_sheet('Opportunity')
    title(ws, 'What closing the gap could be worth - oil-change visits only', f'Period = {per["label"]}; annual = period x {12 / nm:g}. Edit the blue cells.')
    rows = [('Oil-change visits', OIL_ONLY + OIL_MULTI, N0, False), ('Oil change only today', OIL_ONLY, N0, False), ('Oil-change-only rate today', '=B5/B4', PCT, False),
            (f'Target oil-change-only rate (default = best advisor, {first_last(BEST_OO)})', round(oo[BEST_OO], 4), PCT, True),
            ('Additional visits with an add-on per period', '=MAX(0,ROUND(B4*(B6-B7),0))', N0, False),
            ('Add-on sale per converted visit (default = median)', ADDON_MEDIAN, USD, True), ('Gross margin on add-on work', round(ADDON_GM, 4), PCT, True),
            ('Added gross per period', '=B8*B9*B10', USD, False), ('Added gross per year', f'=B11*{12 / nm:g}', USD, False)]
    for i, (k, v, f, inp) in enumerate(rows):
        rr = 4 + i
        ws.cell(rr, 1, k).border = BOX
        c = ws.cell(rr, 2, v); c.number_format = f; c.border = BOX
        if inp: c.fill = INPUT
        if rr >= 11: ws.cell(rr, 1).font = B; c.font = B
    note(ws, 14, f'Add-on sale = customer-pay sale on the non-oil lines of oil-change visits that included more work (median ${ADDON_MEDIAN:,.2f}, average ${ADDON_AVG:,.2f}; gross margin {ADDON_GM:.1%}). '
                 'Using the average instead of the median roughly doubles the result. Does not include repair one-liners or any effect on retention.', 2, 60)
    widths(ws, [58, 16])

    # ---------------- Definitions ----------------
    ws = wb.create_sheet('Definitions')
    title(ws, 'Definitions and limitations', '')
    M = [('Period', f'ROs closed {per["label"]}.'),
         ('Customer-pay line', f'A labor line with pay type C (labor types {lab["cp_types"]}) and labor sale > $0. Parts-only lines do not count.'),
         ('One-line CP visit', 'A visit with exactly one customer-pay line. Warranty and internal lines on the same visit do not change this.'),
         ('Advisor / technician', 'Advisor from the RO header. Technician from the customer-pay line; on a multi-line visit each technician who worked a CP line is credited.'),
         ('Repair vs. maintenance', f'Repair/diagnostic op codes: {lab["repair_ops"]}. Maintenance op codes: {lab["maint_ops"]}. Oil-change op codes: {lab["oil_ops"]}.'),
         ('Warranty visits', f'{lab["ws"]} = {lab["ws_types"]}; {lab["wsc"]} = {lab["wsc_types"]}; {lab["wsm"]} = {lab["wsm_types"]}. PDI excluded. "With CP work" = the visit has a customer-pay line with labor sale > $0.'),
         ('Split tickets', f'An original RO and its companion "C" ticket (e.g. 290486 / 290486C) are one visit: lines are combined and the visit is credited to the original ticket\'s advisor and close month ({N_C} C tickets in the period). '
                           'The mileage counts and the What Was the One Line tab are per ticket.'),
         ('Not included', 'Declined recommendations and MPI results are not in the RO data; a one-line visit may have had recommendations that were declined.')]
    M += [tuple(x) for x in res.get('definitions_extra', [])]
    for i, (k, v) in enumerate(M):
        ws.cell(4 + i, 1, k).font = B; ws.cell(4 + i, 1).alignment = Alignment(vertical='top')
        c = ws.cell(4 + i, 2, v); c.alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[4 + i].height = 45
    widths(ws, [20, 120])
    wb.save(out)
    print(out, 'visits', ONE_TOT + MULTI_TOT, 'one-line', ONE_TOT, f'{ONE_TOT / (ONE_TOT + MULTI_TOT):.1%}', 'C tickets merged', N_C)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--results', required=True)
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    build(json.load(open(a.results)), a.out)


if __name__ == '__main__':
    main()
