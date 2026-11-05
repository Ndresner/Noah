#!/usr/bin/env python3
"""Generate qlik_create_data_object arguments (dimensions/measures/sort/limit JSON) for the parts-matrix audit.

Parts app  = Closed Parts Tickets  d3d25914-349d-4a55-857c-645fccefbecc
RO app     = Closed Repair Orders  c3efc739-b063-47dd-a4d3-d1c4cac07ad0

Bundles
  RO app   : profile   - labor-type mix of part-bearing lines per RO (one row per labor-type combo)
             rolist    - comma list of ROs whose part-bearing lines are ALL in --matrix-types (+ count)
  Parts app: summary   - by source: Lines, ROs, Actual, Matrix, List, Cost, MatchLn
             classify  - by source: under/over/at-list/below-cost/$0-list counts and $
             totals    - no dimension: same as summary + distinct RO count for the whole group
             under     - top-N line exceptions, most under-matrix first
             over      - top-N line exceptions, most over-matrix first
             all_exc   - every exception line (page it with qlik_get_chart_data)

The parts-app bundles restrict ROs through a Qlik alternate state that holds the RO list:
  --state cpt --mode include   -> Parts_Detail.referno = cpt::Parts_Detail.referno

Usage:
  python3 qlik_expr.py --matrix matrix.json --division "TOYOTA PRESCOTT" --year 2026 --months Aug,Sep,Oct \
      --bundle summary --state cpt --out args_summary.json
  python3 qlik_expr.py ... --bundle rolist --matrix-types CPT --out args_rolist.json
"""
import argparse, json, sys

C, L = 'Parts_Detail.unitcost', 'Parts_Detail.listprice'
S = 'Num#(Trim(Parts_Detail.source))'
A = 'Num#(Trim(Parts_Detail.unitsaleprice))'
Q = 'Parts_Detail.quantitysold'
R = 'Parts_Detail.referno'


def num(x):
    return f'{x:.6f}'.rstrip('0').rstrip('.')


def band_formula(b):
    base = C if b['base'] == 'COST' else L
    return base if not b['pct'] else f'{base}*{num(1 + b["pct"] / 100.0)}'


def rule_expr(rule):
    x = C if rule['break_field'] == 'COST' else L
    bands = rule['bands']
    e = band_formula(bands[-1])
    for b in reversed(bands[:-1]):
        e = f'If({x}<={num(b["hi"])},{band_formula(b)},{e})'
    return e


def matrix_price(matrix):
    groups = {}
    for s, r in matrix['sources'].items():
        if r.get('unsupported') or not r['bands'] or not s.isdigit():
            continue
        if len(r['bands']) == 1:   # flat rule (e.g. "LIST" for every range) - break field/range irrelevant
            b = r['bands'][0]
            key = json.dumps(['FLAT', b['base'], b['pct']])
        else:
            key = json.dumps([r['break_field'], [(b['hi'], b['base'], b['pct']) for b in r['bands']]])
        groups.setdefault(key, {'rule': r, 'srcs': []})['srcs'].append(s)
    e = 'Null()'
    for g in reversed(list(groups.values())):
        cond = f'{S}={g["srcs"][0]}' if len(g['srcs']) == 1 else f'Match({S},{",".join(g["srcs"])})'
        e = f'If({cond},{rule_expr(g["rule"])},{e})'
    return f'Round({e},.01)'


def quoted(vals):
    return ','.join("'" + v.strip() + "'" for v in vals)


def base_set(a, ro_app=False):
    parts = [f"Division={{'{a.division}'}}", f"Year={{'{a.year}'}}", f"Month={{{quoted(a.months.split(','))}}}"]
    if not ro_app:
        parts.append("Parts_Detail.invsaletype={'RO'}")
    return ','.join(parts)


def parts_set(a):
    b = base_set(a)
    if a.state:
        op = '=' if a.mode == 'include' else '-='
        b += f',{R}{op}{a.state}::{R}'
    return '{<' + b + '>}'


def m(label, expr):
    return {'label': label, 'expression': expr}


def parts_bundle(a, matrix):
    st, EP = parts_set(a), matrix_price(matrix)
    tol = f'RangeMax(.02,{EP}*.005)'
    atl = f'Fabs({A}-{L})<=RangeMax(.02,{L}*.005)'
    U, O = f'{A}<{EP}-{tol}', f'{A}>{EP}+{tol}'
    miss = f'Fabs({A}-{EP})>{tol}'
    summ = [m('Lines', f'Count({st} %Parts_Detail)'), m('ROs', f'Count({st} distinct {R})'),
            m('Actual', f'Sum({st} Parts_Detail.extendedsale)'), m('Matrix', f'Sum({st} {Q}*{EP})'),
            m('List', f'Sum({st} {Q}*{L})'), m('Cost', f'Sum({st} Parts_Detail.extendedcost)'),
            m('MatchLn', f'Sum({st} If(Fabs({A}-{EP})<={tol},1))')]
    src_dims = [{'field': 'Parts_Detail.source', 'label': 'Source'},
                {'field': 'Source.description', 'label': 'Description'}]
    if a.bundle == 'summary':
        return {'chartType': 'table', 'dimensions': src_dims, 'measures': summ,
                'sort': [{'target': 'measure', 'index': 2, 'order': 'desc'}]}
    if a.bundle == 'totals':
        return {'chartType': 'table', 'dimensions': [], 'measures': summ}
    if a.bundle == 'classify':
        ms = [m('UnderLn', f'Sum({st} If({U},1))'), m('Under$', f'Sum({st} If({U},{Q}*({A}-{EP})))'),
              m('UnderAtListLn', f'Sum({st} If({U} and {atl},1))'),
              m('OverLn', f'Sum({st} If({O},1))'), m('Over$', f'Sum({st} If({O},{Q}*({A}-{EP})))'),
              m('BelowCostLn', f'Sum({st} If({A}<{C}-.01 and {Q}>0,1))'),
              m('BelowCost$', f'Sum({st} If({A}<{C}-.01 and {Q}>0,{Q}*({A}-{C})))'),
              m('ZeroListLn', f'Sum({st} If({L}=0 and {A}>0,1))')]
        return {'chartType': 'table', 'dimensions': [src_dims[0]], 'measures': ms,
                'sort': [{'target': 'measure', 'index': 1, 'order': 'asc'}]}
    if a.bundle in ('under', 'over', 'all_exc'):
        ms = [m('VarExt', f'Sum({st} If({miss},{Q}*({A}-{EP})))'), m('Qty', f'Sum({st} {Q})'),
              m('UnitCost', f'Sum({st} {C})'), m('List', f'Sum({st} {L})'), m('UnitSale', f'Sum({st} {A})')]
        out = {'chartType': 'table',
               'dimensions': [{'field': '%Parts_Detail', 'label': 'Key'},
                              {'field': 'Parts_Detail.partnumber', 'label': 'Part'},
                              {'field': 'Parts_Detail.source', 'label': 'Src'}],
               'measures': ms,
               'sort': [{'target': 'measure', 'index': 0, 'order': 'desc' if a.bundle == 'over' else 'asc'}]}
        if a.bundle != 'all_exc':
            out['limit'] = {'count': a.top, 'showOthers': False}
        else:
            # zero out non-exceptions so they can be dropped while paging
            out['note'] = 'Rows with VarExt null/0 are matches - drop them. Page with qlik_get_chart_data.'
        return out
    sys.exit(f'unknown parts bundle {a.bundle}')


def ro_bundle(a):
    bs = base_set(a, ro_app=True)
    pc = 'RO_Detail.partscost={"<>0"}'
    if a.bundle == 'profile':
        combo = (f"=Aggr(Concat({{<{bs},{pc}>}} distinct RO_Detail.labortype,'+'),RO_Detail.ronumber)")
        return {'chartType': 'table',
                'dimensions': [{'field': combo, 'label': 'LaborTypeCombo'}],
                'measures': [m('ROs', f'Count({{<{bs}>}} distinct RO_Detail.ronumber)'),
                             m('PartsSale', f'Sum({{<{bs},{pc}>}} RO_Detail.partssale)'),
                             m('PartsCost', f'Sum({{<{bs},{pc}>}} RO_Detail.partscost)')],
                'sort': [{'target': 'measure', 'index': 1, 'order': 'desc'}]}
    if a.bundle == 'opcodes':
        lts = quoted(a.matrix_types.split(',')) if a.matrix_types else None
        flt = f',RO_Detail.labortype={{{lts}}}' if lts else ''
        return {'chartType': 'table',
                'dimensions': [{'field': 'RO_Detail.labortype', 'label': 'LT'},
                               {'field': 'RO_Detail.opcodedescription', 'label': 'Op'}],
                'measures': [m('Lines', f'Count({{<{bs},{pc}{flt}>}} %RO_Detail)')],
                'sort': [{'target': 'measure', 'index': 0, 'order': 'desc'}],
                'limit': {'count': 3, 'dimensionIndex': 1, 'showOthers': False}}
    if a.bundle == 'rolist':
        if not a.matrix_types:
            sys.exit('--matrix-types required for rolist')
        lts = quoted(a.matrix_types.split(','))
        inner = (f'RO_Detail.ronumber=P({{<{bs},{pc},RO_Detail.labortype={{{lts}}}>}} RO_Detail.ronumber)'
                 f'-P({{<{bs},{pc},RO_Detail.labortype-={{{lts}}}>}} RO_Detail.ronumber)')
        st = '{<' + bs + ',' + inner + '>}'
        return {'chartType': 'table', 'dimensions': [],
                'measures': [m('n', f'Count({st} distinct RO_Detail.ronumber)'),
                             m('PartsSale', f'Sum({st} RO_Detail.partssale)'),
                             m('ROs', f"Concat({st} distinct RO_Detail.ronumber,',')")]}
    sys.exit(f'unknown RO bundle {a.bundle}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--matrix')
    ap.add_argument('--division', required=True)
    ap.add_argument('--year', required=True)
    ap.add_argument('--months', required=True, help='e.g. Aug,Sep,Oct')
    ap.add_argument('--bundle', required=True)
    ap.add_argument('--state', help='alternate state holding the RO list (parts app bundles)')
    ap.add_argument('--mode', default='include', choices=['include', 'exclude'])
    ap.add_argument('--matrix-types', help='comma list of matrix labor types, e.g. CPT or CP,CPM')
    ap.add_argument('--top', type=int, default=45)
    ap.add_argument('--out')
    a = ap.parse_args()
    if a.bundle in ('profile', 'opcodes', 'rolist'):
        res = ro_bundle(a)
    else:
        if not a.matrix:
            sys.exit('--matrix required for parts-app bundles')
        res = parts_bundle(a, json.load(open(a.matrix)))
    txt = json.dumps(res)
    if a.out:
        open(a.out, 'w').write(txt)
        print(f'{a.bundle}: {len(txt):,} chars -> {a.out}')
    else:
        print(txt)


if __name__ == '__main__':
    main()
