#!/usr/bin/env python3
"""Generate qlik_create_data_object arguments for the parts-matrix peer comparison.

Stores are matrix-library keys from reference/parts_matrices/manifest.json (e.g. toyota_prescott). The first key in
--stores is the TARGET; the rest are PEERS. Each store's manifest entry needs qlik_division, current_parsed,
matrix_labor_types and main_repair_source (the source whose cost-based rule is that store's main OEM repair matrix).

Parts app = Closed Parts Tickets  d3d25914-349d-4a55-857c-645fccefbecc  (bundles: basket, scenarios, pricepos)
RO app    = Closed Repair Orders  c3efc739-b063-47dd-a4d3-d1c4cac07ad0  (bundles: jobs, cpros, broad)

Bundles
  basket     parts app, target state: target's repair-matrix lines by unit-cost band - Lines, ROs, Actual, Cost, List,
             and the price under the target's own matrix and each peer's main repair matrix
  scenarios  parts app, target state: totals - ROs, sale if each line is capped at RangeMin(actual, peer matrix),
             the midpoint (target+first peer)/2, and lines above/below the target's own matrix
  pricepos   parts app, one state per store (--states pr,fl,...): sale/list/cost by price position vs list
  jobs       RO app: Division x job category (op-code description keywords) on matrix labor-type lines with parts
  cpros      RO app: customer-pay ROs (any C* labor line) by Division, for --year/--months (run once per year)
  broad      RO app: ROs with a matrix-type line carrying >= $50 parts cost, parts sale/cost, by Division (--year)

The RO list for the target (and per-store lists for pricepos) comes from the audit skill:
  python3 ../parts-matrix-audit/scripts/qlik_expr.py --division ... --bundle rolist --matrix-types CPT ...

Usage
  python3 peer_expr.py --stores toyota_prescott,toyota_flagstaff,toyota_spokane,toyota_henderson \
      --year 2026 --months Aug,Sep,Oct --bundle basket --state pr3 --out basket.json
"""
import argparse, json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'parts-matrix-audit', 'scripts'))
from qlik_expr import rule_expr  # noqa: E402

LIB = os.path.normpath(os.path.join(HERE, '..', '..', '..', '..', 'reference', 'parts_matrices'))
Q = 'Parts_Detail.quantitysold'
A = 'Num#(Trim(Parts_Detail.unitsaleprice))'
SRC = 'Num#(Trim(Parts_Detail.source))'

# Job categories from op-code description keywords. Order matters: spark plugs before menu (some menus are
# named "MAINTENANCE SERVICE - SPARK PLUGS"); menu before brakes/battery/coolant (menu text mentions them).
JOB_CATS = [
    ('05 Spark plugs', ['*spark plug*']),
    ('01 Menu/maint', ['*blue service*', '*green service*', '*oil change*', '*maintenance service*', '*oil & filter*',
                       '*oil and filter*', '*lube*']),
    ('02 Front brakes', ['*front brake*', '*pads and rotors front*', '*frt brake*', '*front pad*', '*brake*front*']),
    ('03 Rear brakes', ['*rear brake*', '*(rear)*', '*rear pad*', '*brake*rear*']),
    ('04 Battery', ['*battery replace*', '*replace battery*', 'battery', '*battery - replace*', '*replace*battery*']),
    ('06 Trans fluid', ['*transmission flu*', '*trans*flush*', '*transmission service*']),
    ('07 Coolant', ['*coolant*', '*cooling system*']),
    ('08 Fuel/induction', ['*induction*', '*injection clean*', '*direct injection*', '*fuel inj*', '*gdi*']),
    ('09 4x4/diff', ['*4x4*', '*4 wheel drive*', '*drive train*', '*differential*', '*transfer case*']),
    ('10 Tires', ['*tire*']),
]
OTHER_JOB = '11 Diag/general repair'
COST_BANDS = [(10, '1 cost <=$10'), (25, '2 $10-25'), (50, '3 $25-50'), (100, '4 $50-100'), (250, '5 $100-250'),
              (500, '6 $250-500')]
TOP_BAND = '7 >$500'


def load_store(key):
    man = json.load(open(os.path.join(LIB, 'manifest.json')))['stores']
    if key not in man:
        sys.exit(f'{key} not in manifest - add the store to reference/parts_matrices first')
    st = dict(man[key], key=key)
    st['matrix'] = json.load(open(os.path.join(LIB, st['current_parsed'])))
    return st


def main_rule(st, override=None):
    src = override or st.get('main_repair_source')
    if not src:
        sys.exit(f"{st['key']}: no main_repair_source in manifest - pass --sources-override {st['key']}=<source>")
    return st['matrix']['sources'][str(src)]


def repair_sources(st, override=None):
    """Target sources priced by the same rule as its main repair source."""
    rule = main_rule(st, override)
    sig = json.dumps([rule['break_field'], [(b['hi'], b['base'], b['pct']) for b in rule['bands']]])
    out = [s for s, r in st['matrix']['sources'].items()
           if not r.get('unsupported') and s.isdigit()
           and json.dumps([r['break_field'], [(b['hi'], b['base'], b['pct']) for b in r['bands']]]) == sig]
    return sorted(out, key=int)


def price(rule):
    return f'Round({rule_expr(rule)},.01)'


def qmonths(m):
    return ','.join("'" + x.strip() + "'" for x in m.split(','))


def parts_set(div, a, state, srcs=None):
    s = (f"Division={{'{div}'}},Year={{'{a.year}'}},Month={{{qmonths(a.months)}}},Parts_Detail.invsaletype={{'RO'}},"
         f'{Q}={{">0"}},Parts_Detail.referno={state}::Parts_Detail.referno')
    if srcs:
        s += f',Parts_Detail.source={{"=Match({SRC},{",".join(srcs)})"}}'
    return '{<' + s + '>}'


def m(label, expr):
    return {'label': label, 'expression': expr}


def band_dim():
    e = f"'{TOP_BAND}'"
    for hi, lbl in reversed(COST_BANDS):
        e = f"If(Parts_Detail.unitcost<={hi},'{lbl}',{e})"
    return {'field': '=' + e, 'label': 'CostBand'}


def job_dim():
    d = 'RO_Detail.opcodedescription'
    e = f"'{OTHER_JOB}'"
    for name, pats in reversed(JOB_CATS):
        e = f"If(WildMatch({d},{','.join(repr(p) for p in pats)}),'{name}',{e})"
    return {'field': '=' + e, 'label': 'Job'}


def overrides(a):
    out = {}
    for kv in (a.sources_override or '').split(','):
        if '=' in kv:
            k, v = kv.split('=')
            out[k.strip()] = v.strip()
    return out


def parts_bundles(a, stores):
    ov = overrides(a)
    tgt, peers = stores[0], stores[1:]
    if a.bundle == 'pricepos':
        states = a.states.split(',')
        if len(states) != len(stores):
            sys.exit('--states needs one state name per store, same order as --stores')
        lt = (f'If(Parts_Detail.listprice<=0,\'0 No list price\',If({A}<Parts_Detail.listprice*0.995,\'1 Below list\','
              f'If({A}<=Parts_Detail.listprice*1.005,\'2 At list (+/-0.5%)\',If({A}<=Parts_Detail.listprice*1.15,'
              f'\'3 List +0.5% to 15%\',If({A}<=Parts_Detail.listprice*1.5,\'4 List +15% to 50%\',\'5 List +50% or more\')))))')
        ms = []
        for st, sn in zip(stores, states):
            s = parts_set(st['qlik_division'], a, sn)
            ms += [m(f'{sn}_Sale', f'Sum({s} {Q}*{A})'), m(f'{sn}_List', f'Sum({s} {Q}*Parts_Detail.listprice)'),
                   m(f'{sn}_Cost', f'Sum({s} {Q}*Parts_Detail.unitcost)')]
        return {'chartType': 'table', 'dimensions': [{'field': '=' + lt, 'label': 'VsList'}], 'measures': ms,
                'sort': [{'target': 'dimension', 'order': 'asc'}]}
    if not a.state:
        sys.exit('--state (the target store RO-list state) is required for basket/scenarios')
    srcs = repair_sources(tgt, ov.get(tgt['key']))
    s = parts_set(tgt['qlik_division'], a, a.state, srcs)
    own = price(main_rule(tgt, ov.get(tgt['key'])))
    peer_px = [(p['key'], price(main_rule(p, ov.get(p['key'])))) for p in peers]
    if a.bundle == 'basket':
        ms = [m('Lines', f'Count({s} %Parts_Detail)'), m('ROs', f'Count({s} distinct Parts_Detail.referno)'),
              m('Actual', f'Sum({s} {Q}*{A})'), m('Cost', f'Sum({s} {Q}*Parts_Detail.unitcost)'),
              m('List', f'Sum({s} {Q}*Parts_Detail.listprice)'), m(tgt['key'], f'Sum({s} {Q}*{own})')]
        ms += [m(k, f'Sum({s} {Q}*{px})') for k, px in peer_px]
        return {'chartType': 'table', 'dimensions': [band_dim()], 'measures': ms,
                'sort': [{'target': 'dimension', 'order': 'asc'}], 'repair_sources': srcs}
    if a.bundle == 'scenarios':
        mid_key = a.mid_peer or peers[0]['key']
        mid_px = dict(peer_px)[mid_key]
        ms = [m('ROs', f'Count({s} distinct Parts_Detail.referno)'),
              m('Mid_' + mid_key, f'Sum({s} {Q}*RangeMin({A},Round(({own}+{mid_px})/2,.01)))')]
        ms += [m('Cap_' + k, f'Sum({s} {Q}*RangeMin({A},{px}))') for k, px in peer_px]
        ms += [m('LinesAboveOwn', f'Sum({s} If({A}>{own}+.02,1))'), m('LinesBelowOwn', f'Sum({s} If({A}<{own}-.02,1))')]
        return {'chartType': 'table', 'dimensions': [], 'measures': ms, 'repair_sources': srcs}
    sys.exit(f'unknown parts bundle {a.bundle}')


def ro_bundles(a, stores):
    base = f"Year={{'{a.year}'}},Month={{{qmonths(a.months)}}}"
    divs = ','.join("'" + st['qlik_division'] + "'" for st in stores)
    for st in stores:
        if not st.get('matrix_labor_types'):
            sys.exit(f"{st['key']}: matrix_labor_types missing in manifest - confirm with the user, then add it")
    if a.bundle == 'cpros':
        return {'chartType': 'table', 'dimensions': [{'field': 'Division', 'label': 'Div'}],
                'measures': [m('CPROs', f'Count({{<Division={{{divs}}},{base},RO_Detail.labortype={{"C*"}}>}} distinct RO_Detail.ronumber)'),
                             m('CPPartsSale', f'Sum({{<Division={{{divs}}},{base},RO_Detail.labortype={{"C*"}}>}} RO_Detail.partssale)')]}

    def per_store(agg, extra, field):
        return '+'.join(
            f"{agg}({{<Division={{'{st['qlik_division']}'}},{base},"
            f"RO_Detail.labortype={{{','.join(repr(t) for t in st['matrix_labor_types'])}}}{extra}>}} {field})"
            for st in stores)
    if a.bundle == 'jobs':
        ex = ',RO_Detail.partscost={">0"}'
        ms = [m('Lines', per_store('Count', ex, '%RO_Detail')), m('ROs', per_store('Count', ex, 'distinct RO_Detail.ronumber')),
              m('PartsSale', per_store('Sum', ex, 'RO_Detail.partssale')), m('PartsCost', per_store('Sum', ex, 'RO_Detail.partscost')),
              m('LaborSale', per_store('Sum', ex, 'RO_Detail.laborsale')), m('LaborCost', per_store('Sum', ex, 'RO_Detail.laborcost'))]
        return {'chartType': 'table', 'dimensions': [{'field': 'Division', 'label': 'Div'}, job_dim()], 'measures': ms,
                'sort': [{'target': 'dimension', 'order': 'asc'}]}
    if a.bundle == 'broad':
        ex = ',RO_Detail.partscost={">=50"}'
        ms = [m('ROs50', per_store('Count', ex, 'distinct RO_Detail.ronumber')),
              m('PartsSale', per_store('Sum', ex, 'RO_Detail.partssale')), m('PartsCost', per_store('Sum', ex, 'RO_Detail.partscost'))]
        return {'chartType': 'table', 'dimensions': [{'field': 'Division', 'label': 'Div'}], 'measures': ms}
    sys.exit(f'unknown RO bundle {a.bundle}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--stores', required=True, help='manifest keys, target first')
    ap.add_argument('--year', required=True)
    ap.add_argument('--months', required=True, help='e.g. Aug,Sep,Oct')
    ap.add_argument('--bundle', required=True, choices=['basket', 'scenarios', 'pricepos', 'jobs', 'cpros', 'broad'])
    ap.add_argument('--state', help='target store RO-list state (basket, scenarios)')
    ap.add_argument('--states', help='one state per store, same order as --stores (pricepos)')
    ap.add_argument('--mid-peer', help='peer key for the midpoint scenario (default: first peer)')
    ap.add_argument('--sources-override', help='key=source,... to use a different main repair source')
    ap.add_argument('--out')
    a = ap.parse_args()
    stores = [load_store(k.strip()) for k in a.stores.split(',')]
    res = ro_bundles(a, stores) if a.bundle in ('jobs', 'cpros', 'broad') else parts_bundles(a, stores)
    note = res.pop('repair_sources', None)
    txt = json.dumps(res)
    if a.out:
        open(a.out, 'w').write(txt)
        print(f'{a.bundle}: {len(txt):,} chars -> {a.out}' + (f' | target repair sources: {",".join(note)}' if note else ''))
    else:
        print(txt)


if __name__ == '__main__':
    main()
