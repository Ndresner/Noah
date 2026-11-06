#!/usr/bin/env python3
"""Generate qlik_create_data_object arguments for the One-Line Customer Pay RO report.

App: Closed Repair Orders  c3efc739-b063-47dd-a4d3-d1c4cac07ad0  (all bundles)

Bundles (each writes one JSON with dimensions/measures/sort/limit; pass them to qlik_create_data_object as-is):
  profile        labor type x pay type: ROs, lines, labor sale, op codes, avg miles   (confirm roles with the user)
  advisor        advisor x one-line/multi-line flag: visits, CP sale, CP lines, oil visits, per-month visits, 60K+/30K+ counts
  tech           technician x flag: visits, oil visits
  ops            one-line ROs by the single op code (top --top)
  oil_addon      add-on sale/gross/median on oil-change visits that included more work
  repair_adv     advisor x flag: repair visits, with maintenance, CP sale, avg miles
  repair_tech    technician: repair visits, with maintenance, repair-only
  warranty_adv   advisor: warranty visits by type, with CP, CP sale/gross, recalls, avg miles
  warranty_tech  technician on the warranty line: repair-type and maintenance-type visits, with CP
  warranty_median  median CP sale on warranty visits that had CP work
  split_list     every RO ending in "C" in the period (split tickets)
  split_pairs    per-ticket detail for each C ticket AND its original (needs --c-list file from split_list)

Dates: RO_Header.closedate uses M/D/YYYY in set analysis (ISO dates silently return 0).

Usage:
  python3 qlik_queries.py --division "SUBARU ST GEORGE" --start 2026-08-01 --end 2026-10-31 --bundle advisor --out adv.json
"""
import argparse, calendar, datetime as dt, json, sys

D = {  # defaults = Subaru St George (Nov 2026). Override per store after the profile step.
    'oil': 'GREENS,GREEN,LOFS,LOF',
    'repair': 'INFO,DIAG,CEL,LCA,AC,TRANS,BATC,BULB',
    'maint': 'GREENS,GREEN,LOFS,LOF,ILOF,FLOFS,BLUE30,BLUE60,BLUE60T,BLUE606,AF,CAB,2T,CVT,INDS,IND,CF,DIFF,BF,ROT,EVROT,ROTBAL,BAL,4WA,4WAC,4WAT,BRR,BFRT,T4,TU4,MB4,MB1',
    'ws': 'WS,WR', 'wsc': 'WSC', 'wsm': 'WSM', 'pdi': 'WSPDI',
}


def q(csv):
    return ','.join("'" + x.strip() + "'" for x in csv.split(',') if x.strip())


def mdy(d):
    return f'{d.month}/{d.day}/{d.year}'


def m(label, expr):
    return {'label': label, 'expression': expr}


def months(start, end):
    out, d = [], start.replace(day=1)
    while d <= end:
        last = d.replace(day=calendar.monthrange(d.year, d.month)[1])
        out.append((d.strftime('%b'), max(d, start), min(last, end)))
        d = last + dt.timedelta(days=1)
    return out


def build(a):
    s, e = dt.date.fromisoformat(a.start), dt.date.fromisoformat(a.end)
    DIV = f"Division={{'{a.division}'}}"
    PER = f'RO_Header.closedate={{">={mdy(s)}<={mdy(e)}"}}'
    B = f'{DIV},{PER}'
    CP = f"{B},PayType.PayType={{'C'}},RO_Detail.laborsale={{\">0\"}}"
    CPall = f"{B},PayType.PayType={{'C'}}"
    OIL, REP, MNT = q(a.oil), q(a.repair), q(a.maint)
    WS, WSC, WSM, PDI = q(a.ws), q(a.wsc), q(a.wsm), q(a.pdi)
    WALL = f"{B},PayType.PayType={{'W'}},RO_Detail.labortype-={{{PDI}}}"
    hasCP = f'RO_Detail.ronumber=P({{<{CP}>}} RO_Detail.ronumber)'
    flag = {'field': (f"=Aggr(If(Count({{<{CP}>}} %RO_Detail)=1,'1 One line',If(Count({{<{CP}>}} %RO_Detail)>1,'2 Multi-line')),RO_Detail.ronumber)"),
            'label': 'Type'}
    adv = {'field': 'ServiceAdvisor.Name', 'label': 'Advisor'}
    tech = {'field': 'RO_Detail.techno', 'label': 'Tech'}
    b = a.bundle
    if b == 'profile':
        return {'chartType': 'table', 'dimensions': [{'field': 'RO_Detail.labortype', 'label': 'LT'}, {'field': 'PayType.PayType', 'label': 'PT'}],
                'measures': [m('ROs', f'Count({{<{B}>}} distinct RO_Detail.ronumber)'), m('Lines', f'Count({{<{B}>}} %RO_Detail)'),
                             m('LinesLabor', f'Count({{<{B},RO_Detail.laborsale={{">0"}}>}} %RO_Detail)'),
                             m('LaborSale', f'Sum({{<{B}>}} RO_Detail.laborsale)'), m('AvgMiles', f'Avg({{<{B}>}} RO_Header.mileage)'),
                             m('Ops', f"Concat({{<{B}>}} distinct RO_Detail.opcode,'/')"),
                             m('MaxClose', f"Date(Max({{<{DIV}>}} RO_Header.closedate))")],
                'sort': [{'target': 'measure', 'index': 0, 'order': 'desc'}]}
    if b == 'advisor':
        ms = [m('Visits', f'Count({{<{CP}>}} distinct RO_Detail.ronumber)'), m('CPSale', f'Sum({{<{CPall}>}} RO_Detail.laborsale+RO_Detail.partssale)'),
              m('CPLines', f'Count({{<{CP}>}} %RO_Detail)'), m('OilVisits', f'Count({{<{CP},RO_Detail.opcode={{{OIL}}}>}} distinct RO_Detail.ronumber)'),
              m('Visits60K', f'Count({{<{CP},RO_Header.mileage={{">=60000"}}>}} distinct RO_Detail.ronumber)'),
              m('Oil30K', f'Count({{<{CP},RO_Detail.opcode={{{OIL}}},RO_Header.mileage={{">=30000"}}>}} distinct RO_Detail.ronumber)')]
        for lab, ms_, me in months(s, e):
            ms.append(m(lab, f"Count({{<{DIV},RO_Header.closedate={{\">={mdy(ms_)}<={mdy(me)}\"}},PayType.PayType={{'C'}},RO_Detail.laborsale={{\">0\"}}>}} distinct RO_Detail.ronumber)"))
        return {'chartType': 'table', 'dimensions': [adv, flag], 'measures': ms, 'sort': [{'target': 'dimension', 'order': 'asc'}],
                'note': 'advisor JSON: one/multi = Visits by Type; sale_one/sale_multi = CPSale; cp_lines = CPLines (both rows); oil_only = OilVisits on the one-line row, oil_multi on the multi row; month_one/month_multi from the month columns; one_60k = Visits60K (one-line row); oil_only_30k = Oil30K (one-line row). Rows with a null Type are parts-only - ignore.'}
    if b == 'tech':
        return {'chartType': 'table', 'dimensions': [tech, flag],
                'measures': [m('Visits', f'Count({{<{CP}>}} distinct RO_Detail.ronumber)'),
                             m('OilVisits', f'Count({{<{CP},RO_Detail.opcode={{{OIL}}}>}} distinct RO_Detail.ronumber)')],
                'sort': [{'target': 'dimension', 'order': 'asc'}]}
    if b == 'ops':
        one = dict(flag, field=f"=Aggr(If(Count({{<{CP}>}} %RO_Detail)=1,'1 One line'),RO_Detail.ronumber)")
        return {'chartType': 'table', 'dimensions': [one, {'field': 'RO_Detail.opcode', 'label': 'Op'}],
                'measures': [m('ROs', f'Count({{<{CP}>}} distinct RO_Detail.ronumber)'), m('Desc', f'MaxString({{<{CP}>}} Left(RO_Detail.opcodedescription,60))'),
                             m('CPSale', f'Sum({{<{CP}>}} RO_Detail.laborsale+RO_Detail.partssale)'), m('AvgMiles', f'Avg({{<{CP}>}} RO_Header.mileage)')],
                'sort': [{'target': 'measure', 'index': 0, 'order': 'desc'}], 'limit': {'count': a.top, 'dimensionIndex': 1, 'showOthers': False},
                'note': 'Keep only rows where Type = "1 One line". Write a plain-language desc for each op in results.json.'}
    if b == 'oil_addon':
        P = f'RO_Detail.ronumber=P({{<{CP},RO_Detail.opcode={{{OIL}}}>}} RO_Detail.ronumber)'
        A = f"{CPall},RO_Detail.opcode-={{{OIL}}},{P}"
        return {'chartType': 'table', 'dimensions': [],
                'measures': [m('OilVisitsWithAddon', f'Count({{<{CP},RO_Detail.opcode-={{{OIL}}},{P}>}} distinct RO_Detail.ronumber)'),
                             m('AddOnSale', f'Sum({{<{A}>}} RO_Detail.laborsale+RO_Detail.partssale)'), m('AddOnGross', f'Sum({{<{A}>}} RO_Detail.laborgross+RO_Detail.partsgross)'),
                             m('MedianAddOn', f'Median(Aggr(Sum({{<{A}>}} RO_Detail.laborsale+RO_Detail.partssale),RO_Detail.ronumber))')],
                'note': 'avg_sale = AddOnSale / OilVisitsWithAddon; gross_margin = AddOnGross / AddOnSale; median_sale = MedianAddOn'}
    if b == 'repair_adv':
        R = f'RO_Detail.ronumber=P({{<{CP},RO_Detail.opcode={{{REP}}}>}} RO_Detail.ronumber)'
        return {'chartType': 'table', 'dimensions': [adv, flag],
                'measures': [m('RepairVisits', f'Count({{<{CP},RO_Detail.opcode={{{REP}}}>}} distinct RO_Detail.ronumber)'),
                             m('WithMaint', f'Count({{<{CP},RO_Detail.opcode={{{MNT}}},{R}>}} distinct RO_Detail.ronumber)'),
                             m('CPSale', f'Sum({{<{CPall},{R}>}} RO_Detail.laborsale+RO_Detail.partssale)'),
                             m('AvgMiles', f'Avg({{<{CP},RO_Detail.opcode={{{REP}}}>}} RO_Header.mileage)')],
                'sort': [{'target': 'dimension', 'order': 'asc'}],
                'note': 'visits = sum of both Type rows; with_maint = sum; one_line = RepairVisits on the one-line row; cp_sale = sum; avg_miles = weighted average'}
    if b == 'repair_tech':
        return {'chartType': 'table', 'dimensions': [tech],
                'measures': [m('RepairVisits', f'Count({{<{CP},RO_Detail.opcode={{{REP}}}>}} distinct RO_Detail.ronumber)'),
                             m('WithMaint', f'Count({{<{CP},RO_Detail.opcode={{{REP}}},RO_Detail.ronumber=P({{<{CP},RO_Detail.opcode={{{MNT}}}>}} RO_Detail.ronumber)>}} distinct RO_Detail.ronumber)'),
                             m('RepairOnly', f'Count({{<{CP},RO_Detail.opcode={{{REP}}},RO_Detail.ronumber=E({{<{CP},RO_Detail.opcode-={{{REP}}}>}} RO_Detail.ronumber)>}} distinct RO_Detail.ronumber)')],
                'sort': [{'target': 'measure', 'index': 0, 'order': 'desc'}]}
    if b == 'warranty_adv':
        def wt(types, extra=''):
            return (f'Count({{<{B},RO_Detail.labortype={{{types}}}{extra}>}} distinct RO_Detail.ronumber)')
        Pw = f'RO_Detail.ronumber=P({{<{WALL}>}} RO_Detail.ronumber)'
        return {'chartType': 'table', 'dimensions': [adv],
                'measures': [m('WS', wt(WS)), m('WS_CP', wt(WS, ',' + hasCP)), m('WSC', wt(WSC)), m('WSC_CP', wt(WSC, ',' + hasCP)),
                             m('WSM', wt(WSM)), m('WSM_CP', wt(WSM, ',' + hasCP)),
                             m('All', f'Count({{<{WALL}>}} distinct RO_Detail.ronumber)'), m('All_CP', f'Count({{<{WALL},{hasCP}>}} distinct RO_Detail.ronumber)'),
                             m('CPSale', f'Sum({{<{CPall},{Pw}>}} RO_Detail.laborsale+RO_Detail.partssale)'),
                             m('CPGross', f'Sum({{<{CPall},{Pw}>}} RO_Detail.laborgross+RO_Detail.partsgross)'),
                             m('Recall', f"Count({{<{WALL},RO_Detail.opcode={{'RECALL'}}>}} distinct RO_Detail.ronumber)"),
                             m('Recall_CP', f"Count({{<{WALL},RO_Detail.opcode={{'RECALL'}},{hasCP}>}} distinct RO_Detail.ronumber)"),
                             m('AvgMiles', f'Avg({{<{WALL}>}} RO_Header.mileage)')],
                'sort': [{'target': 'measure', 'index': 6, 'order': 'desc'}]}
    if b == 'warranty_tech':
        def wt(types, extra=''):
            return f'Count({{<{B},RO_Detail.labortype={{{types}}}{extra}>}} distinct RO_Detail.ronumber)'
        rt = ','.join(x for x in (WS, WSC) if x)
        return {'chartType': 'table', 'dimensions': [tech],
                'measures': [m('RepairType', wt(rt)), m('RepairType_CP', wt(rt, ',' + hasCP)), m('MaintType', wt(WSM)), m('MaintType_CP', wt(WSM, ',' + hasCP))],
                'sort': [{'target': 'measure', 'index': 0, 'order': 'desc'}]}
    if b == 'warranty_median':
        Pw = f'RO_Detail.ronumber=P({{<{WALL}>}} RO_Detail.ronumber)'
        return {'chartType': 'table', 'dimensions': [],
                'measures': [m('MedianCP', f'Median(Aggr(Sum({{<{CPall},{Pw}>}} RO_Detail.laborsale+RO_Detail.partssale),RO_Detail.ronumber))')]}
    if b == 'split_list':
        return {'chartType': 'table', 'dimensions': [{'field': 'RO_Detail.ronumber', 'label': 'RO'}],
                'measures': [m('Lines', f'Count({{<{B},RO_Detail.ronumber={{"*C"}}>}} %RO_Detail)')],
                'sort': [{'target': 'dimension', 'order': 'asc'}],
                'note': 'Page with qlik_get_chart_data if over 50 rows. Also check other suffixes with dimension =Right(RO_Detail.ronumber,1).'}
    if b == 'split_pairs':
        if not a.c_list:
            sys.exit('--c-list (JSON list of C ticket numbers from split_list) is required')
        cs = json.load(open(a.c_list))
        allro = cs + [c[:-1] for c in cs]
        # Originals are stored as numbers: selecting them as values fails, and a long set literal blows up every
        # measure, so the list goes once into a calculated dimension instead.
        dim = {'field': "=If(Match(RO_Detail.ronumber," + ','.join("'" + r + "'" for r in allro) + "),RO_Detail.ronumber)", 'label': 'RO'}
        X = f"{DIV},PayType.PayType={{'C'}},RO_Detail.laborsale={{\">0\"}}"
        return {'chartType': 'table', 'dimensions': [dim],
                'measures': [m('Adv', f'Only({{<{DIV}>}} ServiceAdvisor.Name)'), m('CPn', f'Count({{<{X}>}} %RO_Detail)'),
                             m('Oil', f'Count({{<{X},RO_Detail.opcode={{{OIL}}}>}} %RO_Detail)'), m('Rep', f'Count({{<{X},RO_Detail.opcode={{{REP}}}>}} %RO_Detail)'),
                             m('Mnt', f'Count({{<{X},RO_Detail.opcode={{{MNT}}}>}} %RO_Detail)'),
                             m('WS', f'Count({{<{DIV},RO_Detail.labortype={{{WS}}}>}} %RO_Detail)'), m('WSC', f'Count({{<{DIV},RO_Detail.labortype={{{WSC}}}>}} %RO_Detail)'),
                             m('WSM', f'Count({{<{DIV},RO_Detail.labortype={{{WSM}}}>}} %RO_Detail)'),
                             m('Recall', f"Count({{<{DIV},PayType.PayType={{'W'}},RO_Detail.opcode={{'RECALL'}}>}} %RO_Detail)"),
                             m('CPSale', f"Sum({{<{DIV},PayType.PayType={{'C'}}>}} RO_Detail.laborsale+RO_Detail.partssale)"),
                             m('CPGross', f"Sum({{<{DIV},PayType.PayType={{'C'}}>}} RO_Detail.laborgross+RO_Detail.partsgross)"),
                             m('CPTech', f"Concat({{<{X}>}} distinct RO_Detail.techno,'/')"),
                             m('RepTech', f"Concat({{<{X},RO_Detail.opcode={{{REP}}}>}} distinct RO_Detail.techno,'/')"),
                             m('WTech', f"Concat({{<{DIV},PayType.PayType={{'W'}},RO_Detail.labortype-={{{PDI}}}>}} distinct RO_Detail.techno,'/')"),
                             m('Close', f"Date(Only({{<{DIV}>}} RO_Header.closedate),'YYYY-MM-DD')")],
                'sort': [{'target': 'dimension', 'order': 'asc'}],
                'note': 'Page with qlik_get_chart_data. results.json split_rows = one list per row in this column order; replace NaN techs with "".'}
    sys.exit(f'unknown bundle {b}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--division', required=True)
    ap.add_argument('--start', required=True, help='YYYY-MM-DD (first day of the first full month)')
    ap.add_argument('--end', required=True, help='YYYY-MM-DD (last day of the last full month)')
    ap.add_argument('--bundle', required=True)
    for k, v in D.items():
        ap.add_argument('--' + k, default=v)
    ap.add_argument('--top', type=int, default=25)
    ap.add_argument('--c-list')
    ap.add_argument('--out')
    a = ap.parse_args()
    res = build(a)
    note = res.pop('note', None)
    txt = json.dumps(res)
    if a.out:
        open(a.out, 'w').write(txt)
        print(f'{a.bundle}: {len(txt):,} chars -> {a.out}' + (f'\n  NOTE: {note}' if note else ''))
    else:
        print(txt)


if __name__ == '__main__':
    main()
