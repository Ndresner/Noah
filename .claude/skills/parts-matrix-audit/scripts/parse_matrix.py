#!/usr/bin/env python3
"""Parse a CDK "Price Breaks by Source" matrix export (+ optional source-description export) into matrix.json.

Matrix export layout (one sheet, header row 1):
    Source | Price Field | Price Break Field | Price Break Range | Formula
A new source starts on a row with a value in Source; following rows with a blank Source are more bands of
that source. Rows whose Source starts with '*' are separators. A source can repeat its Source/Price Field
values mid-block (page breaks in the CDK print) - that is the same source, not a new one.

Formula semantics (CDK price code): "COST + 275.00" = cost x (1 + 275%); "LIST + 30.00" = list x 1.30;
"LIST" = list; "COST - 10.00" = cost x 0.90. The band is chosen by testing the Price Break Field
(COST = unit cost, LIST = list price) against the range.

Usage:
    python3 parse_matrix.py --matrix matrix.xlsx [--sources sources.xlsx] --out matrix.json
"""
import argparse, json, re, sys
import openpyxl

FORMULA_RE = re.compile(r'^(COST|LIST)\s*(?:([+-])\s*([\d.]+))?$')


def _s(v):
    return '' if v is None else str(v).strip()


def parse_matrix(path, price_field=None):
    """Blocks are keyed by (source, price field): one source can carry a rule that sets LIST (how list is
    derived) and a separate rule that sets the customer-pay price field (e.g. PRICE10). Only the chosen
    price field is tested; default = the price field used by the most sources."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    blocks, order, warnings = {}, [], []
    cur = None
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        src, pfield, bfield, rng, formula = (list(row) + [None] * 5)[:5]
        src, pfield, bfield, rng, formula = map(_s, (src, pfield, bfield, rng, formula))
        if src.startswith('*'):
            cur = None
            continue
        if src:
            src = str(int(float(src))) if re.fullmatch(r'[\d.]+', src) else src
            key = (src, pfield or (cur[1] if cur and cur[0] == src else ''))
            if key != cur:
                if key in blocks:
                    if cur is None or cur[0] != src:
                        warnings.append(f'row {i}: source {src} / {key[1]} appears in two separate blocks; bands merged')
                else:
                    blocks[key] = {'price_field': key[1], 'break_field': bfield.upper(), 'bands': []}
                    order.append(key)
                cur = key
        if cur is None or not rng:
            continue
        m = re.match(r'^\s*([\d.]+)\s*-\s*([\d.]+)\s*$', rng)
        if not m:
            warnings.append(f'row {i}: source {cur[0]}: cannot read range "{rng}"')
            continue
        lo, hi = float(m.group(1)), float(m.group(2))
        f = FORMULA_RE.match(re.sub(r'\s+', ' ', formula.upper()))
        if not f:
            warnings.append(f'row {i}: source {cur[0]}: unsupported formula "{formula}" - source will be left untested')
            blocks[cur]['unsupported'] = True
            continue
        pct = float(f.group(3) or 0) * (-1 if f.group(2) == '-' else 1)
        blocks[cur]['bands'].append({'lo': lo, 'hi': hi, 'base': f.group(1), 'pct': pct, 'raw': formula})
    for k in order:
        blocks[k]['bands'].sort(key=lambda b: (b['hi'], b['lo']))
        if blocks[k]['break_field'] not in ('COST', 'LIST'):
            warnings.append(f'source {k[0]}: unknown break field "{blocks[k]["break_field"]}" - left untested')
            blocks[k]['unsupported'] = True
    counts = {}
    for _, pf in order:
        counts[pf] = counts.get(pf, 0) + 1
    chosen = price_field or max(counts, key=counts.get)
    sources, other = {}, {}
    for k in order:
        if k[1] == chosen:
            sources[k[0]] = blocks[k]
        else:
            other.setdefault(k[0], {})[k[1]] = blocks[k]
    for s in other:
        if s not in sources:
            warnings.append(f'source {s} has no {chosen} rule (only {list(other[s])}) - left untested')
    if len(counts) > 1:
        warnings.append(f'price fields in export: {counts}; testing {chosen} (pass --price-field to change). '
                        f'Other-field rules (e.g. how LIST is derived) are kept under other_price_fields.')
    return {'sources': sources, 'price_field': chosen, 'price_fields': sorted(counts), 'other_price_fields': other,
            'warnings': warnings}


def parse_sources(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None and re.fullmatch(r'[\d.]+', _s(row[0])):
            out[str(int(float(_s(row[0]))))] = {'mfg': _s(row[1]) if len(row) > 1 else '',
                                               'description': _s(row[2]) if len(row) > 2 else ''}
    return out


def unit_price(src_rule, cost, lst):
    """Python mirror of the Qlik expression (used for self-tests and spot checks)."""
    x = cost if src_rule['break_field'] == 'COST' else lst
    bands = src_rule['bands']
    band = next((b for b in bands if x <= b['hi'] + 1e-9), bands[-1])
    base = cost if band['base'] == 'COST' else lst
    return round(base * (1 + band['pct'] / 100.0) + 1e-9, 2)


def describe(rule):
    parts = []
    for b in rule['bands']:
        tail = f" {'+' if b['pct'] >= 0 else '-'}{abs(b['pct']):g}%" if b['pct'] else ''
        parts.append(f"{rule['break_field'].lower()} {b['lo']:g}-{b['hi']:g}: {b['base'].lower()}{tail}")
    return '; '.join(parts)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--matrix', required=True)
    ap.add_argument('--sources')
    ap.add_argument('--out', required=True)
    ap.add_argument('--price-field', help='customer-pay price field to test (default: most common)')
    a = ap.parse_args()
    res = parse_matrix(a.matrix, a.price_field)
    res['source_info'] = parse_sources(a.sources) if a.sources else {}
    if a.sources:
        missing = [s for s in res['source_info'] if s not in res['sources']]
        res['sources_without_matrix'] = missing
    for s, r in res['sources'].items():
        r['summary'] = describe(r)
    json.dump(res, open(a.out, 'w'), indent=1)
    print(f"{len(res['sources'])} sources parsed -> {a.out}")
    if res.get('sources_without_matrix'):
        print('sources with no matrix row:', ', '.join(res['sources_without_matrix']))
    for w in res['warnings']:
        print('WARNING:', w, file=sys.stderr)


if __name__ == '__main__':
    main()
