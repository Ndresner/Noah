"""Store gross records for the Day 2 deck, Jan-Sep 2026.

Source: inputs/September_2026_Fixed_Ops_Gross_Records.xlsx, Data tab - Qlik Fixed Ops - Accting TB
(e077f027), pulled Oct 8 2026, Family=FIXED GROSS, Class SERVICE / PARTS, body shop excluded.
History Jan 2021 - Sep 2026. Combined = service + parts for the same store-month.

Rules (Noah, Oct 21 2026):
  - TB is the records source. It matches the P&L to the dollar for every store except GMC PRESCOTT
    service (+$153,746 Jan-Sep 2026, +$286,430 Jan-Sep 2025, body shop). GMC PRESCOTT's
    Jan-Sep combined record holds on both bases.
  - Stores with shorter histories count the same as everyone else.
  - $0 store-months before a store opened are ignored. INEOS is excluded (same-store, 33 stores).

Writes records.json:
  ytd    best Jan-Sep ever, per store and category (svc / pts / tot), vs the best prior Jan-Sep
  broken new all-time monthly records set Jan-Sep 2026, per store and category
  group  Jan-Sep by year and the group's best months
Run from build/: python3 records/compute.py
"""
import json
import re
from collections import defaultdict

import openpyxl

SRC = "../inputs/September_2026_Fixed_Ops_Gross_Records.xlsx"
CATS = ("svc", "pts", "tot")
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
RENAME = {"BUICK GMC": "GMC PRESCOTT", "CHEVY GMC": "CHEVY GMC BULLHEAD"}   # TB name -> deck name


def deck_name(tb):
    base = re.sub(r" [A-Z]+-\d+$", "", tb)
    return RENAME.get(base, base)


wb = openpyxl.load_workbook(SRC, read_only=True)
H = defaultdict(dict)   # store -> (year, month) -> {svc, pts, tot}
for r in list(wb["Data"].iter_rows(values_only=True))[1:]:
    if not r[0]:
        continue
    s = deck_name(r[0])
    if s == "INEOS":
        continue
    svc, pts = float(r[4] or 0), float(r[5] or 0)
    if svc == 0 and pts == 0:
        continue   # before the store opened
    H[s][(int(r[1]), int(r[2]))] = {"svc": svc, "pts": pts, "tot": svc + pts}
assert len(H) == 33, len(H)

ytd, broken = {}, {}
for s, h in H.items():
    keys = sorted(h)
    # new all-time monthly records set in 2026 (each time the record was beaten)
    b = {}
    for c in CATS:
        best, hits = None, []
        for k in keys:
            v = h[k][c]
            if best and v > best[0] and k[0] == 2026 and k[1] <= 9:
                hits.append({"m": f"{MON[k[1]-1]} 2026", "v": v, "prev": best[0], "prevAt": f"{MON[best[1][1]-1]} {best[1][0]}"})
            if best is None or v > best[0]:
                best = (v, k)
        b[c] = hits
    broken[s] = b
    # Jan-Sep by year; a year counts only if all nine months are present
    years = {}
    for y in range(2021, 2027):
        if all((y, m) in h for m in range(1, 10)):
            years[y] = {c: sum(h[(y, m)][c] for m in range(1, 10)) for c in CATS}
    if 2026 not in years:
        raise SystemExit(f"{s}: Jan-Sep 2026 incomplete")
    prior = {y: v for y, v in years.items() if y < 2026}
    rec = {}
    for c in CATS:
        if not prior:
            continue
        py = max(prior, key=lambda y: prior[y][c])
        rec[c] = {"v": years[2026][c], "prev": prior[py][c], "prevYear": py, "record": years[2026][c] > prior[py][c]}
    ytd[s] = {"years": sorted(years), "rec": rec}

grp = {y: sum(sum(h[(y, m)]["tot"] for m in range(1, 10) if (y, m) in h) for h in H.values()) for y in range(2021, 2027)}
gm = defaultdict(float)
for h in H.values():
    for k, v in h.items():
        gm[k] += v["tot"]
top = sorted(gm.items(), key=lambda kv: -kv[1])
out = {
    "ytd": ytd, "broken": broken,
    "group": {"janSep": grp, "topMonths": [[f"{MON[k[1]-1]} {k[0]}", v] for k, v in top[:12]]},
    "_note": "Qlik Fixed Ops - Accting TB, pulled 2026-10-08; see records/compute.py.",
}
json.dump(out, open("records/records.json", "w"), indent=1)
n = sum(1 for s in ytd if ytd[s]["rec"].get("tot", {}).get("record"))
print(f"records: {n} of 33 stores best Jan-Sep combined; group Jan-Sep 2026 {grp[2026]:,.0f}")
