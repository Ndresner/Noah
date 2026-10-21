"""Turn two saved Qlik pulls (by Month, by Division; the same N measures for 2025 then 2026,
plus a trailing pad column) into the {m25, m26, S} block JSON that blocks.load() reads.
Checks every $/hours field: store rows sum to the group months, both years. Field 0 may be a
distinct count (ROs), which is checked month-sum vs store-sum too (an RO closes in one month
at one store). Usage: python3 pulls/to_block.py MONTH.json DIVISION.json OUT.json N "note" """
import json, sys
mfile, dfile, out, n, note = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), sys.argv[5]
MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep"]
num = lambda x: float(x) if x not in (None, "", "-") else 0.0
mo = json.load(open(mfile)); dv = json.load(open(dfile))
assert not mo.get("selections") and not dv.get("selections"), ("selections active", mo.get("selections"), dv.get("selections"))
m = {r[0]: [num(x) for x in r[1:1 + 2 * n]] for r in mo["data"]["data"]}
m25 = [m[k][:n] for k in MON]; m26 = [m[k][n:] for k in MON]
S = []
for r in dv["data"]["data"]:
    v = [num(x) for x in r[1:1 + 2 * n]]
    if r[0] == "INEOS" or not any(v): continue
    S.append([r[0]] + v)
assert len(S) == 33, len(S)
for k in range(2 * n):
    g = sum((m25 if k < n else m26)[i][k % n] for i in range(9)); s = sum(x[1 + k] for x in S)
    assert abs(g - s) < 0.5, ("does not reconcile", k, g, s)
json.dump({"_note": note, "m25": m25, "m26": m26, "S": S}, open(out, "w"))
G25 = [sum(r[k] for r in m25) for k in range(n)]; G26 = [sum(r[k] for r in m26) for k in range(n)]
print("33 stores; reconciles on every field, both years")
print("G25", [round(x, 2) for x in G25]); print("G26", [round(x, 2) for x in G26])
