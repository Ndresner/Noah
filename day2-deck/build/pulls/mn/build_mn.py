"""Moving-now inputs (Oct 23): Jan-Mar and Jul-Sep 2026 per store, the same input vector `a` as each block file's
S rows, written as key `mn` = [[store, ...a_Jan-Mar, ...a_Jul-Sep]]. Raw Qlik pulls (set identifier 1, same basis as
the original pulls) are in this folder; each was reconciled to its file's m26 months before use (checks below).
Productivity's mn comes from hc/build_prod.py. Run from build/: python3 pulls/mn/build_mn.py"""
import json
P = "pulls/mn/"
def ro(f, n):   # Closed RO: Division x [q1 fields..., q3 fields...]
    d = json.load(open(P + f)); assert not d["selections"]
    return {r[0]: [float(v or 0) for v in r[1:1 + 2 * n]] for r in d["data"]["data"] if r[0] != "INEOS"}
def pl(f):      # P&L: Division x 2 measures, 9 monthly values each, as stored
    d = json.load(open(P + f))
    return {r[0]: [[float(v) for v in c.split("|")] for c in r[1:3]] for r in d["data"]["data"] if r[0] != "INEOS"}
q = lambda m: [sum(m[0:3]), sum(m[6:9])]
CP, I, W, OP = ro("cp_raw.json", 6), ro("i_raw.json", 6), ro("w_raw.json", 3), ro("op_raw.json", 5)
A, D = pl("pl_abs_raw.json"), pl("pl_dept_raw.json")
G = {r[0]: [[-float(v) for v in c.split("|")] for c in r[1:3]] for r in json.load(open("pulls/gross_m26_store.json"))["rows"] if r[0] != "INEOS"}
split = lambda v, n: (v[:n], v[n:2 * n])
def put(file, rows):
    d = json.load(open(file)); names = [s[0] for s in d["S"]]
    assert set(names) == set(rows), (file, set(names) ^ set(rows))
    w = len(d["m26"][0])
    for k in range(w):   # reconcile Q1 / Q3 store sums to the file's own group months
        for qi, ms in enumerate([(0, 3), (6, 9)]):
            g = sum(d["m26"][m][k] for m in range(*ms)); s = sum(rows[n][qi * w + k] for n in names)
            assert abs(g - s) < 0.5, (file, k, qi, g, s)
    d["mn"] = [[n] + rows[n] for n in names]; d["_mn_note"] = "Jan-Mar and Jul-Sep 2026 per store, same inputs as S (pulls/mn, Oct 23)."
    json.dump(d, open(file, "w")); print(file, "mn reconciles")
put("m5.json", CP)
put("prof/cphrs.json", {n: [v[2], v[8]] for n, v in CP.items()})
put("internal/internal.json", I)
put("wi.json", {n: W[n][:3] + I[n][1:3] + W[n][3:6] + I[n][7:9] for n in W})
put("abs.json", {n: [q([-x for x in A[n][0]])[0], q(A[n][1])[0], q([-x for x in A[n][0]])[1], q(A[n][1])[1]] for n in A})
put("ntgs.json", {n: [q([-x for x in D[n][0]])[0], q(G[n][0])[0], q([-x for x in D[n][0]])[1], q(G[n][0])[1]] for n in D})
put("ntgp.json", {n: [q([-x for x in D[n][1]])[0], q(G[n][1])[0], q([-x for x in D[n][1]])[1], q(G[n][1])[1]] for n in D})
# opcodes: all.json keeps per-measure blocks; mn = [store, num_q1, den_q1, num_q3, den_q3] (S order)
al = json.load(open("all.json"))
for i, k in enumerate(["align", "bal", "bf", "fuel"]):
    al[k]["mn"] = [[n, v[1 + i], v[0], v[6 + i], v[5]] for n, v in OP.items()]
json.dump(al, open("all.json", "w")); print("all.json mn written (reconciled in the pull check)")
