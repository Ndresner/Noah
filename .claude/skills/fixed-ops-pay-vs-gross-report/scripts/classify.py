"""Shared internal-advisor classification used by build_base.py and add_gross_workload.py.

Input rows come from the Qlik JSON "advisor_internal" block (one row per DMS advisor name per store),
user decisions from its "advisor_overrides" block: {"<ADP code>|<DMS advisor name>": "Internal" | "Customer"}.
"""
import re

ADV = "SERVICE ADVISOR"
INTERNAL_TITLE = "INTERNAL SERVICE ADVISOR"


def _nm(x):
    return [t for t in re.sub(r"[^A-Z ]", " ", (x or "").upper()).split() if t not in ("JR", "SR", "II", "III", "IV")]


def _split(x):
    last, _, first = (x or "").partition(",")
    return _nm(last), _nm(first)


def match(adp, code, dms):
    """adp: list of (code, store, payroll name, title, key). Returns (record or None, match quality)."""
    ql, qf = _split(dms)
    if not ql:
        return None, "None"
    pool = [a for a in adp if a[0] == code and _split(a[2])[0][:len(ql)] == ql]
    for test, label in ((lambda a: qf and _split(a[2])[1][:1] == qf[:1], "Exact"),
                        (lambda a: qf and _split(a[2])[1][:1] and _split(a[2])[1][0][:1] == qf[0][:1], "First initial")):
        hit = [a for a in pool if test(a)]
        if len(hit) == 1:
            return hit[0], label
    if len(pool) == 1:
        return pool[0], "Last name only"
    other = [a for a in adp if a[0] != code and _split(a[2])[0][:len(ql)] == ql and qf and _split(a[2])[1][:1] == qf[:1]]
    if len(other) == 1:
        return other[0], "Other store"
    return None, "None"


def classify(rows, adp, months, overrides=None):
    """Returns one dict per DMS row with suggestion, override, final class and whether it moves an ADP
    Service Advisor into INTERNAL SERVICE ADVISOR ("affects")."""
    overrides = overrides or {}
    out = []
    for x in rows:
        tot = x["int_ros"] + x["cpw_ros"]; share = x["int_ros"] / tot if tot else 0; per = x["int_ros"] / months
        a, q = match(adp, x["code"], x["name"])
        if a is None or q == "Other store":
            sug = "No change"
            why = ("Not on this store's Parts & Service payroll (recon/used-car staff, shared login, or left)"
                   if a is None else f"Payroll record is at {str(a[1]).strip()}")
        elif share >= 0.60 and per >= 25:
            sug, why = "Internal", f"{share:.0%} of ROs internal, {per:.0f} internal ROs/mo"
        elif share >= 0.60:
            sug, why = "Review", f"Mostly internal ({share:.0%}) but low volume ({per:.0f}/mo)"
        else:
            sug, why = "Review", f"Mixed: {share:.0%} internal, {per:.0f} internal ROs/mo"
        if q in ("First initial", "Last name only") and sug != "No change":
            why += " - confirm name match"
        ov = overrides.get(f"{x['code']}|{x['name']}")
        final = "No change" if sug == "No change" else (ov or sug)
        same_store = a is not None and q != "Other store"
        affects = same_store and a[3] == ADV and final == "Internal"
        out.append(dict(x, adp_rec=(a if same_store else None), adp=(a[2] if a else ""), title=(a[3] if a else ""),
                        q=q, share=share, per=per, sug=sug, why=why, override=ov or "", final=final,
                        is_adp_advisor=(same_store and a[3] == ADV), affects=affects))
    return out
