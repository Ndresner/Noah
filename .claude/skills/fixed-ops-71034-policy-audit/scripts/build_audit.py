"""Build the monthly Service Policy (GL 71034) Employee & Advisor-Owned RO audit workbook.

Inputs
  --qlik     Closed RO pull for the month. Either the saved result file of the one-call Qlik
             Concat pull (see SKILL.md / qlik_pull_expr.py), or a CSV with columns
             logon,ro,closedate,custno,custname,advisor_no,advisor_name,labortype,
             laborsale,partssale,laborcost,partscost (one row per RO x labor type).
  --adp      ADP "Employee List - Active/Leave" export (.xlsx): Location Description,
             Store ID, CDK Employee ID #, Legal Last Name, Legal First Name, Position Status.
  --month    YYYY-MM (the close-date month audited).
  --out      output .xlsx. Run recalc.py on it afterwards (mandatory).

Store crosswalk: <repo>/reference/stores.json (shared). Config (../config): policy_labor_types.json
(which labor types post to 71034 per store), common_surnames.txt.

Tests, in priority order (each RO lands in exactly one row, the highest-priority hit):
  1 Advisor = Customer (number)   cust # == numeric part of the RO's advisor ID      Confirmed
  2 Advisor = Customer (name)     cust name == advisor name, on a regular cust #     Confirmed
  3 Employee # = Customer #       cust # == CDK ID of a SAME-store employee, names agree  Confirmed
  4 Employee name, same store     cust LAST,FIRST == same-store employee             Likely
  5 Same last name as advisor     possible relative                                  Review
  6 Employee name, other store    distinctive surname                                Review
  7 Employee name, other store    common surname (config/common_surnames.txt)        Low
"""
import argparse, calendar, csv, datetime as dt, json, os, re, sys
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = os.path.join(HERE, "..", "config")

def _find_crosswalk():
    """reference/stores.json at the repo root (shared store crosswalk)."""
    d = os.path.dirname(os.path.abspath(__file__))
    while d != os.path.dirname(d):
        p = os.path.join(d, "reference", "stores.json")
        if os.path.exists(p):
            return p
        d = os.path.dirname(d)
    raise SystemExit("STOP: reference/stores.json (shared store crosswalk) not found above " + os.path.abspath(__file__))

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument("--qlik", required=True)
ap.add_argument("--adp", required=True)
ap.add_argument("--month", required=True, help="YYYY-MM")
ap.add_argument("--out", required=True)
ap.add_argument("--policy-types", default=os.path.join(CFG, "policy_labor_types.json"))
ap.add_argument("--no-secondary", action="store_true", help="exclude the unconfirmed secondary labor types")
ap.add_argument("--qlik-other", default=None, help="saved result of qlik_pull_expr.py --other-lines (enables the Cross-Type Offsets tab)")
ap.add_argument("--qa", default=None, help="write QA details (JSON) here; default <out>.qa.json")
A = ap.parse_args()

Y, M = (int(x) for x in A.month.split("-"))
M_START, M_END = dt.date(Y, M, 1), dt.date(Y, M, calendar.monthrange(Y, M)[1])
MON_NAME = calendar.month_name[M]

STORES = {str(r["logon"]): dict(name=r["name"], adp=r["adp_code"], qlik=r["qlik_company"])
          for r in json.load(open(_find_crosswalk()))["stores"] if r["logon"] is not None}
PT = json.load(open(A.policy_types))
PRIMARY = {k: set(v) for k, v in PT["primary"].items()}
SECONDARY = {} if A.no_secondary else {k: set(v) for k, v in PT.get("secondary", {}).items()}
OOS_TYPES = set(PT.get("out_of_scope_types", ["ISP"]))
COMMON = {l.strip().upper() for l in open(os.path.join(CFG, "common_surnames.txt")) if l.strip() and not l.startswith("#")}
ADP_TO_LOGON = {v["adp"]: k for k, v in STORES.items()}

# ---------------- name normalisation ----------------
SUFFIX = {"JR", "SR", "II", "III", "IV", "V"}
def _clean(s):
    s = str(s or "").upper().replace("'", "").replace(".", "").replace("-", " ")
    return re.sub(r"[^A-Z ,]", " ", s)
def _toks(s):
    return [t for t in s.split() if t not in SUFFIX]
def ro_key(name):
    """'LAST,FIRST MIDDLE' -> ('LAST', 'FIRST'); None for business / blank names."""
    s = _clean(name)
    if "," not in s:
        return None
    last, first = s.split(",", 1)
    lt, ft = _toks(last), _toks(first)
    return (" ".join(lt), ft[0]) if lt and ft else None
def adp_key(last, first):
    lt, ft = _toks(_clean(last)), _toks(_clean(first))
    return (" ".join(lt), ft[0]) if lt and ft else None
def suffixes(s):
    return {t for t in _clean(s).replace(",", " ").split() if t in SUFFIX}
def num(s):
    s = str(s or "").strip()
    return s.lstrip("0") or s

# ---------------- load ADP ----------------
wb = openpyxl.load_workbook(A.adp, read_only=True, data_only=True)
ws = wb.worksheets[0]
rows = ws.iter_rows(values_only=True)
hdr = [str(h).strip() if h else "" for h in next(rows)]
ix = {h: hdr.index(h) for h in ("Location Description", "Store ID", "CDK Employee ID #", "Legal Last Name", "Legal First Name")}
emps, unmapped_loc = [], Counter()
for r in rows:
    if not r or r[ix["CDK Employee ID #"]] in (None, ""):
        continue
    code = str(r[ix["Store ID"]] or "").strip()
    e = dict(loc=str(r[ix["Location Description"]]).strip(), code=code, logon=ADP_TO_LOGON.get(code),
             id=num(r[ix["CDK Employee ID #"]]), last=str(r[ix["Legal Last Name"]] or "").strip(),
             first=str(r[ix["Legal First Name"]] or "").strip())
    e["key"] = adp_key(e["last"], e["first"])
    e["full"] = f"{e['first']} {e['last']}".strip()
    if not e["logon"]:
        unmapped_loc[f"{code} {e['loc']}"] += 1
    emps.append(e)
ADP_N, ADP_LOCS = len(emps), len({e["code"] for e in emps})
by_store_id = {(e["logon"], e["id"]): e for e in emps if e["logon"]}
by_key = defaultdict(list)
for e in emps:
    if e["key"]:
        by_key[e["key"]].append(e)

# ---------------- load Qlik ----------------
FIELDS = ["ro", "closedate", "custno", "custname", "advisor_no", "advisor_name", "labortype",
          "laborsale", "partssale", "laborcost", "partscost"]
lines = []
if A.qlik.lower().endswith(".csv"):
    for r in csv.DictReader(open(A.qlik, newline="", encoding="utf-8-sig")):
        r["logon"] = str(r["logon"]).split("|")[-1].strip()
        lines.append(r)
else:
    d = json.load(open(A.qlik))
    data = d["data"] if "data" in d else d
    cols = data["columns"]
    for rec in data["data"]:
        rec = dict(zip(cols, rec))
        logon = str(rec["logon"]).split("|")[-1].strip()
        blob = rec.get("rows") or ""
        expected = int(float(rec["n"])) if rec.get("n") not in (None, "") else None
        parts = [p for p in blob.split("~") if p]
        if expected is not None and len(parts) != expected:
            sys.exit(f"ERROR logon {logon}: Concat returned {len(parts)} rows but n={expected} (truncated?)")
        for p in parts:
            f = p.split("|")
            if len(f) != len(FIELDS):
                sys.exit(f"ERROR logon {logon}: malformed row ({len(f)} fields): {p[:120]}")
            r = dict(zip(FIELDS, f)); r["logon"] = logon
            lines.append(r)
for r in lines:
    for k in ("laborsale", "partssale", "laborcost", "partscost"):
        r[k] = float(r[k] or 0)
    r["closedate"] = dt.date.fromisoformat(str(r["closedate"])[:10])
    r["labortype"] = r["labortype"].strip().upper()
    r["custno"] = num(r["custno"]); r["advisor_no"] = num(r["advisor_no"])
bad_dates = [r for r in lines if not (M_START <= r["closedate"] <= M_END)]
if bad_dates:
    sys.exit(f"ERROR {len(bad_dates)} rows close outside {A.month} (e.g. RO {bad_dates[0]['ro']} {bad_dates[0]['closedate']}); check the Qlik Year/Month filter")

# ---------------- scope + RO roll-up ----------------
def types_for(logon):
    return PRIMARY.get(logon, set()) | SECONDARY.get(logon, set())
ros = {}                       # (logon, ro) -> RO record
unmapped_types = Counter()     # policy-family types present at a mapped store but not mapped there
for r in lines:
    lg = r["logon"]
    if lg in PRIMARY:
        if r["labortype"] not in types_for(lg):
            unmapped_types[(lg, r["labortype"])] += 1
            continue
        scope = "in"
    else:
        if r["labortype"] not in OOS_TYPES:
            continue
        scope = "out"
    k = (lg, r["ro"])
    o = ros.setdefault(k, dict(logon=lg, ro=r["ro"], closedate=r["closedate"], custno=r["custno"],
                               custname=r["custname"].strip(), advisor_no=r["advisor_no"],
                               advisor_name=r["advisor_name"].strip(), types=[], ls=0.0, ps=0.0, cost=0.0, scope=scope))
    if r["labortype"] not in o["types"]:
        o["types"].append(r["labortype"])
    o["ls"] += r["laborsale"]; o["ps"] += r["partssale"]; o["cost"] += r["laborcost"] + r["partscost"]
    o["closedate"] = max(o["closedate"], r["closedate"])
POP = sum(1 for o in ros.values() if o["scope"] == "in")
POP_STORES = len({o["logon"] for o in ros.values() if o["scope"] == "in"})
sec_counts = {lg: sum(1 for o in ros.values() if o["logon"] == lg and set(o["types"]) & t) for lg, t in SECONDARY.items()}

# Credit reclass pattern: policy parts charged while policy labor is credited (net negative) on the same RO.
OFFSET_TOL = 1.00   # |labor + parts| within this = "Net $0"
OFFSET_ORDER = {"Net $0 — labor credit offsets parts": 1, "Labor credit exceeds parts": 2, "Partial offset": 3}
for o in ros.values():
    o["offset"] = None
    if o["ps"] > 0.005 and o["ls"] < -0.005:
        net = o["ls"] + o["ps"]
        o["offset"] = ("Net $0 — labor credit offsets parts" if abs(net) <= OFFSET_TOL
                       else "Labor credit exceeds parts" if net < 0 else "Partial offset")

# Cross-type offsets: policy parts charged, with a labor credit on a DIFFERENT labor type on the same RO.
PAYNAME = {"I": "Internal", "C": "Customer", "W": "Warranty"}
XORDER = {"Net $0 — other-type credit offsets policy": 1, "Other-type credit exceeds policy charge": 2, "Partial offset": 3}
other_lines = 0
for o in ros.values():
    o["xcredit"], o["xtypes"], o["xoffset"] = 0.0, [], None
if A.qlik_other:
    d2 = json.load(open(A.qlik_other)); d2 = d2["data"] if "data" in d2 else d2
    for rec in d2["data"]:
        rec = dict(zip(d2["columns"], rec)); lg = str(rec["logon"]).split("|")[-1].strip()
        parts = [p for p in (rec.get("rows") or "").split("~") if p]
        if rec.get("n") not in (None, "") and len(parts) != int(float(rec["n"])):
            sys.exit(f"ERROR other-lines logon {lg}: Concat returned {len(parts)} rows but n={rec['n']} (truncated?)")
        for p in parts:
            f = p.split("|")
            if len(f) != 7:
                sys.exit(f"ERROR other-lines logon {lg}: malformed row: {p[:120]}")
            ro_no, lt, pay, ls = f[0], f[1].strip().upper(), f[2], float(f[3] or 0)
            o = ros.get((lg, ro_no))
            if not o or o["scope"] != "in" or lt in types_for(lg):
                continue
            other_lines += 1
            if ls < -0.005:
                o["xcredit"] += ls
                o["xtypes"].append((lt, PAYNAME.get(pay, pay or "?"), ls))
    for o in ros.values():
        if o["scope"] == "in" and o["ps"] > 0.005 and o["xcredit"] < -0.005:
            net = o["ls"] + o["ps"] + o["xcredit"]
            o["xoffset"] = ("Net $0 — other-type credit offsets policy" if abs(net) <= OFFSET_TOL
                            else "Other-type credit exceeds policy charge" if net < 0 else "Partial offset")

# ---------------- tests ----------------
FLAGS = {1: ("Advisor = Customer (number)", "Confirmed"),
         2: ("Advisor = Customer (name; RO on regular cust #)", "Confirmed"),
         3: ("Employee # = Customer #", "Confirmed"),
         4: ("Employee name, same store (regular cust #)", "Likely"),
         5: ("Same last name as advisor", "Review"),
         6: ("Employee name, other store (distinctive name)", "Review"),
         7: ("Employee name, other store (common name)", "Low")}
MEANING = {1: "Advisor wrote a policy RO to themselves on their own employee customer record.",
           2: "Advisor wrote a policy RO to themselves under a regular customer number — bypasses the employee record.",
           3: "Customer # equals a same-store CDK employee ID and the names agree.",
           4: "Customer name matches a same-store employee, but RO is on a regular customer number.",
           5: "Customer shares the advisor's last name — possible relative. Needs a store to confirm.",
           6: "Name matches an employee at a different Findlay store; name is uncommon.",
           7: "Name matches an employee elsewhere, but the name is common — likely coincidence."}

def names_agree(ck, e):
    return ck is not None and e["key"] is not None and (ck[0] == e["key"][0] or set(ck[0].split()) & set(e["key"][0].split()))

qa = dict(id_hits_name_mismatch=[], multi_other_store=[])
flagged = []
for o in ros.values():
    lg, ck, ak = o["logon"], ro_key(o["custname"]), ro_key(o["advisor_name"])
    hit, emp, note = None, None, ""
    adv_emp = by_store_id.get((lg, o["advisor_no"]))
    if o["custno"] and o["custno"] == o["advisor_no"]:
        hit, emp = 1, adv_emp
    elif ck and ak and ck == ak:
        hit = 2
        emp = adv_emp if adv_emp and names_agree(ck, adv_emp) else next((e for e in by_key.get(ck, []) if e["logon"] == lg), None)
        if emp and emp["id"] == o["advisor_no"]:
            note = f"ADP lists as {emp['full']}; advisor # {o['advisor_no']} = their CDK employee ID" if adp_key(emp["last"], emp["first"]) and _clean(emp["last"]).strip() != ck[0] else f"Advisor # {o['advisor_no']} = their CDK employee ID"
    else:
        e = by_store_id.get((lg, o["custno"]))
        if e and names_agree(ck, e):
            hit, emp = 3, e
        else:
            if e:
                qa["id_hits_name_mismatch"].append(dict(logon=lg, ro=o["ro"], custno=o["custno"], custname=o["custname"], adp=e["full"]))
            cands = by_key.get(ck, []) if ck else []
            same = [x for x in cands if x["logon"] == lg]
            other = [x for x in cands if x["logon"] != lg]
            if same:
                hit, emp = 4, same[0]
            elif ck and ak and ck[0] == ak[0]:
                hit, note = 5, "Possible relative"
            elif other:
                hit, emp = (7 if ck[0] in COMMON or any(t in COMMON for t in ck[0].split()) else 6), other[0]
                if len(other) > 1:
                    note = f"{len(other)} ADP matches: " + "; ".join(f"{x['full']} ({x['loc'].title()})" for x in other)
                    qa["multi_other_store"].append(dict(logon=lg, ro=o["ro"], note=note))
    if hit in (3, 4, 6, 7) and emp and not note:
        rs, es = suffixes(o["custname"]), suffixes(emp["last"] + " " + emp["first"])
        if rs != es:
            note = f"Suffix differs (RO: {'/'.join(sorted(rs)) or 'none'}, ADP: {'/'.join(sorted(es)) or 'none'}) — could be a relative, not the employee"
    if hit:
        o.update(sort=hit, flag=FLAGS[hit][0], conf=FLAGS[hit][1], emp=emp, note=note)
        flagged.append(o)

inscope = sorted([o for o in flagged if o["scope"] == "in"], key=lambda o: (o["sort"], int(o["logon"]), o["closedate"], o["ro"]))
oos = sorted([o for o in flagged if o["scope"] == "out" and o["sort"] <= 4], key=lambda o: (int(o["logon"]), o["ro"]))
id_hits = sum(1 for o in flagged if o["sort"] == 3 or (o["sort"] == 1 and o["emp"]))

# ---------------- workbook ----------------
NAVY, GOLD, RED, GRID = "1A2744", "C9A04B", "F8CBAD", "BFBFBF"
AMBER, AMBER_FONT = "FFEB9C", "9C6500"   # credit reclass pattern
USD = '\\$#,##0.00;"($"#,##0.00\\);\\-'
thin = Side(style="thin", color=GRID); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def F(**k): return Font(name="Arial", size=k.pop("size", 10), **k)
def fill(c): return PatternFill("solid", fgColor=c, bgColor=c)
def header(ws, row, labels, height=None):
    for i, t in enumerate(labels, 1):
        c = ws.cell(row, i, t); c.font = F(bold=True, color="FFFFFF"); c.fill = fill(NAVY)
        c.alignment = Alignment(vertical="center", wrap_text=True); c.border = BOX
    if height: ws.row_dimensions[row].height = height
def widths(ws, d):
    for k, v in d.items(): ws.column_dimensions[k].width = v

out = openpyxl.Workbook()
S = out.active; S.title = "Summary"
FR = out.create_sheet("Flagged ROs"); PO = out.create_sheet("Policy Offsets"); XO = out.create_sheet("Cross-Type Offsets")
NS = out.create_sheet("Not in 71034 Scope"); ML = out.create_sheet("Method & Limits")

# Flagged ROs
cols = ["Sort", "Flag", "Confidence", "Logon #", "Store", "RO #", "Close Date", "Customer #", "Customer Name (RO)",
        "Advisor #", "Advisor Name", "Policy Labor Type", "Policy Labor Sale", "Policy Parts Sale", "Policy Total Sale",
        "Policy Cost", "Matched Employee (ADP)", "Employee Store (ADP)", "CDK Employee ID", "Notes",
        "Credit Reclass Pattern", "Labor Credit Offset", "Other-Type Credit Pattern", "Other-Type Labor Credit"]
header(FR, 1, cols, 35.05)
for i, o in enumerate(inscope, 2):
    emp = o["emp"]
    vals = [o["sort"], o["flag"], o["conf"], int(o["logon"]), STORES[o["logon"]]["name"], o["ro"], o["closedate"],
            o["custno"], o["custname"], o["advisor_no"], o["advisor_name"], " + ".join(o["types"]),
            round(o["ls"], 2), round(o["ps"], 2), f"=M{i}+N{i}", round(o["cost"], 2),
            emp["full"] if emp else None, emp["loc"].title() if emp else None, emp["id"] if emp else None, o["note"] or None,
            o["offset"], round(o["ls"], 2) if o["offset"] else None,
            (o["xoffset"] + " (" + ", ".join(sorted({t for t, _, _ in o["xtypes"]})) + ")") if o["xoffset"] else None,
            round(o["xcredit"], 2) if o["xoffset"] else None]
    red = o["sort"] in (1, 2)
    for j, v in enumerate(vals, 1):
        c = FR.cell(i, j, v); c.border = BOX
        c.font = F() if j == 15 else F(color="0000FF")
        if red: c.fill = fill(RED)
        if j in (21, 22) and o["offset"]: c.fill = fill(AMBER); c.font = F(color=AMBER_FONT)
        if j in (23, 24) and o["xoffset"]: c.fill = fill(AMBER); c.font = F(color=AMBER_FONT)
        if j == 7: c.number_format = "m/d/yyyy"
        if j in (13, 14, 15, 16, 22, 24): c.number_format = USD
T = len(inscope) + 2; L = T - 1
FR.cell(T, 2, "TOTAL"); FR.cell(T, 6, f"=COUNTA(F2:F{L})")
for col in "MNOPVX":
    FR[f"{col}{T}"] = f"=SUM({col}2:{col}{L})"
FR[f"U{T}"] = f"=COUNTA(U2:U{L})"; FR[f"W{T}"] = f"=COUNTA(W2:W{L})"
for j in range(1, 25):
    c = FR.cell(T, j); c.fill = fill(GOLD); c.border = BOX; c.font = F(bold=True)
    if j in (13, 14, 15, 16, 22, 24): c.number_format = USD
FR.freeze_panes = "G2"; FR.auto_filter.ref = f"A1:X{L}"
widths(FR, dict(A=6, B=34, C=11, D=8, E=20, F=10, G=11, H=13, I=30, J=10, K=24, L=10, M=12, N=13, O=13, P=11, Q=24, R=26, S=12, T=30, U=34, V=13, W=44, X=13))

# Summary
S["A1"] = f"Service Policy (71034) Audit — Employee & Advisor-Owned ROs, {MON_NAME} {Y}"; S["A1"].font = F(size=14, bold=True, color=NAVY)
S["A2"] = (f"Closed ROs (close date {M}/1–{M}/{M_END.day}/{Y}) carrying each store's 71034 policy labor type, "
           f"matched to the ADP Active/Leave employee list ({ADP_N:,} employees).")
S["A2"].font = F(size=9, italic=True)
header(S, 4, ["Flag", "Confidence", "ROs", "Policy Sale", "Policy Cost", "What it means"])
for r, k in enumerate(range(1, 8), 5):
    vals = [FLAGS[k][0], FLAGS[k][1], f"=COUNTIFS('Flagged ROs'!$B:$B,A{r})",
            f"=SUMIFS('Flagged ROs'!$O:$O,'Flagged ROs'!$B:$B,A{r})",
            f"=SUMIFS('Flagged ROs'!$P:$P,'Flagged ROs'!$B:$B,A{r})", MEANING[k]]
    for j, v in enumerate(vals, 1):
        c = S.cell(r, j, v); c.font = F(); c.border = BOX; c.alignment = Alignment(vertical="top", wrap_text=True)
        if k in (1, 2): c.fill = fill(RED)
        if j in (4, 5): c.number_format = USD
    S.row_dimensions[r].height = 23.85 if len(MEANING[k]) > 70 or len(FLAGS[k][0]) > 40 else 15
S["A12"], S["C12"], S["D12"], S["E12"] = "TOTAL FLAGGED", "=SUM(C5:C11)", "=SUM(D5:D11)", "=SUM(E5:E11)"
for j in range(1, 7):
    c = S.cell(12, j); c.fill = fill(GOLD); c.font = F(bold=True); c.border = BOX
    if j in (4, 5): c.number_format = USD
S["A13"] = "Check: ties to Flagged ROs tab (should be 0)"
S["C13"] = f"=C12-'Flagged ROs'!F{T}"; S["D13"] = f"=ROUND(D12-'Flagged ROs'!O{T},2)"; S["E13"] = f"=ROUND(E12-'Flagged ROs'!P{T},2)"
for a in ("A13", "C13", "D13", "E13"): S[a].font = F()
header(S, 15, ["Store", "Logon #", "Confirmed + Likely ROs", "Policy Sale", "Policy Cost", "All Flagged ROs"], 46.25)
store_rows = sorted({int(o["logon"]) for o in inscope})
r = 16
for lg in store_rows:
    def cl(col): return (f"SUMIFS('Flagged ROs'!${col}:${col},'Flagged ROs'!$D:$D,B{r},'Flagged ROs'!$C:$C,\"Confirmed\")+"
                         f"SUMIFS('Flagged ROs'!${col}:${col},'Flagged ROs'!$D:$D,B{r},'Flagged ROs'!$C:$C,\"Likely\")")
    vals = [STORES[str(lg)]["name"], lg,
            f"=COUNTIFS('Flagged ROs'!$D:$D,B{r},'Flagged ROs'!$C:$C,\"Confirmed\")+COUNTIFS('Flagged ROs'!$D:$D,B{r},'Flagged ROs'!$C:$C,\"Likely\")",
            "=" + cl("O"), "=" + cl("P"), f"=COUNTIFS('Flagged ROs'!$D:$D,B{r})"]
    for j, v in enumerate(vals, 1):
        c = S.cell(r, j, v); c.font = F(); c.border = BOX
        if j in (4, 5): c.number_format = USD
    r += 1
S.cell(r, 1, "TOTAL")
for col in "CDEF":
    S[f"{col}{r}"] = f"=SUM({col}16:{col}{r - 1})"
for j in range(1, 7):
    c = S.cell(r, j); c.fill = fill(GOLD); c.font = F(bold=True); c.border = BOX
    if j in (4, 5): c.number_format = USD
widths(S, dict(A=46, B=12, C=8, D=13, E=13, F=70))
S.row_dimensions[1].height = 17.35

# Policy Offsets (credit reclass pattern)
offs = sorted([o for o in ros.values() if o["scope"] == "in" and o["offset"]],
              key=lambda o: (int(o["logon"]), OFFSET_ORDER[o["offset"]], o["ls"], o["ro"]))
off_stores = sorted({int(o["logon"]) for o in offs})
PO["A1"] = f"Policy Offsets — Credit Reclass Pattern, {MON_NAME} {Y}"; PO["A1"].font = F(size=14, bold=True, color=NAVY)
PO["A2"] = ("Policy ROs where parts were charged to policy while policy labor was credited (negative) on the same RO. "
            f"Net $0 = labor credit within ${OFFSET_TOL:.2f} of the parts charge. Labor Credit is the net policy labor sale on the policy labor type(s).")
PO["A2"].font = F(size=9, italic=True)
SH = 4; D0 = SH + len(off_stores) + 4          # store summary header row; detail header row
D1, D2 = D0 + 1, D0 + max(len(offs), 1)       # detail data rows
def rng(col): return f"${col}${D1}:${col}${D2}"
header(PO, SH, ["Store", "Logon #", "ROs", "Net $0 ROs", "Policy Parts Sale", "Labor Credit", "Net Policy Sale",
                "Policy Cost", "Employee-Flagged ROs"], 35.05)
r = SH + 1
for lg in off_stores:
    vals = [STORES[str(lg)]["name"], lg, f"=COUNTIFS({rng('A')},B{r})",
            f"=COUNTIFS({rng('A')},B{r},{rng('N')},\"Net $0*\")",
            f"=SUMIFS({rng('J')},{rng('A')},B{r})", f"=SUMIFS({rng('K')},{rng('A')},B{r})",
            f"=E{r}+F{r}", f"=SUMIFS({rng('M')},{rng('A')},B{r})", f"=COUNTIFS({rng('A')},B{r},{rng('O')},\"?*\")"]
    for j, v in enumerate(vals, 1):
        c = PO.cell(r, j, v); c.font = F(); c.border = BOX
        if j in (5, 6, 7, 8): c.number_format = USD
    r += 1
PO.cell(r, 1, "TOTAL")
for col in "CDEFGHI":
    PO[f"{col}{r}"] = f"=SUM({col}{SH + 1}:{col}{r - 1})" if off_stores else 0
for j in range(1, 10):
    c = PO.cell(r, j); c.fill = fill(GOLD); c.font = F(bold=True); c.border = BOX
    if j in (5, 6, 7, 8): c.number_format = USD
PO_TOT = r
header(PO, D0, ["Logon #", "Store", "RO #", "Close Date", "Customer #", "Customer Name (RO)", "Advisor #", "Advisor Name",
                "Policy Labor Type", "Policy Parts Sale", "Labor Credit", "Net Policy Sale", "Policy Cost",
                "Pattern", "Employee Flag (Flagged ROs)", "Confidence"], 35.05)
for i, o in enumerate(offs, D1):
    vals = [int(o["logon"]), STORES[o["logon"]]["name"], o["ro"], o["closedate"], o["custno"], o["custname"],
            o["advisor_no"], o["advisor_name"], " + ".join(o["types"]), round(o["ps"], 2), round(o["ls"], 2),
            f"=J{i}+K{i}", round(o["cost"], 2), o["offset"], o.get("flag") if o in inscope else None,
            o.get("conf") if o in inscope else None]
    for j, v in enumerate(vals, 1):
        c = PO.cell(i, j, v); c.border = BOX
        c.font = F() if j == 12 else F(color="0000FF")
        if j == 4: c.number_format = "m/d/yyyy"
        if j in (10, 11, 12, 13): c.number_format = USD
        if j == 14 and o["offset"].startswith("Net $0"): c.fill = fill(AMBER); c.font = F(color=AMBER_FONT)
        if j in (15, 16) and o in inscope: c.fill = fill(RED) if o["sort"] in (1, 2) else fill(AMBER)
if not offs:
    PO.cell(D1, 1, "None this month.").font = F(italic=True)
else:
    PO.auto_filter.ref = f"A{D0}:P{D2}"
widths(PO, dict(A=8, B=22, C=10, D=11, E=11, F=28, G=10, H=24, I=11, J=13, K=13, L=13, M=12, N=34, O=40, P=12))

# Summary pointer to the offsets tab
r = S.max_row + 2
S.cell(r, 1, "Credit reclass pattern (policy parts charged, policy labor credited)").font = F(bold=True)
S.cell(r + 1, 1, "ROs with the pattern / of which net $0").font = F()
S.cell(r + 1, 3, f"='Policy Offsets'!C{PO_TOT}").font = F(); S.cell(r + 1, 4, f"='Policy Offsets'!D{PO_TOT}").font = F()
S.cell(r + 2, 1, "Policy parts charged / labor credit offsetting it").font = F()
S.cell(r + 2, 4, f"='Policy Offsets'!E{PO_TOT}"); S.cell(r + 2, 5, f"='Policy Offsets'!F{PO_TOT}")
S.cell(r + 3, 1, "Of these, ROs also on Flagged ROs (employee / advisor)").font = F()
S.cell(r + 3, 3, f"='Policy Offsets'!I{PO_TOT}").font = F()
for a in (S.cell(r + 2, 4), S.cell(r + 2, 5)): a.number_format = USD; a.font = F()
S.cell(r + 1, 6, "Details on the Policy Offsets tab. On Flagged ROs these rows are marked in columns U–V.").font = F(size=9, italic=True)

# Cross-Type Offsets
xoffs = sorted([o for o in ros.values() if o["xoffset"]],
               key=lambda o: (int(o["logon"]), XORDER[o["xoffset"]], o["xcredit"], o["ro"]))
x_stores = sorted({int(o["logon"]) for o in xoffs})
XO["A1"] = f"Cross-Type Offsets — Policy Charged, Credit on Another Labor Type, {MON_NAME} {Y}"; XO["A1"].font = F(size=14, bold=True, color=NAVY)
XO["A2"] = ("Policy ROs where parts were charged to policy and a DIFFERENT labor type on the same RO carries a labor credit (negative net labor sale). "
            "Weaker evidence than a same-type offset: comebacks, goodwill and corrections also post credits. Filter Credit Pay Type to focus (Internal / Warranty / Customer)."
            + ("" if A.qlik_other else " NOT RUN: build_audit.py was run without --qlik-other."))
XO["A2"].font = F(size=9, italic=True)
XSH = 4; XD0 = XSH + len(x_stores) + 4
XD1, XD2 = XD0 + 1, XD0 + max(len(xoffs), 1)
def xr(col): return f"${col}${XD1}:${col}${XD2}"
header(XO, XSH, ["Store", "Logon #", "ROs", "Net $0 ROs", "Policy Parts Sale", "Policy Total Sale", "Other-Type Credit",
                 "Net After Credit", "Employee-Flagged ROs", "Credit Labor Types (ROs)"], 35.05)
r = XSH + 1
for lg in x_stores:
    tc = Counter(t for o in xoffs if int(o["logon"]) == lg for t in {f"{t} ({p})" for t, p, _ in o["xtypes"]})
    vals = [STORES[str(lg)]["name"], lg, f"=COUNTIFS({xr('A')},B{r})", f"=COUNTIFS({xr('A')},B{r},{xr('R')},\"Net $0*\")",
            f"=SUMIFS({xr('J')},{xr('A')},B{r})", f"=SUMIFS({xr('L')},{xr('A')},B{r})", f"=SUMIFS({xr('P')},{xr('A')},B{r})",
            f"=F{r}+G{r}", f"=COUNTIFS({xr('A')},B{r},{xr('T')},\"?*\")", ", ".join(f"{k} ×{v}" for k, v in tc.most_common())]
    for j, v in enumerate(vals, 1):
        c = XO.cell(r, j, v); c.font = F(); c.border = BOX
        if j in (5, 6, 7, 8): c.number_format = USD
    r += 1
XO.cell(r, 1, "TOTAL")
for col in "CDEFGHI":
    XO[f"{col}{r}"] = f"=SUM({col}{XSH + 1}:{col}{r - 1})" if x_stores else 0
for j in range(1, 11):
    c = XO.cell(r, j); c.fill = fill(GOLD); c.font = F(bold=True); c.border = BOX
    if j in (5, 6, 7, 8): c.number_format = USD
XO_TOT = r
header(XO, XD0, ["Logon #", "Store", "RO #", "Close Date", "Customer #", "Customer Name (RO)", "Advisor #", "Advisor Name",
                 "Policy Labor Type", "Policy Parts Sale", "Policy Labor Sale", "Policy Total Sale", "Policy Cost",
                 "Credit Labor Type(s)", "Credit Pay Type", "Other-Type Credit", "Net After Credit", "Pattern",
                 "Same-Type Offset Too", "Employee Flag (Flagged ROs)", "Confidence"], 35.05)
for i, o in enumerate(xoffs, XD1):
    vals = [int(o["logon"]), STORES[o["logon"]]["name"], o["ro"], o["closedate"], o["custno"], o["custname"],
            o["advisor_no"], o["advisor_name"], " + ".join(o["types"]), round(o["ps"], 2), round(o["ls"], 2),
            f"=J{i}+K{i}", round(o["cost"], 2),
            ", ".join(f"{t} ({v:,.2f})" for t, _, v in o["xtypes"]), "/".join(sorted({p for _, p, _ in o["xtypes"]})),
            round(o["xcredit"], 2), f"=L{i}+P{i}", o["xoffset"], "Yes" if o["offset"] else None,
            o.get("flag") if o in inscope else None, o.get("conf") if o in inscope else None]
    for j, v in enumerate(vals, 1):
        c = XO.cell(i, j, v); c.border = BOX
        c.font = F() if j in (12, 17) else F(color="0000FF")
        if j == 4: c.number_format = "m/d/yyyy"
        if j in (10, 11, 12, 13, 16, 17): c.number_format = USD
        if j == 18 and o["xoffset"].startswith("Net $0"): c.fill = fill(AMBER); c.font = F(color=AMBER_FONT)
        if j in (20, 21) and o in inscope: c.fill = fill(RED) if o["sort"] in (1, 2) else fill(AMBER)
if not xoffs:
    XO.cell(XD1, 1, "None this month." if A.qlik_other else "Not run (no --qlik-other file).").font = F(italic=True)
else:
    XO.auto_filter.ref = f"A{XD0}:U{XD2}"
widths(XO, dict(A=8, B=22, C=10, D=11, E=11, F=28, G=10, H=24, I=11, J=13, K=13, L=13, M=12, N=30, O=14, P=13, Q=13, R=40, S=11, T=40, U=12))

r = S.max_row + 2
S.cell(r, 1, "Cross-type offsets (policy parts charged, credit on a different labor type)").font = F(bold=True)
S.cell(r + 1, 1, "ROs with the pattern / of which net $0").font = F()
S.cell(r + 1, 3, f"='Cross-Type Offsets'!C{XO_TOT}").font = F(); S.cell(r + 1, 4, f"='Cross-Type Offsets'!D{XO_TOT}").font = F()
S.cell(r + 2, 1, "Policy total sale / other-type credit").font = F()
S.cell(r + 2, 4, f"='Cross-Type Offsets'!F{XO_TOT}"); S.cell(r + 2, 5, f"='Cross-Type Offsets'!G{XO_TOT}")
S.cell(r + 3, 1, "Of these, ROs also on Flagged ROs (employee / advisor)").font = F()
S.cell(r + 3, 3, f"='Cross-Type Offsets'!I{XO_TOT}").font = F()
for a in (S.cell(r + 2, 4), S.cell(r + 2, 5)): a.number_format = USD; a.font = F()
S.cell(r + 1, 6, "Details on the Cross-Type Offsets tab; employee ROs marked on Flagged ROs columns W–X." if A.qlik_other
       else "Not run this month (no --qlik-other file).").font = F(size=9, italic=True)

# Not in 71034 Scope
NS["A1"] = f"Hits at stores with no 71034 activity in {MON_NAME} ({'/'.join(sorted(OOS_TYPES))} there does not post to 71034). Listed for awareness only."
NS["A1"].font = F(size=9, italic=True)
header(NS, 3, ["Logon #", "Store", "RO #", "Customer #", "Customer Name", "Advisor Name", "Labor Type", "Sale", "Flag"], 23.85)
for i, o in enumerate(oos, 4):
    vals = [int(o["logon"]), STORES[o["logon"]]["name"], o["ro"], o["custno"], o["custname"], o["advisor_name"],
            " + ".join(o["types"]), round(o["ls"] + o["ps"], 2), o["flag"]]
    for j, v in enumerate(vals, 1):
        c = NS.cell(i, j, v); c.font = F(color="0000FF"); c.border = BOX
        if j == 8: c.number_format = USD
if not oos:
    NS["A4"] = "None this month."; NS["A4"].font = F(italic=True)
widths(NS, dict(A=8, B=22, C=10, D=11, E=26, F=24, G=10, H=13, I=32))

# Method & Limits
sec_txt = ", ".join(f"{'/'.join(sorted(t))} ({lg})" for lg, t in sorted(SECONDARY.items(), key=lambda x: int(x[0])))
sec_n = ", ".join(f"{'/'.join(sorted(SECONDARY[lg]))} {n}" for lg, n in sorted(sec_counts.items(), key=lambda x: int(x[0])))
oos_stores = sorted({int(o["logon"]) for o in oos})
name_verified = "All %d hits were name-verified against ADP." % id_hits
if qa["id_hits_name_mismatch"]:
    name_verified += f" {len(qa['id_hits_name_mismatch'])} number matches whose names did not agree were dropped."
method = [
    ("Purpose", f"Identify {MON_NAME} {Y} policy ROs (GL 71034) written to employees, and highlight ROs where the service advisor is also the customer."),
    ("Population", f"{POP:,} closed ROs (close date {M}/1–{M}/{M_END.day}/{Y}) carrying each store's 71034 policy labor type(s) per the 71034 Labor Type Mapping ({len(PRIMARY)} stores, as of {PT.get('as_of', '?')})."
                   + (f" Includes unconfirmed secondary types {sec_txt}; ROs carrying them: {sec_n}." if SECONDARY else " Unconfirmed secondary types excluded.")),
    ("Sources", "Qlik Closed Repair Orders app (c3efc739…): %Logon, RO_Header.ronumber/closedate/custno/name1, %serviceadvisor, ServiceAdvisor.Name, "
                f"RO_Detail.labortype/laborsale/partssale/laborcost/partscost. ADP \"Employee List – Active/Leave\" export ({ADP_N:,} rows, {ADP_LOCS} locations, CDK Employee ID #). "
                "ADP store codes mapped to Qlik logon via Store_ADP_PayrollCompanyCode."),
    ("Test 1 – Advisor = Customer", "Customer # equals the numeric part of the advisor ID on the same RO, OR customer name equals advisor name (catches advisors using a regular customer #)."),
    ("Test 2 – Employee # match", f"Customer # equals a CDK Employee ID at the SAME store (IDs repeat across stores, so store is part of the key). {name_verified}"),
    ("Test 3 – Employee name match", "RO customer name (\"LAST,FIRST…\") matched to every ADP employee \"LAST,FIRST\" group-wide (whole-word match on last name and first given name; "
                                     "JR/SR/II/III ignored). Same-store = Likely; other store = Review (distinctive name) or Low (common name: surname in the ~200 most frequent US surnames). "
                                     "Customer sharing only the advisor's last name = Review (possible relative)."),
    ("Dollar basis", "Policy Sale = labor + parts sale on the policy labor-type lines only. Policy Cost = labor + parts cost on the same lines. Several stores (notably Toyota Henderson ISPT) "
                     "sell policy at $0, so cost is the better exposure measure. MLS/misc and tax are not attributable by labor type and are excluded."),
    ("Credit reclass pattern", f"Policy ROs where policy parts sale > $0 and policy labor sale < $0 on the store's policy labor type(s): parts charged to policy with a labor credit offsetting it. "
                               f"Net $0 = labor + parts within ${OFFSET_TOL:.2f}; 'Labor credit exceeds parts' = the RO nets negative; 'Partial offset' = some parts left charged. "
                               "Listed on Policy Offsets; employee/advisor ROs with the pattern are also marked on Flagged ROs (columns U–V: pattern and the labor credit). "
                               "Policy Cost is not reduced by the credit. Offsets on other labor types or on MLS lines are not visible here."),
    ("Cross-type offsets", "Policy ROs with policy parts sale > $0 where a labor type that is NOT the store's policy type, on the same RO, nets to a labor credit (any pay type: "
                           "Internal, Warranty, Customer). Net After Credit = policy total sale + that credit; Net $0 = within $1.00. Listed on Cross-Type Offsets with a store × credit-labor-type count; "
                           "employee/advisor ROs with the pattern are marked on Flagged ROs (columns W–X). Credits on other labor types also arise from comebacks, goodwill and corrections, "
                           "so this is a pattern screen, not proof. MLS lines are still not covered."),
    ("Limitations", "Current employees only — anyone who termed between the 1st of the month and the ADP export date is not caught. Name matching misses nicknames, maiden names and typos. "
                    "\"Same last name\" rows are possible relatives, not proven. Qlik amounts are RO-level and have not been tied to the GL for these specific ROs."
                    + (f" Hits at stores {' and '.join(map(str, oos_stores))} are on labor types that do not post to 71034 there — see \"Not in 71034 Scope\"." if oos_stores else "")),
    ("Legend", "Red fill = advisor wrote the RO to themselves. Amber = credit reclass pattern. Blue font = values pulled from Qlik / ADP. Gold rows = totals."),
]
for i, (a, b) in enumerate(method, 1):
    ML.cell(i, 1, a).font = F(bold=True); ML.cell(i, 1).alignment = Alignment(vertical="top")
    ML.cell(i, 2, b).font = F(); ML.cell(i, 2).alignment = Alignment(vertical="top", wrap_text=True)
    ML.row_dimensions[i].height = 15 * max(1, -(-len(b) // 125))
widths(ML, dict(A=24, B=110))

out.active = 0
out.save(A.out)

# ---------------- QA ----------------
by_flag = Counter(o["sort"] for o in inscope)
qa.update(month=A.month, population=POP, population_stores=POP_STORES, adp_rows=ADP_N, adp_locations=ADP_LOCS,
          adp_unmapped_locations=dict(unmapped_loc), flagged=len(inscope), by_flag={FLAGS[k][0]: by_flag.get(k, 0) for k in FLAGS},
          employee_id_hits=id_hits, out_of_scope=len(oos), secondary_ro_counts=sec_counts,
          cross_type_offsets=dict(run=bool(A.qlik_other), other_lines=other_lines, ros=len(xoffs),
                                  net_zero=sum(1 for o in xoffs if o["xoffset"].startswith("Net $0")),
                                  credit=round(sum(o["xcredit"] for o in xoffs), 2), on_flagged=sum(1 for o in xoffs if o in inscope),
                                  by_pay=dict(Counter(p for o in xoffs for p in {p for _, p, _ in o["xtypes"]})),
                                  by_store=dict(Counter(o["logon"] for o in xoffs))),
          offsets=dict(ros=len(offs), net_zero=sum(1 for o in offs if o["offset"].startswith("Net $0")),
                       parts=round(sum(o["ps"] for o in offs), 2), labor_credit=round(sum(o["ls"] for o in offs), 2),
                       on_flagged=sum(1 for o in offs if o in inscope),
                       by_store=dict(Counter(o["logon"] for o in offs))),
          policy_sale=round(sum(o["ls"] + o["ps"] for o in inscope), 2), policy_cost=round(sum(o["cost"] for o in inscope), 2),
          mapped_stores_without_rows=sorted(set(PRIMARY) - {o["logon"] for o in ros.values()}, key=int),
          unmapped_policy_types={f"{lg}:{lt}": n for (lg, lt), n in sorted(unmapped_types.items(), key=lambda x: (int(x[0][0]), x[0][1]))})
qa_path = A.qa or os.path.splitext(A.out)[0] + ".qa.json"
json.dump(qa, open(qa_path, "w"), indent=1, default=str)
print(f"Wrote {A.out}")
print(f"Population {POP:,} ROs / {POP_STORES} stores | ADP {ADP_N:,} rows, {ADP_LOCS} locations | flagged {len(inscope)} "
      f"(sale ${qa['policy_sale']:,.2f}, cost ${qa['policy_cost']:,.2f}) | out of scope {len(oos)}")
for k in FLAGS:
    print(f"  {k} {FLAGS[k][0]:<50} {by_flag.get(k, 0)}")
q = qa["offsets"]
x = qa["cross_type_offsets"]
print(f"Cross-type offsets: {x['ros']} ROs ({x['net_zero']} net $0) | credit ${x['credit']:,.2f} | by pay {x['by_pay']} | {x['on_flagged']} also on Flagged ROs"
      if x["run"] else "Cross-type offsets: NOT RUN (pass --qlik-other)")
print(f"Policy offsets: {q['ros']} ROs ({q['net_zero']} net $0) | parts ${q['parts']:,.2f} | labor credit ${q['labor_credit']:,.2f} | {q['on_flagged']} also on Flagged ROs")
if unmapped_loc: print("WARN ADP locations with no Qlik logon (not matchable same-store):", dict(unmapped_loc))
if qa["unmapped_policy_types"]: print("INFO labor types at mapped stores that are NOT in the mapping (excluded):", qa["unmapped_policy_types"])
if qa["id_hits_name_mismatch"]: print(f"INFO {len(qa['id_hits_name_mismatch'])} cust#=employee-ID hits dropped (names disagree) — see {qa_path}")
print(f"QA -> {qa_path}\nNEXT: run recalc.py on the output (mandatory)")
