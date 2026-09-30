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
FR = out.create_sheet("Flagged ROs"); NS = out.create_sheet("Not in 71034 Scope"); ML = out.create_sheet("Method & Limits")

# Flagged ROs
cols = ["Sort", "Flag", "Confidence", "Logon #", "Store", "RO #", "Close Date", "Customer #", "Customer Name (RO)",
        "Advisor #", "Advisor Name", "Policy Labor Type", "Policy Labor Sale", "Policy Parts Sale", "Policy Total Sale",
        "Policy Cost", "Matched Employee (ADP)", "Employee Store (ADP)", "CDK Employee ID", "Notes"]
header(FR, 1, cols, 35.05)
for i, o in enumerate(inscope, 2):
    emp = o["emp"]
    vals = [o["sort"], o["flag"], o["conf"], int(o["logon"]), STORES[o["logon"]]["name"], o["ro"], o["closedate"],
            o["custno"], o["custname"], o["advisor_no"], o["advisor_name"], " + ".join(o["types"]),
            round(o["ls"], 2), round(o["ps"], 2), f"=M{i}+N{i}", round(o["cost"], 2),
            emp["full"] if emp else None, emp["loc"].title() if emp else None, emp["id"] if emp else None, o["note"] or None]
    red = o["sort"] in (1, 2)
    for j, v in enumerate(vals, 1):
        c = FR.cell(i, j, v); c.border = BOX
        c.font = F() if j == 15 else F(color="0000FF")
        if red: c.fill = fill(RED)
        if j == 7: c.number_format = "m/d/yyyy"
        if j in (13, 14, 15, 16): c.number_format = USD
T = len(inscope) + 2; L = T - 1
FR.cell(T, 2, "TOTAL"); FR.cell(T, 6, f"=COUNTA(F2:F{L})")
for col in "MNOP":
    FR[f"{col}{T}"] = f"=SUM({col}2:{col}{L})"
for j in range(1, 21):
    c = FR.cell(T, j); c.fill = fill(GOLD); c.border = BOX; c.font = F(bold=True)
    if j in (13, 14, 15, 16): c.number_format = USD
FR.freeze_panes = "G2"; FR.auto_filter.ref = f"A1:T{L}"
widths(FR, dict(A=6, B=34, C=11, D=8, E=20, F=10, G=11, H=13, I=30, J=10, K=24, L=10, M=12, N=13, O=13, P=11, Q=24, R=26, S=12, T=30))

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
    ("Limitations", "Current employees only — anyone who termed between the 1st of the month and the ADP export date is not caught. Name matching misses nicknames, maiden names and typos. "
                    "\"Same last name\" rows are possible relatives, not proven. Qlik amounts are RO-level and have not been tied to the GL for these specific ROs."
                    + (f" Hits at stores {' and '.join(map(str, oos_stores))} are on labor types that do not post to 71034 there — see \"Not in 71034 Scope\"." if oos_stores else "")),
    ("Legend", "Red fill = advisor wrote the RO to themselves. Blue font = values pulled from Qlik / ADP. Gold rows = totals."),
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
if unmapped_loc: print("WARN ADP locations with no Qlik logon (not matchable same-store):", dict(unmapped_loc))
if qa["unmapped_policy_types"]: print("INFO labor types at mapped stores that are NOT in the mapping (excluded):", qa["unmapped_policy_types"])
if qa["id_hits_name_mismatch"]: print(f"INFO {len(qa['id_hits_name_mismatch'])} cust#=employee-ID hits dropped (names disagree) — see {qa_path}")
print(f"QA -> {qa_path}\nNEXT: run recalc.py on the output (mandatory)")
