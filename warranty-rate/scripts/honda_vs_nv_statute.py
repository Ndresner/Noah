"""Honda Henderson + Honda North: Honda SOM 6.15 vs Nevada NRS 482.36385 warranty rate comparison.
python3 -I build_honda_nv.py hh_lines.csv hn_lines.csv aftermkt.json out.xlsx qa.json
"""
import csv, sys, json, re, statistics, collections as C
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HH, HN, AFT, OUT, QA = sys.argv[1:6]
FROM = "2026-05-28"; DAYS = 180
STORES = [
 {"key": "HH", "name": "Honda Henderson", "div": "HONDA HENDERSON", "file": HH, "cp": {"C", "CE"},
  "employee": {"CXC"}, "third": {"CHC", "CHCC", "CXW"}, "policy": {"ISP"}, "warranty": {"W"}},
 {"key": "HN", "name": "Honda North", "div": "HONDA NORTH", "file": HN, "cp": {"C", "CE"},
  "employee": {"CEMP"}, "third": {"CH", "CHC", "CGWW"}, "policy": {"IP"}, "warranty": {"W"}},
]

# aftermarket (source 200) repair parts by RO, per division
aft = json.load(open(AFT))
CONSUM = re.compile(r"\bBG\b|OIL|FILTER|\bOF\b|\bAF\b|\bCF\b|DRAIN|GASKET|WASHER|MOA|FLUSH|PROTECT|TOTE|FLUID|ADDITIVE|COOL|\bRAD\b|REFRIG|FREON|R134|1234|WIP|BULB|KIT|0W|5W|80W|ATF|CVT|DEF|SHOP|SUPPL|CLEAN|TREAT|CONDITION|LUBE|GREASE|SPRAY|DYE|^\S{0,3}$")
aft_rep = {}
for row in aft["data"]["data"]:
    div = row[0]; cell = row[2]
    for x in (cell or "").split("~"):
        if not x: continue
        ro, amt, desc = x.split("|", 2)
        rep = [p.strip() for p in desc.split("/") if p.strip() and not CONSUM.search(p.strip().upper())]
        if rep and ro.isdigit(): aft_rep[(div, int(ro))] = ", ".join(rep)[:60]

Q, X, R = "Qualifies", "Excluded", "Review"
MAINT = ["B1", "B12", "B13", "B15", "B16", "B17", "B123", "FA", "FA2", "OC", "CF", "AF", "2", "FC", "FC2", "5", "BF", "3AT", "3AT9", "3AT10",
         "3CVT", "3MT", "3AWD", "6", "PF", "FI", "ACR", "ACS", "4SP4", "4SP6", "4VA4", "4VA6", "4TB4", "4TB6", "4TB6DC", "5IMA", "BS", "BI",
         "DB", "B14", "B124", "B134", "B1234", "B16"]
OPMAP = {op: (X, "Scheduled / Maintenance Minder / fluid service") for op in MAINT}
OPMAP.update({
 "10": (Q, "General repair"), "GENERAL": (Q, "General repair"), "GEN": (Q, "General repair"), "10W": (Q, "General repair"),
 "DIAG": (Q, "Diagnosis"), "CEL": (Q, "Warning light diagnosis"), "LEAKDIAG": (Q, "Leak diagnosis"), "BTEST": (Q, "Battery/starting diagnosis"),
 "BT": (Q, "Battery/starting diagnosis"), "D": (R, "Declined-line diagnosis"),
 "BAT": (Q, "Battery replacement"), "KEY": (Q, "Key / fob"), "WP": (Q, "Water pump repair"), "WP15T": (Q, "Water pump repair"),
 "R134": (Q, "A/C evacuate & recharge"), "R1234": (Q, "A/C evacuate & recharge"), "AC": (Q, "A/C evacuate & recharge"),
 "SPD": (Q, "Splash shield repair"), "SPM": (Q, "Splash shield repair"), "SOP": (Q, "Special-order part repair"),
 "A": (Q, "Alignment"), "FB": (Q, "Brake pads/rotors"), "RB": (Q, "Brake pads/rotors"), "RBE": (Q, "Brake pads/rotors"),
 "RBD": (Q, "Brake shoes"), "RBR23": (Q, "Brake pads/rotors"), "RFR": (Q, "Rotor resurfacing"), "RRR": (Q, "Rotor resurfacing"),
 "B": (X, "Tires - balance"), "T1": (X, "Tires"), "T2": (X, "Tires"), "T3": (X, "Tires"), "T4": (X, "Tires"), "1": (X, "Tires - rotation"),
 "W": (X, "Wipers"), "WR": (X, "Wipers"), "WB": (X, "Wipers"), "BULB": (X, "Bulbs"), "S": (X, "Smog inspection"), "DETAIL": (X, "Detail / cosmetic"),
})
KW = [
 (r"\bTIRES?\b|BALANCE|ROTATE", X, "Tires (keyword)"), (r"WIPER", X, "Wipers (keyword)"), (r"\bBULBS?\b|\bFUSES?\b", X, "Bulb/fuse (keyword)"),
 (r"OIL AND FILTER|OIL & FILTER|OIL CHANGE|MAINTENANCE MINDER", X, "Maintenance (keyword)"), (r"SPARK PLUG", X, "Spark plugs (keyword)"),
 (r"CABIN|AIR FILTER", X, "Filter (keyword)"), (r"FLUSH|FLUID EXCHANGE|DRAIN AND FILL|FLUID SERVICE", X, "Fluid service (keyword)"),
 (r"ACCESSOR|HITCH|ROOF RACK|RUNNING BOARD|TINT", X, "Accessory (keyword)"), (r"DETAIL|PAINT|BUFF|HEADLIGHT RESTOR", X, "Cosmetic (keyword)"),
 (r"SMOG|EMISSION TEST|PRE-?PURCHASE|APPRAIS", X, "Inspection (keyword)"), (r"SUBLET|WINDSHIELD (REPLACE|REPAIR)|\bGLASS\b", X, "Sublet/glass (keyword)"),
 (r"TIMING BELT", X, "Timing belt (Maintenance Minder)"), (r"ALIGN", Q, "Alignment (keyword)"),
 (r"BRAKE PAD|PADS|ROTOR", Q, "Brake pads/rotors (keyword)"),
]
ALIGN = lambda r: r["opcode"] in ("A",) or "Alignment" in r["why"]
BRAKE = lambda r: r["opcode"] in ("FB", "RB", "RBE", "RBD", "RBR23", "RFR", "RRR") or "Brake pads" in r["why"]
GENERIC = {"10", "GENERAL", "GEN", "10W", "DIAG", "CEL", "D", "SOP"}

def classify(r, st):
    if r["lt"] not in st["cp"]: return X, f"Not customer pay ({r['lt']})"
    if r["opcode"] in OPMAP: cls, why = OPMAP[r["opcode"]]
    else: cls, why = Q, "Unmapped op code - treated as repair (check)"
    if r["opcode"] in GENERIC or r["opcode"] not in OPMAP:
        for pat, c, w in KW:
            if re.search(pat, r["desc"].upper()): cls, why = c, w; break
    if cls != X and r["hrs"] <= 0: cls, why = X, "No sold hours"
    return cls, why

EXCL_CUST = re.compile(r"COLLISION|BODY SHOP|FINDLAY|HONDA OF|RENT A CAR|ENTERPRISE|HERTZ|AVIS|BUDGET")
def tot(ql):
    H = sum(x["hrs"] for x in ql); D = sum(x["labor"] for x in ql)
    PS = sum(x["parts"] for x in ql if x["pcost"] > 0); PC = sum(x["pcost"] for x in ql if x["pcost"] > 0)
    return H, D, PS, PC

def run_store(st):
    L = list(csv.DictReader(open(st["file"])))
    for r in L:
        r["ro"] = int(r["ro"])
        for k in ("hrs", "labor", "parts", "pcost"): r[k] = float(r[k])
        r["cls"], r["why"] = classify(r, st)
        r["rate"] = r["labor"] / r["hrs"] if r["hrs"] > 0 else 0
    ROS = C.OrderedDict()
    for r in sorted(L, key=lambda x: x["ro"]): ROS.setdefault(r["ro"], []).append(r)
    cpl = [x for x in L if x["lt"] in st["cp"]]
    door = C.Counter(round(x["rate"]) for x in cpl if x["opcode"] in ("10", "GENERAL", "GEN") and x["hrs"] > 0).most_common(1)[0][0]

    def common(lines):
        h = lines[0]; lts = {x["lt"] for x in lines}
        if not any(x["lt"] in st["cp"] for x in lines): return "No customer-pay labor"
        if lts & st["policy"]: return "Goodwill/policy on RO"
        if lts & st["third"]: return "Extended service contract / third party on RO"
        if lts & st["employee"]: return "Employee RO"
        if EXCL_CUST.search(h["name"].upper()): return "Excluded customer (dealer/body shop/rental)"
        return None

    def honda_ro(lines):
        rsn = common(lines); h = lines[0]
        if rsn is None and h["make"] != "HOND": rsn = f"Non-Honda vehicle ({h['make']})"
        if rsn is None and (st["div"], h["ro"]) in aft_rep: rsn = "Aftermarket repair parts on RO"
        ql = [x for x in lines if x["lt"] in st["cp"] and x["cls"] == Q and x["parts"] > 0 and not (x["opcode"] == "DIAG" and x["parts"] <= 0)] if rsn is None else []
        if rsn is None and not ql: rsn = "No qualifying lines (needs parts; no maintenance, diag-only, tires, flushes)"
        return rsn, ql

    def nv_ro(lines, align="with_repair", brakes=False, drop_disc=False):
        rsn = common(lines)
        if rsn is None and lines[0]["make"] != "HOND": rsn = f"Non-Honda vehicle ({lines[0]['make']})"
        cp = [x for x in lines if x["lt"] in st["cp"]]
        core = [x for x in cp if x["cls"] == Q and not ALIGN(x) and not BRAKE(x)] if rsn is None else []
        ql = list(core)
        if rsn is None:
            al = [x for x in cp if x["cls"] == Q and ALIGN(x)]
            if align == "all" or (align == "with_repair" and core): ql += al
            if brakes: ql += [x for x in cp if x["cls"] == Q and BRAKE(x)]
        if rsn is None and not ql: rsn = "No warranty-like repair lines"
        if rsn is None and drop_disc and any(x["rate"] < 0.9 * door for x in core if x["hrs"] > 0):
            rsn = "Repair line under 90% of door rate (possible special)"; ql = []
        return rsn, ql

    def windows(fn, size, **kw):
        ev = []
        for ro, lines in ROS.items():
            if lines[0]["open"] < FROM: continue
            rsn, ql = fn(lines, **kw)
            if rsn is None: ev.append({"ro": ro, "open": lines[0]["open"], "ql": ql, **dict(zip(("H", "D", "PS", "PC"), tot(ql)))})
        out = []
        for i in range(len(ev) - size + 1):
            s = ev[i:i + size]
            if size == 100:
                d0 = date.fromisoformat(s[0]["open"])
                s = [e for e in s if (date.fromisoformat(e["open"]) - d0).days <= 89]
            H = sum(e["H"] for e in s); D = sum(e["D"] for e in s); PS = sum(e["PS"] for e in s); PC = sum(e["PC"] for e in s)
            if H <= 0: continue
            out.append({"first": s[0]["ro"], "last": s[-1]["ro"], "d0": s[0]["open"], "d1": s[-1]["open"], "n": len(s), "H": H, "D": D,
                        "elr": D / H, "PS": PS, "PC": PC, "mk": (PS / PC - 1) if PC else 0, "set": s})
        return ev, out

    res = {"st": st, "L": L, "door": door}
    ev, w = windows(honda_ro, 20)
    res["honda"] = {"ev": ev, "wins": w}
    res["nv"] = {}
    for name, kw in [("base", {}), ("all_align_brakes", dict(align="all", brakes=True)), ("no_align", dict(align="none")),
                     ("drop_specials", dict(drop_disc=True))]:
        ev, w = windows(lambda lines, **k: nv_ro(lines, **k), 100, **kw)
        res["nv"][name] = {"ev": ev, "wins": w}
    # literal statute: all customer-pay retail labor & parts (180 days, store-wide, excluding employee/policy/third-party ROs)
    lit = []
    for ro, lines in ROS.items():
        if lines[0]["open"] < FROM or common(lines) or lines[0]["make"] != "HOND": continue
        lit += [x for x in lines if x["lt"] in st["cp"] and x["hrs"] > 0]
    H, D, PS, PC = tot(lit); res["literal"] = {"elr": D / H, "mk": PS / PC - 1, "H": H}
    wt = [x for x in L if x["lt"] in st["warranty"]]
    wh, wd, wps, wpc = tot(wt)
    res["cur"] = {"lr": wd / wh, "mk_qlik": wps / wpc - 1, "mk_som": 0.40, "ann_h": wh * 365 / DAYS, "ann_pc": wpc * 365 / DAYS}
    res["ROS"] = ROS; res["honda_ro"] = honda_ro; res["nv_ro"] = nv_ro
    return res

def stats(wins, k): v = [w[k] for w in wins]; return (min(v), statistics.median(v), max(v)) if v else (0, 0, 0)

RES = [run_store(s) for s in STORES]

# ---------------- workbook ----------------
NAVY = "1F3864"; GOLD = "C9A04B"
HF = Font(name="Arial", bold=True, color="FFFFFF", size=10); BF = Font(name="Arial", size=10); BB = Font(name="Arial", bold=True, size=10)
TF = Font(name="Arial", bold=True, size=14, color=NAVY); IT = Font(name="Arial", italic=True, size=9)
HFILL = PatternFill("solid", fgColor=NAVY); GFILL = PatternFill("solid", fgColor="FFF2CC"); GREEN = PatternFill("solid", fgColor="C6EFCE")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
MONEY = '"$"#,##0.00'; MONEY0 = '"$"#,##0'; PCT = '0.0%'; HRS = '#,##0.0'
wb = Workbook()
def hdr(ws, row, cols, widths=None, col0=1):
    for i, c in enumerate(cols):
        x = ws.cell(row=row, column=col0 + i, value=c); x.font = HF; x.fill = HFILL; x.border = BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    if widths:
        for i, w in enumerate(widths): ws.column_dimensions[get_column_letter(col0 + i)].width = w
def put(ws, row, vals, fmts=None, font=BF, fill=None):
    for i, v in enumerate(vals, 1):
        x = ws.cell(row=row, column=i, value=v); x.font = font; x.border = BOX
        if fmts and i <= len(fmts) and fmts[i - 1]: x.number_format = fmts[i - 1]
        if fill: x.fill = fill

ws = wb.active; ws.title = "Summary"; ws.sheet_properties.tabColor = GOLD
ws["A1"] = "Honda Henderson & Honda North - Warranty Rate: Honda SOM 6.15 vs. Nevada Statute (NRS 482.36385)"; ws["A1"].font = TF
ws["A2"] = "Qlik Closed ROs opened May 28 - Nov 23, 2026. Estimates for decision-making, not a filing. Current-rate figures are what the DMS booked on warranty (W) lines."; ws["A2"].font = IT
hdr(ws, 4, ["Measure", "Henderson", "North", "Notes"], [64, 18, 18, 80])
rowsdef = []
for label, fn, fmt, note in [
 ("Customer-pay labor types used", lambda r: ", ".join(sorted(r["st"]["cp"])), None, "C + CE (express) at both stores."),
 ("Posted repair door rate (most common $/hr on general repair lines)", lambda r: r["door"], MONEY, None),
 ("CURRENT: warranty labor $ per sold hour (W lines)", lambda r: r["cur"]["lr"], MONEY, "Proxy for the approved Honda rate - replace with the rate on file"),
 ("CURRENT: warranty parts markup booked in DMS (W lines)", lambda r: r["cur"]["mk_qlik"], PCT, "SOM 6.2 says Honda pays dealer net + 40%. The DMS shows more - confirm which applies"),
 ("Warranty labor hours per year (annualized)", lambda r: r["cur"]["ann_h"], "#,##0", None),
 ("Warranty parts cost per year (annualized)", lambda r: r["cur"]["ann_pc"], MONEY0, None),
 ("HONDA 6.15: labor rate - lowest 20-RO window", lambda r: stats(r["honda"]["wins"], "elr")[0], MONEY, "20 consecutive qualifying CP ROs; parts required; maintenance, diag-only, flushes, tires, aftermarket excluded"),
 ("HONDA 6.15: labor rate - median 20-RO window", lambda r: stats(r["honda"]["wins"], "elr")[1], MONEY, "Dealer picks the window (from 50+ sequential ROs)"),
 ("HONDA 6.15: labor rate - best 20-RO window", lambda r: stats(r["honda"]["wins"], "elr")[2], MONEY, "Honda caps the request at the calculated rate; DPSM checks 'reasonable and consistent with local trade practices'"),
 ("NEVADA (literal): all customer-pay retail labor $ / hour", lambda r: r["literal"]["elr"], MONEY, "Statute text: not less than what retail customers are charged. Broadest reading; includes maintenance and tires"),
 ("NEVADA (literal): all customer-pay retail parts markup", lambda r: r["literal"]["mk"], PCT, "Statute: parts at not less than the dealer's normal retail markup"),
 ("NEVADA (warranty-like, 100 ROs): labor rate - median window", lambda r: stats(r["nv"]["base"]["wins"], "elr")[1], MONEY, "Arizona-style method as a defensible proxy; alignments only with a repair; brake pads/rotors out"),
 ("NEVADA (warranty-like, 100 ROs): labor rate - best window", lambda r: stats(r["nv"]["base"]["wins"], "elr")[2], MONEY, None),
 ("NEVADA (warranty-like, 100 ROs): parts markup - median window", lambda r: stats(r["nv"]["base"]["wins"], "mk")[1], PCT, None),
 ("NEVADA (warranty-like, 100 ROs): parts markup - best window", lambda r: stats(r["nv"]["base"]["wins"], "mk")[2], PCT, None),
]:
    rowsdef.append(label)
    rr = 4 + len(rowsdef)
    put(ws, rr, [label, fn(RES[0]), fn(RES[1]), note], [None, fmt, fmt, None])
idx = {l: 4 + i + 1 for i, l in enumerate(rowsdef)}
cl = idx["CURRENT: warranty labor $ per sold hour (W lines)"]; cm = idx["CURRENT: warranty parts markup booked in DMS (W lines)"]
ah = idx["Warranty labor hours per year (annualized)"]; ap = idx["Warranty parts cost per year (annualized)"]
r0 = 5 + len(rowsdef) + 1
ws.cell(row=r0, column=1, value="Annual impact vs. current warranty pay").font = BB
hdr(ws, r0 + 1, ["Scenario", "Henderson labor rate", "Henderson $/yr", "North labor rate", "North $/yr", "Basis"])
for c, w in zip("EF", (16, 60)): ws.column_dimensions[c].width = w
scen = [
 ("Honda 6.15 - median window (labor only)", lambda r: stats(r["honda"]["wins"], "elr")[1], None),
 ("Honda 6.15 - best window (labor only)", lambda r: stats(r["honda"]["wins"], "elr")[2], None),
 ("Nevada warranty-like - median labor + median parts", lambda r: stats(r["nv"]["base"]["wins"], "elr")[1], lambda r: stats(r["nv"]["base"]["wins"], "mk")[1]),
 ("Nevada warranty-like - best labor + best parts", lambda r: stats(r["nv"]["base"]["wins"], "elr")[2], lambda r: stats(r["nv"]["base"]["wins"], "mk")[2]),
 ("Nevada literal - all retail labor + parts", lambda r: r["literal"]["elr"], lambda r: r["literal"]["mk"]),
]
for j, (nm, lf, pf) in enumerate(scen):
    rr = r0 + 2 + j
    def f(col_rate_letter, store_col):
        parts = f"+MAX(0,{pf(RES[0 if store_col=='B' else 1]):.6f}-{store_col}{cm})*{store_col}{ap}" if pf else ""
        return f"=({col_rate_letter}{rr}-{store_col}{cl})*{store_col}{ah}{parts}"
    put(ws, rr, [nm, lf(RES[0]), f("B", "B"), lf(RES[1]), f("D", "C"), "Labor gap x annual warranty hours" + (" + parts markup gap x annual warranty parts cost (only if higher than today)" if pf else "")],
        [None, MONEY, MONEY0, MONEY, MONEY0, None])
r1 = r0 + 3 + len(scen)
ws.cell(row=r1, column=1, value="Nevada warranty-like sensitivity (median labor rate)").font = BB
hdr(ws, r1 + 1, ["Scope", "Henderson", "North"])
for j, (k, nm) in enumerate([("base", "Base: alignments only with a repair, no brakes"), ("all_align_brakes", "All alignments + brake jobs"),
                             ("no_align", "No alignments"), ("drop_specials", "Base, dropping ROs with repair lines under 90% of door rate")]):
    put(ws, r1 + 2 + j, [nm, stats(RES[0]["nv"][k]["wins"], "elr")[1], stats(RES[1]["nv"][k]["wins"], "elr")[1]], [None, MONEY, MONEY])
r2 = r1 + 7
ws.cell(row=r2, column=1, value="Read this first").font = BB
for i, t in enumerate([
 "Honda SOM 6.15 itself says the state statute is the primary route and Honda's procedure is an alternative.",
 "Nevada's statute sets a floor (retail rate and retail parts markup) but no method; enforcement is a district-court action for treble damages and fees (NRS 482.36423). Counsel before approaching Honda.",
 "Honda 6.15 eligibility: iExam 12-month warranty mechanical labor GP% must be 80% or lower, and no 6.15 adjustment in the prior 12 months - not checked here.",
 "Honda 6.15 excludes labor lines without parts, so alignments, rotor resurfacing and diagnosis-only lines drop out.",
 "Qlik sale amounts are net of discounts. Advertised service specials can't be identified; the 'drop specials' sensitivity approximates them.",
 "CE (express) counts as customer pay at both stores; employee ROs (CXC / CEMP) are excluded at both stores.",
]):
    ws.cell(row=r2 + 1 + i, column=1, value="- " + t).font = BF

for R_ in RES:
    st = R_["st"]; k = st["key"]
    # windows
    for meth, lab, data in [("Honda 6.15", "Honda", R_["honda"]["wins"]), ("Nevada", "NV", R_["nv"]["base"]["wins"])]:
        w = wb.create_sheet(f"{k} {lab} Windows"); w.sheet_properties.tabColor = NAVY
        hdr(w, 1, ["First RO", "Last RO", "From", "To", "ROs", "Hours", "Labor $", "Labor Rate", "Parts Sale", "Parts Cost", "Parts Markup"],
            [11, 11, 11, 11, 7, 10, 12, 11, 12, 12, 11]); w.freeze_panes = "A2"
        best = max(data, key=lambda x: x["elr"])
        for i, x in enumerate(data):
            rr = 2 + i
            put(w, rr, [x["first"], x["last"], x["d0"], x["d1"], x["n"], round(x["H"], 2), round(x["D"], 2), round(x["elr"], 2), round(x["PS"], 2), round(x["PC"], 2), round(x["mk"], 4)],
                [None] * 5 + [HRS, MONEY, MONEY, MONEY, MONEY, PCT], fill=GREEN if x is best else None)
        w2 = wb.create_sheet(f"{k} {lab} Best Set"); w2.sheet_properties.tabColor = GOLD
        w2["A1"] = f"{st['name']} - {meth} best window, ROs {best['first']}-{best['last']} ({best['d0']} to {best['d1']})"; w2["A1"].font = BB
        hdr(w2, 3, ["#", "RO Number", "Opened", "Qualified Ops", "Labor Hours", "Labor $", "Parts Sale", "Parts Cost"], [5, 11, 11, 40, 11, 12, 12, 12]); w2.freeze_panes = "A4"
        for i, e in enumerate(best["set"]):
            put(w2, 4 + i, [i + 1, e["ro"], e["open"], "; ".join(sorted({x["opcode"] for x in e["ql"]}))[:50], e["H"], e["D"], e["PS"], e["PC"]], [None] * 4 + [HRS, MONEY, MONEY, MONEY])
        t = 4 + len(best["set"])
        put(w2, t, ["", "Total", "", "", f"=SUM(E4:E{t-1})", f"=SUM(F4:F{t-1})", f"=SUM(G4:G{t-1})", f"=SUM(H4:H{t-1})"], [None] * 4 + [HRS, MONEY, MONEY, MONEY], font=BB, fill=GFILL)
        put(w2, t + 1, ["", "Labor rate", "", "", "", f"=F{t}/E{t}"], [None] * 5 + [MONEY], font=BB)
        put(w2, t + 2, ["", "Parts markup", "", "", "", f"=IF(H{t}=0,0,G{t}/H{t}-1)"], [None] * 5 + [PCT], font=BB)
    # op code map
    w3 = wb.create_sheet(f"{k} Op Code Map"); w3.sheet_properties.tabColor = NAVY
    hdr(w3, 1, ["Op Code", "Sample Description", "Class", "Reason", "Lines", "Hours", "Labor $", "$/Hr", "% Lines w/ Parts"], [10, 50, 11, 46, 8, 10, 12, 9, 10]); w3.freeze_panes = "A2"
    agg = C.defaultdict(lambda: [0, 0.0, 0.0, 0, C.Counter(), C.Counter()])
    for x in R_["L"]:
        if x["lt"] not in st["cp"]: continue
        a = agg[x["opcode"]]; a[0] += 1; a[1] += x["hrs"]; a[2] += x["labor"]; a[3] += x["parts"] > 0; a[4][x["desc"][:60]] += 1; a[5][(x["cls"], x["why"])] += 1
    for i, (op, a) in enumerate(sorted(agg.items(), key=lambda z: -z[1][1])):
        (c_, why), _ = a[5].most_common(1)[0]
        put(w3, 2 + i, [op, a[4].most_common(1)[0][0], c_, why + (" (mixed by keyword)" if len(a[5]) > 1 else ""), a[0], round(a[1], 2), round(a[2], 2), round(a[2] / a[1], 2) if a[1] else 0, a[3] / a[0]],
            [None] * 4 + [None, HRS, MONEY, MONEY, PCT])

w = wb.create_sheet("Method & Limits"); w.sheet_properties.tabColor = NAVY; w.column_dimensions["A"].width = 160
for i, (t, b) in enumerate([
 ("Sources", True),
 ("Qlik Closed ROs (Division HONDA HENDERSON / HONDA NORTH, RO open May 28 - Nov 23, 2026) and Closed Parts Tickets (source 200 = aftermarket / general).", False),
 ("2026 Honda Service Operations Manual sections 6.1, 6.2, 6.15; Honda forms AWL 3068-26804, MWL 4990 8209, AWL 3069 4735 (rev 10/16/2025); NRS 482.36385, 482.36423, 482.36425 (rev 9/9/2026).", False),
 ("Honda 6.15 method", True),
 ("20 consecutive qualifying customer-pay ROs (dealer-chosen window from 50+ sequential ROs). Rate = labor $ / hours on qualifying lines. Qualifying line: customer-pay, repair with parts, Honda vehicle. Excluded: advertised specials, non-Honda parts, fluid flushes, fleet discounts, lines without parts, diagnosis only, tires, insurance, Owner's Manual / Maintenance Minder / dealer-recommended maintenance, extended service contracts. Brake pad/rotor replacement (with parts) is not excluded by 6.15 and is included.", False),
 ("Nevada method", True),
 ("Literal: all customer-pay retail labor and parts on Honda vehicles across 180 days (excluding employee, policy, third-party and dealer/body-shop ROs). Warranty-like: 100 sequential Honda-vehicle ROs within 90 days, Arizona-style exclusions; alignments only with another warranty-like repair; brake pads/rotors excluded as wear items; diagnosis included.", False),
 ("Limits", True),
 ("Current warranty rate and parts markup are what the DMS booked on W lines, not Honda's approval letter. SOM 6.2 states parts at dealer net + 40%; the DMS shows higher booked markups, so the parts gap may be larger than shown.", False),
 ("Discounts are netted in Qlik. Advertised specials, fleet discounts and insurance-paid repairs can't be identified directly. Aftermarket-part screening is at RO level. Unmapped Honda labor op codes are treated as repairs (listed on the Op Code Map tabs).", False),
 ("CE (express) included as customer pay at both stores (confirmed by Noah).", False),
]):
    c = w.cell(row=1 + i, column=1, value=t); c.font = BB if b else BF; c.alignment = Alignment(wrap_text=True)
wb.save(OUT)

qa = {}
for R_ in RES:
    k = R_["st"]["key"]
    qa[k] = {"door": R_["door"], "cur": R_["cur"], "literal": R_["literal"],
             "honda": {"n": len(R_["honda"]["ev"]), "elr": stats(R_["honda"]["wins"], "elr")},
             "nv": {kk: {"n": len(v["ev"]), "elr": stats(v["wins"], "elr"), "mk": stats(v["wins"], "mk")} for kk, v in R_["nv"].items()},
             "unmapped": sorted({x["opcode"] for x in R_["L"] if x["lt"] in R_["st"]["cp"] and x["why"].startswith("Unmapped")})[:60]}
json.dump(qa, open(QA, "w"), indent=1, default=str); print(json.dumps(qa, indent=1, default=str))
