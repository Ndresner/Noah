"""Toyota Flagstaff: Toyota Policy 7.4 (Option B) vs Arizona statutory (ARS 28-4451 E-F) warranty rate comparison.
python3 -I build_compare.py lines180.csv q_aftermkt.json out.xlsx qa.json
"""
import csv, sys, json, re, statistics, collections as C
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC, AFT, OUT, QA = sys.argv[1:5]
SUBMIT = "2026-11-24"
AZ_FROM = "2026-05-28"     # 180 days before submission
TY_FROM = "2026-07-27"     # Toyota: most recent 120 days
DOOR = 190.0               # posted repair rate seen on GR/DIAG lines (most lines $190; some $199.26)
L = list(csv.DictReader(open(SRC)))
for r in L:
    r["ro"] = int(r["ro"])
    for k in ("hrs", "labor", "parts", "pcost"): r[k] = float(r[k])

# aftermarket (source 200) parts by RO, split consumables vs repair parts
aft = json.load(open(AFT))
cell = aft["data"]["data"][0][1]
rows = cell if isinstance(cell, str) else aft["data"]["text"][0][1]
CONSUM = re.compile(r"\bBG\b|OIL|FILTER|\bOF\b|\bAF\b|\bCF\b|\bCAF\b|DRAIN PLUG|GASKET|MOA|FLUSH|PROTECT|TOTE|FLUID|ADDITIVE|COOL|\bRAD\b|REFRIG|WIP|BULB|KIT|PLATINUM|0W|5W|80W|W$|^P$|^S KIT$|^530$|^020$|^08$|^44K$|^RA$|^W$|^CAP$|^$")
aft_repair = {}
for x in rows.split("~"):
    if not x: continue
    ro, amt, desc = x.split("|", 2)
    parts = [p.strip() for p in desc.split("/")]
    rep = [p for p in parts if not CONSUM.search(p.upper())]
    if rep and ro.isdigit(): aft_repair[int(ro)] = ", ".join(rep)[:60]

# ---------- line classification ----------
Q, X, R = "Qualifies", "Excluded", "Review"
OPMAP = {
 "GR": (Q, "General repair"), "DIAG": (Q, "Diagnosis"), "CEL": (Q, "Warning light diagnosis"), "BRC": (Q, "Brake concern diagnosis"),
 "ET": (Q, "Exterior trim concern"), "NO-OP": (Q, "Free-text repair line"), "MISC": (R, "Miscellaneous"),
 "4WA": (Q, "Alignment"), "4WA+": (Q, "Alignment"), "BSR": (Q, "Brake pads/rotors"), "BSF": (Q, "Brake pads/rotors"),
 "BAT": (Q, "Battery replacement"),
 "GREEN": (X, "Maintenance - oil/rotation"), "GREENP": (X, "Maintenance - oil/rotation"), "GREEN08": (X, "Maintenance - oil/rotation"),
 "RED": (X, "Maintenance - oil"), "BLUE": (X, "Maintenance"), "GSI": (X, "Maintenance - intermediate"), "MOA": (X, "Oil additive"),
 "BGBF": (X, "Brake fluid service"), "BGCFL": (X, "Coolant service"), "HCDR": (X, "Coolant service"), "LCDR": (X, "Coolant service"),
 "BGTFWS": (X, "Transmission fluid service"), "BFTFCVT": (X, "Transmission fluid service"), "TDF": (X, "Transmission fluid service"),
 "TUN": (X, "Spark plugs"), "AIR": (X, "Air filter"), "CABIN": (X, "Cabin filter"), "FILTER": (X, "Filter"), "HBF": (X, "Hybrid battery fan service"),
 "BAL": (X, "Tires - balance"), "BUY3": (X, "Tires - promotion"), "4T": (X, "Tires"), "3T": (X, "Tires"), "2T": (X, "Tires"), "1T": (X, "Tires"), "TR": (X, "Tires - repair"),
 "WIPER": (X, "Wipers"), "BULB": (X, "Bulbs"),
 "4WD": (R, "4WD fluid service"), "AWD": (R, "AWD fluid service"), "BGDSR": (R, "Differential fluid service"), "BGPSF": (R, "Power steering fluid service"),
 "BGFIS": (R, "Fuel induction service"), "HTB": (R, "Throttle body service"), "BGBS": (R, "Battery terminal service"),
}
KW = [
 (r"\bTIRES?\b|BALANCE|ROTAT", X, "Tires (keyword)"), (r"WIPER|REPLACE INSERTS?$", X, "Wipers (keyword)"),
 (r"ADDITIVE|ADDATIVE", X, "Oil additive (keyword)"), (r"\bBULBS?\b|\bFUSES?\b", X, "Bulb/fuse (keyword)"),
 (r"OIL CHANGE|OIL AND FILTER|OIL & FILTER|ENGINE OIL", X, "Oil service (keyword)"), (r"SPARK PLUG", X, "Spark plugs (keyword)"),
 (r"CABIN|AIR FILTER|\bFILTER\b", X, "Filter (keyword)"), (r"FLUSH|FLUID EXCHANGE|DRAIN AND FILL|DRAIN & FILL", X, "Fluid service (keyword)"),
 (r"ACCESSOR|HITCH|ROOF RACK|RUNNING BOARD|TINT", X, "Accessory (keyword)"), (r"DETAIL|PAINT|BUFF|HEADLIGHT RESTOR", X, "Cosmetic (keyword)"),
 (r"APPRAIS|PRE-?PURCHASE|EMISSION|SMOG|SAFETY INSPECTION|BUYERS? CHECK", X, "Inspection (keyword)"),
 (r"SUBLET|REPLACE (THE )?WINDSHIELD\b(?! COWL)|WINDSHIELD (REPLACEMENT|REPAIR)|\bGLASS\b", X, "Sublet/glass (keyword)"),
 (r"RODENT|MOUSE|MICE|NEST", R, "Rodent"), (r"\bKEY\b|FOB", Q, "Key/fob"), (r"CALIBRAT", Q, "Calibration"),
 (r"\bCLUTCH", Q, "Clutch"), (r"\bBELT\b", R, "Belt"), (r"ALIGN", Q, "Alignment (keyword)"),
 (r"BRAKE (PAD|SERVICE)|PADS|ROTOR", Q, "Brake pads/rotors (keyword)"),
]
def base(r):
    cls, why = OPMAP.get(r["opcode"], (R, "Op code not mapped"))
    if r["opcode"] in ("GR", "NO-OP", "MISC"):
        for pat, c, w in KW:
            if re.search(pat, r["desc"].upper()): cls, why = c, w; break
    if cls != X and r["hrs"] <= 0: cls, why = X, "No sold hours"
    return cls, why
ALIGN = lambda r: r["opcode"] in ("4WA", "4WA+") or "Alignment" in r["why"]
BRAKE = lambda r: r["opcode"] in ("BSR", "BSF") or "Brake pads" in r["why"]

for r in L:
    if r["lt"] != "CPT":
        r["cls"], r["why"] = X, f"Not CPT ({r['lt']})"
    else:
        r["cls"], r["why"] = base(r)
    r["rate"] = r["labor"] / r["hrs"] if r["hrs"] > 0 else 0

ROS = C.OrderedDict()
for r in sorted(L, key=lambda x: x["ro"]): ROS.setdefault(r["ro"], []).append(r)

EXCL_CUST = re.compile(r"COLLISION|BODY SHOP|FINDLAY|TOYOTA OF|RENT A CAR|ENTERPRISE|HERTZ|AVIS|BUDGET")
def ro_common(lines):
    h = lines[0]; lts = {x["lt"] for x in lines}
    if not any(x["lt"] == "CPT" for x in lines): return "No CPT labor"
    if "IPS" in lts: return "Goodwill/policy on RO (IPS)"
    if "CTEC" in lts: return "Extended warranty on RO (CTEC)"
    if "CPTE" in lts: return "Employee (CPTE)"
    if EXCL_CUST.search(h["name"].upper()): return "Excluded customer"
    return None

def toyota_ro(lines):
    h = lines[0]
    rsn = ro_common(lines)
    if rsn is None and h["make"] not in ("TOYO", "SCIO"): rsn = f"Non-Toyota ({h['make']})"
    if rsn is None and h["ro"] in aft_repair: rsn = "Aftermarket repair parts on RO (genuine-parts rule)"
    ql = [x for x in lines if x["lt"] == "CPT" and x["cls"] == Q] if rsn is None else []
    if rsn is None and not ql: rsn = "No qualified repair lines"
    return rsn, ql

def az_ro(lines, align_mode="with_repair", brakes=False, drop_disc=False):
    rsn = ro_common(lines)
    if rsn is None and lines[0]["make"] not in ("TOYO", "SCIO"): rsn = f"Non-Toyota ({lines[0]['make']})"
    cp = [x for x in lines if x["lt"] == "CPT"]
    core = [x for x in cp if x["cls"] == Q and not ALIGN(x) and not BRAKE(x)] if rsn is None else []
    ql = list(core)
    if rsn is None:
        al = [x for x in cp if x["cls"] == Q and ALIGN(x)]
        if align_mode == "all" or (align_mode == "with_repair" and core): ql += al
        if brakes: ql += [x for x in cp if x["cls"] == Q and BRAKE(x)]
    if rsn is None and not ql: rsn = "No warranty-like repair lines"
    if rsn is None and drop_disc and any(x["rate"] < 0.9 * DOOR for x in core if x["hrs"] > 0):
        rsn = "Repair line below 90% of door rate (possible promo/discount)"; ql = []
    return rsn, ql

def tot(ql):
    H = sum(x["hrs"] for x in ql); D = sum(x["labor"] for x in ql)
    PS = sum(x["parts"] for x in ql if x["pcost"] > 0); PC = sum(x["pcost"] for x in ql if x["pcost"] > 0)
    return H, D, PS, PC

# ---------- Toyota 7.4 ----------
ty = []
for ro, lines in ROS.items():
    if lines[0]["open"] < TY_FROM: continue
    rsn, ql = toyota_ro(lines)
    ty.append({"ro": ro, "open": lines[0]["open"], "rsn": rsn, "ql": ql, **dict(zip(("H", "D", "PS", "PC"), tot(ql)))})
tyq = [e for e in ty if e["rsn"] is None]
ty_starts = []
for d in sorted({e["open"] for e in ty}):
    pool = [e for e in tyq if e["open"] >= d]
    if len(pool) < 30: continue
    s = pool[:30]; H = sum(e["H"] for e in s); D = sum(e["D"] for e in s)
    ty_starts.append({"start": d, "first": s[0]["ro"], "last": s[-1]["ro"], "H": H, "D": D, "elr": D / H, "set": s})
ty_elr = [x["elr"] for x in ty_starts]
ty_best = max(ty_starts, key=lambda x: x["elr"])

# ---------- Arizona 28-4451 ----------
def az_windows(**kw):
    ev = []
    for ro, lines in ROS.items():
        if lines[0]["open"] < AZ_FROM: continue
        rsn, ql = az_ro(lines, **kw)
        if rsn is None:
            ev.append({"ro": ro, "open": lines[0]["open"], "ql": ql, **dict(zip(("H", "D", "PS", "PC"), tot(ql)))})
    wins = []
    for i in range(len(ev)):
        s = ev[i:i + 100]
        if len(s) < 100: break
        # lesser of 100 sequential ROs or 90 consecutive days
        from datetime import date
        d0 = date.fromisoformat(s[0]["open"]); d99 = date.fromisoformat(s[-1]["open"])
        if (d99 - d0).days > 89:
            s = [e for e in s if (date.fromisoformat(e["open"]) - d0).days <= 89]
        H = sum(e["H"] for e in s); D = sum(e["D"] for e in s); PS = sum(e["PS"] for e in s); PC = sum(e["PC"] for e in s)
        wins.append({"first": s[0]["ro"], "last": s[-1]["ro"], "d0": s[0]["open"], "d1": s[-1]["open"], "n": len(s),
                     "H": H, "D": D, "elr": D / H, "PS": PS, "PC": PC, "mk": (PS / PC - 1) if PC else 0, "set": s})
    return ev, wins

SCEN = [
 ("AZ base: alignments only with a repair, no brakes", dict(align_mode="with_repair", brakes=False, drop_disc=False)),
 ("AZ: all alignments + brakes (Toyota-like scope)", dict(align_mode="all", brakes=True, drop_disc=False)),
 ("AZ: no alignments, no brakes", dict(align_mode="none", brakes=False, drop_disc=False)),
 ("AZ base, excluding ROs with below-door-rate repair lines", dict(align_mode="with_repair", brakes=False, drop_disc=True)),
]
az = {}
for name, kw in SCEN:
    ev, wins = az_windows(**kw)
    el = [w["elr"] for w in wins]; mk = [w["mk"] for w in wins]
    az[name] = {"ev": ev, "wins": wins, "elr": (min(el), statistics.median(el), max(el)), "mk": (min(mk), statistics.median(mk), max(mk)),
                "best_l": max(wins, key=lambda w: w["elr"]), "best_p": max(wins, key=lambda w: w["mk"])}
BASE = SCEN[0][0]

# ---------- current warranty (proxy) and volume ----------
wt = [x for x in L if x["lt"] == "WT"]
days = 180
wt_h = sum(x["hrs"] for x in wt); wt_d = sum(x["labor"] for x in wt)
wt_ps = sum(x["parts"] for x in wt if x["pcost"] > 0); wt_pc = sum(x["pcost"] for x in wt if x["pcost"] > 0)
cur_lr = wt_d / wt_h; cur_mk = wt_ps / wt_pc - 1
ann_h = wt_h * 365 / days; ann_pc = wt_pc * 365 / days

# ---------- workbook ----------
NAVY = "1F3864"; GOLD = "C9A04B"
HF = Font(name="Arial", bold=True, color="FFFFFF", size=10); BF = Font(name="Arial", size=10); BB = Font(name="Arial", bold=True, size=10)
TF = Font(name="Arial", bold=True, size=14, color=NAVY); IT = Font(name="Arial", italic=True, size=9)
HFILL = PatternFill("solid", fgColor=NAVY); GFILL = PatternFill("solid", fgColor="FFF2CC"); GREEN = PatternFill("solid", fgColor="C6EFCE")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
MONEY = '"$"#,##0.00'; MONEY0 = '"$"#,##0'; PCT = '0.0%'; HRS = '#,##0.0'
wb = Workbook()
def header(ws, row, cols, widths):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c); x.font = HF; x.fill = HFILL; x.border = BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
def put(ws, row, vals, fmts=None, font=BF, fill=None):
    for i, v in enumerate(vals, 1):
        x = ws.cell(row=row, column=i, value=v); x.font = font; x.border = BOX
        if fmts and i <= len(fmts) and fmts[i - 1]: x.number_format = fmts[i - 1]
        if fill: x.fill = fill

# Summary
ws = wb.active; ws.title = "Summary"; ws.sheet_properties.tabColor = GOLD
ws["A1"] = "Toyota Flagstaff - Warranty Rate: Toyota Policy 7.4 vs. Arizona Statute (ARS 28-4451)"; ws["A1"].font = TF
ws["A2"] = f"Qlik Closed ROs, ROs opened May 28 - Nov 23, 2026. Assumed submission date {SUBMIT}. Estimate for decision-making, not a filing."; ws["A2"].font = IT
header(ws, 4, ["Measure", "Toyota 7.4 (Option B)", "Arizona statute (base)", "Notes"], [62, 20, 20, 70])
ws.freeze_panes = None
b = az[BASE]
rows_ = [
 ("Look-back window", "120 days (Jul 27 - Nov 23)", "180 days (May 28 - Nov 23)", None, "Toyota: Region sets start date. Arizona: dealer picks the window."),
 ("Sample", "First 30 qualified ROs", "Lesser of 100 sequential ROs / 90 days", None, "All 100-RO windows fit inside 90 days at Flagstaff's volume"),
 ("Qualified ROs available", len(tyq), len(b["ev"]), "#,##0", "After exclusions"),
 ("Labor rate - lowest possible window", min(ty_elr), b["elr"][0], MONEY, None),
 ("Labor rate - median window", statistics.median(ty_elr), b["elr"][1], MONEY, "Toyota: you can't pick, so plan on the median"),
 ("Labor rate - best window", max(ty_elr), b["elr"][2], MONEY, f"Toyota best start {ty_best['start']}; AZ best window ROs {b['best_l']['first']}-{b['best_l']['last']} ({b['best_l']['d0']} to {b['best_l']['d1']})"),
 ("Parts markup - best window", "Not part of 7.4", b["mk"][2], PCT, f"AZ best parts window ROs {b['best_p']['first']}-{b['best_p']['last']}; labor and parts may be declared from different windows (to confirm with counsel)"),
 ("Parts markup - median window", "Not part of 7.4", b["mk"][1], PCT, "Parts sale / parts cost - 1 on qualified repair lines (cost from Qlik)"),
 ("Current warranty labor $ per sold hour (proxy)", cur_lr, cur_lr, MONEY, "WT lines, 180 days. Replace with Toyota's approved rate when Tyler sends it."),
 ("Current warranty parts markup (proxy)", cur_mk, cur_mk, PCT, "WT parts sale / cost, 180 days. Replace with Toyota's approved parts reimbursement."),
 ("Warranty labor hours per year (annualized)", ann_h, ann_h, "#,##0", "From 180 days of WT sold hours"),
 ("Warranty parts cost per year (annualized)", ann_pc, ann_pc, MONEY0, "From 180 days of WT parts cost"),
]
for i, (m, t, a, f, n) in enumerate(rows_):
    put(ws, 5 + i, [m, t, a, n], [None, f, f, None])
r = 5 + len(rows_) + 1
ws.cell(row=r, column=1, value="Annual impact vs. current warranty pay (proxy rates)").font = BB
header_cells = ["Scenario", "Labor rate", "Labor $/yr vs now", "Parts markup", "Parts $/yr vs now", "Total $/yr vs now"]
for i, c in enumerate(header_cells, 1):
    x = ws.cell(row=r + 1, column=i, value=c); x.font = HF; x.fill = HFILL; x.border = BOX; x.alignment = Alignment(wrap_text=True, horizontal="center")
for c, w in zip("EF", (18, 18)): ws.column_dimensions[c].width = w
imp = [
 ("Toyota 7.4 - median start date", statistics.median(ty_elr), None),
 ("Toyota 7.4 - best start date", max(ty_elr), None),
 ("Arizona base - median window", b["elr"][1], b["mk"][1]),
 ("Arizona base - best labor window + best parts window", b["elr"][2], b["mk"][2]),
]
cl = f"$B${5+8}"; cm = f"$B${5+9}"; ah = f"$B${5+10}"; ap = f"$B${5+11}"
for j, (nm, lr, mk) in enumerate(imp):
    rr = r + 2 + j
    put(ws, rr, [nm, lr, f"=(B{rr}-{cl})*{ah}", mk if mk is not None else "n/a", f"=(D{rr}-{cm})*{ap}" if mk is not None else 0, f"=C{rr}+E{rr}"],
        [None, MONEY, MONEY0, PCT, MONEY0, MONEY0])
rr = r + 2 + len(imp) + 1
ws.cell(row=rr, column=1, value="Arizona sensitivity (labor rate: low / median / best; parts markup best)").font = BB
for i, c in enumerate(["Scenario", "Low", "Median", "Best", "Best parts markup"], 1):
    x = ws.cell(row=rr + 1, column=i, value=c); x.font = HF; x.fill = HFILL; x.border = BOX
for j, (name, _) in enumerate(SCEN):
    z = az[name]; put(ws, rr + 2 + j, [name, z["elr"][0], z["elr"][1], z["elr"][2], z["mk"][2]], [None, MONEY, MONEY, MONEY, PCT])
rr = rr + 3 + len(SCEN)
ws.cell(row=rr, column=1, value="Read this first").font = BB
for i, t in enumerate([
 "Qlik sale amounts are already net of discounts (per Noah), so discounts are reflected in both methods; Toyota's proportional-discount step can't be reproduced line by line.",
 "Arizona excludes 'manufacturer or distributor special events, specials or promotional discounts' entirely. Qlik can't tell a promo from an advisor rate change, so the last sensitivity drops any RO with a repair line under 90% of the $190 door rate.",
 "'Warranty-like' is not defined in the statute. Base case counts alignments only on ROs with another warranty-like repair and leaves out brake pads/rotors (wear items, not covered by Toyota warranty). Counsel should confirm.",
 "Current-rate figures are proxies from warranty lines in Qlik (what Toyota actually paid per hour and on parts), not the approved rate letter.",
 "Toyota method excludes ROs carrying aftermarket repair parts (parts source 200, excluding fluids/filters/consumables) under its genuine-parts rule. The statute has no genuine-parts requirement.",
 "Toyota can rebut an Arizona declaration within 30 days by showing the rate is unreasonable vs. similarly situated Arizona Toyota dealers - Prescott's rate is the obvious comparison.",
]):
    c = ws.cell(row=rr + 1 + i, column=1, value="- " + t); c.font = BF

# Toyota scenarios
w2 = wb.create_sheet("Toyota 7.4 Start Dates"); w2.sheet_properties.tabColor = NAVY
header(w2, 1, ["Start Date", "First RO", "30th RO", "Hours", "Labor $", "ELR"], [12, 11, 11, 10, 12, 11])
for i, x in enumerate(ty_starts):
    put(w2, 2 + i, [x["start"], x["first"], x["last"], x["H"], x["D"], f"=E{2+i}/D{2+i}"], [None, None, None, HRS, MONEY, MONEY],
        fill=GREEN if x is ty_best else None)

# AZ windows (base)
w3 = wb.create_sheet("AZ Windows (base)"); w3.sheet_properties.tabColor = NAVY
header(w3, 1, ["First RO", "Last RO", "From", "To", "ROs", "Hours", "Labor $", "Labor Rate", "Parts Sale", "Parts Cost", "Parts Markup"],
       [11, 11, 11, 11, 7, 10, 12, 11, 12, 12, 11])
for i, x in enumerate(b["wins"]):
    k = 2 + i
    put(w3, k, [x["first"], x["last"], x["d0"], x["d1"], x["n"], x["H"], x["D"], f"=G{k}/F{k}", x["PS"], x["PC"], f"=IF(J{k}=0,0,I{k}/J{k}-1)"],
        [None, None, None, None, None, HRS, MONEY, MONEY, MONEY, MONEY, PCT],
        fill=GREEN if x is b["best_l"] else None)

# AZ best labor window RO list
w4 = wb.create_sheet("AZ Best Labor Window"); w4.sheet_properties.tabColor = GOLD
w4["A1"] = f"Arizona submission draft - 100 sequential warranty-like customer-pay ROs {b['best_l']['first']}-{b['best_l']['last']}"; w4["A1"].font = BB
header(w4, 3, ["#", "RO Number", "Opened", "Qualified Ops", "Labor Hours", "Labor Sales", "Parts Sale", "Parts Cost"], [5, 11, 11, 40, 11, 12, 12, 12])
w4.freeze_panes = "A4"
for i, e in enumerate(b["best_l"]["set"]):
    put(w4, 4 + i, [i + 1, e["ro"], e["open"], "; ".join(sorted({x["opcode"] for x in e["ql"]}))[:50], e["H"], e["D"], e["PS"], e["PC"]],
        [None, None, None, None, HRS, MONEY, MONEY, MONEY])
n = len(b["best_l"]["set"]); t = 4 + n
put(w4, t, ["", "Total", "", "", f"=SUM(E4:E{t-1})", f"=SUM(F4:F{t-1})", f"=SUM(G4:G{t-1})", f"=SUM(H4:H{t-1})"], [None]*4 + [HRS, MONEY, MONEY, MONEY], font=BB, fill=GFILL)
put(w4, t + 1, ["", "Labor rate", "", "", "", f"=F{t}/E{t}"], [None]*5 + [MONEY], font=BB)
put(w4, t + 2, ["", "Parts markup", "", "", "", f"=G{t}/H{t}-1"], [None]*5 + [PCT], font=BB)

# Toyota best worksheet
w5 = wb.create_sheet("Toyota Best Start Worksheet"); w5.sheet_properties.tabColor = GOLD
w5["A1"] = f"Toyota 7.4 RO Worksheet draft - start date {ty_best['start']} (Region sets the date; this is the best case)"; w5["A1"].font = BB
header(w5, 3, ["#", "RO Number", "RO Date", "Qualified Ops", "Labor Hours", "Labor Amount", "ELR"], [5, 11, 11, 40, 11, 12, 11])
w5.freeze_panes = "A4"
for i, e in enumerate(ty_best["set"]):
    k = 4 + i
    put(w5, k, [i + 1, e["ro"], e["open"], "; ".join(sorted({x["opcode"] for x in e["ql"]}))[:50], e["H"], e["D"], f"=IFERROR(F{k}/E{k},0)"],
        [None, None, None, None, HRS, MONEY, MONEY])
put(w5, 34, ["", "Total", "", "", "=SUM(E4:E33)", "=SUM(F4:F33)", "=F34/E34"], [None]*4 + [HRS, MONEY, MONEY], font=BB, fill=GFILL)

# Line detail (CPT only)
w6 = wb.create_sheet("CPT Line Detail"); w6.sheet_properties.tabColor = NAVY
header(w6, 1, ["RO Number", "Opened", "Make", "Op Code", "Description", "Sold Hours", "Labor $", "$/Hr", "Parts Sale", "Parts Cost", "Line Class", "Reason",
               "In Toyota Sample?", "Toyota RO Exclusion", "In AZ Base?", "AZ RO Exclusion", "Aftermarket Repair Parts on RO"],
       [11, 11, 7, 10, 50, 10, 11, 9, 11, 11, 11, 26, 12, 34, 10, 34, 30])
ty_map = {e["ro"]: e for e in ty}; az_ql = {id(x) for e in b["ev"] for x in e["ql"]}
az_rsn = {}
for ro, lines in ROS.items():
    if lines[0]["open"] >= AZ_FROM: az_rsn[ro] = az_ro(lines, align_mode="with_repair")[0]
ty_ql = {id(x) for e in tyq for x in e["ql"]}
k = 2
for r_ in L:
    if r_["lt"] != "CPT": continue
    te = ty_map.get(r_["ro"])
    put(w6, k, [r_["ro"], r_["open"], r_["make"], r_["opcode"], r_["desc"], r_["hrs"], r_["labor"], round(r_["rate"], 2), r_["parts"], r_["pcost"], r_["cls"], r_["why"],
                "Yes" if id(r_) in ty_ql else "No", (te["rsn"] or "") if te else "Outside 120 days",
                "Yes" if id(r_) in az_ql else "No", az_rsn.get(r_["ro"]) or "", aft_repair.get(r_["ro"], "")],
        [None]*5 + [HRS, MONEY, MONEY, MONEY, MONEY])
    k += 1
w6.auto_filter.ref = f"A1:Q{k-1}"

# Method
w7 = wb.create_sheet("Method & Limits"); w7.sheet_properties.tabColor = NAVY; w7.column_dimensions["A"].width = 160
for i, (t, bold) in enumerate([
 ("Sources", True),
 ("Qlik Closed Repair Orders (Division TOYOTA FLAGSTAFF, RO open May 28 - Nov 23, 2026; 21,649 lines) and Closed Parts Tickets (parts source 200 = 'AFTERMARKET NON TO').", False),
 ("Toyota Warranty Policy 7.4 (rev 1/26) and ARS 28-4451 (azleg.gov, read Nov 24, 2026).", False),
 ("Toyota 7.4 method", True),
 ("ROs opened in the last 120 days; for each possible start date, first 30 qualified Toyota customer-pay (CPT) ROs in RO-number order. Qualified = Toyota 7.4 table (alignments and brake pads/rotors included; maintenance, tires, wear items, inspections, non-Toyota, goodwill, extended warranty, employee, body-shop/dealer customers excluded). ROs with aftermarket repair parts excluded (genuine-parts rule).", False),
 ("Arizona method (ARS 28-4451 E-F)", True),
 ("Lesser of 100 sequential customer-paid ROs for warranty-like repairs or 90 consecutive days, within 180 days of submission; dealer chooses the window. Labor rate = labor sales / labor hours; parts markup = parts sales / parts cost - 1, both on the warranty-like repair lines of the submitted ROs.", False),
 ("Excluded per F: promotional/manufacturer specials, wholesale parts, routine maintenance (fluids, filters, belts not in the course of repairs), reconditioning. Tires, wipers, bulbs, inspections and fluid services treated as not warranty-like.", False),
 ("Base case interpretation: alignments only on ROs with another warranty-like repair; brake pads/rotors excluded as wear items. Same RO-level exclusions as Toyota for goodwill, extended warranty, employee and body-shop/dealer customers. Toyota/Scion vehicles only (on-brand, as submissions require).", False),
 ("Limits", True),
 ("Discounts are already netted in Qlik sale amounts; promo ROs can't be identified directly (sensitivity provided). The CDK discount feed would let us drop true promotional ROs.", False),
 ("Current warranty rates are proxies from WT lines (labor $ / sold hour, parts sale / cost). Toyota's approved rate letter should replace them.", False),
 ("Closed ROs only. Genuine-parts and nut/bolt exclusions are approximated at RO level. Diagnosis documentation not checked. Each RO in a final submission needs a manual review.", False),
 ("Toyota may rebut an Arizona declaration within 30 days by comparison to similarly situated Arizona same-line-make dealers; a protest goes to the MVD director and an ALJ hearing.", False),
]):
    c = w7.cell(row=1 + i, column=1, value=t); c.font = BB if bold else BF; c.alignment = Alignment(wrap_text=True)

wb.save(OUT)
qa = {"ty_starts": len(ty_starts), "ty_q": len(tyq), "ty_elr": [min(ty_elr), statistics.median(ty_elr), max(ty_elr)], "ty_best": ty_best["start"],
      "ty_excl": C.Counter(e["rsn"] for e in ty if e["rsn"]).most_common(10),
      "cur_lr": cur_lr, "cur_mk": cur_mk, "ann_h": ann_h, "ann_pc": ann_pc,
      "az": {k: {"n": len(v["ev"]), "wins": len(v["wins"]), "elr": v["elr"], "mk": v["mk"],
                 "best_l": [v["best_l"]["first"], v["best_l"]["last"], v["best_l"]["d0"], v["best_l"]["d1"], v["best_l"]["n"]],
                 "best_p": [v["best_p"]["first"], v["best_p"]["last"]]} for k, v in az.items()},
      "aft_repair_ros": len(aft_repair)}
json.dump(qa, open(QA, "w"), indent=1, default=str); print(json.dumps(qa, indent=1, default=str))
