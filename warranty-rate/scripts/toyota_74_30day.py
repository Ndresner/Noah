"""Toyota Flagstaff - Warranty Labor Rate (Option B / Policy 7.4) in-house replication.
python3 -I build.py lines.csv out.xlsx qa.json
"""
import csv, sys, json, re, statistics, collections as C
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC, OUT, QA = sys.argv[1:4]
L = list(csv.DictReader(open(SRC)))
for r in L:
    r["hrs"] = float(r["hrs"]); r["labor"] = float(r["labor"]); r["parts"] = float(r["parts"]); r["ro"] = int(r["ro"])

# ---------------- classification ----------------
Q, X, R = "Qualifies", "Excluded", "Review"
OPMAP = {  # opcode -> (class, Toyota 7.4 category / reason)
 "GR": (Q, "General repair"), "DIAG": (Q, "Diagnosis (documented concern - exception)"),
 "CEL": (Q, "Warning light diagnosis (exception)"), "BRC": (Q, "Brake concern diagnosis"),
 "ET": (Q, "Exterior trim concern (all other services)"),
 "4WA": (Q, "Alignment (exception)"), "4WA+": (Q, "Alignment (exception)"),
 "BSR": (Q, "Brake pads/rotors (exception)"), "BSF": (Q, "Brake pads/rotors (exception)"),
 "BAT": (Q, "Battery replacement (not a listed exclusion)"),
 "NO-OP": (Q, "Free-text repair line (keyword-screened)"), "MISC": (R, "Miscellaneous - check RO"),
 "GREEN": (X, "Scheduled maintenance - oil/rotation"), "GREENP": (X, "Scheduled maintenance - oil/rotation"),
 "GREEN08": (X, "Scheduled maintenance - oil/rotation"), "RED": (X, "Scheduled maintenance - oil change"),
 "BLUE": (X, "Scheduled maintenance"), "GSI": (X, "Scheduled maintenance - intermediate service"),
 "MOA": (X, "Oil additive with maintenance"),
 "BGBF": (X, "Brake fluid service (listed exclusion)"), "BGCFL": (X, "Coolant service (listed exclusion)"),
 "HCDR": (X, "Coolant service (listed exclusion)"), "LCDR": (X, "Coolant service (listed exclusion)"),
 "BGTFWS": (X, "Automatic transmission service (listed exclusion)"),
 "BFTFCVT": (X, "Automatic transmission service (listed exclusion)"),
 "TDF": (X, "Transmission fluid service (listed exclusion)"),
 "TUN": (X, "Spark plugs (listed exclusion)"), "AIR": (X, "Air filter (listed exclusion)"),
 "CABIN": (X, "Cabin air filter (listed exclusion)"), "FILTER": (X, "Filter replacement (maintenance)"),
 "HBF": (X, "Hybrid battery fan service (listed exclusion)"),
 "BAL": (X, "Tire operation - balance"), "BUY3": (X, "Tire operation - replacement"),
 "4T": (X, "Tire operation - replacement"), "2T": (X, "Tire operation - replacement"),
 "1T": (X, "Tire operation - replacement"), "TR": (X, "Tire operation - repair"),
 "WIPER": (X, "Wear item - wipers"), "BULB": (X, "Wear item - bulbs"),
 # maintenance-type services that Toyota's list does not name: scenario-tested
 "4WD": (R, "Fluid service not named in 7.4 (4WD)"), "AWD": (R, "Fluid service not named in 7.4 (AWD)"),
 "BGDSR": (R, "Fluid service not named in 7.4 (differential)"),
 "BGPSF": (R, "Fluid service not named in 7.4 (power steering)"),
 "BGFIS": (R, "Fuel induction service not named in 7.4"),
 "HTB": (R, "Throttle body service not named in 7.4"), "BGBS": (R, "Battery terminal service not named in 7.4"),
}
KW = [  # free-text overrides for GR / NO-OP / MISC lines (first match wins)
 (r"\bTIRES?\b|BALANCE|ROTAT", X, "Tire operation (keyword)"),
 (r"WIPER|REPLACE INSERTS?$", X, "Wear item - wipers (keyword)"),
 (r"ADDITIVE|ADDATIVE", X, "Oil additive (keyword)"),
 (r"\bBULBS?\b|\bFUSES?\b", X, "Wear item - bulb/fuse (keyword)"),
 (r"OIL CHANGE|OIL AND FILTER|OIL & FILTER|ENGINE OIL", X, "Maintenance - oil (keyword)"),
 (r"SPARK PLUG", X, "Spark plugs (keyword)"),
 (r"CABIN|AIR FILTER|\bFILTER\b", X, "Filter (keyword)"),
 (r"FLUSH|FLUID EXCHANGE|DRAIN AND FILL|DRAIN & FILL", X, "Fluid service (keyword)"),
 (r"ACCESSOR|HITCH|ROOF RACK|RUNNING BOARD|TINT", X, "Accessory (keyword)"),
 (r"DETAIL|PAINT|BUFF|HEADLIGHT RESTOR", X, "Cosmetic (keyword)"),
 (r"APPRAIS|PRE-?PURCHASE|EMISSION|SMOG|SAFETY INSPECTION|BUYERS? CHECK", X, "Inspection (keyword)"),
 (r"SUBLET|REPLACE (THE )?WINDSHIELD\b(?! COWL)|WINDSHIELD (REPLACEMENT|REPAIR)|\bGLASS\b", X, "Sublet / glass (keyword)"),
 (r"RODENT|MOUSE|MICE|NEST", R, "Rodent - repair qualifies, cleaning only does not"),
 (r"\bKEY\b|FOB", Q, "Key / fob (exception)"),
 (r"CALIBRAT", Q, "Calibration (exception)"),
 (r"\bCLUTCH", Q, "Clutch (exception)"),
 (r"\bBELT\b", R, "Belt - timing/serpentine qualify, other belts do not"),
]
EXCL_CUSTOMER = re.compile(r"COLLISION|BODY SHOP|FINDLAY|TOYOTA OF|RENT A CAR|ENTERPRISE|HERTZ|AVIS|BUDGET")
TOYOTA_MAKES = {"TOYO", "SCIO"}
GOODWILL_TYPES = {"IPS"}          # GL 71034 service policy at Toyota Flagstaff (config/policy_labor_types.json)
THIRD_PARTY_TYPES = {"CTEC"}      # assumed Toyota Extra Care / service contract -> whole RO excluded
EMPLOYEE_TYPES = {"CPTE"}         # assumed employee customer pay -> whole RO excluded

def classify(r):
    if r["lt"] != "CPT":
        return X, f"Not true customer pay (labor type {r['lt']})"
    cls, why = OPMAP.get(r["opcode"], (R, "Op code not mapped - check RO"))
    if r["opcode"] in ("GR", "NO-OP", "MISC"):
        for pat, c, w in KW:
            if re.search(pat, r["desc"].upper()):
                cls, why = c, w; break
    if cls != X and r["hrs"] <= 0:
        return X, "No sold hours on line"
    return cls, why

for r in L:
    r["cls"], r["why"] = classify(r)

# ---------------- RO level ----------------
ROS = C.OrderedDict()
for r in sorted(L, key=lambda x: x["ro"]):
    ROS.setdefault(r["ro"], []).append(r)

def ro_eval(lines, include_review):
    h = lines[0]
    lts = {x["lt"] for x in lines}
    cp = [x for x in lines if x["lt"] == "CPT"]
    reason = None
    if h["make"] not in TOYOTA_MAKES: reason = f"Non-Toyota vehicle ({h['make']})"
    elif not cp: reason = "No true customer-pay (CPT) labor"
    elif lts & GOODWILL_TYPES: reason = "Goodwill / policy on RO (IPS)"
    elif lts & THIRD_PARTY_TYPES: reason = "Extended warranty / third party on RO (CTEC)"
    elif lts & EMPLOYEE_TYPES: reason = "Employee customer pay (CPTE)"
    elif EXCL_CUSTOMER.search(h["name"].upper()): reason = "Excluded customer (body shop / dealership / rental)"
    ok = {Q, R} if include_review else {Q}
    ql = [x for x in cp if x["cls"] in ok] if reason is None else []
    if reason is None and not ql: reason = "No qualified repair lines (maintenance / tires / other excluded work only)"
    return {
        "ro": h["ro"], "open": h["open"], "close": h["close"], "make": h["make"], "custno": h["custno"],
        "status": "Qualified" if reason is None else "Excluded", "reason": reason or "",
        "q_hrs": round(sum(x["hrs"] for x in ql), 2), "q_lab": round(sum(x["labor"] for x in ql), 2),
        "cp_lab": round(sum(x["labor"] for x in cp), 2), "cp_parts": round(sum(x["parts"] for x in cp), 2),
        "desc": "; ".join(sorted({x["opcode"] for x in ql}))[:60],
        "lts": ",".join(sorted(lts)),
    }

def scenario(include_review):
    ev = [ro_eval(v, include_review) for v in ROS.values()]
    qual = [e for e in ev if e["status"] == "Qualified"]
    dates = sorted({e["open"] for e in ev})
    rows = []
    for d in dates:
        pool = [e for e in qual if e["open"] >= d]  # RO-number order
        if len(pool) < 30:
            rows.append({"start": d, "n": len(pool), "ok": False}); continue
        s = pool[:30]
        reviewed = [e for e in ev if e["open"] >= d and e["ro"] <= s[-1]["ro"]]
        H = sum(e["q_hrs"] for e in s); D = sum(e["q_lab"] for e in s)
        rows.append({"start": d, "n": 30, "ok": True, "first": s[0]["ro"], "last": s[-1]["ro"],
                     "last_open": s[-1]["open"], "reviewed": len(reviewed), "hrs": round(H, 2),
                     "lab": round(D, 2), "elr": round(D / H, 2) if H else 0, "set": s})
    allH = sum(e["q_hrs"] for e in qual); allD = sum(e["q_lab"] for e in qual)
    return ev, qual, rows, (round(allD / allH, 2) if allH else 0, len(qual), allH, allD)

evS, qualS, rowsS, totS = scenario(False)
evB, qualB, rowsB, totB = scenario(True)
okS = [r for r in rowsS if r["ok"]]; okB = [r for r in rowsB if r["ok"]]
best = max(okS, key=lambda r: r["elr"])

wt = [x for x in L if x["lt"] == "WT"]
wt_rate = round(sum(x["labor"] for x in wt) / sum(x["hrs"] for x in wt), 2)
cpt = [x for x in L if x["lt"] == "CPT"]
cpt_rate = round(sum(x["labor"] for x in cpt) / sum(x["hrs"] for x in cpt), 2)

# ---------------- workbook ----------------
NAVY = "1F3864"; GOLD = "FFC9A04B"
HF = Font(name="Arial", bold=True, color="FFFFFF", size=10); BF = Font(name="Arial", size=10)
BB = Font(name="Arial", bold=True, size=10); TF = Font(name="Arial", bold=True, size=14, color=NAVY)
HFILL = PatternFill("solid", fgColor=NAVY); GFILL = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE"); RED = PatternFill("solid", fgColor="FFC7CE"); YEL = PatternFill("solid", fgColor="FFEB9C")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
MONEY = '"$"#,##0.00'; HRS = '#,##0.00'

wb = Workbook()
def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c); x.font = HF; x.fill = HFILL; x.border = BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
def put(ws, row, vals, fmts=None, font=BF):
    for i, v in enumerate(vals, 1):
        x = ws.cell(row=row, column=i, value=v); x.font = font; x.border = BOX
        if fmts and fmts[i - 1]: x.number_format = fmts[i - 1]

# Summary
ws = wb.active; ws.title = "Summary"; ws.sheet_properties.tabColor = GOLD[2:]
ws["A1"] = "Toyota Flagstaff - Warranty Labor Rate, In-House Calculation (Toyota Policy 7.4, Option B)"; ws["A1"].font = TF
ws["A2"] = "ROs opened Oct 24 - Nov 21, 2026 (last 30 days, closed ROs in Qlik). Draft for method testing - not a submission."; ws["A2"].font = Font(name="Arial", italic=True, size=10)
ws.column_dimensions["A"].width = 58; ws.column_dimensions["B"].width = 18; ws.column_dimensions["C"].width = 18; ws.column_dimensions["D"].width = 60
header_row = 4
for i, c in enumerate(["Measure", "Strict", "Incl. Review lines", "Notes"], 1):
    x = ws.cell(row=header_row, column=i, value=c); x.font = HF; x.fill = HFILL; x.border = BOX
sc = "'Start Date Scenarios'"
nS = len(rowsS); last = 5 + nS
lines = [
 ("Qualified customer-pay ROs in window", totS[1], totB[1], None, "Toyota ROs with at least one qualified CPT repair line after RO-level exclusions"),
 ("Qualified labor hours (all qualified ROs)", totS[2], totB[2], HRS, "Hours billed to the customer (Qlik sold hours)"),
 ("Qualified labor $ (all qualified ROs)", totS[3], totB[3], MONEY, "No discounts found in Qlik - see Method & Limits"),
 ("Effective labor rate - all qualified ROs in window", totS[0], totB[0], MONEY, "Total qualified labor $ / total qualified hours"),
 ("30-RO ELR - highest start date", f"=MAX({sc}!G6:G{last})", f"=MAX({sc}!K6:K{last})", MONEY, "Best case if the Region's start date landed here"),
 ("30-RO ELR - median start date", f"=MEDIAN({sc}!G6:G{last})", f"=MEDIAN({sc}!K6:K{last})", MONEY, "Typical outcome"),
 ("30-RO ELR - lowest start date", f"=MIN({sc}!G6:G{last})", f"=MIN({sc}!K6:K{last})", MONEY, "Worst case"),
 ("Warranty labor $ per sold hour (WT lines, same window)", wt_rate, wt_rate, MONEY, "Proxy for the current warranty rate - confirm the approved rate with Tyler / DSPM"),
 ("All CPT labor $ per sold hour (incl. maintenance/tires)", cpt_rate, cpt_rate, MONEY, "Blended retail; Toyota caps the warranty rate at the Current Retail Labor Rate on the form"),
 ("Gap: median 30-RO ELR vs warranty proxy", "=B10-B12", "=C10-C12", MONEY, "Per warranty hour; positive = Option B likely beats current rate"),
]
for i, (m, s, b, f, n) in enumerate(lines):
    rr = header_row + 1 + i
    put(ws, rr, [m, s, b, n], [None, f, f, None])
ws.cell(row=header_row + 1 + len(lines) + 1, column=1, value="Highest-ELR start date (strict):").font = BB
ws.cell(row=header_row + 1 + len(lines) + 1, column=2, value=best["start"]).font = BB
ws.cell(row=header_row + 1 + len(lines) + 2, column=1, value="The Region sets the start date (an RO open date before the notice date) - we cannot pick it. Use the range above to judge timing and upside.").font = Font(name="Arial", italic=True, size=10)
r0 = header_row + len(lines) + 5
ws.cell(row=r0, column=1, value="Key assumptions (details on Method & Limits)").font = BB
for i, t in enumerate([
 "Only labor type CPT is true customer pay (per Noah). CTX, CPTC, CTEC, CPTE lines never qualify.",
 "Whole RO excluded if it carries IPS (goodwill/policy), CTEC (assumed extended warranty/third party) or CPTE (assumed employee).",
 "Non-Toyota makes (incl. Lexus), body shops, Findlay and rental customers excluded.",
 "'Strict' = only clearly qualified lines. 'Incl. Review' adds lines Toyota's list doesn't name (e.g., 4WD/AWD, PS fluid, fuel induction, throttle body).",
 "Discounts: none visible in Qlik line or header data, so proportional discount = $0. Verify on the accounting copies.",
 "Genuine parts and diagnosis documentation can't be checked in Qlik - assumed met; confirm on the RO copies.",
]):
    ws.cell(row=r0 + 1 + i, column=1, value="- " + t).font = BF

# Start Date Scenarios
ws2 = wb.create_sheet("Start Date Scenarios"); ws2.sheet_properties.tabColor = NAVY
ws2["A1"] = "First 30 qualified ROs for every possible Region start date (RO open date)"; ws2["A1"].font = TF
ws2["A2"] = "ROs taken in RO-number order from the start date, none skipped. Blank = fewer than 30 qualified closed ROs after that date."; ws2["A2"].font = Font(name="Arial", italic=True, size=10)
header(ws2, 5, ["Start Date", "First RO", "30th RO", "30th RO Opened", "ROs Reviewed", "Strict Hours", "Strict ELR", "Strict Labor $", "Incl-Review Hours", "Incl-Review Labor $", "Incl-Review ELR"],
       [12, 11, 11, 14, 11, 12, 12, 14, 13, 15, 14])
bmap = {r["start"]: r for r in rowsB}
for i, r in enumerate(rowsS):
    rr = 6 + i; b = bmap[r["start"]]
    if r["ok"]:
        put(ws2, rr, [r["start"], r["first"], r["last"], r["last_open"], r["reviewed"], r["hrs"], f"=H{rr}/F{rr}", r["lab"],
                      b["hrs"] if b["ok"] else None, b["lab"] if b["ok"] else None, f"=J{rr}/I{rr}" if b["ok"] else None],
            [None, None, None, None, None, HRS, MONEY, MONEY, HRS, MONEY, MONEY])
        if r["start"] == best["start"]:
            for c in range(1, 12): ws2.cell(row=rr, column=c).fill = GREEN
    else:
        put(ws2, rr, [r["start"], f"only {r['n']} qualified", None, None, None, None, None, None, None, None, None])
ws2.conditional_formatting  # (no CF; best row shaded green)

# RO Worksheet (Toyota layout) for best strict start date
ws3 = wb.create_sheet("RO Worksheet"); ws3.sheet_properties.tabColor = GOLD[2:]
ws3["B1"] = f"DEALER'S WARRANTY LABOR RATE - REPAIR ORDER WORKSHEET (draft)   Toyota Flagstaff   Start date tested: {best['start']}"; ws3["B1"].font = Font(name="Arial", bold=True, size=12, color=NAVY)
ws3["B2"] = "Same columns and formulas as Toyota's RO Worksheet tab. Columns G/H are the RO's customer-pay (CPT) labor and parts totals. Discount (J) = 0 - none found in Qlik; verify against accounting copies."
ws3["B2"].font = Font(name="Arial", italic=True, size=9); ws3["B2"].alignment = Alignment(wrap_text=True); ws3.merge_cells("B2:M2"); ws3.row_dimensions[2].height = 30
for i, c in enumerate("ABCDEFGHIJKLM", 1):
    x = ws3.cell(row=3, column=i, value=c); x.font = BB; x.alignment = Alignment(horizontal="center")
header(ws3, 4, ["#", "R.O. Description", "RO Number", "RO Date", "Labor Hours", "Labor Amount", "RO Total for Labor", "RO Total for Parts", "Repair Order Total", "RO Discount", "Proportional Discount Amount", "Net Customer Paid Labor After Discount", "Effective Labor Rate"],
       [5, 30, 11, 12, 11, 13, 13, 13, 13, 11, 14, 16, 12])
for i, e in enumerate(best["set"]):
    rr = 5 + i
    put(ws3, rr, [i + 1, e["desc"], e["ro"], e["open"], e["q_hrs"], e["q_lab"], e["cp_lab"], e["cp_parts"], f"=SUM(G{rr}+H{rr})", 0,
                  f"=IFERROR(SUM((J{rr}/I{rr})*F{rr}),0)", f"=IFERROR(SUM(F{rr}-K{rr}),0)", f"=IFERROR(SUM(L{rr}/E{rr}),0)"],
        [None, None, None, None, HRS, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY, MONEY])
put(ws3, 35, ["Total", None, None, None, "=IFERROR(SUM(E5:E34),0)", "=SUM(F5:F34)", None, None, None, "=SUM(J5:J34)", "=SUM(K5:K34)", "=IFERROR(SUM(L5:L34),0)", "=IFERROR(SUM(L35/E35),0)"],
    [None, None, None, None, HRS, MONEY, None, None, None, MONEY, MONEY, MONEY, MONEY], font=BB)
for c in range(1, 14): ws3.cell(row=35, column=c).fill = GFILL

# RO Review (every RO in sequence)
ws4 = wb.create_sheet("RO Review"); ws4.sheet_properties.tabColor = NAVY
header(ws4, 1, ["RO Number", "Opened", "Closed", "Make", "Cust #", "Labor Types on RO", "Status (Strict)", "Exclusion Reason", "Qualified Ops", "Qualified Hours", "Qualified Labor $", "CPT Labor $", "CPT Parts $", "Status (Incl. Review)"],
       [11, 11, 11, 8, 11, 26, 14, 52, 30, 12, 14, 13, 13, 18])
bstat = {e["ro"]: e["status"] for e in evB}
for i, e in enumerate(evS):
    rr = 2 + i
    put(ws4, rr, [e["ro"], e["open"], e["close"], e["make"], e["custno"], e["lts"], e["status"], e["reason"], e["desc"], e["q_hrs"], e["q_lab"], e["cp_lab"], e["cp_parts"], bstat[e["ro"]]],
        [None] * 9 + [HRS, MONEY, MONEY, MONEY, None])
    ws4.cell(row=rr, column=7).fill = GREEN if e["status"] == "Qualified" else RED
ws4.auto_filter.ref = f"A1:N{len(evS) + 1}"

# Line Detail
ws5 = wb.create_sheet("Line Detail"); ws5.sheet_properties.tabColor = NAVY
header(ws5, 1, ["RO Number", "Opened", "Pay Type", "Labor Type", "Op Code", "Description", "Sold Hours", "Labor $", "Parts $", "Line Class", "Reason"],
       [11, 11, 9, 10, 11, 60, 11, 12, 12, 12, 46])
for i, r in enumerate(sorted(L, key=lambda x: (x["ro"], x["lt"], x["opcode"]))):
    rr = 2 + i
    put(ws5, rr, [r["ro"], r["open"], r["pt"], r["lt"], r["opcode"], r["desc"], r["hrs"], r["labor"], r["parts"], r["cls"], r["why"]],
        [None] * 6 + [HRS, MONEY, MONEY, None, None])
    ws5.cell(row=rr, column=10).fill = {Q: GREEN, X: RED, R: YEL}[r["cls"]]
ws5.auto_filter.ref = f"A1:K{len(L) + 1}"

# Op Code Map
ws6 = wb.create_sheet("Op Code Map"); ws6.sheet_properties.tabColor = NAVY
header(ws6, 1, ["Op Code", "Default Class", "Toyota 7.4 Category / Reason", "CPT Lines", "CPT Hours", "CPT Labor $"], [11, 14, 52, 11, 12, 14])
agg = C.defaultdict(lambda: [0, 0.0, 0.0])
for r in cpt:
    a = agg[r["opcode"]]; a[0] += 1; a[1] += r["hrs"]; a[2] += r["labor"]
for i, (op, a) in enumerate(sorted(agg.items(), key=lambda x: -x[1][2])):
    cls, why = OPMAP.get(op, (R, "Op code not mapped - check RO"))
    if op in ("GR", "NO-OP", "MISC"): why += " - description keywords can override"
    put(ws6, 2 + i, [op, cls, why, a[0], round(a[1], 2), round(a[2], 2)], [None, None, None, None, HRS, MONEY])
    ws6.cell(row=2 + i, column=2).fill = {Q: GREEN, X: RED, R: YEL}[cls]

# Method & Limits
ws7 = wb.create_sheet("Method & Limits"); ws7.sheet_properties.tabColor = NAVY
ws7.column_dimensions["A"].width = 150
txt = [
 ("Source", True),
 ("Qlik Closed Repair Orders app, Division = TOYOTA FLAGSTAFF, RO open date Oct 24 - Nov 21, 2026 (pulled Nov 23, 2026). Line grain = RO_Detail. Hours = sold hours (billed to customer), not flag time.", False),
 ("Toyota rules applied (Policies 7.1-7.4, Policy 7.4 rev 1/26)", True),
 ("Start date = RO open date set by the Region. From it, ROs are taken in numerical order and the first 30 qualified Toyota customer-pay ROs are used; none may be skipped.", False),
 ("ELR = sum of net qualified labor $ / sum of qualified hours across the 30 ROs (Toyota RO Worksheet formulas, reproduced on the RO Worksheet tab).", False),
 ("Excluded: warranty, goodwill, internal, non-Toyota vehicles; scheduled maintenance, tire work, wear items, inspections, accessories, cosmetic, sublet body/glass, hybrid battery fan; extended-warranty / third-party ROs (whole RO); employee, other-dealer, body-shop, rental customers.", False),
 ("Included exceptions: alignments, brake pads/rotors, documented diagnosis, warning lights, keys/fobs, calibrations, clutch, timing/serpentine belts, damage-related mechanical repair.", False),
 ("Store-specific assumptions", True),
 ("CPT is the only true customer-pay labor type (confirmed by Noah). Lines on CTX, CPTC (ToyotaCare), CTEC, CPTE never qualify.", False),
 ("CTEC assumed = Toyota Extra Care / service contract -> whole RO excluded (3 ROs). CPTE assumed = employee -> whole RO excluded. IPS = GL 71034 service policy (goodwill) -> whole RO excluded (30 ROs).", False),
 ("GR, NO-OP and MISC lines are classified from the op code description with keyword rules (tires, wipers, bulbs, oil, filters, fluid services, inspections, etc. -> excluded; rodent and belt lines -> Review).", False),
 ("Review lines are maintenance-type services Toyota's table does not name (4WD/AWD/differential, power steering fluid, fuel induction, throttle body, battery terminal service). 7.4 says 'all other services are to be included', but the DSPM may treat them as maintenance. Both scenarios are shown.", False),
 ("Limits - verify before relying on the number", True),
 ("Discounts: Qlik shows no negative or discount lines. If CDK discounts post outside the labor/parts lines, the RO Discount column must be filled from the accounting copies.", False),
 ("Genuine parts: Qlik RO lines don't identify aftermarket or used parts. Repairs with non-genuine parts would drop out.", False),
 ("Diagnosis documentation (CCR, codes, tech notes) and highlighting must be checked on the RO copies.", False),
 ("Closed ROs only: ROs still open are missing, so late-window start dates have gaps and fewer than 30 qualified ROs.", False),
 ("Employee check uses labor type only; an ADP name match was not run. Customer exclusions use name keywords (collision, Findlay, rental companies).", False),
 ("The warranty proxy (WT labor $ / sold hour) is not the approved Toyota warranty rate. Get the current approved rate from Tyler or the DSPM.", False),
 ("Flagstaff is on Option A (CPI) in the last year of its 3-year term. Check the CPI Participation Agreement's renewal/opt-out terms before planning an Option B request. Requesting a start date and not completing (or declining the offer) can block a new request for 12 months (Policy 7.3).", False),
]
for i, (t, b) in enumerate(txt):
    c = ws7.cell(row=1 + i, column=1, value=t); c.font = BB if b else BF; c.alignment = Alignment(wrap_text=True)

wb.save(OUT)
qa = {"lines": len(L), "ros": len(ROS), "qualified_strict": totS[1], "qualified_review": totB[1],
      "elr_all_strict": totS[0], "elr_all_review": totB[0], "wt_rate": wt_rate, "cpt_rate": cpt_rate,
      "starts_ok": len(okS), "starts_total": len(rowsS),
      "elr30_strict": [min(r["elr"] for r in okS), statistics.median(r["elr"] for r in okS), max(r["elr"] for r in okS)],
      "elr30_review": [min(r["elr"] for r in okB), statistics.median(r["elr"] for r in okB), max(r["elr"] for r in okB)],
      "best_start": best["start"], "best_ros": [best["first"], best["last"]], "best_reviewed": best["reviewed"],
      "excl_reasons": C.Counter(e["reason"] for e in evS if e["status"] == "Excluded"),
      "line_class": C.Counter(r["cls"] for r in cpt),
      "unmapped_ops": sorted({r["opcode"] for r in cpt if r["opcode"] not in OPMAP}),
      "zero_dollar_qual_lines": sum(1 for r in cpt if r["cls"] == Q and r["labor"] == 0),
      "review_lines": [(r["opcode"], r["desc"][:50]) for r in cpt if r["cls"] == R][:40]}
json.dump(qa, open(QA, "w"), indent=1, default=str)
print(json.dumps(qa, indent=1, default=str))
