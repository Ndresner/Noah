"""Historical check: our statutory rebuild vs. DealerUplift submissions.
python3 -I build_history.py <hist_dir> <approved_status.xlsx> <out.xlsx> <qa.json>
hist_dir has cycles.json, lines/<DIVISION>.csv, raw/aft_ut.json
"""
import csv, json, re, sys, statistics, collections as C
from datetime import date, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

H, STATUS, OUT, QA = sys.argv[1:5]
cfg = json.load(open(f"{H}/cycles.json"))
STORES, CYCLES = cfg["stores"], cfg["cycles"]

# ---------- line categories (description-driven, brand-agnostic) ----------
RX = [
 ("OTHER", r"SMS DISCLAIMER|^TEXT$|DETAIL|ACCESSOR|INSTALL (HITCH|ROOF|RUNNING|TOW)|SMOG|EMISSION TEST|SAFETY INSPECTION|PRE-?PURCHASE|APPRAIS|RECONDITION|CAR WASH|SHUTTLE|RENTAL|LOANER|PAINT|PPF|CERAMIC|TINT"),
 ("TIRE", r"\bTIRES?\b|BALANCE|ROTAT|MOUNT AND BAL|MOUNT & BAL|MOUNT \d|POTHOLE|TIRE REPAIR|FLAT REPAIR"),
 ("ALIGN", r"ALIGN"),
 ("BRAKE", r"BRAKE PAD|BRAKES? (JOB|REPAIR|SERVICE W)|\bPADS\b|ROTORS?\b|BRAKE SHOE|REPLACE (FRONT|REAR) BRAKES|(FRONT|REAR) BRAKES?\b(?! FLUID)|RESURFACE"),
 ("MAINT", r"OIL (AND|&) FILTER|OIL CHANGE|\bLOF\b|LUBE|SYNTHETIC|SYNPOWER|\b\d{1,3}K\b|\d{1,3},\d{3} MILE|MILE (SERVICE|MAINT)|INTERVAL|MAINTENANCE|SERVICE MENU|ANNUAL SERVICE|YEAR SERVICE|\bA\d+[- ]?B\d+|B1 SERVICE|MAJOR SERVICE|MINOR SERVICE|TUNE[- ]?UP|VALVE ADJUST|INSPECTION LINE|MULTI[- ]?POINT|\bMPI\b|FILTER|FLUID|FLUSH|EXCHANGE|DRAIN AND FILL|DIFFERENTIAL SERVICE|TRANSFER CASE SERVICE|TRANSMISSION SERVICE|CVT SERVICE|DSG SERVICE|4X4 SERVICE|AWD SERVICE|COOLANT|COOLING SYSTEM (SERVICE|FLUSH)|\bBG\b|INDUCTION|FUEL SYSTEM SERVICE|FUEL SERVICE|THROTTLE BODY (SERVICE|CLEAN)|INTAKE VALVE|EGR SERVICE|DECARB|REFRESH|SPARK PLUG|TIMING BELT|WIPER|BULB|BATTERY (TERMINAL|SERVICE)|SERVICE BATTERY|TERMINALS|BODY LUBE|GREEN SERVICE|BLUE SERVICE|RED SERVICE|EXPRESS|XPRESS|VALVOLINE"),
 ("BATTERY", r"BATTER"),
]
def category(desc, hrs, labor=None):
    d = desc.upper()
    if hrs <= 0: return "NOHOURS"
    if labor is not None and hrs >= 15 and labor / hrs < 30: return "BADHOURS"   # clerical hours errors (e.g. 150.0 for 1.5)
    for cat, pat in RX:
        if re.search(pat, d): return cat
    return "REPAIR"   # diagnosis, concerns, repairs, unmapped

EXCL_CUST = re.compile(r"COLLISION|BODY SHOP|FINDLAY|RENT A CAR|ENTERPRISE|HERTZ|AVIS|BUDGET|AUTO BODY|MOTORS INC|AUTO SALES")

def state_rules(state):
    # which categories count; alignments only with a repair on the same RO (NV/AZ/ID/WA); UT excludes alignments
    base = {"REPAIR"}
    r = {"cats": set(base), "battery": True, "align": "with_repair", "aftermarket": False}
    if state == "UT": r.update(battery=False, align="none", aftermarket=True)
    if state == "WA": r.update(battery=False)
    return r

# aftermarket repair parts on RO (Utah only)
aft = {}
d = json.load(open(f"{H}/raw/aft_ut.json"))
CONSUM = re.compile(r"\bBG\b|OIL|FILTER|DRAIN|GASKET|WASHER|MOA|FLUSH|PROTECT|TOTE|FLUID|ADDITIVE|COOL|REFRIG|WIP|BULB|KIT|0W|5W|ATF|CVT|CLEAN|TREAT|LUBE|GREASE|SPRAY|^\S{0,3}$")
for div, n, cell in d["data"]["data"]:
    for x in (cell or "").split("~"):
        if not x: continue
        ro, amt, desc = x.split("|", 2)
        if any(p.strip() and not CONSUM.search(p.strip().upper()) for p in desc.split("/")): aft[(div, ro)] = 1

def load(div):
    L = list(csv.DictReader(open(f"{H}/lines/{div}.csv")))
    for r in L:
        for k in ("hrs", "labor", "parts", "pcost"): r[k] = float(r[k])
        r["cat"] = category(r["desc"], r["hrs"], r["labor"])
    return L

def build_cycle(div, L, typ, end, make_pref):
    st = STORES[div]; rules = state_rules(st["state"])
    e = date.fromisoformat(end); s = e - timedelta(days=180)
    makes = [make_pref] if make_pref else st["make"]
    ros = C.OrderedDict()
    for r in sorted(L, key=lambda x: (int(x["ro"]) if x["ro"].isdigit() else 0)):
        if not (s.isoformat() <= r["open"] <= end): continue
        ros.setdefault(r["ro"], []).append(r)
    ev = []; excl = C.Counter()
    for ro, lines in ros.items():
        h = lines[0]; lts = {x["lt"] for x in lines}
        cp = [x for x in lines if x["lt"] in st["cp"]]
        why = None
        if not cp: why = "no customer pay"
        elif not any(h["make"].startswith(m) for m in makes): why = "other make"
        elif lts & set(st["policy"]): why = "goodwill/policy"
        elif lts & set(st["third"]): why = "extended warranty/third party"
        elif lts & set(st["emp"]): why = "employee"
        elif EXCL_CUST.search(h["name"].upper()): why = "dealer/body shop/rental customer"
        elif rules["aftermarket"] and (div, ro) in aft: why = "aftermarket parts (UT)"
        if why: excl[why] += 1; continue
        core = [x for x in cp if x["cat"] == "REPAIR" or (x["cat"] == "BATTERY" and rules["battery"])]
        q = list(core)
        if rules["align"] == "with_repair" and any(x["cat"] == "REPAIR" for x in core):
            q += [x for x in cp if x["cat"] == "ALIGN"]
        if not q: excl["no qualifying repair"] += 1; continue
        H_ = sum(x["hrs"] for x in q); D_ = sum(x["labor"] for x in q)
        PS = sum(x["parts"] for x in q if x["pcost"] > 0); PC = sum(x["pcost"] for x in q if x["pcost"] > 0)
        ev.append({"ro": ro, "open": h["open"], "H": H_, "D": D_, "PS": PS, "PC": PC})
    wins = []
    for i in range(len(ev)):
        w = ev[i:i + 100]
        d0 = date.fromisoformat(w[0]["open"])
        w = [x for x in w if (date.fromisoformat(x["open"]) - d0).days <= 89]
        # lesser of 100 sequential ROs or 90 days: a short window is valid only if its 90 days fit in the look-back
        if len(w) < 100 and d0 + timedelta(days=89) > e: break
        Hh = sum(x["H"] for x in w); Dd = sum(x["D"] for x in w); PS = sum(x["PS"] for x in w); PC = sum(x["PC"] for x in w)
        if Hh <= 0 or PC <= 0: continue
        wins.append({"first": w[0]["ro"], "last": w[-1]["ro"], "d0": w[0]["open"], "d1": w[-1]["open"], "n": len(w),
                     "lr": Dd / Hh, "mk": PS / PC - 1})
    key = "lr" if typ == "Labor" else "mk"
    vals = [w[key] for w in wins]
    if not vals: return None
    best = max(wins, key=lambda w: w[key])
    return {"n_ro": len(ev), "n_win": len(wins), "excl": excl, "min": min(vals), "med": statistics.median(vals), "max": max(vals),
            "latest": wins[-1][key], "best": best, "window": f"{s.isoformat()} to {end}"}

# approved status (Armatus account status file)
ap = {}
ws = openpyxl.load_workbook(STATUS, data_only=True, read_only=True)["Findlay"]
for r in ws.iter_rows(min_row=3, values_only=True):
    if r[1]: ap[str(r[1]).strip()] = r
NAME = {("HONDA HENDERSON", None): "Findlay Honda Henderson", ("HONDA NORTH", None): "Findlay Honda North", ("NISSAN", None): "Findlay Nissan Henderson",
        ("LR HENDERSON", "LAND"): "Land Rover Henderson", ("LR HENDERSON", "JAGU"): "Jaguar Henderson Service", ("LR LV", "LAND"): "Land Rover Las Vegas",
        ("LR LV", "JAGU"): "Jaguar Las Vegas", ("SUBARU ST GEORGE", None): "Findlay Subaru St George", ("KIA ST GEORGE", None): "Findlay Kia St George",
        ("VW ST GEORGE", None): "Findlay Volkswagen St George", ("SUBARU PRESCOTT", None): "Findlay Subaru Prescott",
        ("HONDA FLAGSTAFF", None): "Findlay Honda Flagstaff", ("GMC PRESCOTT", None): "Findlay BGMC Prescott",
        ("CDJR POST FALLS", None): "Findlay Chrysler Post Falls", ("CHEVY LV", None): "Findlay Chevrolet", ("LEXUS", None): "Findlay Lexus Spokane"}

cache = {}; res = []
for c in CYCLES:
    div, typ, end, du, cur, outcome = c[:6]; mk = c[6] if len(c) > 6 else None
    if div not in cache: cache[div] = load(div)
    r = build_cycle(div, cache[div], typ, end, mk)
    nm = NAME[(div, mk)]; a = ap.get(nm)
    if typ == "Labor": appr, appr_d = (a[3], a[4]) if a else (None, None)
    else: appr, appr_d = (a[7], a[8]) if a else (None, None)
    pct = typ == "Parts"
    res.append({"store": nm, "state": STORES[div]["state"], "type": typ, "date": end, "du": (du / 100 if pct else du) if du is not None else None,
                "cur": cur / 100 if pct else cur, "outcome": outcome, "approved": appr, "approved_date": appr_d, "r": r})

# ---------- workbook ----------
NAVY = "1F3864"; GOLD = "C9A04B"
HF = Font(name="Arial", bold=True, color="FFFFFF", size=10); BF = Font(name="Arial", size=10); BB = Font(name="Arial", bold=True, size=10)
TF = Font(name="Arial", bold=True, size=14, color=NAVY); IT = Font(name="Arial", italic=True, size=9)
HFILL = PatternFill("solid", fgColor=NAVY); GREEN = PatternFill("solid", fgColor="C6EFCE"); YEL = PatternFill("solid", fgColor="FFEB9C"); RED = PatternFill("solid", fgColor="FFC7CE")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook(); ws = wb.active; ws.title = "Comparison"; ws.sheet_properties.tabColor = GOLD
ws["A1"] = "DealerUplift submissions vs. our statutory rebuild from Qlik (same 180-day look-back)"; ws["A1"].font = TF
ws["A2"] = "Our rate = 100 sequential qualifying customer-pay ROs (or 90 days, whichever fewer), state rules applied. 'Best' is the highest window in the 180 days before DealerUplift's results date - the comparable number, since they pick the window."; ws["A2"].font = IT
cols = ["Store", "State", "Type", "DU results date", "Prior rate", "DU submitted", "Approved (status file)", "Approved eff.", "Outcome (email)",
        "Our best window", "Our median window", "Our latest window", "Our low window", "Best minus DU", "Qualified ROs", "Windows", "Best window ROs", "Best window dates"]
for i, h in enumerate(cols, 1):
    x = ws.cell(row=4, column=i, value=h); x.font = HF; x.fill = HFILL; x.border = BOX; x.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws.row_dimensions[4].height = 42
widths = [28, 6, 7, 12, 10, 11, 12, 11, 18, 11, 11, 11, 10, 11, 9, 8, 20, 22]
for i, w_ in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w_
for j, x in enumerate(res):
    rr = 5 + j; r = x["r"]; pct = x["type"] == "Parts"; f = '0.0%' if pct else '"$"#,##0.00'
    vals = [x["store"], x["state"], x["type"], x["date"], x["cur"], x["du"], x["approved"] if isinstance(x["approved"], (int, float)) else x["approved"],
            str(x["approved_date"])[:10] if x["approved_date"] else None, x["outcome"]]
    if r:
        vals += [r["max"], r["med"], r["latest"], r["min"], (f"=J{rr}-F{rr}" if x["du"] is not None else None), r["n_ro"], r["n_win"], f"{r['best']['first']}-{r['best']['last']}", f"{r['best']['d0']} to {r['best']['d1']}"]
    else:
        vals += ["insufficient data"] + [None] * 8
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=rr, column=i, value=v); c.font = BF; c.border = BOX
        if i in (5, 6, 7, 10, 11, 12, 13, 14): c.number_format = f
    if r and x["du"] is not None:
        gap = (r["max"] - x["du"]) / (x["du"] if not pct else 1)
        tol = 0.03 if not pct else 0.05
        fill = GREEN if abs(gap) <= tol else (YEL if abs(gap) <= 2 * tol else RED)
        ws.cell(row=rr, column=14).fill = fill
ws.freeze_panes = "B5"
n = len(res); rr = 6 + n
ws.cell(row=rr, column=1, value="Color on 'Best minus DU': green = within 3% (labor) / 5 pts (parts); yellow = within 6% / 10 pts; red = further.").font = IT
w2 = wb.create_sheet("Method & Limits"); w2.column_dimensions["A"].width = 160
for i, (t, b) in enumerate([
 ("Method", True),
 ("For each DealerUplift cycle with a stated final/prelim rate, ROs opened in the 180 days up to DealerUplift's results date were pulled from Qlik Closed ROs (line grain).", False),
 ("Qualifying lines: store's customer-pay labor types (config), franchise make only, ROs with policy/goodwill, extended-warranty/third-party or employee labor types excluded, dealer/body-shop/rental customers excluded.", False),
 ("Line categories are assigned from the op-code description (brand-agnostic): maintenance/fluids/filters/menus, tires, alignments, brakes, batteries, other (inspections, detail, accessories), else repair/diagnosis.", False),
 ("NV, AZ, ID: repair + diagnosis + batteries; alignments only on ROs with another repair; brake pads/rotors, maintenance and tires excluded. WA: same, batteries excluded. UT (13-14-204): repair + diagnosis only; alignments, brakes, non-EV batteries excluded; ROs with aftermarket repair parts excluded.", False),
 ("Rates: labor = labor sales / sold hours; parts markup = parts sale / parts cost - 1, both on qualifying lines, per window of 100 sequential qualifying ROs (lesser of 100 or 90 days).", False),
 ("Limits", True),
 ("Qlik sales are net of discounts; DealerUplift notes discount losses in some results. Promotional ROs can't be identified. Description-based classification will differ from a line-by-line manual review.", False),
 ("DealerUplift's exact sample windows are not known (only results dates), so 'best window' is the fairest comparison; median/latest show how much depends on window choice.", False),
 ("Approved figures come from the Armatus account status file (latest approved rate per store, which may postdate the cycle shown).", False),
 ("Parts cycles compare markup percentages; DealerUplift parts results may use a different line set (e.g., all customer-pay parts on qualifying ROs).", False),
]):
    c = w2.cell(row=1 + i, column=1, value=t); c.font = BB if b else BF; c.alignment = Alignment(wrap_text=True)
wb.save(OUT)
json.dump([{**{k: v for k, v in x.items() if k != "r"}, "r": ({k: v for k, v in x["r"].items() if k not in ("best", "excl")} | {"excl": dict(x["r"]["excl"]), "best": {k: v for k, v in x["r"]["best"].items()}}) if x["r"] else None} for x in res],
          open(QA, "w"), indent=1, default=str)
for x in res:
    r = x["r"]; p = x["type"] == "Parts"
    fmt = (lambda v: f"{v*100:6.1f}%") if p else (lambda v: f"${v:7.2f}")
    print(f"{x['store'][:26]:26s} {x['state']} {x['type'][:1]} {x['date']} DU={fmt(x['du']) if x['du'] is not None else '  (none)'} appr={x['approved']} | "
          + (f"best={fmt(r['max'])} med={fmt(r['med'])} latest={fmt(r['latest'])} ros={r['n_ro']} wins={r['n_win']}" if r else "NO DATA"))
