import openpyxl, re, json
from config import load_config

_cfg = load_config()
ROSTER_PATH = _cfg["roster_path"]
QLIK_DATA_PATH = _cfg["qlik_data_path"]

# Roster location -> Division (Qlik) mapping
LOC_TO_DIVISION = {
    "ACURA": "ACURA",
    "AUDI HENDERSON": "AUDI HENDERSON",
    "AUDI RENO TAHOE": "AUDI RENO",
    "CADILLAC": "CADILLAC",
    "CHEVROLET LAS VEGAS": "CHEVY LV",
    "CHRYSLER POST FALLS": "CDJR POST FALLS",
    "FINDLAY HONDA SPOKANE": "HONDA SPOKANE",
    "GMC PRESCOTT": "GMC PRESCOTT",
    "HONDA FLAGSTAFF": "HONDA FLAGSTAFF",
    "HONDA HENDERSON": "HONDA HENDERSON",
    "HONDA NORTH": "HONDA NORTH",
    "HYUNDAI PRESCOTT": "HYUNDAI PRESCOTT",
    "HYUNDAI ST GEORGE": "HYUNDAI ST GEORGE",
    "INEOS GRENADIER": "INEOS",
    "JAGUAR LAND ROVER HENDERSON": "LR HENDERSON",
    "KIA LAS VEGAS": "KIA LV",
    "KIA ST GEORGE": "KIA ST GEORGE",
    "LAND ROVER LAS VEGAS": "LR LV",
    "LAND ROVER RENO": "LR RENO",
    "LINCOLN": "LINCOLN",
    "LEXUS SPOKANE": "LEXUS",
    "MOTOR COMPANY": "CHEVY GMC BULLHEAD",
    "MAZDA HENDERSON": "MAZDA",
    "NISSAN HENDERSON": "NISSAN",
    "SUBARU LAS VEGAS": "SUBARU LV",
    "SUBARU PRESCOTT": "SUBARU PRESCOTT",
    "SUBARU ST GEORGE": "SUBARU ST GEORGE",
    "TOYOTA FLAGSTAFF": "TOYOTA FLAGSTAFF",
    "TOYOTA HENDERSON": "TOYOTA HENDERSON",
    "TOYOTA PRESCOTT": "TOYOTA PRESCOTT",
    "TOYOTA SPOKANE": "TOYOTA SPOKANE",
    "VW HENDERSON": "VW HENDERSON",
    "VW ST. GEORGE": "VW ST GEORGE",
    "VOLVO CARS LAS VEGAS": "VOLVO",
}

DIVISION_TO_SHEET = {
    "ACURA": "Acura", "AUDI HENDERSON": "Audi Henderson", "AUDI RENO": "Audi Reno",
    "CADILLAC": "Cadillac", "CDJR POST FALLS": "Cdjr Post Falls", "CHEVY GMC BULLHEAD": "Chevy Gmc Bullhead",
    "CHEVY LV": "Chevy Lv", "GMC PRESCOTT": "Gmc Prescott", "HONDA FLAGSTAFF": "Honda Flagstaff",
    "HONDA HENDERSON": "Honda Henderson", "HONDA NORTH": "Honda North", "HONDA SPOKANE": "Honda Spokane",
    "HYUNDAI PRESCOTT": "Hyundai Prescott", "HYUNDAI ST GEORGE": "Hyundai St George", "INEOS": "Ineos",
    "KIA LV": "Kia Lv", "KIA ST GEORGE": "Kia St George", "LEXUS": "Lexus", "LINCOLN": "Lincoln",
    "LR HENDERSON": "Lr Henderson", "LR LV": "Lr Lv", "LR RENO": "Lr Reno", "MAZDA": "Mazda",
    "NISSAN": "Nissan", "SUBARU LV": "Subaru Lv", "SUBARU PRESCOTT": "Subaru Prescott",
    "SUBARU ST GEORGE": "Subaru St George", "TOYOTA FLAGSTAFF": "Toyota Flagstaff",
    "TOYOTA HENDERSON": "Toyota Henderson", "TOYOTA PRESCOTT": "Toyota Prescott",
    "TOYOTA SPOKANE": "Toyota Spokane", "VOLVO": "Volvo", "VW HENDERSON": "Vw Henderson",
    "VW ST GEORGE": "Vw St George",
}

POSITION_MAP = {
    "SERVICE TECHNICIAN": "Service Technician",
    "EXPRESS TECHNICIAN": "Express Technician",
    "SHOP FOREMAN": "Shop Foreman",
    "SERVICE TEAM LEADER": "Service Team Leader",
    "BODY SHOP TECH": "Body Shop Tech",
    "BODY SHOP PAINTER": "Body Shop Painter",
    "BODY SHOP HELPER": "Body Shop Helper",
    "DETAILER": "Detailer",
    "DETAIL MANAGER": "Detailer",
    "TINTER": "Tinter",
}

CATEGORY_ORDER = ["Service Technician", "Express Technician", "Service Team Leader", "Shop Foreman",
                   "Body Shop Tech", "Body Shop Painter", "Body Shop Helper", "Detailer", "Tinter"]


def tech_no_from_position_id(pid):
    # Position IDs are always 9 chars: a 3-char store/prefix code + a 6-digit tech number
    # (the digits may include leading zeros, which Qlik's techno field drops).
    suffix = pid[3:]
    if not suffix.isdigit():
        return None
    return str(int(suffix))


def load_roster():
    wb = openpyxl.load_workbook(ROSTER_PATH, data_only=True)
    ws = wb["1"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    roster = []
    for r in rows:
        loc, pid, last, first, title, status, reg, ot = r
        if not isinstance(pid, str):
            continue
        loc_u = loc.strip().upper()
        division = LOC_TO_DIVISION.get(loc_u)
        if not division:
            raise ValueError(f"Unmapped location: {loc}")
        tech_no = tech_no_from_position_id(pid)
        position = POSITION_MAP.get((title or "").strip().upper())
        name = f"{first} {last}".strip()
        actual_hours = round((reg or 0) + (ot or 0), 2)
        roster.append({
            "division": division, "tech_no": tech_no, "name": name,
            "position": position, "actual_hours": actual_hours,
        })
    return roster


def load_qlik():
    with open(QLIK_DATA_PATH) as fh:
        tech_data = json.load(fh)
    qlik = {}
    for division, tech_no, sold_hours, ro_count, labor_sale, labor_gross in tech_data:
        qlik.setdefault(division, {})[tech_no] = {
            "sold_hours": sold_hours, "ro_count": ro_count,
            "labor_sale": labor_sale, "labor_gross": labor_gross,
        }
    return qlik


def build():
    roster = load_roster()
    qlik = load_qlik()

    stores = {}  # sheet_name -> {"named": {category: [rows]}, "unmapped": [rows]}
    matched_keys = set()

    for emp in roster:
        division = emp["division"]
        sheet = DIVISION_TO_SHEET[division]
        qd = qlik.get(division, {}).get(emp["tech_no"])
        stores.setdefault(sheet, {"named": {}, "unmapped": []})
        if qd is not None:
            matched_keys.add((division, emp["tech_no"]))
        # Every roster employee has a real, mapped position -> always a named row,
        # even with zero Qlik production (no closed ROs this period).
        row = {
            "tech_no": emp["tech_no"], "name": emp["name"], "position": emp["position"],
            "sold_hours": qd["sold_hours"] if qd else 0.0,
            "ro_count": qd["ro_count"] if qd else 0,
            "actual_hours": emp["actual_hours"],
            "labor_sale": qd["labor_sale"] if qd else 0.0,
            "labor_gross": qd["labor_gross"] if qd else 0.0,
        }
        stores[sheet]["named"].setdefault(emp["position"], []).append(row)

    # Any Qlik tech codes with NO roster match at all -> pooled/unmapped (e.g. 77777, MULT, 9999, DS)
    for division, techs in qlik.items():
        sheet = DIVISION_TO_SHEET.get(division)
        if not sheet:
            continue
        stores.setdefault(sheet, {"named": {}, "unmapped": []})
        for tech_no, qd in techs.items():
            if (division, tech_no) in matched_keys:
                continue
            row = {
                "tech_no": tech_no, "name": None, "position": None,
                "sold_hours": qd["sold_hours"], "ro_count": qd["ro_count"],
                "actual_hours": None,
                "labor_sale": qd["labor_sale"], "labor_gross": qd["labor_gross"],
            }
            stores[sheet]["unmapped"].append(row)

    return stores


if __name__ == "__main__":
    stores = build()
    # sanity print
    total_named = sum(len(v) for s in stores.values() for v in s["named"].values())
    total_unmapped = sum(len(s["unmapped"]) for s in stores.values())
    print("stores:", len(stores))
    print("named rows:", total_named, "unmapped rows:", total_unmapped)
    for s in ["Cadillac", "Acura", "Chevy Gmc Bullhead"]:
        print(s, {k: len(v) for k, v in stores[s]["named"].items()}, "unmapped:", len(stores[s]["unmapped"]))
