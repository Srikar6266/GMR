#!/usr/bin/env python3
"""
Build August 2026 Duty Roster for full 281 manpower.
- Electrical: Client_Electrical_Requirement + Sample layout
- Civil: Client_Civil_Requirement Final (ALS & PTB)
- Other departments: Client_Manpower_Requirement (skip Electrical & Civil sections)
- Existing people: names/phones/locations from Exisiting tab
Rules: exact daily A/B/C; reliever = W/O; 6 work + 1 WO where math allows;
       after C next day must not be A or G.
"""
from __future__ import annotations

import calendar
import re
from collections import Counter, defaultdict
from copy import deepcopy
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import pulp

YEAR, MONTH = 2026, 8
DAYS = calendar.monthrange(YEAR, MONTH)[1]  # 31
DATES = [date(YEAR, MONTH, d) for d in range(1, DAYS + 1)]
SHIFTS = ["A", "B", "C", "G", "W/O"]
SOURCE = "/workspace/281_ShiftRoster.xlsx"
OUT = "/workspace/281_August_2026_Duty_Roster.xlsx"


def nz(v):
    if v is None or v == "-" or v == "":
        return 0
    try:
        return int(float(v))
    except Exception:
        return 0


def rest_window(n_shift: int, R: int):
    if R <= 0 or n_shift <= 0:
        return None
    if R * 7 >= n_shift:
        return 7
    for W in range(8, n_shift + 2):
        if R * W >= n_shift:
            return W
    return None


# ---------------------------------------------------------------------------
# Requirement definitions (281 slots)
# Each group: section, area, role_label, required, plan{G,A,B,C,R}, default_location
# ---------------------------------------------------------------------------

def build_requirement_groups():
    groups = []

    def add(section, area, role, req, g=0, a=0, b=0, c=0, r=0, loc=""):
        total = g + a + b + c + r
        if req <= 0:
            return
        # Prefer explicit req; align plan to req if mismatch by adjusting G or R
        plan = {"G": g, "A": a, "B": b, "C": c, "R": r}
        if total != req:
            # Keep A/B/C; put remainder into G if no shifts, else adjust R/G
            if a + b + c == 0:
                plan = {"G": req, "A": 0, "B": 0, "C": 0, "R": 0}
            else:
                core = a + b + c
                rem = req - core
                if rem < 0:
                    # scale down not expected; force R=0 and trim C/B/A
                    plan = {"G": 0, "A": a, "B": b, "C": c, "R": 0}
                    # if still over, keep as stated A/B/C and set req=core (caller shouldn't)
                else:
                    # rem = G + R; prefer keeping stated R if rem>=R else all R
                    if r and rem >= r:
                        plan = {"G": rem - r, "A": a, "B": b, "C": c, "R": r}
                    else:
                        plan = {"G": 0, "A": a, "B": b, "C": c, "R": rem}
        groups.append({
            "section": section,
            "area": area,
            "role": role,
            "required": req,
            "plan": plan,
            "location": loc,
        })

    # ===== A. Site Management (15) =====
    add("Site Management", "ADMIN", "Site Head (GM)", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "Sr. Manager (Mechanical)", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "Sr. Manager (Electrical)", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "Sr. Manager (Civil)", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "HSE & Compliance Manager", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "HR Manager", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "MIS Executive", 1, g=1, loc="ADMIN")
    add("Site Management", "ADMIN", "Materials Manager", 1, g=1, loc="ADMIN")
    add("Site Management", "Stores", "Stores Exe", 7, a=2, b=2, c=2, r=1, loc="Stores")

    # ===== B. SCADA & Helpdesk (8) =====
    add("SCADA & Helpdesk", "APOC", "BMS", 4, a=1, b=1, c=1, r=1, loc="APOC")
    add("SCADA & Helpdesk", "APOC", "Help Desk Executive", 4, a=1, b=1, c=1, r=1, loc="APOC")

    # ===== C. Mechanical & HVAC (85) =====
    add("Mechanical & HVAC", "MECH", "Asst. Manager", 1, g=1, loc="PTB")
    add("Mechanical & HVAC", "MECH", "Shift Engineer", 4, a=1, b=1, c=1, r=1, loc="UTILITY")
    add("Mechanical & HVAC", "MECH", "Sr Technician", 4, a=1, b=1, c=1, r=1, loc="PTB")
    add("Mechanical & HVAC", "MECH", "Technician (HVAC) L", 21, g=1, a=6, b=6, c=5, r=3, loc="UTILITY")
    add("Mechanical & HVAC", "MECH", "Tech Associate Mech.", 11, g=2, a=3, b=3, c=2, r=1, loc="PTB")
    add("Mechanical & HVAC", "MECH", "PM Technician HVAC", 2, g=2, loc="PTB")
    add("Mechanical & HVAC", "MECH", "Fitter Mechanical", 1, g=1, loc="ATC")
    add("Mechanical & HVAC", "MECH", "WTP/STP Chemist", 1, g=1, loc="WTP")
    add("Mechanical & HVAC", "MECH", "WTP Operator", 4, a=1, b=1, c=1, r=1, loc="WTP")
    add("Mechanical & HVAC", "MECH", "STP Operator", 3, a=1, b=1, c=1, r=0, loc="UTILITY")
    add("Mechanical & HVAC", "MECH", "WTP/STP Helper", 4, a=1, b=1, c=1, r=1, loc="WTP")
    add("Mechanical & HVAC", "MECH", "Plumber", 14, a=4, b=4, c=4, r=2, loc="UTILITY")
    add("Mechanical & HVAC", "MECH", "Pump Operator", 9, a=3, b=3, c=2, r=1, loc="CFR MAIN PUMP HOUSE")
    add("Mechanical & HVAC", "MECH", "Sewerman & Drainage", 6, a=2, b=2, c=1, r=1, loc="UTILITY")

    # ===== D. Electrical from Client Electrical + Sample (71) =====
    # PTB
    add("Electrical", "PTB", "Asst Manager", 1, g=1, loc="PTB")
    add("Electrical", "PTB", "Shift Engineer", 4, a=1, b=1, c=1, r=1, loc="PTB")
    add("Electrical", "PTB", "Floor Technician", 6, a=2, b=2, c=1, r=1, loc="PTB")
    add("Electrical", "PTB", "Substation Operator", 7, a=2, b=2, c=2, r=1, loc="PTB")
    add("Electrical", "PTB", "PPM Technician Sr", 2, g=2, loc="PTB")
    add("Electrical", "PTB", "PPM Technician", 3, g=3, loc="PTB")  # Sample=3 (plan sheet=2)
    add("Electrical", "PTB", "Helper (Tech-Ass-ELE)", 2, g=2, loc="PTB")
    # ALS
    add("Electrical", "ALS", "Shift Engineer", 4, a=1, b=1, c=1, r=1, loc="Airside")
    add("Electrical", "ALS", "Technician (MST)", 4, a=1, b=1, c=1, r=1, loc="Airside")
    add("Electrical", "ALS", "High Mast Operator", 2, b=1, c=1, loc="High Mast")  # Sample=2
    add("Electrical", "ALS", "High Mast PPM Sr Technician", 1, g=1, loc="High Mast")
    add("Electrical", "ALS", "PPM Technician", 1, g=1, loc="ALS")
    add("Electrical", "ALS", "Helper (Tech-Ass-ELE)", 1, g=1, loc="ALS")
    # Landside / MRSS / PPM (Sample layout)
    add("Electrical", "Landside", "Sr Technician", 1, g=1, loc="Landside")
    add("Electrical", "Landside", "Technician", 5, a=1, b=1, c=1, r=1, g=1, loc="Landside")
    add("Electrical", "Landside", "Helper (Tech-Ass-ELE)", 5, a=1, b=1, c=1, r=1, g=1, loc="Landside")
    add("Electrical", "MRSS", "MRSS Operator Sr Technician", 4, a=1, b=1, c=1, r=1, loc="MRSS")
    add("Electrical", "MRSS", "MRSS Operator Technician", 4, a=1, b=1, c=1, r=1, loc="MRSS")
    add("Electrical", "MRSS", "Helper (Tech-Ass-ELE)", 4, a=1, b=1, c=1, r=1, loc="MRSS")
    add("Electrical", "HT & LT PPM", "Engineer PM", 1, g=1, loc="HT & LT PPM")
    add("Electrical", "HT & LT PPM", "HT PPM Technician Sr", 3, g=3, loc="HT & LT PPM")
    add("Electrical", "HT & LT PPM", "LT PPM Technician Sr", 1, g=1, loc="HT & LT PPM")
    add("Electrical", "HT & LT PPM", "LT PPM Technician", 2, g=2, loc="HT & LT PPM")
    add("Electrical", "HT & LT PPM", "Solar Technician", 1, a=1, loc="Solar")
    add("Electrical", "HT & LT PPM", "Helper (Tech-Ass-ELE)", 2, g=2, loc="HT & LT PPM")

    # ===== E. Fire (17) =====
    add("Fire", "CFR", "Safety Officer", 2, g=2, loc="CFR SATELLITE")
    add("Fire", "CFR", "Supervisor", 4, a=1, b=1, c=1, r=1, loc="CFR SATELLITE")
    add("Fire", "CFR", "Fire Technician", 11, a=3, b=3, c=3, r=2, loc="CFR SATELLITE")

    # ===== F. Civil from Client Civil Final (57) =====
    add("Civil", "PTB/ALS", "Engr", 3, g=3, loc="PTB/LANDSIDE")
    add("Civil", "PTB/ALS", "Supervisor Civil", 8, g=4, a=1, b=1, c=1, r=1, loc="PTB/LANDSIDE")
    add("Civil", "PTB/ALS", "Mason", 7, g=3, a=1, b=1, c=1, r=1, loc="PTB/LANDSIDE")
    add("Civil", "PTB/ALS", "Carpenter", 5, g=1, a=1, b=1, c=1, r=1, loc="PTB/LANDSIDE")
    add("Civil", "PTB/ALS", "Civil Helper", 28, g=20, a=2, b=2, c=2, r=2, loc="PTB/LANDSIDE")
    add("Civil", "PTB/ALS", "Welder", 2, g=2, loc="PTB/LANDSIDE")
    add("Civil", "PTB/ALS", "Painter", 4, g=4, loc="PTB/LANDSIDE")

    # ===== G. Motor Transport (17) — use Total column (Boom/Sweeper Total=0 omitted) =====
    add("Motor Transport", "ADMIN", "MT Supervisor", 1, g=1, loc="ADMIN")
    add("Motor Transport", "ADMIN", "Motor Technician", 2, g=2, loc="ADMIN")
    add("Motor Transport", "ADMIN", "Motor Electrician", 2, g=2, loc="ADMIN")
    add("Motor Transport", "ADMIN", "Tyre Technician", 1, g=1, loc="ADMIN")
    add("Motor Transport", "ADMIN", "Driver (For CLR Vehicles)", 11, g=1, a=3, b=3, c=3, r=1, loc="ADMIN")

    # ===== H. Landside Facilities (11) =====
    add("Landside Facilities", "Landside", "Technician (HVAC) L", 1, g=1, loc="Landside")
    add("Landside Facilities", "Landside", "Electrician L", 4, g=2, b=1, r=1, loc="Landside")
    add("Landside Facilities", "Landside", "Plumber L", 1, g=1, loc="Landside")
    add("Landside Facilities", "Landside", "Sewerman & Drainage L", 2, g=1, r=1, loc="Landside")
    add("Landside Facilities", "Landside", "Helper L", 3, g=2, r=1, loc="Landside")

    return groups


def normalize(s: str) -> str:
    if not s:
        return ""
    s = str(s).upper().strip()
    s = s.replace("&", " AND ")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Mapping Existing designation (+ optional location hint) → list of (section, role) preferences
EXISTING_MAP_RULES = [
    # (match_fn(desig, loc) -> score, target section, target role)
]


def load_existing(path):
    wb = load_workbook(path, data_only=True)
    ws = wb["Exisiting"]
    people = []
    for r in range(4, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name:
            continue
        mobile = ws.cell(r, 3).value
        if isinstance(mobile, float) and mobile == int(mobile):
            mobile = str(int(mobile))
        elif mobile is not None:
            mobile = str(mobile).replace(" ", "")
        desig = ws.cell(r, 4).value
        loc = ws.cell(r, 5).value
        people.append({
            "name": str(name).strip(),
            "mobile": mobile,
            "desig": desig,
            "loc": loc,
            "norm_desig": normalize(desig),
            "norm_loc": normalize(loc),
            "used": False,
        })
    return people


def match_score(person, section, role, area, loc_default):
    """Heuristic score for assigning an existing person to a requirement slot."""
    nd = person["norm_desig"]
    nl = person["norm_loc"]
    nr = normalize(role)
    ns = normalize(section)
    na = normalize(area)
    nloc = normalize(loc_default)
    if not nd:
        return -1

    score = 0

    # Direct / strong designation matches
    rules = [
        # Site management
        ("SITE HEAD", "SITE HEAD", 100),
        ("HR MANAGER", "HR MANAGER", 100),
        ("HSE", "HSE AND COMPLIANCE MANAGER", 95),
        ("MIS EXECUTIVE", "MIS EXECUTIVE", 100),
        ("ADMIN SUPERVISOR", "MATERIALS MANAGER", 40),  # weak fallback only
        ("STORE", "STORES EXE", 90),
        # Helpdesk / BMS
        ("HELP DESK", "HELP DESK EXECUTIVE", 100),
        ("BMS", "BMS", 100),
        # Mechanical
        ("SHIFT ENGINEER MECHANICAL", "SHIFT ENGINEER", 95),
        ("SR TECHNICIAN MEC", "SR TECHNICIAN", 95),
        ("TECHNICIAN HVAC", "TECHNICIAN HVAC L", 90),
        ("TECH ASSOCIATE MECH", "TECH ASSOCIATE MECH", 100),
        ("FITTER MECHANICAL", "FITTER MECHANICAL", 100),
        ("PLUMBER", "PLUMBER", 90),
        ("STP OPERATOR", "STP OPERATOR", 100),
        ("WTP OPERATOR", "WTP OPERATOR", 100),
        ("STP HELPER", "WTP STP HELPER", 90),
        ("WTP HELPER", "WTP STP HELPER", 90),
        ("MECHANICAL HELPERS", "TECH ASSOCIATE MECH", 50),
        ("FITTER HELPER", "SEWERMAN AND DRAINAGE", 75),
        ("FITTER HELPER", "TECH ASSOCIATE MECH", 40),
        ("PUMP", "PUMP OPERATOR", 90),
        # Fire
        ("FIRE TECHNICIAN", "FIRE TECHNICIAN", 100),
        ("FIRE SUPERVISOR", "SUPERVISOR", 90),
        ("SAFETY OFFICER", "SAFETY OFFICER", 100),
        # Civil
        ("CIVIL SUPERVISOR", "SUPERVISOR CIVIL", 100),
        ("MASON", "MASON", 100),
        ("CARPENTER", "CARPENTER", 100),
        ("CIVIL HELPER", "CIVIL HELPER", 100),
        ("WELDER", "WELDER", 100),
        ("PAINTER", "PAINTER", 100),
        ("ENGR", "ENGR", 80),
        # Transport
        ("MT SUPERVISOR", "MT SUPERVISOR", 100),
        ("MT ELECTRICIAN", "MOTOR ELECTRICIAN", 100),
        ("MT TECHNICIAN", "MOTOR TECHNICIAN", 100),
        ("DRIVER", "DRIVER FOR CLR VEHICLES", 85),
        # Electrical AM
        ("AM ELECTRICAL", "ASST MANAGER", 100),
    ]

    for key, target, pts in rules:
        if key in nd and normalize(target) == nr:
            score = max(score, pts)
        if key in nd and normalize(target) in nr:
            score = max(score, pts - 5)

    # Electrical OPS / Sr OPS / Tech Assats — location based
    if nd in ("OPS", "SR OPS", "TECH ASSATS") or nd == "SHIFT ENGINEER":
        if section != "Electrical":
            return -1 if score < 50 else score
        # Asst already handled
        if nd == "SHIFT ENGINEER":
            if "SHIFT ENGINEER" in nr:
                score = max(score, 90)
                if nl and (nl in nloc or nloc in nl or (nl in ("PTB",) and area == "PTB")
                           or (nl in ("AIRSIDE", "AIR SIDE") and area == "ALS")
                           or (nl in ("LAND SIDE", "LANDSIDE") and area == "ALS")):
                    score += 15
            else:
                return -1
        elif nd == "SR OPS":
            if "MRSS OPERATOR SR" in nr or ("SR TECHNICIAN" in nr and area == "MRSS"):
                score = max(score, 95)
            elif "SR TECHNICIAN" in nr:
                score = max(score, 70)
            elif "PPM TECHNICIAN SR" in nr or "HIGH MAST PPM" in nr:
                score = max(score, 60)
            else:
                return -1
        elif nd == "TECH ASSATS":
            if "HELPER" in nr or "TECH ASS" in nr:
                score = max(score, 80)
                if nl == "MRSS" and area == "MRSS":
                    score += 20
                elif nl == "PTB" and area == "PTB":
                    score += 20
                elif nl == "ALL" and area in ("HT & LT PPM", "Landside", "ALS"):
                    score += 5
            else:
                return -1
        elif nd == "OPS":
            # Map by location to electrical roles
            loc_role_prefs = []
            if nl == "MRSS":
                loc_role_prefs = ["MRSS OPERATOR TECHNICIAN", "MRSS OPERATOR SR", "HELPER"]
            elif nl == "PTB":
                loc_role_prefs = ["FLOOR TECHNICIAN", "SUBSTATION OPERATOR", "TECHNICIAN", "HELPER"]
            elif nl == "ADMIN":
                loc_role_prefs = ["TECHNICIAN", "SUBSTATION", "HELPER"]
            elif "AGL" in nl:
                loc_role_prefs = ["TECHNICIAN MST", "HIGH MAST OPERATOR", "HELPER"]
            elif nl == "ATC":
                loc_role_prefs = ["TECHNICIAN", "SUBSTATION", "HELPER"]
            elif nl == "UTILITY":
                loc_role_prefs = ["TECHNICIAN", "HELPER", "FLOOR"]
            elif nl == "ALL":
                loc_role_prefs = ["TECHNICIAN", "HELPER", "PPM", "SR TECHNICIAN"]
            else:
                loc_role_prefs = ["TECHNICIAN", "HELPER", "FLOOR", "SUBSTATION", "HIGH MAST"]

            matched = False
            for i, pref in enumerate(loc_role_prefs):
                if pref in nr or any(p in nr for p in pref.split()):
                    # stronger if primary tokens match
                    if pref in nr:
                        score = max(score, 85 - i * 5)
                        matched = True
                    elif "TECHNICIAN" in pref and "TECHNICIAN" in nr and "PPM" not in nr and "MST" not in pref:
                        score = max(score, 75 - i * 5)
                        matched = True
                    elif "FLOOR" in pref and "FLOOR" in nr:
                        score = max(score, 90)
                        matched = True
                    elif "SUBSTATION" in pref and "SUBSTATION" in nr:
                        score = max(score, 90)
                        matched = True
                    elif "HIGH MAST" in pref and "HIGH MAST OPERATOR" in nr:
                        score = max(score, 90)
                        matched = True
                    elif "MST" in pref and "MST" in nr:
                        score = max(score, 90)
                        matched = True
                    elif "HELPER" in pref and "HELPER" in nr:
                        score = max(score, 70 - i * 3)
                        matched = True
                    elif "MRSS OPERATOR TECHNICIAN" in nr and nl == "MRSS":
                        score = max(score, 95)
                        matched = True
            if not matched:
                # generic OPS into technician/floor/substation/helper only
                if any(x in nr for x in ("FLOOR", "SUBSTATION", "TECHNICIAN", "HELPER", "HIGH MAST", "MST")):
                    score = max(score, 40)
                else:
                    return -1

    # Section gates for non-electrical
    if section == "Mechanical & HVAC":
        mech_keys = ("MECH", "HVAC", "PLUMBER", "WTP", "STP", "FITTER", "PUMP", "SEWER",
                     "TECH ASSOCIATE", "SHIFT ENGINEER MECHANICAL", "SR TECHNICIAN MEC",
                     "TECHNICIAN HVAC", "MECHANICAL HELPER", "FITTER HELPER")
        if not any(k in nd for k in mech_keys) and nd not in ("SHIFT ENGINEER",):
            # Shift Engineer without MECHANICAL tag — only if already scored for mech SE
            if score < 50:
                return -1
        if nd == "SHIFT ENGINEER" and "MECHANICAL" not in nd:
            # electrical/airside shift engineers — don't steal for mech unless weak
            if person["norm_loc"] in ("PTB", "AIRSIDE", "LAND SIDE", "LANDSIDE"):
                if "SHIFT ENGINEER" in nr and section == "Mechanical & HVAC":
                    return -1  # those are electrical SEs in Existing
        if nd == "SHIFT ENGINEER MECHANICAL" or "SHIFT ENGINEER MECHANICAL" in nd:
            if "SHIFT ENGINEER" in nr and section == "Mechanical & HVAC":
                score = max(score, 100)
        if nd == "PLUMBER" and "PLUMBER L" in nr:
            score = max(score, 40)  # prefer main plumber first
        if nd == "PLUMBER" and nr == normalize("Plumber"):
            score = max(score, 95)
        if "TECHNICIAN HVAC" in nd:
            if "TECHNICIAN HVAC L" in nr and section == "Mechanical & HVAC":
                score = max(score, 92)
            if section == "Landside Facilities" and "TECHNICIAN HVAC L" in nr:
                score = max(score, 55)

    if section == "Civil":
        if not any(k in nd for k in ("CIVIL", "MASON", "CARPENTER", "PAINTER", "WELDER", "ENGR")):
            return -1

    if section == "Fire":
        if "FIRE" not in nd and "SAFETY" not in nd:
            return -1

    if section == "Motor Transport":
        if not any(k in nd for k in ("MT ", "MT", "DRIVER", "TYRE", "BOOM")):
            # DRIVER alone ok
            if "DRIVER" not in nd and "MT" not in nd:
                return -1

    if section == "Site Management":
        mgr_ok = any(k in nd for k in (
            "SITE HEAD", "HR MANAGER", "HSE", "MIS", "STORE", "MATERIALS",
            "SR MANAGER", "ADMIN SUPERVISOR"
        ))
        if not mgr_ok:
            return -1

    if section == "SCADA & Helpdesk":
        if not any(k in nd for k in ("HELP DESK", "BMS", "SCADA")):
            return -1

    if section == "Landside Facilities":
        # only leftover / L-tagged; prefer low score unless explicit
        if "PLUMBER" in nd and "PLUMBER L" in nr:
            score = max(score, 45)
        if "TECHNICIAN HVAC" in nd and "TECHNICIAN HVAC L" in nr:
            score = max(score, 50)

    # Location bonus
    if nl and nloc and (nl in nloc or nloc in nl):
        score += 8
    if nl and na and (nl in na or na in nl):
        score += 5

    return score if score > 0 else -1


def assign_people(groups, people):
    """Greedy assign existing people to slots by best score."""
    slots = []  # flat list of slot dicts
    for gi, g in enumerate(groups):
        for i in range(1, g["required"] + 1):
            slots.append({
                "gi": gi,
                "idx": i,
                "section": g["section"],
                "area": g["area"],
                "role": g["role"],
                "client_req": f"{g['role']}-{i}",
                "required": g["required"],
                "plan": g["plan"],
                "location": g["location"],
                "person": None,
            })

    # Build candidate scores
    candidates = []  # (score, si, pi)
    for si, slot in enumerate(slots):
        for pi, person in enumerate(people):
            sc = match_score(person, slot["section"], slot["role"], slot["area"], slot["location"])
            if sc > 0:
                candidates.append((sc, si, pi))
    candidates.sort(reverse=True)

    used_p = set()
    used_s = set()
    for sc, si, pi in candidates:
        if si in used_s or pi in used_p:
            continue
        slots[si]["person"] = people[pi]
        people[pi]["used"] = True
        used_s.add(si)
        used_p.add(pi)

    return slots


# ---------------------------------------------------------------------------
# Shift generation
# ---------------------------------------------------------------------------

def solve_shift_team(plan, n_people, label=""):
    """Return list of dict day->shift for n_people covering plan A/B/C/G/R."""
    need = {s: int(plan.get(s, 0)) for s in ["A", "B", "C", "G"]}
    R = int(plan.get("R", 0))
    assert n_people == sum(need.values()) + R, (label, n_people, plan)

    a, b, c = need["A"], need["B"], need["C"]
    n_g = need["G"]
    results = [None] * n_people

    # Case: no A/B/C — everyone is General; R means that many WO per day (staggered)
    if a + b + c == 0:
        n_all = n_g + R
        assert n_all == n_people
        # Exactly R people on WO each day (or 1 if R=0 and we still want weekly rest)
        offs_per_day = R if R > 0 else (1 if n_all > 1 else 0)
        for i in range(n_all):
            seq = {}
            for d in range(DAYS):
                if offs_per_day == 0:
                    # single G person: WO every 7th day
                    seq[d] = "W/O" if (n_all == 1 and d % 7 == 6) else "G"
                else:
                    # rotating block of offs_per_day
                    off_set = {(d * offs_per_day + k) % n_all for k in range(offs_per_day)}
                    seq[d] = "W/O" if i in off_set else "G"
            results[i] = seq
        return results

    # Mixed / pure shift teams
    n_shift = a + b + c + R
    # G staff first (staggered WO among G only)
    if n_g:
        for i in range(n_g):
            seq = {}
            for d in range(DAYS):
                if n_g == 1:
                    seq[d] = "W/O" if (d % 7 == 6) else "G"
                else:
                    seq[d] = "W/O" if (d % n_g == i) else "G"
            results[i] = seq

    if n_shift == 0:
        return results

    shift_idx = list(range(n_g, n_g + n_shift))

    # Fixed ABC no reliever
    if R == 0 and a + b + c == n_shift:
        pool = ["A"] * a + ["B"] * b + ["C"] * c
        for j, sh in enumerate(pool):
            results[shift_idx[j]] = {d: sh for d in range(DAYS)}
        return results

    # Case: some A/B/C with R, possibly mixed with G already handled
    # Special: Electrician L style G + B + R without full ABC — treat non-G as mini shift team
    assign = _solve_ilp(a, b, c, R, n_shift, label)
    if assign is None:
        assign = _heuristic_roster(a, b, c, R, n_shift)

    for j in range(n_shift):
        results[shift_idx[j]] = assign[j]
    return results


def _solve_ilp(a, b, c, R, n, label):
    need = {"A": a, "B": b, "C": c}
    W = rest_window(n, R)
    try:
        prob = pulp.LpProblem(f"R_{re.sub(r'[^A-Za-z0-9]', '_', label)[:30]}", pulp.LpMinimize)
        x = {(i, d, s): pulp.LpVariable(f"x{i}_{d}_{s}", cat="Binary")
             for i in range(n) for d in range(DAYS) for s in ["A", "B", "C", "W/O"]}
        for i in range(n):
            for d in range(DAYS):
                prob += pulp.lpSum(x[i, d, s] for s in ["A", "B", "C", "W/O"]) == 1
        for d in range(DAYS):
            for s in ["A", "B", "C"]:
                prob += pulp.lpSum(x[i, d, s] for i in range(n)) == need[s]
            if R > 0:
                prob += pulp.lpSum(x[i, d, "W/O"] for i in range(n)) == R
        # no C→A
        for i in range(n):
            for d in range(DAYS - 1):
                prob += x[i, d, "C"] + x[i, d + 1, "A"] <= 1
                # no consecutive WO when possible
                if R * 2 <= n:  # enough people
                    prob += x[i, d, "W/O"] + x[i, d + 1, "W/O"] <= 1
        # fair WO
        if R > 0:
            total_wo = R * DAYS
            base, rem = divmod(total_wo, n)
            for i in range(n):
                wc = pulp.lpSum(x[i, d, "W/O"] for d in range(DAYS))
                if rem:
                    prob += wc >= base
                    prob += wc <= base + 1
                else:
                    prob += wc == base
        if W is not None:
            for i in range(n):
                for start in range(0, DAYS - W + 1):
                    prob += pulp.lpSum(x[i, start + k, "W/O"] for k in range(W)) >= 1
        # soft prefer A→B→C→WO→A
        obj = []
        for i in range(n):
            for d in range(DAYS - 1):
                for s1, s2, w in [("A", "B", -10), ("B", "C", -10), ("C", "W/O", -12), ("W/O", "A", -8),
                                  ("C", "B", 15), ("B", "A", 12)]:
                    y = pulp.LpVariable(f"y{i}_{d}_{s1}_{s2}", cat="Binary")
                    prob += y <= x[i, d, s1]
                    prob += y <= x[i, d + 1, s2]
                    obj.append(w * y)
        prob += pulp.lpSum(obj)
        # time limit scales with size
        tlim = 30 if n <= 6 else (60 if n <= 12 else 120)
        st = pulp.LpStatus[prob.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=tlim))]
        if st not in ("Optimal", "Feasible"):
            return None
        out = []
        for i in range(n):
            seq = {}
            for d in range(DAYS):
                for s in ["A", "B", "C", "W/O"]:
                    if pulp.value(x[i, d, s]) and pulp.value(x[i, d, s]) > 0.5:
                        seq[d] = s
                        break
            out.append(seq)
        return out
    except Exception as e:
        print("ILP fail", label, e)
        return None


def _heuristic_roster(a, b, c, R, n):
    """Forward-cycle heuristic guaranteeing coverage and no C→A."""
    # Classic: people cycle through positions in blocks.
    # Build daily assignment lists.
    people_seq = [{d: None for d in range(DAYS)} for _ in range(n)]
    # Order people on a ring; each day take A,B,C from consecutive, next R off
    # Rotate starting offset each day by 1 for fairness when R small, or by block.

    # Better: use repeating pattern length = n
    # Position roles in a circle of size n: first a are A, next b B, next c C, next R WO
    # Each day rotate by 1.
    base = ["A"] * a + ["B"] * b + ["C"] * c + ["W/O"] * R
    assert len(base) == n
    for d in range(DAYS):
        rot = base[d % n:] + base[: d % n]
        # Ensure no person gets C then A: check person i
        # With rotate-by-1, person moves base[(i-d)%n] — transition from role at offset to offset-1
        for i in range(n):
            people_seq[i][d] = rot[i]

    # Fix C→A violations by swapping with someone on B or WO next day
    for i in range(n):
        for d in range(DAYS - 1):
            if people_seq[i][d] == "C" and people_seq[i][d + 1] == "A":
                # find j with d+1 in {B, W/O, C} and preferably d != making new C→A
                swapped = False
                for j in range(n):
                    if j == i:
                        continue
                    if people_seq[j][d + 1] in ("B", "W/O", "C"):
                        # swap day d+1
                        people_seq[i][d + 1], people_seq[j][d + 1] = people_seq[j][d + 1], people_seq[i][d + 1]
                        # verify counts still ok automatically since swap
                        if people_seq[i][d] == "C" and people_seq[i][d + 1] == "A":
                            # revert
                            people_seq[i][d + 1], people_seq[j][d + 1] = people_seq[j][d + 1], people_seq[i][d + 1]
                            continue
                        if d + 1 < DAYS - 1 and people_seq[j][d] == "C" and people_seq[j][d + 1] == "A":
                            people_seq[i][d + 1], people_seq[j][d + 1] = people_seq[j][d + 1], people_seq[i][d + 1]
                            continue
                        swapped = True
                        break
                if not swapped:
                    # force to WO or B by swapping with any B/WO
                    for j in range(n):
                        if people_seq[j][d + 1] in ("B", "W/O"):
                            people_seq[i][d + 1], people_seq[j][d + 1] = people_seq[j][d + 1], people_seq[i][d + 1]
                            break
    return people_seq


def generate_all_shifts(groups, slots):
    """Generate shifts per group and attach to slots."""
    # group slots by gi
    by_gi = defaultdict(list)
    for si, slot in enumerate(slots):
        by_gi[slot["gi"]].append(si)

    for gi, g in enumerate(groups):
        sis = by_gi[gi]
        assert len(sis) == g["required"]
        plan = g["plan"]
        label = f"{g['section']}_{g['area']}_{g['role']}"
        print(f"  Roster {label}: n={g['required']} plan={plan}")
        seqs = solve_shift_team(plan, g["required"], label=label)
        for j, si in enumerate(sis):
            slots[si]["shifts"] = seqs[j]
            # validate no C→A/G for this person
            for d in range(DAYS - 1):
                if seqs[j][d] == "C" and seqs[j][d + 1] in ("A", "G"):
                    print(f"    WARN C→{seqs[j][d+1]} in {label} person {j} day {d+1}")


def validate_coverage(groups, slots):
    issues = []
    ca = 0
    by_gi = defaultdict(list)
    for slot in slots:
        by_gi[slot["gi"]].append(slot)
        seq = slot["shifts"]
        for d in range(DAYS - 1):
            if seq[d] == "C" and seq[d + 1] in ("A", "G"):
                ca += 1
                issues.append(f"C→{seq[d+1]} {slot['client_req']}")

    for gi, g in enumerate(groups):
        plan = g["plan"]
        for d in range(DAYS):
            cnt = Counter(slots[si]["shifts"][d] for si in range(len(slots)) if slots[si]["gi"] == gi)
            # recount via by_gi
            cnt = Counter(s["shifts"][d] for s in by_gi[gi])
            for s in ["A", "B", "C"]:
                if cnt[s] != plan[s]:
                    issues.append(f"COV {g['role']}/{g['area']} day{d+1} {s}={cnt[s]} want {plan[s]}")
            # G on-duty may be plan G or plan G-1 when staggered WO among G-only or mixed
            # For mixed G+shift, G people are first n_g — their WO reduces G count
            # Accept G count between max(0, planG-1) and planG when G staff stagger
            if plan["G"]:
                if cnt["G"] > plan["G"] or cnt["G"] < max(0, plan["G"] - (1 if plan["G"] else 0)):
                    # only flag if wildly off
                    if abs(cnt["G"] - plan["G"]) > 1:
                        issues.append(f"GCOV {g['role']} day{d+1} G={cnt['G']} plan={plan['G']}")
            if plan["R"]:
                if cnt["W/O"] != plan["R"] + (1 if plan["G"] > 1 and False else 0):
                    # shift WO should equal R; G WO is extra among G staff
                    # Count WO among shift people only
                    n_g = plan["G"]
                    shift_slots = by_gi[gi][n_g:]
                    wo_shift = sum(1 for s in shift_slots if s["shifts"][d] == "W/O")
                    if wo_shift != plan["R"]:
                        issues.append(f"WO {g['role']}/{g['area']} day{d+1} shiftWO={wo_shift} want R={plan['R']}")

    return ca, issues


# ---------------------------------------------------------------------------
# Excel writer (Sample format)
# ---------------------------------------------------------------------------

def write_workbook(groups, slots, people, ca, issues):
    wb = Workbook()
    ws = wb.active
    ws.title = "August 2026 Roster"

    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    fills = {
        "A": PatternFill("solid", fgColor="C6EFCE"),
        "B": PatternFill("solid", fgColor="FFEB9C"),
        "C": PatternFill("solid", fgColor="BDD7EE"),
        "G": PatternFill("solid", fgColor="E2D5F1"),
        "W/O": PatternFill("solid", fgColor="F4B183"),
    }
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    title_font = Font(bold=True, size=13, color="1F4E79", name="Calibri")
    section_fill = PatternFill("solid", fgColor="D6DCE4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")

    # Row 1 headers
    headers = [
        "Sr No", "Required", "Filled", "Client Requirement",
        "Name of the Person", "Contact Number", "Designation", "Location", "Month",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin
    ws.cell(1, 9, "Month")
    ws.cell(1, 10, date(YEAR, MONTH, 1))
    ws.merge_cells(start_row=1, start_column=10, end_row=1, end_column=10 + DAYS - 1)
    for col in range(10, 10 + DAYS):
        ws.cell(1, col).fill = header_fill
        ws.cell(1, col).font = header_font
        ws.cell(1, col).border = thin
    ws.cell(1, 10).alignment = center
    ws.cell(1, 10).number_format = "MMM-YYYY"

    # Row 2 Date numbers, Row 3 Weekdays — but Sample puts first person on row 2 with Date labels.
    # We'll use: Row2 = Date, Row3 = Week, then data from row 4 (clearer than Sample quirk).
    # User asked to follow Sample format — Sample has person1 on row2 with Date in col9.
    # We'll do: Row2 Date labels, Row3 Week labels, data from 4 — still Sample columns.

    ws.cell(2, 9, "Date").font = Font(bold=True, size=9)
    ws.cell(3, 9, "Week").font = Font(bold=True, size=9)
    for i, dt in enumerate(DATES):
        c = ws.cell(2, 10 + i, dt.day)
        c.alignment = center
        c.border = thin
        c.fill = header_fill
        c.font = header_font
        w = ws.cell(3, 10 + i, calendar.day_abbr[dt.weekday()].upper()[:3])
        # Sample uses WEN for Wed
        if dt.weekday() == 2:
            w.value = "WEN"
        w.alignment = center
        w.border = thin
        w.fill = PatternFill("solid", fgColor="2E75B6")
        w.font = header_font

    for col in range(1, 10):
        for r in (2, 3):
            ws.cell(r, col).border = thin
            ws.cell(r, col).fill = PatternFill("solid", fgColor="D6DCE4")

    # Precompute filled per group
    filled_by_gi = Counter()
    for slot in slots:
        if slot["person"]:
            filled_by_gi[slot["gi"]] += 1

    row = 4
    sno = 1
    prev_section = None
    by_gi = defaultdict(list)
    for slot in slots:
        by_gi[slot["gi"]].append(slot)

    for gi, g in enumerate(groups):
        if prev_section is not None and g["section"] != prev_section:
            # blank separator row
            row += 1
        if g["section"] != prev_section:
            ws.cell(row, 1, g["section"]).font = Font(bold=True, size=11, color="1F4E79")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
            for col in range(1, 10 + DAYS):
                ws.cell(row, col).fill = section_fill
            row += 1
            # area note for electrical
            prev_section = g["section"]

        # area subheader for electrical / when area changes within section
        group_slots = by_gi[gi]
        first_row = row
        for j, slot in enumerate(group_slots):
            ws.cell(row, 1, sno).alignment = center
            if j == 0:
                ws.cell(row, 2, g["required"]).alignment = center
                ws.cell(row, 3, filled_by_gi[gi]).alignment = center
            ws.cell(row, 4, slot["client_req"]).alignment = left
            person = slot["person"]
            if person:
                ws.cell(row, 5, person["name"])
                ws.cell(row, 6, person["mobile"])
                # Designation: use client role (clean) but keep existing desig in mind
                ws.cell(row, 7, g["role"])
                ws.cell(row, 8, person["loc"] or g["location"])
            else:
                ws.cell(row, 5, None)
                ws.cell(row, 6, None)
                ws.cell(row, 7, g["role"])
                ws.cell(row, 8, g["location"])
            for col in range(1, 10):
                ws.cell(row, col).border = thin
                ws.cell(row, col).font = Font(size=9, name="Calibri")
            for d in range(DAYS):
                sh = slot["shifts"][d]
                cell = ws.cell(row, 10 + d, sh)
                cell.fill = fills[sh]
                cell.alignment = center
                cell.border = thin
                cell.font = Font(size=8, bold=True, name="Calibri")
            sno += 1
            row += 1
        # merge Required / Filled for group
        if g["required"] > 1:
            ws.merge_cells(start_row=first_row, start_column=2, end_row=first_row + g["required"] - 1, end_column=2)
            ws.merge_cells(start_row=first_row, start_column=3, end_row=first_row + g["required"] - 1, end_column=3)
            ws.cell(first_row, 2).alignment = center
            ws.cell(first_row, 3).alignment = center

    widths = [6, 10, 8, 28, 24, 14, 22, 16, 8] + [4] * DAYS
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "J4"
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 18

    # Coverage Summary
    ws2 = wb.create_sheet("Coverage Summary", 1)
    ws2["A1"] = "GVIAL Bhogapuram — 281 Manpower August 2026 Duty Roster — Coverage Summary"
    ws2["A1"].font = title_font
    ws2.merge_cells("A1:L1")
    ws2["A2"] = (
        "Electrical from Client_Electrical_Requirement + Sample layout | "
        "Civil from Client_Civil_Requirement Final | "
        "Other departments from Client_Manpower_Requirement | "
        "Existing names from Exisiting tab"
    )
    ws2["A2"].font = Font(italic=True, size=9)

    headers2 = ["Section", "Area", "Designation", "G", "A", "B", "C", "R", "Required",
                "Filled", "Vacant", "Rest window", "Notes"]
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(4, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin

    r = 5
    total_req = total_filled = 0
    for gi, g in enumerate(groups):
        plan = g["plan"]
        filled = filled_by_gi[gi]
        n_shift = plan["A"] + plan["B"] + plan["C"] + plan["R"]
        W = rest_window(n_shift, plan["R"]) if n_shift else None
        if plan["R"] == 0 and n_shift:
            note = "No reliever — fixed A/B/C (no WO) to keep 24x7 cover"
            Wstr = "N/A"
        elif W == 7:
            note = "6 work + 1 WO hard target"
            Wstr = 7
        elif W:
            note = f"Reliever short for weekly off all; WO every {W} days"
            Wstr = W
        else:
            note = "General shift staggered WO"
            Wstr = "G-stag"
        vals = [g["section"], g["area"], g["role"], plan["G"], plan["A"], plan["B"], plan["C"],
                plan["R"], g["required"], filled, g["required"] - filled, Wstr, note]
        for col, v in enumerate(vals, 1):
            cell = ws2.cell(r, col, v)
            cell.border = thin
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        total_req += g["required"]
        total_filled += filled
        r += 1

    r += 1
    ws2.cell(r, 1, f"TOTAL REQUIRED = {total_req} | FILLED (from Existing) = {total_filled} | VACANT = {total_req - total_filled}")
    ws2.cell(r, 1).font = Font(bold=True, size=11, color="C00000")
    r += 2
    ws2.cell(r, 1, "RULES").font = Font(bold=True, color="C00000")
    r += 1
    for line in [
        "1. Daily A/B/C headcount matches the plan for each designation group (24x7).",
        "2. Reliever (R) = number of people on W/O that day among the shift team.",
        "3. HARD: After C, next day is never A or G.",
        "4. Preferred rotation A → B → C → W/O → A.",
        "5. Where Reliever×7 ≥ shift team size → everyone gets ≥1 Week Off in every 7 days.",
        "6. Electrical follows Client Electrical / Sample (not Client_Manpower Electrical rollup).",
        "7. Civil follows Client Civil Final Requirement ALS & PTB (57).",
        "8. Vacant rows left with blank Name/Contact so new joiners can be filled later.",
        f"9. C→A/G violations in generated roster: {ca}",
        f"10. Unassigned Existing people: {sum(1 for p in people if not p['used'])}",
    ]:
        ws2.cell(r, 1, line)
        ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        r += 1

    if issues[:20]:
        r += 1
        ws2.cell(r, 1, "Validation notes (first 20):").font = Font(bold=True)
        r += 1
        for iss in issues[:20]:
            ws2.cell(r, 1, iss)
            r += 1

    for col in range(1, 14):
        ws2.column_dimensions[get_column_letter(col)].width = 14
    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["C"].width = 28
    ws2.column_dimensions["M"].width = 55

    # Unassigned existing
    ws3 = wb.create_sheet("Unassigned Existing", 2)
    ws3["A1"] = "Existing people not auto-matched (review & place manually if needed)"
    ws3["A1"].font = title_font
    for col, h in enumerate(["Name", "Mobile", "Existing Designation", "Location"], 1):
        cell = ws3.cell(3, col, h)
        cell.fill = header_fill
        cell.font = header_font
    rr = 4
    for p in people:
        if not p["used"]:
            ws3.cell(rr, 1, p["name"])
            ws3.cell(rr, 2, p["mobile"])
            ws3.cell(rr, 3, p["desig"])
            ws3.cell(rr, 4, p["loc"])
            rr += 1
    if rr == 4:
        ws3.cell(4, 1, "(All existing people were matched)")

    # How to use
    ws4 = wb.create_sheet("How to use", 3)
    ws4["A1"] = "How to use"
    ws4["A1"].font = title_font
    for i, line in enumerate([
        "",
        "Shifts: A=Shift-1 (06:00-14:00) | B=Shift-2 (14:00-22:00) | C=Shift-3 (22:00-06:00) | G=General (09:00-18:00) | W/O=Weekly Off",
        "Fill blank Name / Contact for vacant positions when staff join — do not change daily A/B/C counts.",
        "Required / Filled columns show client requirement vs names placed from Existing tab.",
        "Never put A or G the day after C for the same person.",
        "Source file: 281_ShiftRoster.xlsx (Client Electrical, Client Civil, Client_Manpower_Requirement, Exisiting, Sample).",
    ], 1):
        ws4.cell(i, 1, line)
    ws4.column_dimensions["A"].width = 110

    # Copy source reference sheets (values only summary already done) — attach original tabs
    src = load_workbook(SOURCE, data_only=False)
    for name in src.sheetnames:
        new_name = f"SRC_{name}"[:31]
        src_ws = src[name]
        dst = wb.create_sheet(new_name)
        for row in src_ws.iter_rows():
            for cell in row:
                dst.cell(cell.row, cell.column, cell.value)

    wb.save(OUT)
    print("Saved", OUT)
    print(f"Total slots={len(slots)} filled={sum(1 for s in slots if s['person'])} ca={ca} issues={len(issues)}")


def main():
    groups = build_requirement_groups()
    total = sum(g["required"] for g in groups)
    print("Total required slots:", total)
    assert total == 281, total

    people = load_existing(SOURCE)
    print("Existing people:", len(people))

    slots = assign_people(groups, people)
    filled = sum(1 for s in slots if s["person"])
    print(f"Matched {filled}/{len(slots)}; unmatched existing={sum(1 for p in people if not p['used'])}")

    print("Generating shifts...")
    generate_all_shifts(groups, slots)
    ca, issues = validate_coverage(groups, slots)
    print(f"Validation C→A/G={ca} coverage_issues={len(issues)}")
    for iss in issues[:30]:
        print(" ", iss)

    write_workbook(groups, slots, people, ca, issues)


if __name__ == "__main__":
    main()
