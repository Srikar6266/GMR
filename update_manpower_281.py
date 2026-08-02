#!/usr/bin/env python3
"""
Update MANPOWER.xlsx from 281_DailySheet (source of truth).

Rules:
- Do not invent people, designations, or shift patterns.
- Use intelligent fuzzy matching for human spelling variants.
- Recalculate 281ManPowerSummary filled/shortfall/surplus from DailySheet.
- Refresh 172Phase1Deployment contacts/joining; fill vacant slots only with
  clear designation mappings from unused DailySheet people.
- Recalculate 172Phase1Summary from the updated 172 deployment.
- Create 281ManpowerDeployment as a format-replica of 172Phase1Deployment
  with exactly 281 required slots from Plan281 / 281ManPowerSummary.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = "MANPOWER.xlsx"
OUT = "MANPOWER.xlsx"

# ───────────────────────── name / designation helpers ─────────────────────────

def norm_name(n) -> str:
    if n is None:
        return ""
    n = str(n).strip().lower()
    n = re.sub(r"[^a-z0-9\s]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n


def name_tokens(n) -> set[str]:
    # drop ultra-short tokens that cause false matches (initials)
    return {t for t in norm_name(n).split() if len(t) >= 2}


def edit_distance(a: str, b: str) -> int:
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    # banded DP for short strings
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        cur = [i]
        for j, cb in enumerate(b, start=1):
            ins = cur[j - 1] + 1
            delete = prev[j] + 1
            sub = prev[j - 1] + (0 if ca == cb else 1)
            cur.append(min(ins, delete, sub))
        prev = cur
    return prev[-1]


def names_match(a, b) -> bool:
    """Fuzzy match tolerant of human spelling variants; guards against false merges."""
    na, nb = norm_name(a), norm_name(b)
    if not na or not nb:
        return False
    if na == nb:
        return True

    ca, cb = na.replace(" ", ""), nb.replace(" ", "")
    if ca == cb:
        return True

    # Close spelling (Thallapudi/Tallapudi, Dhakamarri/Dakamarri, vasudeva/vasuseva)
    # Allow up to 3 edits for longer compacted names (extra letters + typo combos)
    if min(len(ca), len(cb)) >= 10 and abs(len(ca) - len(cb)) <= 3:
        max_ed = 2 if min(len(ca), len(cb)) < 18 else 3
        if edit_distance(ca, cb) <= max_ed:
            return True

    # Compact containment for concatenated tokens
    if min(len(ca), len(cb)) >= 8 and abs(len(ca) - len(cb)) <= 2:
        if ca in cb or cb in ca:
            return True

    ta, tb = name_tokens(a), name_tokens(b)
    if not ta or not tb:
        return False

    fa, fb = na.split()[0], nb.split()[0]

    def first_names_conflict() -> bool:
        # Different substantial first names => not the same person
        if len(fa) >= 3 and len(fb) >= 3 and fa != fb:
            if fa not in nb and fb not in na and edit_distance(fa, fb) > 1:
                return True
        # Initial vs full first name: initial must match first letter
        if len(fa) <= 2 and len(fb) >= 3 and not fb.startswith(fa[0]):
            return True
        if len(fb) <= 2 and len(fa) >= 3 and not fa.startswith(fb[0]):
            return True
        return False

    # Full-string containment (Pothirendi Eswara Rao vs Eswara Rao)
    if na in nb or nb in na:
        shorter = ta if len(ta) <= len(tb) else tb
        if len(shorter) >= 2 and not first_names_conflict():
            return True

    shorter, longer = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    if shorter <= longer and len(shorter) >= 2 and not first_names_conflict():
        return True

    inter = ta & tb
    if (
        len(inter) >= 2
        and len(inter) / min(len(ta), len(tb)) >= 0.8
        and not first_names_conflict()
    ):
        return True

    # Same long surname + other-token edit-close (Thallapudi~Tallapudi)
    if ta and tb and not first_names_conflict():
        la, lb = sorted(ta, key=len)[-1], sorted(tb, key=len)[-1]
        if la == lb and len(la) >= 5:
            rest_a, rest_b = ta - {la}, tb - {lb}
            for xa in rest_a:
                for xb in rest_b:
                    if xa == xb or (min(len(xa), len(xb)) >= 4 and edit_distance(xa, xb) <= 1):
                        return True
                    if min(len(xa), len(xb)) >= 4 and (xa.startswith(xb[:4]) or xb.startswith(xa[:4])):
                        return True

    # Initial + surname vs full given name + same surname (S. Vasu <-> Sunkara vasu)
    pa, pb = na.split(), nb.split()
    if len(pa) >= 2 and len(pb) >= 2:
        if len(pa[0]) <= 2 and len(pa[-1]) >= 3 and pa[-1] == pb[-1]:
            if pb[0].startswith(pa[0][0]) and len(pb[0]) >= 3:
                return True
        if len(pb[0]) <= 2 and len(pb[-1]) >= 3 and pb[-1] == pa[-1]:
            if pa[0].startswith(pb[0][0]) and len(pa[0]) >= 3:
                return True
    return False


def canon_desig(d):
    if not d:
        return None
    s = re.sub(r"\s+", " ", str(d).strip())
    key = s.lower().replace("( ", "(").replace(" )", ")")
    key2 = re.sub(r"\s+", " ", key.replace("(", " (")).strip()
    key_compact = key.replace(" ", "")
    MAP = {
        "hr manager": "HR Manager",
        "mis executive": "MIS Executive",
        "wtp operator": "WTP Operator",
        "supervisor(fire)": "Supervisor (Fire)",
        "supervisor (fire)": "Supervisor (Fire)",
        "technician (hvac) l": "Technician (HVAC) L",
        "sr.technician ops": "Sr. Technician Ops",
        "sr technician ops": "Sr. Technician Ops",
        "sr. technician ops": "Sr. Technician Ops",
        "sr hvac technician mechanical": "Sr Technician (Mechanical)",
        "hvac technician mechanical": "HVAC Technician Mechanical",
        "ups technician": "UPS Technician",
    }
    if key in MAP:
        return MAP[key]
    if key2 in MAP:
        return MAP[key2]
    if key_compact == "supervisor(fire)":
        return "Supervisor (Fire)"
    return s


# DailySheet designation → Phase-1 172 roster designations it may fill.
# Conservative: only clear role equivalents used in prior matching notes.
DAILY_TO_172 = {
    "Site Head (GM)": ["Site Head (GM)"],
    "Sr. Manager (Mechanical)": ["Sr. Manager (Mechanical)"],
    "Sr. Manager (Electrical)": ["Sr. Manager (Electrical)"],
    "Sr. Manager (Civil)": ["Sr. Manager (Civil)"],
    "HSE & Compliance Manager": ["HSE & Compliance Manager"],
    "HR Manager": ["HR Manager"],
    "MIS Executive": ["MIS Executive"],
    "Materials Manager": ["Materials Manager"],
    "Civil Engineer": ["Engr"],
    "Supervisor Civil": ["Supervisor Civil"],
    "Carpenter": ["Carpenter"],
    "Welder": ["Welder"],
    "Civil Fitter": ["Civil Fitter"],
    "Civil Helper": ["Civil Helper", "Helper Utility building"],
    "Mason": ["Mason"],
    "Asst. Manager (Electrical)": ["Asst Manager"],
    "Shift Engineer (Electrical)": ["Shift Engineers"],
    "Electrician - Ops": [
        "Floor Technicians",
        "Substation Operators",
        "Ops Technicians (ELE -Ops)",
        "Technicians (ELE -Ops)",
        "Technicians",
        "MRSS Operators (Technician)",
    ],
    "Tech Associate Electrical": [
        "Floor Technicians",
        "Substation Operators",
        "Ops Technicians (ELE -Ops)",
        "Technicians (ELE -Ops)",
        "Technicians",
        "MRSS Operators (Technician)",
    ],
    "MST": ["Technicians (MST)"],
    "Sr. Technician Ops": [
        "Sr Technicians",
        "MRSS Operators (Sr Technician)",
        "High Mast Operator (Technicians)",
    ],
    "Sr Electrician OPS": ["Sr Technicians", "MRSS Operators (Sr Technician)"],
    "Technician": ["Technicians", "Technicians (ELE -Ops)"],
    "Asst. Manager (Mechanical)": ["Asst Manager"],
    "Shift Engineer (Mechanical)": ["Shift Engineers"],
    "HVAC Technician Mechanical": ["HVAC Floor Technicians"],
    "Sr Technician (Mechanical)": ["HVAC Sr Technicians (chiller plant)"],
    "Tech Associate Mech.": [
        "Helper Tech Associate Mech.",
        "Technician-mech",
        "Helper (Tech-Ass-Mec) chiller plant",
    ],
    "Fitter Mechanical": ["Fitter Mechanical"],
    "Plumber": ["Plumbing Technican", "Plumbing -Tech"],
    "Pump Operator": ["Pump Operator", "Pump operator"],
    "WTP Operator": ["WTP technician"],
    "STP Operator": ["STP operator"],
    "WTP/STP Helper": ["WTP Helper", "STP helper"],
    "WTP/STP Chemist": ["STP incharge"],
    "Safety Officer": ["Safety Officer"],
    "Motor Technician": ["Motor Technician"],
    "Motor Electrician": ["Motor Electrician"],
    "Tyre Technician": ["Tyre Technician"],
    "MT Supervisor": ["Associate mgr/JM"],
    "Driver (CLR Vehicles)": ["Road sweeper operator"],
    "Boom Lift Operator": ["High raise machine operators 21mt& 18mt"],
    "Desk Engineer (SCADA)": ["Engineers", "Supervisior"],
    "Sewerman & Drainage": ["Sewerman & Drainage"],
}

TEAM_LABELS = {
    "A": "ADMIN / Site Management",
    "B": "SCADA & Helpdesk",
    "C": "MECHANICAL & HVAC",
    "D": "ELECTRICAL",
    "E": "FIRE",
    "F": "CIVIL",
    "G": "MOTOR TRANSPORT",
    "H": "LANDSIDE",
}

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
fill_header_blue = PatternFill("solid", fgColor="9FC5E8")
fill_header_green = PatternFill("solid", fgColor="B6D7A8")
fill_header_navy = PatternFill("solid", fgColor="4B5C92")
fill_short = PatternFill("solid", fgColor="FCE4D6")
fill_match = PatternFill("solid", fgColor="E2EFDA")
fill_surp = PatternFill("solid", fgColor="FFF2CC")
fill_total = PatternFill("solid", fgColor="D9E2F3")
fill_vacant = PatternFill("solid", fgColor="FFF2CC")
fill_white = PatternFill("solid", fgColor="FFFFFF")


def load_daily(wb):
    ws = wb["281_DailySheet"]
    people = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name:
            continue
        people.append(
            {
                "name": str(name).strip(),
                "contact": ws.cell(r, 4).value,
                "desig_raw": ws.cell(r, 5).value,
                "desig": canon_desig(ws.cell(r, 5).value),
                "join": ws.cell(r, 6).value,
                "remarks": ws.cell(r, 7).value,
                "shortlisted": ws.cell(r, 8).value,
                "used_172": False,
                "used_281": False,
            }
        )
    return people


def find_daily_match(name, daily, used_flag="used_172"):
    for p in daily:
        if p[used_flag]:
            continue
        if names_match(name, p["name"]):
            return p
    return None


# ───────────────────────── 1) 281ManPowerSummary ─────────────────────────

def update_281_summary(wb, daily):
    ws = wb["281ManPowerSummary"]
    counts = Counter(p["desig"] for p in daily)

    rows = []
    for r in range(2, ws.max_row + 1):
        desig = ws.cell(r, 2).value
        if not desig or str(desig).strip().upper() == "TOTAL":
            continue
        rows.append(
            {
                "desig": desig,
                "team": ws.cell(r, 3).value,
                "req": ws.cell(r, 4).value,
            }
        )

    existing = {r["desig"] for r in rows}
    for d in counts:
        if d and d not in existing and not any(d.lower() == e.lower() for e in existing):
            rows.append({"desig": d, "team": "—", "req": "—"})

    def metrics(desig, req):
        filled = counts.get(desig, 0)
        if filled == 0:
            for k, v in counts.items():
                if k and desig and k.lower() == desig.lower():
                    filled = v
                    break
        if req in ("—", None, ""):
            short, surp = 0, filled
            status = "SURPLUS" if filled else "—"
        else:
            reqn = float(req)
            short = max(0, reqn - filled)
            surp = max(0, filled - reqn)
            if short > 0:
                status = "SHORTFALL"
            elif surp > 0:
                status = "SURPLUS"
            else:
                status = "MATCHED"
        return filled, short, surp, status

    enriched = []
    for s in rows:
        filled, short, surp, status = metrics(s["desig"], s["req"])
        enriched.append({**s, "filled": filled, "short": short, "surp": surp, "status": status})

    order = {"SHORTFALL": 0, "MATCHED": 1, "SURPLUS": 2, "—": 3}
    enriched.sort(key=lambda e: (order.get(e["status"], 9), -e["short"], -e["surp"], e["desig"] or ""))

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for c in range(1, 9):
        cell = ws.cell(1, c)
        cell.fill = fill_header_navy
        cell.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.alignment = center
        cell.border = thin

    tot_req = tot_f = tot_sh = tot_su = 0
    for i, e in enumerate(enriched, start=1):
        r = i + 1
        vals = [i, e["desig"], e["team"], e["req"], e["filled"], e["short"], e["surp"], e["status"]]
        fill = {"SHORTFALL": fill_short, "MATCHED": fill_match, "SURPLUS": fill_surp}.get(e["status"])
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell.border = thin
            cell.alignment = center
            cell.font = Font(name="Calibri", size=10)
            if fill:
                cell.fill = fill
        if isinstance(e["req"], (int, float)):
            tot_req += e["req"]
        tot_f += e["filled"]
        tot_sh += e["short"]
        tot_su += e["surp"]

    tr = len(enriched) + 2
    for c, v in enumerate([None, "TOTAL", None, tot_req, tot_f, tot_sh, tot_su, "—"], start=1):
        cell = ws.cell(tr, c, v)
        cell.border = thin
        cell.alignment = center
        cell.fill = fill_total
        cell.font = Font(name="Calibri", size=10, bold=True)

    print(f"[281ManPowerSummary] Req={tot_req} Filled={tot_f} Short={tot_sh} Surp={tot_su}")
    return enriched


# ───────────────────────── 2) 172Phase1Deployment ─────────────────────────

def can_fill_172(person, slot_desig, slot_dept) -> bool:
    targets = DAILY_TO_172.get(person["desig"], [])
    if slot_desig not in targets:
        return False
    dept = str(slot_dept or "")
    if slot_desig in ("Asst Manager", "Shift Engineers"):
        if person["desig"] == "Asst. Manager (Electrical)" and dept != "ELECTRICAL":
            return False
        if person["desig"] == "Asst. Manager (Mechanical)" and not dept.startswith("MECHANICAL"):
            return False
        if person["desig"] == "Shift Engineer (Electrical)" and dept != "ELECTRICAL":
            return False
        if person["desig"] == "Shift Engineer (Mechanical)" and not (
            dept.startswith("MECHANICAL") or dept == "PHE"
        ):
            return False
    if person["desig"] == "Desk Engineer (SCADA)" and dept != "ELV / SCADA":
        return False
    # Do not dump electrical manpower into ELV vacancies
    if dept == "ELV / SCADA" and person["desig"] in (
        "Electrician - Ops",
        "Tech Associate Electrical",
        "Technician",
    ):
        return False
    return True


def update_172_deployment(wb, daily):
    ws = wb["172Phase1Deployment"]
    refreshed = 0
    kept_external = []

    # Pass 1: refresh existing named rows from DailySheet
    for r in range(5, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name:
            continue
        match = find_daily_match(name, daily, "used_172")
        if match:
            match["used_172"] = True
            ws.cell(r, 2).value = match["name"]  # canonical DailySheet spelling
            if match["contact"] is not None:
                ws.cell(r, 3).value = match["contact"]
            if match["join"] is not None:
                ws.cell(r, 5).value = match["join"]
            refreshed += 1
        else:
            # duplicate rows of already-matched people, or external names
            # Try match ignoring used flag for duplicate detection
            any_match = None
            for p in daily:
                if names_match(name, p["name"]):
                    any_match = p
                    break
            if any_match:
                # Duplicate assignment already on roster — keep, sync contact/join, don't consume again
                ws.cell(r, 2).value = any_match["name"]
                if any_match["contact"] is not None:
                    ws.cell(r, 3).value = any_match["contact"]
                if any_match["join"] is not None:
                    ws.cell(r, 5).value = any_match["join"]
                refreshed += 1
            else:
                kept_external.append((r, name, ws.cell(r, 6).value))

    print(f"[172Deployment] refreshed/synced: {refreshed}")
    print(f"[172Deployment] named not in DailySheet (kept): {len(kept_external)}")
    for r, n, d in kept_external:
        print(f"    row{r}: {n!r} | {d}")

    # Existing named people on roster (to avoid introducing near-duplicate rows)
    existing_names = []
    for r in range(5, ws.max_row + 1):
        n = ws.cell(r, 2).value
        if n:
            existing_names.append(n)

    def already_on_roster(person_name: str) -> bool:
        for n in existing_names:
            if names_match(person_name, n):
                return True
        return False

    # Pass 2: fill vacant slots
    filled_new = 0
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value:
            continue
        slot_desig = ws.cell(r, 6).value
        slot_dept = ws.cell(r, 4).value
        if not slot_desig:
            continue
        candidate = None
        for p in daily:
            if p["used_172"]:
                continue
            if already_on_roster(p["name"]):
                # Near-duplicate of someone already named on 172 — skip
                p["used_172"] = True  # consume so we don't try elsewhere awkwardly
                continue
            if can_fill_172(p, slot_desig, slot_dept):
                candidate = p
                break
        if candidate:
            candidate["used_172"] = True
            ws.cell(r, 2).value = candidate["name"]
            ws.cell(r, 3).value = candidate["contact"]
            ws.cell(r, 5).value = candidate["join"]
            existing_names.append(candidate["name"])
            # shifts left blank — do not invent
            filled_new += 1
            print(
                f"    + row{r}: {slot_dept}/{slot_desig} ← {candidate['name']} ({candidate['desig']})"
            )

    # Stats
    req = filled = vacant = 0
    by_dept = defaultdict(lambda: [0, 0, 0])
    slots = []
    for r in range(5, ws.max_row + 1):
        desig = ws.cell(r, 6).value
        dept = ws.cell(r, 4).value or "?"
        loc = ws.cell(r, 7).value
        name = ws.cell(r, 2).value
        if not desig:
            continue
        req += 1
        by_dept[dept][0] += 1
        is_filled = bool(name)
        if is_filled:
            filled += 1
            by_dept[dept][1] += 1
        else:
            vacant += 1
            by_dept[dept][2] += 1
        slots.append(
            {
                "dept": dept,
                "desig": desig,
                "loc": loc,
                "filled": is_filled,
                "name": name,
            }
        )

    print(f"[172Deployment] Required={req} Filled={filled} Vacant={vacant}")
    for d, (a, b, c) in by_dept.items():
        print(f"    {d}: req={a} filled={b} vacant={c}")
    return slots, by_dept


# ───────────────────────── 3) 172Phase1Summary ─────────────────────────

def update_172_summary(wb, slots, by_dept):
    """Rebuild 172Phase1Summary from updated deployment (no invented numbers)."""
    # Remove old sheet content and rewrite structure similar to existing
    if "172Phase1Summary" in wb.sheetnames:
        idx = wb.sheetnames.index("172Phase1Summary")
        del wb["172Phase1Summary"]
        ws = wb.create_sheet("172Phase1Summary", idx)
    else:
        ws = wb.create_sheet("172Phase1Summary")

    ws.merge_cells("B2:I2")
    ws["B2"] = "PHASE-1 172 ROSTER — REQUIRED / FILLED / VACANT (CROSS-CHECKED FROM 281_DailySheet)"
    ws["B2"].font = Font(bold=True, name="Calibri", size=14, color="FFFFFF")
    ws["B2"].fill = fill_header_navy
    ws["B2"].alignment = center

    ws.merge_cells("B3:I4")
    filled = sum(1 for s in slots if s["filled"])
    vacant = sum(1 for s in slots if not s["filled"])
    ws["B3"] = (
        f"Source: 172Phase1Deployment updated from 281_DailySheet. "
        f"Required 172 | Filled {filled} | Vacant {vacant}. "
        f"Vacant = designation present, name empty. No invented data."
    )
    ws["B3"].alignment = Alignment(wrap_text=True, vertical="center")

    ws.merge_cells("B6:G6")
    ws["B6"] = "A. SECTION SUMMARY"
    ws["B6"].font = Font(bold=True, size=12)
    ws["B6"].fill = fill_header_green

    headers = ["Section", "Required Count", "Filled Count", "Vacant", "Fill %", "Status"]
    for c, h in enumerate(headers, start=2):
        cell = ws.cell(7, c, h)
        cell.font = Font(bold=True)
        cell.fill = fill_header_blue
        cell.border = thin
        cell.alignment = center

    # Collapse MECHANICAL / HVAC and PHE into MECHANICAL for section view (as prior summary did)
    section_map = {
        "ADMIN": "ADMIN",
        "CIVIL": "CIVIL",
        "ELV / SCADA": "ELV / SCADA",
        "ELECTRICAL": "ELECTRICAL",
        "MECHANICAL": "MECHANICAL",
        "MECHANICAL / HVAC": "MECHANICAL",
        "PHE": "MECHANICAL",
        "TRANSPORT": "TRANSPORT",
    }
    sec = defaultdict(lambda: [0, 0, 0])
    for s in slots:
        key = section_map.get(s["dept"], s["dept"])
        sec[key][0] += 1
        if s["filled"]:
            sec[key][1] += 1
        else:
            sec[key][2] += 1

    order_secs = ["ADMIN", "CIVIL", "ELV / SCADA", "ELECTRICAL", "MECHANICAL", "TRANSPORT"]
    r = 8
    tot = [0, 0, 0]
    for name in order_secs:
        a, b, c = sec.get(name, [0, 0, 0])
        pct = (b / a) if a else 0
        if pct >= 0.95:
            status = "Complete"
        elif pct >= 0.65:
            status = "On Track"
        elif pct == 0:
            status = "Critical"
        else:
            status = "Low"
        for col, val in enumerate([name, a, b, c, round(pct, 2), status], start=2):
            cell = ws.cell(r, col, val)
            cell.border = thin
            cell.alignment = center
        tot[0] += a
        tot[1] += b
        tot[2] += c
        r += 1

    for col, val in enumerate(
        ["TOTAL", tot[0], tot[1], tot[2], round(tot[1] / tot[0], 2) if tot[0] else 0, "—"],
        start=2,
    ):
        cell = ws.cell(r, col, val)
        cell.border = thin
        cell.alignment = center
        cell.fill = fill_total
        cell.font = Font(bold=True)

    # Designation detail
    r += 2
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    ws.cell(r, 2).value = "B. DESIGNATION DETAIL — Required Count / Filled Count / Vacant Count"
    ws.cell(r, 2).font = Font(bold=True, size=12)
    ws.cell(r, 2).fill = fill_header_green

    r += 1
    detail_h = [
        "Sr No",
        "Section",
        "Designation",
        "Location / Area",
        "Required Count",
        "Filled Count",
        "Vacant Count",
        "Hiring Priority",
    ]
    for c, h in enumerate(detail_h, start=2):
        cell = ws.cell(r, c, h)
        cell.font = Font(bold=True)
        cell.fill = fill_header_blue
        cell.border = thin
        cell.alignment = center

    # Aggregate by section/desig/loc
    agg = defaultdict(lambda: [0, 0])
    for s in slots:
        key = (section_map.get(s["dept"], s["dept"]), s["desig"], s["loc"] or "")
        agg[key][0] += 1
        if s["filled"]:
            agg[key][1] += 1

    detail_rows = []
    for (section, desig, loc), (reqn, fil) in agg.items():
        vac = reqn - fil
        if vac > 0 and fil == 0:
            pri = "CRITICAL"
        elif vac >= 3:
            pri = "HIGH"
        elif vac > 0:
            pri = "MEDIUM"
        else:
            pri = "OK"
        detail_rows.append((section, desig, loc, reqn, fil, vac, pri))

    pri_order = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3, "OK": 4}
    detail_rows.sort(key=lambda x: (pri_order.get(x[6], 9), -x[5], x[0], x[1], x[2]))

    r += 1
    start = r
    for i, (section, desig, loc, reqn, fil, vac, pri) in enumerate(detail_rows, start=1):
        for c, val in enumerate([i, section, desig, loc, reqn, fil, vac, pri], start=2):
            cell = ws.cell(r, c, val)
            cell.border = thin
            cell.alignment = center
            if pri == "CRITICAL":
                cell.fill = fill_short
            elif pri == "OK":
                cell.fill = fill_match
        r += 1

    for c, val in enumerate(
        [None, "TOTAL", None, None, tot[0], tot[1], tot[2], None], start=2
    ):
        cell = ws.cell(r, c, val)
        cell.border = thin
        cell.alignment = center
        cell.fill = fill_total
        cell.font = Font(bold=True)

    for col in range(2, 10):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 22

    print(f"[172Phase1Summary] rebuilt: Req={tot[0]} Filled={tot[1]} Vacant={tot[2]}")


# ───────────────────────── 4) 281ManpowerDeployment (NEW) ─────────────────────────

def build_281_deployment(wb, daily, summary_enriched):
    """Create/replace 281ManpowerDeployment as replica of 172Phase1Deployment format."""
    # Copy column widths / freeze style from 172 sheet
    src = wb["172Phase1Deployment"]

    if "281ManpowerDeployment" in wb.sheetnames:
        del wb["281ManpowerDeployment"]
    # Place after 281ManPowerSummary
    idx = wb.sheetnames.index("281ManPowerSummary") + 1
    ws = wb.create_sheet("281ManpowerDeployment", idx)

    # Column widths from source
    for col_letter, dim in src.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[col_letter].width = dim.width
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["F"].width = 32
    ws.column_dimensions["G"].width = 14

    # Title / legend / dates / headers — mirror 172
    ws.merge_cells("A1:AD1")
    ws["A1"] = (
        "DUTY ROSTER FOR THE MONTH OF JULY 2026 — GVIAL BHOGAPURAM — FULL 281 MANPOWER DEPLOYMENT"
    )
    ws["A1"].font = Font(name="Calibri", size=15, bold=True, color="000000")
    ws["A1"].fill = fill_header_blue
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:AD2")
    ws["A2"] = (
        "Shift Codes: A = Shift-1 (06:00-14:00) | B = Shift-2 (14:00-22:00) | "
        "C = Shift-3 (22:00-06:00) | G = General (09:00-18:00) | W/O = Weekly Off"
    )
    ws["A2"].font = Font(name="Calibri", size=11, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    # Date numbers 9..31 in H3:AD3 (same as 172)
    for i, day in enumerate(range(9, 32)):
        cell = ws.cell(3, 8 + i, day)
        cell.alignment = center
        cell.font = Font(bold=True, size=9)

    headers = [
        "S.NO.",
        "NAME",
        "MOBILE NUMBER",
        "DEPARTMENT",
        "Date Of Joining",
        "DESIGNATION",
        "LOCATION",
    ]
    # weekdays for 9 Jul 2026 (Thu) through 31 Jul 2026 (Fri)
    weekdays = ["THU", "FRI", "SAT", "SUN", "MON", "TUE", "WED"] * 4  # 28 days
    weekdays = weekdays[:23]  # 9..31 inclusive = 23 days
    headers.extend(weekdays)
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(4, c, h)
        cell.font = Font(name="EB Garamond", bold=True)
        cell.fill = fill_header_green
        cell.alignment = center
        cell.border = thin

    # Build 281 required slots from summary (Required > 0 numeric), preserve team order A..H
    team_order = ["A", "B", "C", "D", "E", "F", "G", "H"]
    required_slots = []
    for team in team_order:
        items = [
            e
            for e in summary_enriched
            if e["team"] == team and isinstance(e["req"], (int, float)) and e["req"] > 0
        ]
        # keep shortfall-first within team by designation name for stability
        items.sort(key=lambda e: e["desig"])
        for e in items:
            for _ in range(int(e["req"])):
                required_slots.append(
                    {
                        "team": team,
                        "dept": TEAM_LABELS.get(team, team),
                        "desig": e["desig"],
                        "loc": team,  # Plan281 has no location; use team code
                    }
                )

    assert len(required_slots) == 281, f"Expected 281 slots, got {len(required_slots)}"

    # Index DailySheet people by canonical designation for assignment
    by_desig = defaultdict(list)
    for p in daily:
        by_desig[p["desig"]].append(p)

    # Also build shift lookup from 172 deployment by fuzzy name
    shift_by_name = {}
    ws172 = wb["172Phase1Deployment"]
    for r in range(5, ws172.max_row + 1):
        name = ws172.cell(r, 2).value
        if not name:
            continue
        shifts = [ws172.cell(r, c).value for c in range(8, 31)]
        if any(shifts):
            shift_by_name[norm_name(name)] = shifts

    def shifts_for(person_name):
        nn = norm_name(person_name)
        if nn in shift_by_name:
            return shift_by_name[nn]
        for k, v in shift_by_name.items():
            if names_match(person_name, k):
                return v
        return [None] * 23

    filled = 0
    for i, slot in enumerate(required_slots, start=1):
        r = 4 + i
        person = None
        pool = by_desig.get(slot["desig"], [])
        for p in pool:
            if not p["used_281"]:
                person = p
                break
        # case-insensitive fallback pool
        if person is None:
            for k, pool in by_desig.items():
                if k and k.lower() == slot["desig"].lower():
                    for p in pool:
                        if not p["used_281"]:
                            person = p
                            break
                if person:
                    break

        ws.cell(r, 1).value = i
        ws.cell(r, 4).value = slot["dept"]
        ws.cell(r, 6).value = slot["desig"]
        ws.cell(r, 7).value = slot["loc"]

        if person:
            person["used_281"] = True
            ws.cell(r, 2).value = person["name"]
            ws.cell(r, 3).value = person["contact"]
            ws.cell(r, 5).value = person["join"]
            for c, v in enumerate(shifts_for(person["name"]), start=8):
                ws.cell(r, c).value = v
            filled += 1
            row_fill = fill_white
        else:
            row_fill = fill_vacant  # vacant highlight

        for c in range(1, 31):
            cell = ws.cell(r, c)
            cell.border = thin
            cell.alignment = center
            cell.font = Font(name="Calibri", size=9, bold=True)
            if row_fill:
                cell.fill = row_fill

    # Surplus people from DailySheet not placed into a required slot — appendix note rows?
    # Keep deployment at exactly 281 required slots (replica of 172 which is exactly 172).
    # Surplus remain visible on DailySheet + Summary.

    unused = [p for p in daily if not p["used_281"]]
    print(
        f"[281ManpowerDeployment] slots=281 filled={filled} vacant={281-filled} "
        f"daily_unused_surplus_or_overflow={len(unused)}"
    )
    if unused:
        print("    Unused DailySheet people (surplus / overflow beyond required):")
        for p in unused:
            print(f"      {p['desig']}: {p['name']}")

    return filled


def main():
    wb = openpyxl.load_workbook(SRC)
    daily = load_daily(wb)
    print(f"Loaded DailySheet people: {len(daily)}")

    summary_enriched = update_281_summary(wb, daily)
    slots, by_dept = update_172_deployment(wb, daily)
    update_172_summary(wb, slots, by_dept)
    build_281_deployment(wb, daily, summary_enriched)

    # Keep sheet order sensible
    desired = [
        "281_DailySheet",
        "281ManPowerSummary",
        "281ManpowerDeployment",
        "172Phase1Deployment",
        "172Phase1Summary",
    ]
    for i, name in enumerate(desired):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.save(OUT)
    print(f"\nSaved {OUT}")
    print("Final sheets:", wb.sheetnames)


if __name__ == "__main__":
    main()
