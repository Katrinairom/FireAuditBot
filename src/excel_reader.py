# src/excel_reader.py
# Reads the store Excel (e.g. Z943-ADT-013_...xlsx) using openpyxl.
# Returns a flat dict `store_data` that all other modules consume.
# Does NO calculation. Empty cells return None.
#
# NOTE: This module does NOT need the Drafter's Copy Excel.
#       All reference constants (OLP, calorific values, NBC rules)
#       live in constants.py and are used by formula_engine / nbc_checker.

import os
import re
import openpyxl
from .constants import SHEET_PROPERTY, SHEET_EQUIPMENT, SHEET_CHECKLIST, SHEET_LOCATION, MATERIAL_ORDER


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def _val(ws, row, col):
    """Return stripped string value of a cell, or None if empty."""
    v = ws.cell(row=row, column=col).value
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _find_row(ws, keyword, col=1, start=1):
    """Find first row where the given column contains keyword (case-insensitive). Returns -1 if not found."""
    kw = keyword.upper()
    for row in ws.iter_rows(min_row=start):
        cell_val = row[col - 1].value
        if cell_val and kw in str(cell_val).upper():
            return row[0].row
    return -1


def _parse_number(val):
    """Extract first numeric value from a string like '25000 Sq Mtr' or '2300'. Returns float or None."""
    if val is None:
        return None
    match = re.search(r"[\d,]+\.?\d*", str(val))
    if match:
        return float(match.group().replace(",", ""))
    return None


def _extract_store_code(filepath):
    """Extract store code like Z943 or Z440 from the filename."""
    base = os.path.basename(filepath)
    match = re.match(r"^([A-Z]\d{3,4})", base, re.IGNORECASE)
    return match.group(1).upper() if match else "XXXX"


# ---------------------------------------------------------------------------
# SECTION READERS
# ---------------------------------------------------------------------------

def _read_property_overview(ws):
    """Read section 3.1 — Property Overview. Returns dict."""
    start = _find_row(ws, "PROPERTY OVERVIEW", col=3)
    if start == -1:
        return {}

    fields = [
        "property_name", "address", "owner_occupier_name", "contact_details",
        "oc_number", "registration_jurisdiction", "type_of_property",
        "total_builtup_area", "number_of_floors", "marginal_front",
        "marginal_left", "marginal_right", "marginal_behind",
        "stilt_parking", "basement_details"
    ]
    result = {}
    for i, key in enumerate(fields):
        result[key] = _val(ws, start + i, 5)
    return result


def _read_store_details(ws):
    """Read section 3.4 — Store Details. Returns dict."""
    start = _find_row(ws, "STORE DETAILS", col=3)
    if start == -1:
        return {}

    fields = [
        "name_of_store", "store_address", "nature_of_business", "type_of_use",
        "nature_of_goods", "total_builtup_area", "number_of_floors",
        "floorwise_area", "occupancy_classification", "business_registration",
        "year_of_establishment", "store_contact", "total_employees",
        "hours_of_operations", "security_personnel"
    ]
    result = {}
    for i, key in enumerate(fields):
        result[key] = _val(ws, start + i, 5)
    return result


def _read_building_details(ws):
    """Read section 3.2 — Building Details. Returns dict."""
    start = _find_row(ws, "BUILDING  DETAILS", col=3)
    if start == -1:
        return {}
    fields = [
        "construction_type", "year_of_construction", "building_height",
        "refuge_area", "staircases", "openings_ventilation", "compartmentation",
        "elevator", "fire_exits", "electrical_supply", "gas_flammable",
        "hvac", "water_supply", "security_surveillance", "automated_doors"
    ]
    result = {}
    for i, key in enumerate(fields):
        result[key] = _val(ws, start + i, 5)
    return result


def _read_checklist_section(ws, header, count):
    """Read a Yes/No checklist block of `count` rows after the header. Returns list of dicts."""
    start = _find_row(ws, header, col=2)
    if start == -1:
        start = _find_row(ws, header, col=1)
    if start == -1:
        return []
    items = []
    for r in range(start + 2, start + 2 + count):
        sr = _val(ws, r, 1)
        item = _val(ws, r, 2)
        yn = _val(ws, r, 3)
        rmk = _val(ws, r, 4)
        if item:
            items.append(
                {"sr": sr, "item": item, "yes_no": yn, "remarks": rmk})
    return items


def _read_material_masses(ws):
    """
    Read the 10-row material mass table from Property sheet.
    Returns dict keyed by material name (lowercase), value = mass in kg (float).
    """
    # Find the header row containing "Mass of Material"
    start = _find_row(ws, "Mass of Material", col=5)
    if start == -1:
        # Fallback: scan for the Sr No 1.) pattern after row 60
        for r in range(60, ws.max_row):
            v = _val(ws, r, 1)
            if v and "1.)" in str(v):
                start = r - 1
                break

    masses = {}
    for i, mat_name in enumerate(MATERIAL_ORDER):
        row_num = start + 1 + i
        raw = _val(ws, row_num, 5)
        masses[mat_name] = _parse_number(raw) or 0.0
    return masses


def _read_equipment_section(ws, header):
    """
    Read one firefighting equipment table section (e.g. 4.1.A Water Tank Details).
    Returns list of {label, yes_no, quantity, status, maintenance, remarks}.
    """
    start = _find_row(ws, header, col=2)
    if start == -1:
        return []
    items = []
    section_headers = ["4.1.A", "4.1.B", "4.1.C",
                       "4.1.D", "4.1.E", "4.1.F", "4.1.G"]
    for r in range(start + 1, start + 80):
        label = _val(ws, r, 2)
        if not label:
            continue
        # Stop if we hit the next major section
        if any(h in str(label).upper() for h in section_headers) and r > start + 2:
            break
        items.append({
            "label":       label,
            "yes_no":      _val(ws, r, 3),
            "quantity":    _val(ws, r, 4),
            "status":      _val(ws, r, 5),
            "maintenance": _val(ws, r, 6),
            "remarks":     _val(ws, r, 7),
        })
    return items


def _read_location_sheet(ws):
    """Read Firefighting Equipment Store Wise sheet. Returns dict of category → list of rows."""
    categories = [
        ("hydrant_valves",        "FIRE HYDRANT VALVES"),
        ("hose_accessories",      "FIRE HOSE WITH ACCESSORIES"),
        ("hose_reel",             "FIRE HOSE REEL"),
        ("alarm_panel",           "FIRE ALARM PANEL"),
        ("manual_call_point",     "MANUAL CALL POINT"),
        ("hooter",                "HOOTER"),
        ("smoke_detector",        "SMOKE DETECTOR"),
        ("fire_extinguishers",    "FIRE EXTINGUISHERS"),
    ]
    data = {}
    for key, label in categories:
        start = _find_row(ws, label, col=2)
        if start == -1:
            data[key] = []
            continue
        items = []
        for r in range(start + 1, start + 40):
            loc = _val(ws, r, 2)
            maint = _val(ws, r, 3)
            rmk = _val(ws, r, 4)
            if not loc:
                continue
            # Stop at next category header
            if any(l in str(loc).upper() for _, l in categories):
                break
            items.append(
                {"location": loc, "maintenance": maint, "remarks": rmk})
        data[key] = items
    return data


def _read_maintenance_training(ws):
    """Read maintenance training checklist from Equipment sheet."""
    start = _find_row(ws, "MAINTENACE TRAINING CHECKLIST", col=2)
    if start == -1:
        return []
    items = []
    for r in range(start + 1, start + 15):
        label = _val(ws, r, 2)
        if not label:
            continue
        if "AMC" in str(label).upper() or "ITEM" in str(label).upper():
            break
        items.append({
            "label":     label,
            "yes_no":    _val(ws, r, 5),
            "frequency": _val(ws, r, 6),
            "remarks":   _val(ws, r, 7),
        })
    return items


def _read_amc_table(ws):
    """Read AMC status table from Equipment sheet."""
    start = _find_row(ws, "FIRE PUMPS", col=2)
    if start == -1:
        return []
    items = []
    for r in range(start, start + 8):
        item = _val(ws, r, 2)
        if not item:
            continue
        items.append({
            "item":           item,
            "amc_status":     _val(ws, r, 3),
            "frequency":      _val(ws, r, 4),
            "report_status":  _val(ws, r, 6),
            "remarks":        _val(ws, r, 7),
        })
    return items


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def read_store_data(excel_path, inspection_template=None):
    """
    Load all sheets from store Excel and return a flat store_data dict.

    NOTE: inspection_template parameter is accepted for backward compatibility
    but is NOT used. All constants are in constants.py.
    The old error 'Inspection template sheet missing' is eliminated.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Store Excel not found at: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # --- Read Property sheet ---
    ws_prop = wb[SHEET_PROPERTY]
    property_overview = _read_property_overview(ws_prop)
    store_details = _read_store_details(ws_prop)
    building_details = _read_building_details(ws_prop)
    material_masses = _read_material_masses(ws_prop)

    # Parse numeric area values
    property_area = _parse_number(property_overview.get("total_builtup_area"))
    store_area = _parse_number(store_details.get("total_builtup_area"))

    # --- Read Equipment sheet ---
    ws_equip = wb[SHEET_EQUIPMENT]
    equipment = {
        "water_tanks":  _read_equipment_section(ws_equip, "WATER TANK DETAILS"),
        "pump_room":    _read_equipment_section(ws_equip, "FIRE PUMP ROOM DETAILS"),
        "hydrant":      _read_equipment_section(ws_equip, "FIRE HYDRANT SYSTEM DETAILS"),
        "sprinkler":    _read_equipment_section(ws_equip, "FIRE SPRINKLER SYSTEM DETAILS"),
        "alarm":        _read_equipment_section(ws_equip, "FIRE ALARM SYSTEM DETAILS"),
        "passive":      _read_equipment_section(ws_equip, "PASSIVE FIRE FIGHTING SYSTEM DETAILS"),
        "extinguisher": _read_equipment_section(ws_equip, "PORTABLE FIRE EXTINGUISHER"),
        "maintenance":  _read_maintenance_training(ws_equip),
        "amc":          _read_amc_table(ws_equip),
    }

    # --- Read Checklist sheet ---
    ws_check = wb[SHEET_CHECKLIST]
    checklists = {
        "evacuation":        _read_checklist_section(ws_check, "EMERGENCY EVACUATION CHECKLIST", 15),
        "smoke_ventilation": _read_checklist_section(ws_check, "SMOKE VENTILATION CHECKLIST", 14),
        "electrical_safety": _read_checklist_section(ws_check, "ELECTRICAL SAFETY CHECKLIST", 13),
    }

    # --- Read Location sheet ---
    ws_loc = wb[SHEET_LOCATION]
    locations = _read_location_sheet(ws_loc)

    wb.close()

    return {
        "store_code":         _extract_store_code(excel_path),
        "source_file":        excel_path,
        "property_area":      property_area,
        "store_area":         store_area,
        "property_overview":  property_overview,
        "store_details":      store_details,
        "building_details":   building_details,
        "material_masses":    material_masses,
        "equipment":          equipment,
        "checklists":         checklists,
        "locations":          locations,
    }
