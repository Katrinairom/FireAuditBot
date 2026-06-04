# src/report_builder.py
# Fills every table in the Word template with data from store_data.
# Table index map (from template inspection):
#   3  = 3.1 Property Overview
#   4  = 3.2 Building Details
#   5  = 3.3 Building Fire Equipment Checklist
#   6  = 3.4 Store Details
#   7  = 3.5 Store Fire Equipment Checklist
#   8  = 4.1.A Water Tanks + 4.1.B Pump Room
#   9  = 4.1.C Fire Hydrant
#   10 = 4.1.D Fire Sprinkler
#   11 = 4.1.E Fire Alarm
#   12 = 4.1.F Passive Firefighting
#   13 = 4.1.G Portable Extinguisher
#   15 = 5.3 Occupancy Load
#   16 = 5.4 Fire Load
#   18 = 6.1.A NBC Water Tanks
#   19 = 6.1.B NBC Pumps
#   20 = 6.1.C NBC Hydrant
#   21 = 6.1.D NBC Sprinkler
#   22 = 6.1.E NBC Alarm
#   23 = 6.1.F NBC Extinguisher
#   25 = 7.1 Emergency Evacuation Checklist
#   26 = 7.2 Smoke Ventilation Checklist
#   27 = 7.3 Electrical Safety Checklist
#   28 = 7.4 AMC Maintenance Table
#   29 = 7.4 Training Checklist (part 1)
#   30 = 7.4 Training Checklist (part 2)
#   31 = 9.2 Location-wise Equipment

import os
from docx import Document
from docx.shared import Pt
from .constants import FONT_NAME, FONT_SIZE_PT, MATERIAL_ORDER, CALORIFIC_VALUES


# ---------------------------------------------------------------------------
# CORE CELL WRITER
# ---------------------------------------------------------------------------

def _write(cell, text, caps=True):
    """Write text into a cell, enforce Calibri 12pt, preserve existing formatting."""
    if text is None:
        text = "N/A"
    final = str(text).upper() if caps else str(text)
    para = cell.paragraphs[0]
    # Clear existing runs
    for run in para.runs:
        run.text = ""
    if para.runs:
        run = para.runs[0]
    else:
        run = para.add_run()
    run.text = final
    run.font.name = FONT_NAME
    run.font.size = Pt(FONT_SIZE_PT)


def _na(val):
    """Return N/A if val is None/empty, else return val as string."""
    if val is None:
        return "N/A"
    s = str(val).strip()
    return s if s else "N/A"


def _yn(val):
    """Normalise YES/NO/N/A values."""
    if val is None:
        return "N/A"
    s = str(val).strip().upper()
    if s in ("YES", "Y", "1"):
        return "YES"
    if s in ("NO", "N", "0"):
        return "NO"
    return "N/A"


def _qty(val):
    """Format quantity string. If already has unit, caps it. Otherwise return as-is."""
    if val is None:
        return "N/A"
    s = str(val).strip()
    return s.upper() if s else "N/A"


def _find_equip_val(equip_list, keyword, field="yes_no"):
    """Find first item in equipment list whose label contains keyword, return field value."""
    if not equip_list:
        return None
    for item in equip_list:
        if keyword.upper() in str(item.get("label", "")).upper():
            return item.get(field)
    return None


# ---------------------------------------------------------------------------
# SECTION FILLERS
# ---------------------------------------------------------------------------

def _fill_two_col_table(table, data_dict, key_order):
    """
    Fill a 2-column Field/Description table.
    key_order: list of (row_index, data_key) tuples.
    """
    for row_idx, data_key in key_order:
        if row_idx < len(table.rows):
            _write(table.rows[row_idx].cells[1], _na(data_dict.get(data_key)))


def _fill_31_property_overview(table, prop):
    """Table 3 — 3.1 Property Overview (16 rows)."""
    mapping = [
        (1,  "property_name"),
        (2,  "address"),
        (3,  "owner_occupier_name"),
        (4,  "contact_details"),
        (5,  "oc_number"),
        (6,  "registration_jurisdiction"),
        (7,  "type_of_property"),
        (8,  "total_builtup_area"),
        (9,  "number_of_floors"),
        (10, "marginal_front"),
        (11, "marginal_left"),
        (12, "marginal_right"),
        (13, "marginal_behind"),
        (14, "stilt_parking"),
        (15, "basement_details"),
    ]
    _fill_two_col_table(table, prop, mapping)


def _fill_32_building_details(table, bldg):
    """Table 4 — 3.2 Building Details (16 rows)."""
    mapping = [
        (1,  "construction_type"),
        (2,  "year_of_construction"),
        (3,  "building_height"),
        (4,  "refuge_area"),
        (5,  "staircases"),
        (6,  "openings_ventilation"),
        (7,  "compartmentation"),
        (8,  "elevator"),
        (9,  "fire_exits"),
        (10, "electrical_supply"),
        (11, "gas_flammable"),
        (12, "hvac"),
        (13, "water_supply"),
        (14, "security_surveillance"),
        (15, "automated_doors"),
    ]
    _fill_two_col_table(table, bldg, mapping)


def _fill_checklist_table(table, checklist):
    """Tables 5, 7 — Fire Equipment Checklists (16 rows)."""
    mapping = [
        (1,  "fire_noc_details"),
        (2,  "fire_pump_details"),
        (3,  "fire_hydrant_system_details"),
        (4,  "fire_sprinkler_system_details"),
        (5,  "fire_extinguisher_details"),
        (6,  "fire_alarm_system_details"),
        (7,  "public_address_system_details"),
        (8,  "fire_safety_signage"),
        (9,  "fire_door"),
        (10, "dedicated_evacuation_point"),
        (11, "emergency_exit"),
        (12, "fire_staircase"),
        (13, "emergency_elevator"),
        (14, "fire_refuge_area"),
        (15, "fire_escape_plan"),
    ]
    _fill_two_col_table(table, checklist, mapping)


def _fill_34_store_details(table, store):
    """Table 6 — 3.4 Store Details (16 rows)."""
    mapping = [
        (1,  "name_of_store"),
        (2,  "store_address"),
        (3,  "nature_of_business"),
        (4,  "type_of_use"),
        (5,  "nature_of_goods"),
        (6,  "total_builtup_area"),
        (7,  "number_of_floors"),
        (8,  "floorwise_area"),
        (9,  "occupancy_classification"),
        (10, "business_registration"),
        (11, "year_of_establishment"),
        (12, "store_contact"),
        (13, "total_employees"),
        (14, "hours_of_operations"),
        (15, "security_personnel"),
    ]
    _fill_two_col_table(table, store, mapping)


def _fill_41a_water_tanks(table, equip):
    """Table 8 — 4.1.A Water Tanks + 4.1.B Pump Room."""
    wt = equip.get("water_tanks", [])
    pr = equip.get("pump_room", [])

    def wt_yn(kw):  return _yn(_find_equip_val(wt, kw, "yes_no"))
    def wt_qty(kw): return _qty(_find_equip_val(wt, kw, "quantity"))
    def pr_yn(kw):  return _yn(_find_equip_val(pr, kw, "yes_no"))
    def pr_qty(kw): return _qty(_find_equip_val(pr, kw, "quantity"))

    def w(row, col, val): _write(table.rows[row].cells[col], val)

    # 4.1.A Water Tanks
    w(2,  2, wt_yn("U/G Water Tank Capacity of Building"));  w(2,  3, wt_qty("U/G Water Tank Capacity of Building"))
    w(3,  2, wt_yn("O/H Water Tank Capacity of Building"));  w(3,  3, wt_qty("O/H Water Tank Capacity of Building"))
    w(6,  2, wt_yn("separate fire water tank"));             w(6,  3, wt_qty("separate fire water tank"))
    w(7,  2, wt_yn("U/G Water Tank Capacity of Building"));  w(7,  3, wt_qty("U/G Water Tank Capacity of Building"))
    w(8,  2, wt_yn("O/H Water Tank Capacity of Building"));  w(8,  3, wt_qty("O/H Water Tank Capacity of Building"))
    w(11, 2, wt_yn("separate pump connection"))
    w(12, 2, wt_yn("U/G Water Tank Capacity of Building"));  w(12, 3, wt_qty("U/G Water Tank Capacity of Building"))
    w(13, 2, wt_yn("O/H Water Tank Capacity of Building"));  w(13, 3, wt_qty("O/H Water Tank Capacity of Building"))

    # 4.1.B Pump Room
    w(16, 2, pr_yn("Monoblock or Submersible"))
    w(19, 2, pr_yn("Positive or Negative"))
    w(21, 2, pr_yn("Riser or Downcomer"))
    w(23, 2, pr_yn("Electric Fire Pump – Main"));    w(23, 3, pr_qty("Electric Fire Pump – Main"))
    w(24, 2, pr_yn("Electric Fire Pump – Standby")); w(24, 3, pr_qty("Electric Fire Pump – Standby"))
    w(25, 2, pr_yn("Electric Fire Pump – Sprinkler"));w(25, 3, pr_qty("Electric Fire Pump – Sprinkler"))
    w(26, 2, pr_yn("Diesel Fire Engine"));            w(26, 3, pr_qty("Diesel Fire Engine"))
    w(27, 2, pr_yn("Jockey"));                        w(27, 3, pr_qty("Jockey"))
    w(28, 2, pr_yn("Booster Pump"));                  w(28, 3, pr_qty("Booster Pump"))
    w(29, 2, pr_yn("separate fire pump connection"))


def _fill_41c_hydrant(table, equip):
    """Table 9 — 4.1.C Fire Hydrant System."""
    hyd = equip.get("hydrant", [])
    def yn(kw):  return _yn(_find_equip_val(hyd, kw, "yes_no"))
    def qty(kw): return _qty(_find_equip_val(hyd, kw, "quantity"))
    def w(row, col, val): _write(table.rows[row].cells[col], val)

    w(4,  2, yn("Fire Hydrant Valve"));        w(4,  3, qty("Fire Hydrant Valve"))
    w(5,  2, yn("Fire Hose Box"));              w(5,  3, qty("Fire Hose Box"))
    w(6,  2, yn("Fire Hose including"));        w(6,  3, qty("Fire Hose including"))
    w(7,  2, yn("Hose Reel"));                  w(7,  3, qty("Hose Reel"))
    w(8,  2, yn("Fire Extinguisher"));          w(8,  3, qty("Fire Extinguisher"))
    w(9,  2, yn("Fire Brigade Inlet"));         w(9,  3, qty("Fire Brigade Inlet"))
    w(12, 2, yn("Courtyard"));                  w(12, 3, qty("Courtyard"))
    w(22, 2, yn("Internal Fire Hydrant Valve"));w(22, 3, qty("Internal Fire Hydrant Valve"))
    w(23, 2, yn("Hose Box incl"));              w(23, 3, qty("Hose Box incl"))
    w(24, 2, yn("Fire Hose Reel"));             w(24, 3, qty("Fire Hose Reel"))
    w(27, 2, yn("External Fire Hydrant Valve"));w(27, 3, qty("External Fire Hydrant Valve"))
    w(29, 2, yn("Fire Brigade Inlet"));         w(29, 3, qty("Fire Brigade Inlet"))


def _fill_41d_sprinkler(table, equip):
    """Table 10 — 4.1.D Fire Sprinkler System."""
    spr = equip.get("sprinkler", [])
    def yn(kw):  return _yn(_find_equip_val(spr, kw, "yes_no"))
    def qty(kw): return _qty(_find_equip_val(spr, kw, "quantity"))
    def w(row, col, val): _write(table.rows[row].cells[col], val)

    w(1,  2, yn("property has Fire Sprinkler"))
    w(4,  2, yn("installed in basement"));       w(4,  3, qty("installed in basement"))
    w(5,  2, yn("installed in stilt parking"));  w(5,  3, qty("installed in stilt parking"))
    w(6,  2, yn("common area"))
    w(7,  2, yn("inside complete premises"));    w(7,  3, qty("inside complete premises"))
    w(11, 2, yn("Pendent Sprinkler"));           w(11, 3, qty("Pendent Sprinkler"))
    w(12, 2, yn("Upright Sprinkler"));           w(12, 3, qty("Upright Sprinkler"))
    w(13, 2, yn("Conceal Sprinkler"))
    w(16, 2, yn("Pendent Sprinkler"))
    w(17, 2, yn("Upright Sprinkler"));           w(17, 3, qty("Upright Sprinkler"))
    w(18, 2, yn("Water Curtain"))
    w(21, 2, yn("separate Sprinkler Arrangement"))
    w(23, 2, yn("Isolation Valve"))
    w(24, 2, yn("Pressure Gauge"))
    w(25, 2, yn("Drain Valve"))
    w(27, 2, yn("Pendent Sprinkler"));           w(27, 3, qty("Pendent Sprinkler"))
    w(28, 2, yn("Upright Sprinkler"))
    w(29, 2, yn("Conceal Sprinkler"))


def _fill_41e_alarm(table, equip):
    """Table 11 — 4.1.E Fire Alarm System."""
    alm = equip.get("alarm", [])
    def yn(kw):  return _yn(_find_equip_val(alm, kw, "yes_no"))
    def qty(kw): return _qty(_find_equip_val(alm, kw, "quantity"))
    def w(row, col, val): _write(table.rows[row].cells[col], val)

    w(1,  2, yn("property has Fire Detection"))
    w(4,  2, yn("installed in basement"))
    w(5,  2, yn("installed in parking"))
    w(6,  2, yn("common"))
    w(7,  2, yn("inside complete premises"))
    w(10, 2, yn("Addressable"));   w(10, 3, qty("Addressable"))
    w(11, 2, yn("Conventional"));  w(11, 3, qty("Conventional"))
    w(12, 2, yn("MCP"));           w(12, 3, qty("MCP"))
    w(13, 2, yn("Hooter"));        w(13, 3, qty("Hooter"))
    w(14, 2, yn("Smoke Detector"));w(14, 3, qty("Smoke Detector"))
    w(15, 2, yn("Heat Detector")); w(15, 3, qty("Heat Detector"))
    w(17, 2, yn("Public Address")); w(17, 3, qty("Public Address"))
    w(18, 2, yn("Talkback"))
    # Store section (4.1.E.B rows 20-29)
    w(20, 2, yn("separate Detection Arrangement"))
    w(21, 2, yn("Addressable"));   w(21, 3, qty("Addressable"))
    w(22, 2, yn("Conventional"))
    w(23, 2, yn("MCP"));           w(23, 3, qty("MCP"))
    w(24, 2, yn("Hooter"));        w(24, 3, qty("Hooter"))
    w(25, 2, yn("Smoke Detector"));w(25, 3, qty("Smoke Detector"))
    w(26, 2, yn("Heat Detector"))
    w(28, 2, yn("Public Address"))
    w(29, 2, yn("Talkback"))


def _fill_41f_passive(table, equip):
    """Table 12 — 4.1.F Passive Firefighting."""
    pas = equip.get("passive", [])
    def yn(kw):  return _yn(_find_equip_val(pas, kw, "yes_no"))
    def qty(kw): return _qty(_find_equip_val(pas, kw, "quantity"))
    def w(row, col, val): _write(table.rows[row].cells[col], val)

    # Building (4.1.F.A rows 3-15)
    for row, kw in [
        (3, "Smoke Exhaust"), (4, "Staircase Press"), (5, "Auto Glow"),
        (6, "LED Emergency"), (7, "Emergency Exit Doors"), (8, "Assembly"),
        (9, "Evacuation Plan"), (10, "First Aid"), (11, "Breathing"),
        (12, "Fireman Axe"), (13, "Fire Bucket"), (14, "Fire Beater"), (15, "Fire Blanket")
    ]:
        w(row, 2, yn(kw)); w(row, 3, qty(kw))

    # Store (4.1.F.B rows 18-30)
    for row, kw in [
        (18, "Smoke Exhaust"), (19, "Staircase Press"), (20, "Auto Glow"),
        (21, "LED Emergency"), (22, "Emergency Exit Doors"), (23, "Assembly"),
        (24, "Evacuation Plan"), (25, "First Aid"), (26, "Breathing"),
        (27, "Fireman Axe"), (28, "Fire Bucket"), (29, "Fire Beater"), (30, "Fire Blanket")
    ]:
        w(row, 2, yn(kw)); w(row, 3, qty(kw))


def _fill_41g_extinguisher(table, equip):
    """Table 13 — 4.1.G Portable Fire Extinguisher."""
    ext = equip.get("extinguisher", [])
    def yn(kw):  return _yn(_find_equip_val(ext, kw, "yes_no"))
    def qty(kw): return _qty(_find_equip_val(ext, kw, "quantity"))
    def w(row, col, val): _write(table.rows[row].cells[col], val)

    # Building (rows 3-11)
    for row, kw in [
        (3, "ABC"), (4, "CO2"), (5, "DCP"), (6, "CLEAN AGENT"),
        (7, "K-TYPE"), (8, "MECHANICAL FOAM"), (9, "MODULAR"),
        (10, "AUTOMATIC GAS"), (11, "FIRE BALL")
    ]:
        w(row, 2, yn(kw)); w(row, 3, qty(kw))

    # Store (rows 15-23)
    for row, kw in [
        (15, "ABC"), (16, "CO2"), (17, "DCP"), (18, "CLEAN AGENT"),
        (19, "K-TYPE"), (20, "MECHANICAL FOAM"), (21, "MODULAR"),
        (22, "AUTOMATIC GAS"), (23, "FIRE BALL")
    ]:
        w(row, 2, yn(kw)); w(row, 3, qty(kw))


def _fill_occupancy_table(table, occ_table):
    """Table 15 — 5.3 Occupancy Load (14 data rows + 1 total row)."""
    for i, zone_row in enumerate(occ_table):
        row_idx = i + 1
        if row_idx >= len(table.rows):
            break
        _write(table.rows[row_idx].cells[3], str(round(zone_row["area_occupied"], 1)))
        _write(table.rows[row_idx].cells[5], str(round(zone_row["occupancy_load"], 1)))


def _fill_fire_load_table(table, fire_table, masses):
    """Table 16 — 5.4 Fire Load Calculation (10 material rows)."""
    for i, mat in enumerate(MATERIAL_ORDER):
        row_idx = i + 1
        if row_idx >= len(table.rows):
            break
        mass = masses.get(mat, 0.0)
        cv   = CALORIFIC_VALUES.get(mat, 0.0)
        mj   = mass * cv
        _write(table.rows[row_idx].cells[4], str(mass))
        _write(table.rows[row_idx].cells[5], str(round(mj, 2)))


def _fill_nbc_water_tanks(table, nbc, equip):
    """Table 18 — 6.1.A NBC Water Tank Necessity."""
    reqs     = nbc.get("requirements", {})
    verdicts = nbc.get("verdicts", {})
    wt       = equip.get("water_tanks", [])

    def actual_qty(kw):
        return _qty(_find_equip_val(wt, kw, "quantity"))

    common_ug_req = reqs.get("common_ug_tank")
    fire_ug_req   = reqs.get("fire_ug_tank")
    common_oh_req = reqs.get("common_oh_tank")
    fire_oh_req   = reqs.get("fire_oh_tank")

    def fmt_req(val): return f"{val:,} LTR" if isinstance(val, int) else "NOT COMPULSORY"

    if len(table.rows) > 1:
        r = table.rows[1]
        _write(r.cells[1], f"UG: {fmt_req(common_ug_req)} / OH: {fmt_req(common_oh_req)}")
        _write(r.cells[2], actual_qty("U/G Water Tank") + " UG / " + actual_qty("O/H Water Tank") + " OH")
        _write(r.cells[3], verdicts.get("common_ug_tank", "N/A"))
    if len(table.rows) > 2:
        r = table.rows[2]
        _write(r.cells[1], f"UG: {fmt_req(fire_ug_req)} / OH: {fmt_req(fire_oh_req)}")
        _write(r.cells[2], actual_qty("fire water tank"))
        _write(r.cells[3], verdicts.get("fire_ug_tank", "N/A"))


def _fill_nbc_pumps(table, nbc):
    """Table 19 — 6.1.B NBC Pump Necessity."""
    reqs     = nbc.get("requirements", {})
    verdicts = nbc.get("verdicts", {})

    pump_rows = [
        (1, reqs.get("main_pump",   "N/A"), verdicts.get("main_pump",   "N/A")),
        (3, reqs.get("backup_pump", "N/A"), verdicts.get("backup_pump", "N/A")),
        (5, reqs.get("jockey_pump", "N/A"), verdicts.get("jockey_pump", "N/A")),
        (7, reqs.get("terrace_booster","N/A"), verdicts.get("terrace_booster","N/A")),
    ]
    for row_idx, req, verdict in pump_rows:
        if row_idx < len(table.rows):
            _write(table.rows[row_idx].cells[1], str(req))
            _write(table.rows[row_idx].cells[3], verdict)


def _fill_nbc_four_col(table, items):
    """
    Generic filler for NBC tables 20-24 (Description, NBC Req, Actual, Remarks).
    items: list of (row_idx, actual_val, verdict_str)
    """
    for row_idx, actual, verdict in items:
        if row_idx < len(table.rows):
            _write(table.rows[row_idx].cells[2], _na(actual))
            _write(table.rows[row_idx].cells[3], _na(verdict))


def _fill_checklist_yn(table, checklist_items, yn_col=2, rmk_col=3):
    """Tables 25-27 — fill Yes/No + Remarks columns from checklist data."""
    for i, item in enumerate(checklist_items):
        row_idx = i + 1
        if row_idx >= len(table.rows):
            break
        _write(table.rows[row_idx].cells[yn_col], _yn(item.get("yes_no")))
        rmk = _na(item.get("remarks"))
        _write(table.rows[row_idx].cells[rmk_col], rmk, caps=False)


def _fill_amc_table(table, amc_list):
    """Table 28 — AMC Maintenance Table."""
    item_labels = [
        "FIRE PUMPS", "FIRE HYDRANT SYSTEM", "FIRE SPRINKLER SYSTEM",
        "FIRE ALARM SYSTEM", "FIRE EXTINGUISHERS", "FIRE SUPPRESSION SYSTEM", "FIRE EXIT DOORS"
    ]
    for row_idx, label in enumerate(item_labels, start=1):
        if row_idx >= len(table.rows):
            break
        amc_item = next((a for a in amc_list if label in str(a.get("item","")).upper()), None)
        if amc_item:
            _write(table.rows[row_idx].cells[1], _na(amc_item.get("amc_status")))
            _write(table.rows[row_idx].cells[2], _na(amc_item.get("frequency")))
            _write(table.rows[row_idx].cells[3], _na(amc_item.get("report_status")))
            _write(table.rows[row_idx].cells[4], _na(amc_item.get("remarks")), caps=False)


def _fill_training_tables(table29, table30, training_list):
    """Tables 29-30 — Training and Records Checklists."""
    part1_keys = [
        "induction training", "Evacuation Drills", "QRT", "Extinguisher", "Evacuation Points"
    ]
    part2_keys = [
        "Service Records", "Mock Drill", "Staff Fire Safety Training", "System Testing", "Security Team", "AMC"
    ]
    for i, kw in enumerate(part1_keys, start=1):
        if i >= len(table29.rows): break
        item = next((t for t in training_list if kw.lower() in str(t.get("label","")).lower()), None)
        if item:
            _write(table29.rows[i].cells[1], _yn(item.get("yes_no")))
            _write(table29.rows[i].cells[2], _na(item.get("remarks")), caps=False)

    for i, kw in enumerate(part2_keys, start=1):
        if i >= len(table30.rows): break
        item = next((t for t in training_list if kw.lower() in str(t.get("label","")).lower()), None)
        if item:
            _write(table30.rows[i].cells[1], _yn(item.get("yes_no")))
            _write(table30.rows[i].cells[2], _na(item.get("remarks")), caps=False)


def _fill_location_table(table, locations):
    """Table 31 — 9.2 Location-wise Equipment Report."""
    # Category headers in table match location dict keys
    cat_map = {
        "FIRE HYDRANT VALVES":       "hydrant_valves",
        "FIRE HOSE WITH ACCESSORIES":"hose_accessories",
        "FIRE HOSE REEL":            "hose_reel",
        "FIRE ALARM PANEL":          "alarm_panel",
        "MANUAL CALL POINT":         "manual_call_point",
        "HOOTER":                    "hooter",
        "SMOKE DETECTOR":            "smoke_detector",
        "FIRE EXTINGUISHER":         "fire_extinguishers",
    }

    current_key = None
    item_iter   = {}
    counters    = {k: 0 for k in cat_map.values()}

    for row in table.rows:
        cells = row.cells
        label = cells[0].text.strip().upper() + " " + cells[1].text.strip().upper()

        # Check if this is a category header row
        for cat_label, cat_key in cat_map.items():
            if cat_label in label:
                current_key = cat_key
                break
        else:
            # Data row — fill if we have items for current category
            if current_key and not cells[1].text.strip():
                items = locations.get(current_key, [])
                idx = counters[current_key]
                if idx < len(items):
                    item = items[idx]
                    _write(cells[1], _na(item.get("location")))
                    _write(cells[2], _na(item.get("maintenance")))
                    _write(cells[3], _na(item.get("remarks")), caps=False)
                    counters[current_key] += 1


# ---------------------------------------------------------------------------
# PARAGRAPH TOKEN REPLACER
# ---------------------------------------------------------------------------

def _replace_in_paragraph(para, old, new):
    """Replace token in paragraph across all runs, preserving font."""
    if old not in para.text:
        return
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            run.font.name = FONT_NAME
            run.font.size = Pt(FONT_SIZE_PT)


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def generate_report(template_path, output_path, store_data):
    """Fill every section of the Word template and save to output_path."""
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Word template not found at: {template_path}")

    doc = Document(template_path)
    tables = doc.tables

    store_code   = str(store_data.get("store_code", "XXXX"))
    prop         = store_data.get("property_overview", {})
    bldg         = store_data.get("building_details", {})
    bldg_check   = store_data.get("building_checklist", {})
    store        = store_data.get("store_details", {})
    store_check  = store_data.get("store_checklist", {})
    equip        = store_data.get("equipment", {})
    checklists   = store_data.get("checklists", {})
    locations    = store_data.get("locations", {})
    occ_table    = store_data.get("occupancy_table", [])
    fire_table   = store_data.get("fire_load_table", [])
    masses       = store_data.get("material_masses", {})
    nbc          = store_data.get("nbc", {})
    verdicts     = nbc.get("verdicts", {})

    date_start   = store_data.get("audit_date_start", "N/A")
    date_end     = store_data.get("audit_date_end", "N/A")
    report_date  = store_data.get("report_date", "N/A")
    occ_load     = store_data.get("occupant_load", "N/A")
    fire_density = store_data.get("fire_load_density", "N/A")
    fire_class   = store_data.get("fire_load_classification", "N/A")

    # --- 1. Replace paragraph tokens ---
    for para in doc.paragraphs:
        _replace_in_paragraph(para, "XXXX", store_code)
        _replace_in_paragraph(para, "05/05/2026", report_date)
        # Summary date range
        for fragment in ["25th April and completed on 04th May 2026",
                         "25th April 2026 and completed on 04th May 2026"]:
            _replace_in_paragraph(para, fragment,
                f"{date_start} and completed on {date_end}")

    # Also replace in table cells
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_in_paragraph(para, "XXXX", store_code)
                    _replace_in_paragraph(para, "05/05/2026", report_date)

    # --- 2. Replace footer store code ---
    for section in doc.sections:
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for para in footer.paragraphs:
                    _replace_in_paragraph(para, "1234", store_code)
                    _replace_in_paragraph(para, "XXXX", store_code)

    # --- 3. Fill all data tables ---
    _fill_31_property_overview(tables[3], prop)
    _fill_32_building_details(tables[4], bldg)
    _fill_checklist_table(tables[5], bldg_check)
    _fill_34_store_details(tables[6], store)
    _fill_checklist_table(tables[7], store_check)
    _fill_41a_water_tanks(tables[8], equip)
    _fill_41c_hydrant(tables[9], equip)
    _fill_41d_sprinkler(tables[10], equip)
    _fill_41e_alarm(tables[11], equip)
    _fill_41f_passive(tables[12], equip)
    _fill_41g_extinguisher(tables[13], equip)

    if occ_table:
        _fill_occupancy_table(tables[15], occ_table)
    if fire_table:
        _fill_fire_load_table(tables[16], fire_table, masses)

    # NBC tables
    _fill_nbc_water_tanks(tables[18], nbc, equip)
    _fill_nbc_pumps(tables[19], nbc)

    # NBC Hydrant (table 20) — actual values from equipment
    hyd = equip.get("hydrant", [])
    _fill_nbc_four_col(tables[20], [
        (1,  _find_equip_val(hyd, "Courtyard"),              verdicts.get("hydrant_valve")),
        (3,  _find_equip_val(hyd, "Internal Fire Hydrant"),  verdicts.get("hydrant_valve")),
        (5,  _find_equip_val(hyd, "Fire Hose Box"),          "N/A"),
        (7,  _find_equip_val(hyd, "Fire Hose including"),    "N/A"),
        (9,  _find_equip_val(hyd, "Hose Reel"),              verdicts.get("hose_reel")),
        (11, _find_equip_val(hyd, "Fire Brigade Inlet"),     "N/A"),
    ])

    # NBC Sprinkler (table 21)
    spr = equip.get("sprinkler", [])
    _fill_nbc_four_col(tables[21], [
        (1, _find_equip_val(spr, "Pendent Sprinkler", "quantity"), verdicts.get("sprinkler_installed")),
        (3, _find_equip_val(spr, "Suppression"),                   "N/A"),
    ])

    # NBC Alarm (table 22)
    alm = equip.get("alarm", [])
    _fill_nbc_four_col(tables[22], [
        (1,  _find_equip_val(alm, "Alarm Panel"),     verdicts.get("alarm_panel")),
        (3,  _find_equip_val(alm, "Public Address"),  verdicts.get("public_address")),
        (5,  _find_equip_val(alm, "MCP"),             verdicts.get("mcp")),
        (7,  _find_equip_val(alm, "Hooter"),          verdicts.get("hooter")),
        (9,  _find_equip_val(alm, "Smoke Detector"),  verdicts.get("smoke_detector")),
        (11, _find_equip_val(alm, "Heat Detector"),   verdicts.get("heat_detector")),
    ])

    # NBC Extinguisher (table 23)
    ext = equip.get("extinguisher", [])
    _fill_nbc_four_col(tables[23], [
        (1, _find_equip_val(ext, "CO2", "quantity"),          verdicts.get("co2_extinguisher")),
        (3, _find_equip_val(ext, "ABC", "quantity"),          verdicts.get("abc_extinguisher")),
        (5, _find_equip_val(ext, "MODULAR", "quantity"),      verdicts.get("modular_extinguisher")),
        (7, _find_equip_val(ext, "CLEAN AGENT", "quantity"),  verdicts.get("clean_agent")),
    ])

    # Checklists (tables 25-27) — fill from Excel data
    if checklists.get("evacuation"):
        _fill_checklist_yn(tables[25], checklists["evacuation"])
    if checklists.get("smoke_ventilation"):
        _fill_checklist_yn(tables[26], checklists["smoke_ventilation"])
    if checklists.get("electrical_safety"):
        _fill_checklist_yn(tables[27], checklists["electrical_safety"])

    # AMC + Training (tables 28-30)
    _fill_amc_table(tables[28], equip.get("amc", []))
    _fill_training_tables(tables[29], tables[30], equip.get("maintenance", []))

    # Location table (table 31)
    _fill_location_table(tables[31], locations)

    doc.save(output_path)
    return output_path
