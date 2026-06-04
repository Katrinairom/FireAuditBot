# src/nbc_checker.py
# NBC Necessity checker.
# Uses PROPERTY area (section 3.1) — NOT store area (section 3.4) — to select rule tier.
# Iterates NBC_RULES top-down (highest first) and takes the first match.

from .constants import NBC_RULES, NBC_FIXED_RULES, VERDICT_MEETS, VERDICT_EXCEEDS, VERDICT_FAILS, VERDICT_NA


def _parse_litres(val):
    """Extract numeric litres from strings like '60,000 ltr' or '30000 Ltr'. Returns float or None."""
    if val is None:
        return None
    import re
    match = re.search(r"[\d,]+\.?\d*", str(val).replace(",", ""))
    return float(match.group()) if match else None


def _verdict(actual_litres, required_litres):
    """Compare actual vs required tank/pump capacity and return verdict string."""
    if required_litres is None:
        # Not compulsory — having it is exceeding expectation
        if actual_litres and actual_litres > 0:
            return VERDICT_EXCEEDS
        return VERDICT_NA
    if actual_litres is None or actual_litres == 0:
        return VERDICT_FAILS
    if actual_litres >= required_litres:
        return VERDICT_MEETS if actual_litres == required_litres else VERDICT_EXCEEDS
    return VERDICT_FAILS


def _yn_verdict(actual_yn, is_required):
    """Return verdict for a YES/NO field against a boolean requirement."""
    yn = str(actual_yn or "").strip().upper()
    present = yn == "YES"
    if is_required:
        return VERDICT_MEETS if present else VERDICT_FAILS
    return VERDICT_EXCEEDS if present else VERDICT_NA


def check_compliance(store_data):
    """
    Uses property_area (section 3.1) to find the applicable NBC tier.
    Compares actual equipment data from store_data against requirements.
    Returns dict with:
        nbc_tier_applied  - label string
        verdicts          - dict of item → verdict string
        requirements      - dict of the applied NBC requirements for report display
    """
    property_area = store_data.get("property_area")

    if not isinstance(property_area, (int, float)):
        return {
            "nbc_tier_applied": "UNKNOWN — PROPERTY AREA NOT FOUND",
            "verdicts":         {},
            "requirements":     {},
        }

    # --- Find applicable rule tier (top-down, highest first) ---
    applied_rule = None
    for rule in NBC_RULES:  # Already sorted highest-first in constants.py
        if property_area >= rule["min_area"]:
            applied_rule = rule
            break

    if applied_rule is None:
        return {
            "nbc_tier_applied": "BELOW 500 SQ. MTR — NBC RULES NOT APPLICABLE",
            "verdicts":         {},
            "requirements":     {},
        }

    reqs     = applied_rule["requirements"]
    verdicts = {}
    equip    = store_data.get("equipment", {})

    # ------------------------------------------------------------------
    # 6.1.A — WATER TANK DETAILS
    # Reads actual capacity from equipment.water_tanks list
    # ------------------------------------------------------------------
    tank_data = {item["label"]: item for item in equip.get("water_tanks", [])}

    def _get_tank_litres(label_keyword):
        for label, item in tank_data.items():
            if label_keyword.upper() in label.upper():
                return _parse_litres(item.get("quantity"))
        return None

    common_ug_actual = _get_tank_litres("U/G Water Tank Capacity of Building")
    fire_ug_actual   = _get_tank_litres("U/G Water Tank Capacity")
    common_oh_actual = _get_tank_litres("O/H Water Tank Capacity of Building")
    fire_oh_actual   = _get_tank_litres("O/H Water Tank Capacity")

    verdicts["common_ug_tank"]  = _verdict(common_ug_actual, reqs["common_ug_tank"])
    verdicts["fire_ug_tank"]    = _verdict(fire_ug_actual,   reqs["fire_ug_tank"])
    verdicts["common_oh_tank"]  = _verdict(common_oh_actual, reqs["common_oh_tank"])
    verdicts["fire_oh_tank"]    = _verdict(fire_oh_actual,   reqs["fire_oh_tank"])

    # ------------------------------------------------------------------
    # 6.1.B — FIRE PUMP DETAILS
    # Main pump, backup pump, jockey, booster — YES/NO presence check
    # ------------------------------------------------------------------
    pump_data = {item["label"]: item for item in equip.get("pump_room", [])}

    def _pump_present(keyword):
        for label, item in pump_data.items():
            if keyword.upper() in label.upper():
                return str(item.get("yes_no") or "").upper() == "YES"
        return False

    main_pump_present   = _pump_present("Electric Fire Pump – Main")
    backup_pump_present = _pump_present("Diesel Fire Engine") or _pump_present("Standby")
    jockey_present      = _pump_present("Jockey")
    booster_present     = _pump_present("Booster")

    main_required   = reqs["main_pump"]   != "NOT COMPULSORY"
    backup_required = reqs["backup_pump"] != "NOT COMPULSORY"

    verdicts["main_pump"]        = _yn_verdict("YES" if main_pump_present else "NO",   main_required)
    verdicts["backup_pump"]      = _yn_verdict("YES" if backup_pump_present else "NO", backup_required)
    verdicts["jockey_pump"]      = _yn_verdict("YES" if jockey_present else "NO",      True)
    verdicts["terrace_booster"]  = _yn_verdict("YES" if booster_present else "NO",     False)

    # ------------------------------------------------------------------
    # 6.1.C — FIRE HYDRANT (Fixed rules — required for all >= 500 m²)
    # ------------------------------------------------------------------
    hydrant_data = {item["label"]: item for item in equip.get("hydrant", [])}

    def _hydrant_yn(keyword):
        for label, item in hydrant_data.items():
            if keyword.upper() in label.upper():
                return str(item.get("yes_no") or "").upper() == "YES"
        return False

    verdicts["hydrant_valve"]     = _yn_verdict("YES" if _hydrant_yn("Hydrant Valve") else "NO", True)
    verdicts["hose_reel"]         = _yn_verdict("YES" if _hydrant_yn("Hose Reel") else "NO",     True)
    verdicts["fire_brigade_inlet"]= _yn_verdict("YES" if _hydrant_yn("Fire Brigade Inlet") else "NO", True)

    # ------------------------------------------------------------------
    # 6.1.D — SPRINKLER (Fixed rule — required for all >= 500 m²)
    # ------------------------------------------------------------------
    spr_data = {item["label"]: item for item in equip.get("sprinkler", [])}

    def _spr_yn(keyword):
        for label, item in spr_data.items():
            if keyword.upper() in label.upper():
                return str(item.get("yes_no") or "").upper() == "YES"
        return False

    verdicts["sprinkler_installed"]   = _yn_verdict("YES" if _spr_yn("Sprinkler") else "NO", True)
    verdicts["isolation_valve"]       = _yn_verdict("YES" if _spr_yn("Isolation Valve") else "NO", True)
    verdicts["pressure_gauge"]        = _yn_verdict("YES" if _spr_yn("Pressure Gauge") else "NO", True)

    # ------------------------------------------------------------------
    # 6.1.E — FIRE ALARM (Fixed rule)
    # ------------------------------------------------------------------
    alarm_data = {item["label"]: item for item in equip.get("alarm", [])}

    def _alarm_yn(keyword):
        for label, item in alarm_data.items():
            if keyword.upper() in label.upper():
                return str(item.get("yes_no") or "").upper() == "YES"
        return False

    verdicts["alarm_panel"]    = _yn_verdict("YES" if _alarm_yn("Alarm Panel") else "NO",     True)
    verdicts["smoke_detector"] = _yn_verdict("YES" if _alarm_yn("Smoke Detector") else "NO",  True)
    verdicts["mcp"]            = _yn_verdict("YES" if _alarm_yn("MCP") else "NO",             True)
    verdicts["hooter"]         = _yn_verdict("YES" if _alarm_yn("Hooter") else "NO",          True)
    verdicts["heat_detector"]  = _yn_verdict("YES" if _alarm_yn("Heat Detector") else "NO",   True)
    verdicts["public_address"] = _yn_verdict("YES" if _alarm_yn("Public Address") else "NO",  True)

    # ------------------------------------------------------------------
    # 6.1.G — EXTINGUISHERS (Fixed rule)
    # ------------------------------------------------------------------
    ext_data = {item["label"]: item for item in equip.get("extinguisher", [])}

    def _ext_yn(keyword):
        for label, item in ext_data.items():
            if keyword.upper() in label.upper():
                return str(item.get("yes_no") or "").upper() == "YES"
        return False

    verdicts["abc_extinguisher"]    = _yn_verdict("YES" if _ext_yn("ABC") else "NO",          True)
    verdicts["co2_extinguisher"]    = _yn_verdict("YES" if _ext_yn("CO2") else "NO",          True)
    verdicts["clean_agent"]         = _yn_verdict("YES" if _ext_yn("CLEAN AGENT") else "NO",  True)
    verdicts["modular_extinguisher"]= _yn_verdict("YES" if _ext_yn("MODULAR") else "NO",      True)

    # ------------------------------------------------------------------
    # 6.1.F — PASSIVE (Fixed rules)
    # ------------------------------------------------------------------
    passive_data = {item["label"]: item for item in equip.get("passive", [])}

    def _passive_yn(keyword):
        for label, item in passive_data.items():
            if keyword.upper() in label.upper():
                return str(item.get("yes_no") or "").upper() == "YES"
        return False

    verdicts["autoglow_signage"]    = _yn_verdict("YES" if _passive_yn("Auto Glow") else "NO",     True)
    verdicts["emergency_exit_sign"] = _yn_verdict("YES" if _passive_yn("LED Emergency") else "NO", True)
    verdicts["evacuation_plan"]     = _yn_verdict("YES" if _passive_yn("Evacuation Plan") else "NO", True)
    verdicts["assembly_area"]       = _yn_verdict("YES" if _passive_yn("Assembly Area") else "NO",   True)

    return {
        "nbc_tier_applied": applied_rule["label"],
        "verdicts":         verdicts,
        "requirements":     reqs,
        "fixed_rules":      NBC_FIXED_RULES,
    }
