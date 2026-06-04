# src/validators.py
# Pre-flight checks run before report generation.
# Returns a list of warning strings. Empty list = all checks passed.

def run_preflight_checks(store_data, filename):
    """
    Validates store_data dict before allowing report generation.
    Returns list of warning strings. Empty = safe to proceed.
    """
    warnings = []

    # 1. Store code detected from filename
    store_code = store_data.get("store_code", "XXXX")
    if store_code == "XXXX":
        warnings.append(
            "Store code could not be detected from filename. "
            "Filename should start with store code e.g. 'Z943-...'."
        )

    # 2. Property area (section 3.1) — must be a number, used for NBC lookup
    property_area = store_data.get("property_area")
    if not isinstance(property_area, (int, float)):
        warnings.append(
            "Property Area (section 3.1 'Total Built-up Area') is missing or not numeric. "
            "NBC necessity check will be skipped."
        )

    # 3. Store area (section 3.4) — must be a number, used for formula calculations
    store_area = store_data.get("store_area")
    if not isinstance(store_area, (int, float)):
        warnings.append(
            "Store Area (section 3.4 'Total Built-up Area') is missing or not numeric. "
            "Occupancy load and fire load density will not be calculated."
        )

    # 4. Material masses — needed for fire load calculation
    material_masses = store_data.get("material_masses", {})
    total_mass = sum(v for v in material_masses.values()
                     if isinstance(v, (int, float)))
    if total_mass == 0:
        warnings.append(
            "All material masses are zero or missing. "
            "Fire load density will calculate as 0."
        )

    # 5. Equipment sections — warn if any major section is empty
    equipment = store_data.get("equipment", {})
    for section_key, section_label in [
        ("water_tanks",  "4.1.A Water Tank Details"),
        ("pump_room",    "4.1.B Fire Pump Room Details"),
        ("hydrant",      "4.1.C Fire Hydrant System"),
        ("sprinkler",    "4.1.D Fire Sprinkler System"),
        ("alarm",        "4.1.E Fire Alarm System"),
        ("extinguisher", "4.1.G Fire Extinguisher"),
    ]:
        if not equipment.get(section_key):
            warnings.append(
                f"Section {section_label} could not be read from Excel. "
                "Check sheet name and layout."
            )

    # 6. Checklists — warn if empty
    checklists = store_data.get("checklists", {})
    for key, label in [
        ("evacuation",        "7.1 Emergency Evacuation Checklist"),
        ("smoke_ventilation", "7.2 Smoke Ventilation Checklist"),
        ("electrical_safety", "7.3 Electrical Safety Checklist"),
    ]:
        if not checklists.get(key):
            warnings.append(
                f"Checklist section {label} could not be read. "
                "Check sheet name and layout."
            )

    return warnings
