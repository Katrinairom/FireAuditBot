# src/formula_engine.py
# Calculates Occupant Load and Fire Load Density from store_data.
# All reference constants come from constants.py — no file I/O here.

from .constants import OLP_ZONES, CALORIFIC_VALUES, MATERIAL_ORDER, FIRE_LOAD_THRESHOLDS


def calculate_loads(store_data):
    """
    Calculates Occupant Load and Fire Load Density.
    Returns a dict with keys:
        occupant_load           - int (total people estimate)
        occupancy_table         - list of per-zone breakdown dicts
        fire_load_density       - float (MJ/m²)
        fire_load_classification - str
        fire_load_table         - list of per-material breakdown dicts
    """
    results = {
        "occupant_load":            None,
        "occupancy_table":          [],
        "fire_load_density":        None,
        "fire_load_classification": "N/A",
        "fire_load_table":          [],
    }

    store_area = store_data.get("store_area")

    # ------------------------------------------------------------------
    # 1. OCCUPANT LOAD CALCULATION
    # Uses all 13 zones from Drafter's Copy.
    # Formula: zone_area = (area_pct / 100) * store_area
    #          zone_load = zone_area / olp
    #          total     = sum of all zone_loads
    # ------------------------------------------------------------------
    if isinstance(store_area, (int, float)) and store_area > 0:
        total_load = 0.0
        table_rows = []

        for zone in OLP_ZONES:
            zone_area = (zone["area_pct"] / 100.0) * store_area
            zone_load = zone_area / zone["olp"]
            total_load += zone_load
            table_rows.append({
                "zone":          zone["zone"],
                "olp":           zone["olp"],
                "area_pct":      zone["area_pct"],
                "area_occupied": round(zone_area, 2),
                "occupancy_load": round(zone_load, 2),
            })

        results["occupant_load"] = round(total_load)
        results["occupancy_table"] = table_rows

    # ------------------------------------------------------------------
    # 2. FIRE LOAD DENSITY CALCULATION
    # mass per material is read from store Excel material mass column.
    # Formula: fire_load = sum(mass_i * calorific_value_i)
    #          density   = fire_load / store_area
    # ------------------------------------------------------------------
    material_masses = store_data.get("material_masses", {})

    if isinstance(store_area, (int, float)) and store_area > 0 and material_masses:
        total_mj = 0.0
        fire_table = []

        for mat in MATERIAL_ORDER:
            mass = material_masses.get(mat, 0.0)
            cv = CALORIFIC_VALUES.get(mat, 0.0)
            mj = mass * cv
            total_mj += mj
            fire_table.append({
                "material":         mat.upper(),
                "calorific_value":  cv,
                "mass_kg":          mass,
                "total_mj":         round(mj, 2),
            })

        density = total_mj / store_area
        results["fire_load_density"] = round(density, 2)
        results["fire_load_table"] = fire_table

        # Classification from Drafter's Copy thresholds
        # ≤275 = LOW, 275-550 = MODERATE, >550 = HIGH
        for upper_bound, label in FIRE_LOAD_THRESHOLDS:
            if density <= upper_bound:
                results["fire_load_classification"] = label
                break

    return results
