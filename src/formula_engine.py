from .constants import OLP_ZONES, CALORIFIC_VALUES


def calculate_loads(store_data):
    """
    Calculates Occupant Load and Fire Load Density.
    Returns a dictionary of calculated results.
    """
    results = {
        "occupant_load": None,
        "fire_load_density": None,
        "fire_load_classification": "N/A"
    }

    # 1. Occupant Load Calculation
    store_area = store_data.get("store_area")
    zone_type = store_data.get("zone_type")

    if store_area and isinstance(store_area, (int, float)) and zone_type in OLP_ZONES:
        olp_factor = OLP_ZONES[zone_type]
        # Occupant Load = Area / OLP Factor
        results["occupant_load"] = round(store_area / olp_factor)

    # 2. Fire Load Density Calculation
    materials = store_data.get("materials", [])
    if store_area and isinstance(store_area, (int, float)) and store_area > 0 and materials:
        total_fire_load = 0
        for mat in materials:
            mat_type = mat.get("material_type", "").lower()
            mass = mat.get("mass_kg", 0)

            if mat_type in CALORIFIC_VALUES and isinstance(mass, (int, float)):
                # Heat energy = mass * calorific value
                total_fire_load += (mass * CALORIFIC_VALUES[mat_type])

        # Density = Total Heat / Store Area
        density = total_fire_load / store_area
        results["fire_load_density"] = round(density, 2)

        # Basic Classification (Example thresholds, adjust as needed)
        if density < 100:
            results["fire_load_classification"] = "LOW"
        elif density < 500:
            results["fire_load_classification"] = "MODERATE"
        else:
            results["fire_load_classification"] = "HIGH"

    return results
