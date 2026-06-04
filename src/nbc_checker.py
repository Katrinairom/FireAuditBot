from .constants import NBC_RULES


def check_compliance(store_data):
    """
    Uses property_area to find the strict top-down NBC requirement tier.
    Evaluates store data against these requirements.
    """
    property_area = store_data.get("property_area")

    # Default fallback if no area is provided or area is invalid
    if not isinstance(property_area, (int, float)):
        return {"nbc_tier_applied": "None", "verdicts": {}}

    applied_rules = None

    # 1. Find the applicable rule tier (Strict Top-Down)
    for rule in NBC_RULES:
        if property_area >= rule["min_area"]:
            applied_rules = rule
            break

    if not applied_rules:
        return {"nbc_tier_applied": "< 500 sqm", "verdicts": {}}

    reqs = applied_rules["requirements"]
    verdicts = {}

    # 2. Compare Actuals against Requirements

    # Example: Fire Alarms
    actual_alarms = str(store_data.get("alarms_working", "")).upper()
    if reqs["fire_alarms_required"]:
        if actual_alarms == "YES":
            verdicts["fire_alarms"] = "MEETS EXPECTATION"
        else:
            verdicts["fire_alarms"] = "DOES NOT MEET EXPECTATION"
    else:
        verdicts["fire_alarms"] = "EXCEEDS EXPECTATION" if actual_alarms == "YES" else "N/A"

    # Example: Extinguishers (assuming actual_count / store_area gives ratio)
    store_area = store_data.get("store_area")
    actual_extinguishers = store_data.get("extinguisher_count")

    if store_area and isinstance(actual_extinguishers, (int, float)):
        # Calculate how many sqm per extinguisher we actually have
        actual_ratio = store_area / actual_extinguishers if actual_extinguishers > 0 else 999999
        req_ratio = reqs["extinguishers_per_sqm"]

        # Lower ratio is better (e.g., 1 per 50 sqm is better than 1 per 100 sqm)
        if actual_ratio <= req_ratio:
            verdicts["extinguishers"] = "MEETS EXPECTATION"
        else:
            verdicts["extinguishers"] = "DOES NOT MEET EXPECTATION"

    return {
        "nbc_tier_applied": f">= {applied_rules['min_area']} sqm",
        "verdicts": verdicts
    }
