def run_preflight_checks(store_data, filename):
    """
    Validates data before allowing the report generation to proceed.
    Returns a list of warning strings. If empty, all checks passed.
    """
    warnings = []

    # 1. Check Filename for Store Code (assuming filename should have it)
    store_code = store_data.get("store_code")
    if store_code and str(store_code) not in filename:
        warnings.append(
            f"Store code '{store_code}' not found in filename '{filename}'.")

    # 2. Numeric Checks
    if not isinstance(store_data.get("property_area"), (int, float)):
        warnings.append("Property Area is missing or not a valid number.")

    if not isinstance(store_data.get("store_area"), (int, float)):
        warnings.append("Store Area is missing or not a valid number.")

    # 3. Required Yes/No Fields
    checklist_fields = ["alarms_working", "exits_clear"]
    for field in checklist_fields:
        val = str(store_data.get(field, "")).strip().upper()
        if val not in ["YES", "NO"]:
            warnings.append(
                f"Checklist field '{field}' must be YES or NO. Found: '{val}'.")

    return warnings
