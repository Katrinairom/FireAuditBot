# src/constants.py
# Single source of truth for all reference data from Inspection Sheet - Drafter's Copy.xlsx
# Do NOT hardcode these numbers anywhere else in the project.

# ---------------------------------------------------------------------------
# OCCUPANCY LOAD CALCULATION
# Source: Drafter's Copy > Sheet "Occupancy load Calculation"
# Formula per zone: zone_area = (area_pct / 100) * store_area
#                   zone_load = zone_area / olp
#                   total_load = sum of all zone_loads
# ---------------------------------------------------------------------------
OLP_ZONES = [
    {"zone": "STORE ENTRANCE LOBBY",       "olp": 2.6, "area_pct": 2.5},
    {"zone": "RETAIL OUTLET AREA",         "olp": 2.6, "area_pct": 39.0},
    {"zone": "MERCHANDISE DISPLAY FLOOR",  "olp": 2.6, "area_pct": 18.0},
    {"zone": "CASH COUNTER",               "olp": 4.0, "area_pct": 8.0},
    {"zone": "MAIN EXIT AREA",             "olp": 2.6, "area_pct": 2.5},
    {"zone": "TRIAL ROOMS",                "olp": 4.0, "area_pct": 5.0},
    {"zone": "STOCK STORAGE AREA",         "olp": 5.0, "area_pct": 8.0},
    {"zone": "ELECTRICAL / IT PANEL ROOM", "olp": 5.0, "area_pct": 3.0},
    {"zone": "ADMINISTRATIVE ROOM",        "olp": 4.0, "area_pct": 5.0},
    {"zone": "PANTRY / STAFF KITCHEN",     "olp": 4.0, "area_pct": 2.5},
    {"zone": "EMPLOYEE ROOMS",             "olp": 4.0, "area_pct": 2.5},
    {"zone": "CORRIDOR / STAIRCASES",      "olp": 5.0, "area_pct": 2.0},
    {"zone": "EMERGENCY EXIT",             "olp": 5.0, "area_pct": 2.0},
]

# ---------------------------------------------------------------------------
# FIRE LOAD CALCULATION
# Source: Drafter's Copy > Sheet "Fire Load Calculation"
# Calorific values in MJ/kg. Keys are lowercase for easy matching.
# ---------------------------------------------------------------------------
CALORIFIC_VALUES = {
    "paper":        15.4,
    "cotton":       15.8,
    "leather":      17.6,
    "wood":         17.6,
    "wool":         19.6,
    "nylon":        22.0,
    "polyester":    22.0,
    "rubber":       37.4,
    "ethane":       47.3,
    "polyethylene": 48.4,
}

# Ordered list matching row positions in store Excel material mass column (rows 1-10)
MATERIAL_ORDER = [
    "paper", "cotton", "leather", "wood", "wool",
    "nylon", "polyester", "rubber", "ethane", "polyethylene"
]

# Fire load density thresholds from Drafter's Copy
# Format: (upper_bound_inclusive, classification_label)
FIRE_LOAD_THRESHOLDS = [
    (275,          "LOW FIRE LOAD"),
    (550,          "MODERATE FIRE LOAD"),
    (float("inf"), "HIGH FIRE LOAD"),
]

# ---------------------------------------------------------------------------
# NBC NECESSITY RULES
# Source: Drafter's Copy > Sheet "NBC Necessity"
#
# CRITICAL: These rules use PROPERTY area (section 3.1 of store Excel),
#           NOT store area (section 3.4).
#
# CRITICAL: List is sorted HIGHEST first. The checker iterates top-down
#           and takes the FIRST match. This means at exactly 5000 m²,
#           the stricter ">= 5000" band applies correctly.
# ---------------------------------------------------------------------------
NBC_RULES = [
    {
        "label":    "ABOVE 10000 SQ. MTR",
        "min_area": 10001,
        "requirements": {
            "terrace_booster":  "900 LPM+ REQUIRED IF G+2 & ABOVE",
            "jockey_pump":      "900 LPM & ABOVE",
            "main_pump":        "2250 LPM & ABOVE",
            "backup_pump":      "2250 LPM & ABOVE",
            "common_ug_tank":   150000,
            "fire_ug_tank":     75000,
            "common_oh_tank":   50000,
            "fire_oh_tank":     25000,
        },
    },
    {
        "label":    "5000 TO 10000 SQ. MTR",
        "min_area": 5000,
        "requirements": {
            "terrace_booster":  "900 LPM REQUIRED IF G+2 & ABOVE",
            "jockey_pump":      "900 LPM & ABOVE",
            "main_pump":        "1850 LPM TO 2250 LPM",
            "backup_pump":      "1850 LPM TO 2250 LPM",
            "common_ug_tank":   100000,
            "fire_ug_tank":     50000,
            "common_oh_tank":   30000,
            "fire_oh_tank":     10000,
        },
    },
    {
        "label":    "2000 TO 5000 SQ. MTR",
        "min_area": 2000,
        "requirements": {
            "terrace_booster":  "450 LPM REQUIRED IF G+2 & ABOVE",
            "jockey_pump":      "450 TO 900 LPM",
            "main_pump":        "1350 LPM TO 1850 LPM",
            "backup_pump":      "NOT COMPULSORY",
            "common_ug_tank":   50000,
            "fire_ug_tank":     25000,
            "common_oh_tank":   25000,
            "fire_oh_tank":     5000,
        },
    },
    {
        "label":    "500 TO 2000 SQ. MTR",
        "min_area": 500,
        "requirements": {
            "terrace_booster":  "450 LPM REQUIRED IF G+2 & ABOVE",
            "jockey_pump":      "180 TO 450 LPM",
            "main_pump":        "NOT COMPULSORY",
            "backup_pump":      "NOT COMPULSORY",
            "common_ug_tank":   None,
            "fire_ug_tank":     None,
            "common_oh_tank":   30000,
            "fire_oh_tank":     5000,
        },
    },
]

# Fixed NBC rules applied to ALL properties >= 500 m²
NBC_FIXED_RULES = {
    "sprinkler_system":     "REQUIRED IF PROPERTY AREA EXCEEDS 500 SQ. MTR",
    "hydrant_courtyard":    "REQUIRED IF PROPERTY AREA EXCEEDS 500 SQ. MTR",
    "hose_reel_drum":       "REQUIRED ON GROUND AND EACH FLOOR",
    "riser_downcomer":      "REQUIRED IF FLOORS EXCEED G+2",
    "alarm_panel":          "ALARM PANEL, SMOKE DETECTOR, MCP, HOOTER REQUIRED IN STORE",
    "heat_detector":        "REQUIRED IN ELECTRICAL PANEL ROOM",
    "public_address":       "REQUIRED IN STORE",
    "fire_extinguisher":    "REQUIRED IN STORE",
    "co2_extinguisher":     "REQUIRED IN ELECTRICAL ROOM",
    "clean_agent":          "REQUIRED IN ELECTRICAL ROOM",
    "modular_extinguisher": "REQUIRED IN ELECTRICAL ROOM AND STORAGE ROOM",
}

# NBC verdict strings
VERDICT_MEETS   = "MEETS EXPECTATION"
VERDICT_EXCEEDS = "EXCEEDS EXPECTATION"
VERDICT_FAILS   = "DOES NOT MEET EXPECTATION"
VERDICT_NA      = "N/A"

# ---------------------------------------------------------------------------
# REPORT FORMATTING
# ---------------------------------------------------------------------------
FONT_NAME    = "Calibri"
FONT_SIZE_PT = 12

# Sheet names in the store Excel (exact as they appear)
SHEET_PROPERTY   = "Property and Store Details"
SHEET_EQUIPMENT  = "Firefighting Equipment Details"
SHEET_CHECKLIST  = "Checklist"
SHEET_LOCATION   = "Firefighting Equipment Store Wi"
