# src/constants.py

# 1. Calorific Values (10 items) - Update with exact values from Drafter's Copy
CALORIFIC_VALUES = {
    "wood": 17.5,
    "paper": 16.0,
    "plastic": 40.0,
    "cardboard": 16.5,
    "textiles": 18.0,
    "rubber": 30.0,
    "leather": 19.0,
    "chemicals": 25.0,
    "metals": 0.0,
    "mixed_waste": 15.0
}

# 2. OLP (Occupant Load Factor) Table - 13 Zones
# e.g., square meters per person based on zone type
OLP_ZONES = {
    "retail_sales_basement": 2.8,
    "retail_sales_ground": 2.8,
    "retail_sales_upper": 5.6,
    "office": 10.0,
    "storage": 30.0,
    # Add the remaining 8 zones here...
}

# 3. NBC Area-Band Table (Section 3.1 Property Area rules)
# Enforced strictly top-down in nbc_checker.py
NBC_RULES = [
    {
        "min_area": 10000,
        "requirements": {
            "extinguishers_per_sqm": 50,
            "hose_reels_required": True,
            "sprinkler_system_required": True,
            "fire_alarms_required": True
        }
    },
    {
        "min_area": 5000,
        "requirements": {
            "extinguishers_per_sqm": 100,
            "hose_reels_required": True,
            "sprinkler_system_required": False,
            "fire_alarms_required": True
        }
    },
    {
        "min_area": 2000,
        "requirements": {
            "extinguishers_per_sqm": 200,
            "hose_reels_required": False,
            "sprinkler_system_required": False,
            "fire_alarms_required": True
        }
    },
    {
        "min_area": 500,
        "requirements": {
            "extinguishers_per_sqm": 300,
            "hose_reels_required": False,
            "sprinkler_system_required": False,
            "fire_alarms_required": False
        }
    }
]

# 4. Fixed "Other Notes" Rules
FIXED_NOTES = [
    "Ensure all fire exits are completely unobstructed.",
    "Fire mock drills must be conducted bi-annually.",
    "Electrical panels must have a 1-meter clear radius."
]
