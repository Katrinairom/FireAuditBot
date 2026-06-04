# src/utils.py

def to_caps(val):
    """Uppercases a string. Returns 'N/A' if the value is empty or None."""
    if val is None or str(val).strip() == "":
        return "N/A"
    return str(val).strip().upper()


def format_qty(val, unit="NOS"):
    """Formats a quantity (e.g., 8 -> '8 NOS'). Returns 'N/A' if missing."""
    if val is None or str(val).strip() == "":
        return "N/A"
    return f"{str(val).strip()} {unit}"


def safe_get(d, key, default=None):
    """Pulls a value from a dictionary safely. Prevents KeyError."""
    if not isinstance(d, dict):
        return default
    return d.get(key, default)
