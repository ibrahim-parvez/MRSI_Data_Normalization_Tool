import json

_SETTINGS_CONFIG = {
    # New: Toggle for Stdev
    "STDEV_THRESHOLD_ENABLED": False,
    "STDEV_THRESHOLD": 0.08,
    
    # New: Outlier Configuration
    "OUTLIER_SIGMA": 2,                        # Options: 1, 2, 3
    "OUTLIER_EXCLUSION_MODE": "Individual",   # Options: "Exclude Row", "Individual"
    
    # Updated: Split Calculation Modes
    "CALC_MODE_STEP3": "Last 6",               # Options: "Last 6", "Last 6 Outliers Excluded"
    "CALC_MODE_STEP7": "All Values",           # Options: "All Values", "Outliers Excluded"

    # Updated: Materials split by Type
    "REFERENCE_MATERIALS": {
        "Carbonate": [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "-2.37", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "-46.6", "col_f": "", "col_g": "", "col_h": "-26.7", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-23.01", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "-2.20",  "col_h": "", "color": "darkblue"}
        ],
        "Water": [
            {"col_c": "MRSI-STD-W1", "col_d": "", "col_e": "", "col_f": "-3.52", "col_g": "-0.58", "col_h": "", "color": "red"},
            {"col_c": "MRSI-STD-W2",  "col_d": "", "col_e": "", "col_f": "-214.79", "col_g": "-28.08", "col_h": "", "color": "darkblue"},
            {"col_c": "USGS W-67400",  "col_d": "", "col_e": "", "col_f": "1.25", "col_g": "-1.97", "col_h": "", "color": "orange"},
            {"col_c": "USGS W-64444",  "col_d": "", "col_e": "", "col_f": "-399.1", "col_g": "-51.14", "col_h": "", "color": "green"}
        ]
    },

    # Updated: Slope Groups split by Type
    "SLOPE_INTERCEPT_GROUPS": {
        "Carbonate": [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ],
        "Water": [
            ["MRSI-STD-W1", "MRSI-STD-W2"],
            ["USGS W-67400", "USGS W-64444"]
        ]
    }
}

def get_setting(key, sub_key=None):
    """
    Returns the current value. 
    If sub_key is provided (e.g. 'Carbonate'), returns that specific subset.
    """
    val = _SETTINGS_CONFIG.get(key)
    
    # Return deep copies to prevent accidental reference mutation
    if key in ["REFERENCE_MATERIALS", "SLOPE_INTERCEPT_GROUPS"]:
        if sub_key and isinstance(val, dict):
            return [item[:] if isinstance(item, list) else item.copy() for item in val.get(sub_key, [])]
        return val # Return whole dict if no sub_key
    return val

def set_setting(key, value, sub_key=None):
    """
    Sets the new value. 
    If sub_key is provided (e.g. 'Carbonate'), updates only that entry in the dictionary.
    """
    # New: Handle the enable/disable toggle
    if key == "STDEV_THRESHOLD_ENABLED":
        _SETTINGS_CONFIG[key] = bool(value)
        return True, "Updated"

    elif key == "STDEV_THRESHOLD":
        try:
            new_value = float(value)
            if new_value <= 0: return False, "Must be positive."
            _SETTINGS_CONFIG[key] = new_value
            return True, "Updated"
        except ValueError:
            return False, "Invalid number"

    # New: Handle Outlier Sigma (must be 1, 2, or 3)
    elif key == "OUTLIER_SIGMA":
        if value in [1, 2, 3]:
            _SETTINGS_CONFIG[key] = value
            return True, "Updated"
        return False, "Invalid Sigma"
    
    # New: Handle Exclusion Mode
    elif key == "OUTLIER_EXCLUSION_MODE":
        _SETTINGS_CONFIG[key] = value
        return True, "Updated"

    # Handle the two calc modes separately
    elif key in ["CALC_MODE_STEP3", "CALC_MODE_STEP7"]:
        _SETTINGS_CONFIG[key] = value
        return True, "Updated"

    # Handle Dictionary based settings (Materials & Slope Groups)
    elif key in ["REFERENCE_MATERIALS", "SLOPE_INTERCEPT_GROUPS"]:
        if sub_key:
            if key not in _SETTINGS_CONFIG: _SETTINGS_CONFIG[key] = {}
            _SETTINGS_CONFIG[key][sub_key] = value
            return True, f"Updated {sub_key}"
        else:
            _SETTINGS_CONFIG[key] = value
            return True, "Updated all"

    # Fallback
    _SETTINGS_CONFIG[key] = value
    return True, "Updated"

def get_reference_names(material_type="Carbonate"):
    """Helper to get list of names for a specific material type."""
    mats = get_setting("REFERENCE_MATERIALS", sub_key=material_type)
    return [m["col_c"] for m in mats if m.get("col_c")]