import json
import copy
import datetime

_SETTINGS_CONFIG = {
    # UI Toggles
    "ENABLE_FREEZE_PANE": True,
    "CO2_TEMP_MODE": "25 °C",
    # Free-text label the user types for the Custom CO2 profile (e.g. "70 °C").
    # Empty means fall back to the generic "Custom Temp" wording.
    "CO2_CUSTOM_LABEL": "",

    # Normalization blocks: which are drawn, and their left-to-right order.
    # Yield is never first - it has no reference box for the top-left corner.
    "CALC_CARBONATE": True,
    "CALC_YIELD": False,
    "CALC_CO2": False,
    "BLOCK_ORDER": ["Carbonate", "Yield", "CO2"],

    # Toggle for Stdev
    "STDEV_THRESHOLD_ENABLED": True,
    "STDEV_THRESHOLD": 0.1,
    
    # Outlier Configuration
    "OUTLIER_SIGMA": 2,                        
    "OUTLIER_EXCLUSION_MODE": "Individual",   
    
    # Split Calculation Modes
    "CALC_MODE_STEP3": "Last 6",               
    "CALC_MODE_STEP7": "All Values",           

    # Yield: the compound used for reference materials and for samples
    "YIELD_COMPOUNDS": {"ref": "CaCO3", "samp": "MnCO3"},

    # Materials split by Type
    "REFERENCE_MATERIALS": {
        "Carbonate": [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "-2.37", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "-26.7", "col_h": "", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-23.01", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "-2.20",  "col_h": "", "color": "darkblue"}
        ],
        "Water": [
            {"col_c": "MRSI-STD-W1", "col_d": "", "col_e": "", "col_f": "-3.52", "col_g": "-0.58", "col_h": "", "color": "red"},
            {"col_c": "MRSI-STD-W2",  "col_d": "", "col_e": "", "col_f": "-214.79", "col_g": "-28.08", "col_h": "", "color": "darkblue"},
            {"col_c": "USGS W-67400",  "col_d": "", "col_e": "", "col_f": "1.25", "col_g": "-1.97", "col_h": "", "color": "orange"},
            {"col_c": "USGS W-64444",  "col_d": "", "col_e": "", "col_f": "-399.1", "col_g": "-51.14", "col_h": "", "color": "green"}
        ],
        "CO2": [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "7.86", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "", "col_h": "", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-13.00", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "8.03",  "col_h": "", "color": "darkblue"}
        ]
    },

    # Slope Groups split by Type
    "SLOPE_INTERCEPT_GROUPS": {
        "Carbonate": [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ],
        "Water": [
            ["MRSI-STD-W1", "MRSI-STD-W2"],
            ["USGS W-67400", "USGS W-64444"]
        ],
        "Yield": [
            ["NBS 18", "NBS 19"],
            ["Carrara"]
        ],
        "CO2": [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ],
    }
}

# The values above, as shipped. Reset to Default reads them from here, so the
# defaults live in one place and a reset always matches a fresh launch.
_DEFAULTS = copy.deepcopy(_SETTINGS_CONFIG)


def reset_to_defaults(keys):
    """Put each setting in `keys` back to its shipped default."""
    for key in keys:
        if key in _DEFAULTS:
            _SETTINGS_CONFIG[key] = copy.deepcopy(_DEFAULTS[key])


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

    elif key == "OUTLIER_SIGMA":
        if value in [1, 2, 3]:
            _SETTINGS_CONFIG[key] = value
            return True, "Updated"
        return False, "Invalid Sigma"
    
    elif key == "OUTLIER_EXCLUSION_MODE":
        _SETTINGS_CONFIG[key] = value
        return True, "Updated"

    elif key in ["CALC_MODE_STEP3", "CALC_MODE_STEP7"]:
        _SETTINGS_CONFIG[key] = value
        return True, "Updated"

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

# --- AUTO-SAVE & FILE MANAGEMENT LOGIC ---

_AUTO_SAVES = []
_MAX_AUTO_SAVES = 10


def auto_save(note="Auto-save"):
    """Snapshot the current settings, newest first.

    Skipped when nothing changed since the last snapshot, so leaving the
    Advanced tab repeatedly does not fill the list with identical entries.
    """
    if _AUTO_SAVES and _AUTO_SAVES[0]["data"] == _SETTINGS_CONFIG:
        return
    now = datetime.datetime.now()
    _AUTO_SAVES.insert(0, {
        "time": now,
        "timestamp": now.strftime("%Y-%m-%d %I:%M:%S %p"),
        "note": note,
        "data": copy.deepcopy(_SETTINGS_CONFIG),
    })
    del _AUTO_SAVES[_MAX_AUTO_SAVES:]


def get_auto_saves():
    """Snapshots, newest first: dicts with time, timestamp, note and data."""
    return [dict(s, data=copy.deepcopy(s["data"])) for s in _AUTO_SAVES]


def get_current_settings():
    """A copy of every current setting, for comparing against a snapshot."""
    return copy.deepcopy(_SETTINGS_CONFIG)


def restore_auto_save(index):
    """Restore a snapshot. The current settings are snapshotted first, so a
    restore can itself be undone from the same list."""
    if not 0 <= index < len(_AUTO_SAVES):
        return False
    data = copy.deepcopy(_AUTO_SAVES[index]["data"])
    auto_save("Before restoring")
    _SETTINGS_CONFIG.update(data)
    return True

def export_to_file(filepath):
    """Exports the current settings dictionary to a JSON file."""
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(_SETTINGS_CONFIG, f, indent=4)
        return True, "Settings exported successfully."
    except Exception as e:
        return False, f"Export failed: {str(e)}"

def import_from_file(filepath):
    """Imports settings from a JSON file and updates the current configuration."""
    global _SETTINGS_CONFIG
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        if isinstance(data, dict):
            auto_save("Before import")  # so an import can be undone
            _SETTINGS_CONFIG.update(data)
            return True, "Settings imported successfully."
        return False, "Invalid file format. Expected a JSON dictionary."
    except Exception as e:
        return False, f"Import failed: {str(e)}"
