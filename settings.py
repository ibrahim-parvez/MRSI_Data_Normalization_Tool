_SETTINGS_CONFIG = {
    "STDEV_THRESHOLD": 0.08
}

def get_setting(key):
    """
    Returns the current value of a specific setting.
    """
    return _SETTINGS_CONFIG.get(key)

def set_setting(key, value):
    """
    Sets the new value for a specific setting after validation.
    Returns (True, message) on success, or (False, message) on failure.
    """
    if key == "STDEV_THRESHOLD":
        try:
            # Type validation for STDEV_THRESHOLD
            new_value = float(value)
            if new_value <= 0:
                return False, "STDEV_THRESHOLD must be a positive number."
            _SETTINGS_CONFIG[key] = new_value
            return True, f"STDEV_THRESHOLD updated to {new_value}"
        except ValueError:
            return False, "Invalid input: STDEV_THRESHOLD must be a number."
    
    # Fallback for keys that might be added later
    if key in _SETTINGS_CONFIG:
        _SETTINGS_CONFIG[key] = value
        return True, f"Setting {key} updated."
    else:
        return False, f"Setting key '{key}' not recognized."

def get_advanced_settings_options():
    """
    Returns a dictionary of all settings that can be changed via the UI.
    """
    return _SETTINGS_CONFIG.copy()