"""Reference-material colours, shared by the UI and every processor.

Each name the Color dropdown offers maps to one plain RGB hex code, which is
what openpyxl writes and every Excel version reads. Keeping a single table
means the swatch in the app is the colour that lands in the workbook; before,
each processor had its own partial map and unknown names fell back to black.
All shades are dark enough to read on white and on the grey run rows.
"""

MATERIAL_COLORS = {
    "black":     "000000",
    "red":       "FF0000",
    "green":     "008000",
    "darkblue":  "00008B",
    "lightblue": "3399FF",
    "orange":    "E46C0A",
    "purple":    "7030A0",
    "brown":     "8B4513",
    "magenta":   "CC00CC",
    "teal":      "008080",
    "blue":      "0000FF",
    "cyan":      "00A3C4",
}

DEFAULT_MATERIAL_COLOR = "black"

# Order new reference materials are coloured in. Black is left out so a new
# material always stands out; once every colour is taken the cycle repeats.
AUTO_COLOR_ORDER = (
    "red", "darkblue", "green", "orange", "purple", "lightblue",
    "teal", "magenta", "brown", "blue", "cyan",
)


def material_color_hex(color_name):
    """Colour name (or a raw 6-digit hex) -> hex for openpyxl; black if unknown."""
    text = str(color_name or "").strip().lower().lstrip("#")
    if text in MATERIAL_COLORS:
        return MATERIAL_COLORS[text]
    if len(text) == 6 and all(c in "0123456789abcdef" for c in text):
        return text.upper()
    return MATERIAL_COLORS[DEFAULT_MATERIAL_COLOR]


def next_material_color(materials):
    """Colour for a material about to join `materials` (a list of dicts).

    The first colour in AUTO_COLOR_ORDER nobody uses yet; when all are taken,
    keep cycling so neighbours still differ.
    """
    used = {str(m.get("color", "")).strip().lower() for m in materials or []}
    for name in AUTO_COLOR_ORDER:
        if name not in used:
            return name
    return AUTO_COLOR_ORDER[len(materials) % len(AUTO_COLOR_ORDER)]
