from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment
import settings

def embed_settings_popup(ws, cell_coordinate="AB1", show_popup=True):
    """
    Embeds specific calculation settings into a cell as a clean, hoverable Excel comment.
    If show_popup is False, the function exits without doing anything.
    """
    if not show_popup:
        return 

    config = settings._SETTINGS_CONFIG
    
    clean_text = (
        "--- Run Settings ---\n\n"
        f"Stdev Threshold: {config.get('STDEV_THRESHOLD')}\n"
        f"Outlier Sigma: {config.get('OUTLIER_SIGMA')}\n"
        f"Exclusion Logic: {config.get('OUTLIER_EXCLUSION_MODE')}\n"
        f"Step 3 Selection: {config.get('CALC_MODE_STEP3')}\n"
        f"Step 7 Selection: {config.get('CALC_MODE_STEP7')}"
    )
    
    # Create the Comment object and declare dimensions (in pixels) immediately
    settings_comment = Comment(
        text=clean_text, 
        author="DNT", 
        width=192,  # 2 inches
        height=125  # 1.3 inches
    )
    
    # Target the cell, set the text, and attach the comment
    target_cell = ws[cell_coordinate]
    target_cell.value = "⚙️ Settings"
    target_cell.comment = settings_comment
    
    # Style the cell so it looks distinct (blue, bold, centered)
    target_cell.font = Font(color="0052cc", bold=True)
    target_cell.alignment = Alignment(horizontal="center", vertical="center")

def normalize_name(s):
    if s is None:
        return ''
    return ' '.join(str(s).split()).lower()
