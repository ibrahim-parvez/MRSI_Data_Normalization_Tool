import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment, Color
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter
from datetime import datetime
from copy import copy
import re
import unicodedata
import statistics
import utils.settings as settings
from utils.colors import material_color_hex
from utils.common_utils import embed_settings_popup, save_workbook_atomic
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.views import Selection

# --- Helper Functions ---
def _normalize_text(text):
    if not text:
        return ""
    text = str(text)
    text = unicodedata.normalize("NFKD", text)
    text = re.sub(r"[^A-Za-z0-9]+", "", text)
    return text.lower().strip()

# Yield colour code, shared by the Yield table and the Theoretical column:
# green marks the compound used for reference materials, pink the one used
# for samples. Everything else in the Yield table stays black.
# Display format for the blue box's measured values and slope/intercept cells.
# Three decimals matches the Water processor; the underlying values keep full
# precision, so Excel's increase/decrease-decimal buttons still reveal more.
FMT_3_DEC = "0.000"

YIELD_REF_COLOR = "00B050"   # green
YIELD_SAMP_COLOR = "FF66B2"  # pink


def create_rich_text(parts):
    rt = CellRichText()
    for font, text in parts:
        rt.append(TextBlock(font, text))
    return rt

def get_subscript_label(main_text, sub_text, base_font=None):
    """Creates a CellRichText object with subscripted text."""
    if base_font is None:
        base_font = InlineFont(rFont="Calibri", sz=11, color="000000")
    sub_font = copy(base_font)
    sub_font.vertAlign = "subscript"
    return create_rich_text([
        (base_font, main_text),
        (sub_font, sub_text)
    ])

def get_co2_temp_label():
    """Retrieves the dynamic temperature tag shown beside the CO2 δ¹⁸O header."""
    # Use the exact string saved by the UI
    mode = settings.get_setting("CO2_TEMP_MODE")
    
    if mode == "72 °C":
        return "@ 72 °C"
    elif mode == "25 °C":
        return "@ 25 °C"
    elif mode == "Custom":
        # Whatever the user typed in the Advanced Settings CO2 box, e.g. "70 °C"
        custom = str(settings.get_setting("CO2_CUSTOM_LABEL") or "").strip()
        return f"@ {custom}" if custom else "@ ____"
        
    # Absolute fallback just in case settings fail to load completely
    return "@ 25 °C"

def write_co2_temp_tag(ws, offset_col):
    """Yellow "@ 25 °C" tag in the cell right of the published δ¹⁸O(CO2) header.

    Called again after the blanket box shading so the highlight survives it.
    """
    cell = ws.cell(row=3, column=8 + offset_col, value=get_co2_temp_label())
    cell.font = Font(color="000000", bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def extract_sample_base(identifier):
    if not identifier or not isinstance(identifier, str):
        return ""
    identifier = identifier.strip()
    base = re.sub(r"\s*r\d+(\.\d+)?$", "", identifier, flags=re.IGNORECASE)
    return base.strip()

def _yield_match_key(name):
    """Compact form of a material or sample name for Yield matching.

    Case, spaces, dashes, underscores and a "STD" tag are ignored, so
    "NBS 18", "nbs-18" and "NBS18 std" all compare equal.
    """
    return re.sub(r'[\s\-_]+', '', str(name or "").upper()).replace("STD", "")

def _yield_name_matches(mat, sample_base):
    """True when a sample base name is the Yield reference material `mat`.

    Exact comparison of the compact forms. Substring matching is wrong here:
    a sample called "A" sits inside "CARRARA" and "B" inside "NBS18", which
    pulled whole sample sets into the reference slopes.
    """
    mat_key = _yield_match_key(mat)
    return bool(mat_key) and mat_key == _yield_match_key(sample_base)

def _yield_material_present(mat, present_bases):
    """True when a Yield reference material actually shows up in the run.

    Mirrors the name matching used by the Yield slope/intercept loop so a
    group is drawn only when its slope can really be built from the data.
    """
    return any(_yield_name_matches(mat, key) for key in present_bases)

def filter_yield_groups_by_presence(yield_groups, present_bases):
    """Drop Yield groups whose reference materials were not measured.

    The Yield groups come from Advanced Settings, which students cannot edit,
    so the defaults (NBS 18 + NBS 19, and Carrara) apply to every file. A run
    without Carrara would otherwise get a Carrara column with no slope behind
    it, filling the sheet with zeros. A group survives only when every material
    it names is present.
    """
    return [g for g in yield_groups
            if g and all(_yield_material_present(m, present_bases) for m in g)]

# Sort key for an identifier with no rN suffix: last, and recognisably "none".
UNNUMBERED_RUN = 9999

def extract_run_number(identifier):
    if not identifier or not isinstance(identifier, str):
        return (UNNUMBERED_RUN, 0)
    m = re.search(r"r(\d+)(?:\.(\d+))?", identifier, flags=re.IGNORECASE)
    if m:
        major = int(m.group(1))
        minor = int(m.group(2)) if m.group(2) else 0
        return (major, minor)
    return (UNNUMBERED_RUN, 0)

# --- Splitting a re-measured sample into sessions -------------------------
# Runs are gathered by base name ("Rhodo 74 r2" -> "Rhodo 74"). That is right
# when one sample's runs are spread out and keep counting (r1, r2 ... r3), but
# wrong for a time series that re-measures the same name: "A r1-r3" on DAY 12,
# again on DAY 9, and so on. Each of those is its own sample.
#
# Rule: take one name's runs in measurement order (Time Code) and start a new
# session whenever the run number fails to go up. r1 r2 r3 r1 r2 r3 is two
# sessions; r1 r2 ... r3 is one. Runs with no rN never open a session.
# Identifier 2 cannot do this: it is only the position in the sequence.
# Only samples are split; every run of a standard belongs in its slope.
# step5_group.py imports this too, so Group and Normalization always agree.

_SESSION_RUN_RE = re.compile(r"[rR](\d+)(?:\.(\d+))?\s*[a-zA-Z]*\s*$")
_TIME_CODE_FORMATS = ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M")

def _session_run_number(identifier):
    """'A r2' -> (2, 0), 'HeCO2 r1.3' -> (1, 3), 'HeCO2' -> None."""
    if identifier is None:
        return None
    m = _SESSION_RUN_RE.search(str(identifier).strip())
    return (int(m.group(1)), int(m.group(2) or 0)) if m else None

def _parse_time_code(value):
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in _TIME_CODE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None

def split_into_sessions(items, ident_of, time_of):
    """Split one base name's runs into sessions, earliest first.

    `items` is in sheet order; `ident_of` / `time_of` read Identifier 1 and
    the Time Code from an item. Runs are ordered by Time Code when every run
    has one (some files list the sequence newest first), else by sheet order.
    """
    items = list(items)
    times = [_parse_time_code(time_of(it)) for it in items]
    if items and all(t is not None for t in times):
        order = sorted(range(len(items)), key=lambda i: (times[i], i))
    else:
        order = list(range(len(items)))

    sessions = [[]]
    prev = None
    for i in order:
        run = _session_run_number(ident_of(items[i]))
        if run is not None:
            if prev is not None and run <= prev and sessions[-1]:
                sessions.append([])
            prev = run
        sessions[-1].append(items[i])
    return [s for s in sessions if s]

def _make_fill(hex_color):
    c = hex_color.replace("#", "").upper()
    return PatternFill(start_color=c, end_color=c, fill_type="solid")

def get_excel_range(col_let, rows):
    """Converts a list of row numbers into a valid Excel range string. Uses CHOOSE for non-contiguous."""
    if not rows:
        return ""
    rows = sorted(list(set(rows)))
    
    # Check if rows are perfectly continuous
    is_continuous = all(rows[i] == rows[i-1] + 1 for i in range(1, len(rows)))
    if is_continuous:
        if len(rows) == 1:
            return f"{col_let}{rows[0]}"
        else:
            return f"{col_let}{rows[0]}:{col_let}{rows[-1]}"

    # Workaround for non-continuous arrays compatible with older Excel versions
    indices = "{" + ",".join(str(i+1) for i in range(len(rows))) + "}"
    cells = ",".join(f"{col_let}{r}" for r in rows)
    return f"CHOOSE({indices},{cells})"

# Column M of the grouped layout carries step 1's peak-count flag, e.g.
# "> 11: 18" or "< 11: 5": the run had more or fewer than the expected 11
# peaks, which makes it unreliable.
CO2_PEAK_FLAG_COL = 13
PEAK_FLAG_RE = re.compile(r"^[<>]\s*\d+\s*:\s*\d+$")

def has_peak_count_flag(value):
    """True for a red "> 11: 18" / "< 11: 5" peak-count flag."""
    if value is None:
        return False
    return bool(PEAK_FLAG_RE.match(str(value).strip()))

def _get_valid_co2_rows(rows, col_identifier1):
    """Picks the CO2/HeCO2 runs worth keeping.

    Numbered runs still collapse to one row per rN (lowest sub-run, r1 skipped),
    but runs written without any rN are each their own run: collapsing them
    would silently keep the first and drop every later one.
    """
    seen = {}
    unnumbered = []
    for i, r in enumerate(rows):
        ident_raw = str(r[col_identifier1 - 1] or "").strip()
        ident_lower = ident_raw.lower()
        if ident_lower.startswith("heco2"):
            ident_clean = "co2" + ident_raw[len("heco2"):]
        elif ident_lower.startswith("co2"):
            ident_clean = "co2" + ident_raw[len("co2"):]
        else:
            ident_clean = ident_raw
        major, minor = extract_run_number(ident_clean)
        if major is None or minor is None:
            continue
        if major == UNNUMBERED_RUN:
            unnumbered.append(i)
            continue
        if major == 1:
            continue
        if major not in seen:
            seen[major] = (minor, i)
        else:
            prev_minor, prev_i = seen[major]
            if minor < prev_minor:
                seen[major] = (minor, i)
    
    chosen = sorted([idx for _, idx in seen.values()] + unnumbered)
    
    # The flag is read only once a run has its representative: a bad r2.1 means
    # run 2 is bad, not that r2.2 should stand in for it.
    return [i for i in chosen
            if not has_peak_count_flag(rows[i][CO2_PEAK_FLAG_COL - 1]
                                       if len(rows[i]) >= CO2_PEAK_FLAG_COL else None)]

# --- Normalized output geometry -------------------------------------------
# Each slope group owns a 4-column block: the value column (which doubles as
# Average on the summary rows), Stdev, Count, and a blank spacer so adjacent
# groups stay visually separate. The last group in a section drops the spacer.
OUT_GROUP_STRIDE = 4
OUT_GROUP_WIDTH = 3

# Aragonite only earns its columns when the run actually contains n.arag
# samples. The whole sheet's geometry depends on it, so every caller reads the
# same module-level answer, set once per run by step6_normalization_carbonate.
OUT_INCLUDE_ARAG = True

def out_col(section_start, group_index):
    """First (value / Average) column of a slope group inside a section."""
    return section_start + group_index * OUT_GROUP_STRIDE

def get_output_layout(num_groups, output_offset=0, include_arag=None):
    """Column geometry of one Normalized VPDB + VSMOW block."""
    if include_arag is None:
        include_arag = OUT_INCLUDE_ARAG
    span = max(0, num_groups * OUT_GROUP_STRIDE - 1)
    c_start = 19 + output_offset
    o_calc_start = c_start + span + 1
    o_arag_start = o_calc_start + span + 1 if include_arag else None
    box1_end = (o_arag_start if include_arag else o_calc_start) + span - 1
    box2_start = box1_end + 2
    vsmow_calc_start = box2_start
    vsmow_arag_start = vsmow_calc_start + span + 1 if include_arag else None
    box2_end = (vsmow_arag_start if include_arag else vsmow_calc_start) + span - 1
    return {
        "span": span,
        "include_arag": include_arag,
        "c_start": c_start,
        "o_calc_start": o_calc_start,
        "o_arag_start": o_arag_start,
        "vsmow_calc_start": vsmow_calc_start,
        "vsmow_arag_start": vsmow_arag_start,
        "box1_start": c_start,
        "box1_end": box1_end,
        "box2_start": box2_start,
        "box2_end": box2_end,
    }

def out_section_starts(layout, is_arag, do_c=True, do_o=True):
    """The (C, O, VSMOW) section starts a summary row should touch."""
    starts = []
    if do_c:
        starts.append(layout["c_start"])
    if do_o:
        arag = is_arag and layout["include_arag"]
        starts.append(layout["o_arag_start"] if arag else layout["o_calc_start"])
        starts.append(layout["vsmow_arag_start"] if arag else layout["vsmow_calc_start"])
    return starts

# --- Sheet-wide block layout ----------------------------------------------
# Carbonate, Yield and CO2 are laid out left to right in whatever order the
# user arranged them, and any of them can be switched off. Only the widths and
# the order decide where each one lands, so nothing is pinned to a column.
BLOCK_GAP = 3            # blank columns between two blocks
BLUE_BOX_SPAN = 14       # a reference-material box covers 14 columns
FIRST_OUT_COL = 19       # where the left-most block's output starts
DATA_BLOCKS = ("Carbonate", "CO2")
DEFAULT_BLOCK_ORDER = ("Carbonate", "Yield", "CO2")

def yield_block_end(start_col, num_yield_groups):
    """Right-most column the Yield block uses, table or per-sample columns."""
    return max(start_col + 12, start_col + 1 + num_yield_groups * 4)

def plan_block_layout(order, group_counts, num_yield_groups):
    """Places each enabled block and returns where it goes.

    `order` is the enabled blocks, left to right; `group_counts` maps a data
    block to its slope-group count. A data block gets `out_shift` (its output
    columns) and `top_shift` (its reference box). The left-most block keeps its
    box in the top-left corner, above the copied data, which is where it has
    always been; any block further right carries its box over its own columns.
    """
    plan = {}
    cursor = FIRST_OUT_COL
    first = True
    for name in order:
        if name in DATA_BLOCKS:
            out_shift = cursor - FIRST_OUT_COL
            # A trailing box sits one column right of its own output start:
            # column 3 + top_shift == 19 + out_shift + 1.
            top_shift = 0 if first else out_shift + 17
            end = get_output_layout(group_counts.get(name, 0), out_shift)["box2_end"]
            if not first:
                end = max(end, 3 + top_shift + BLUE_BOX_SPAN - 1)
            plan[name] = {"out_shift": out_shift, "top_shift": top_shift,
                          "start": cursor, "end": end}
        else:
            end = yield_block_end(cursor, num_yield_groups)
            plan[name] = {"start_col": cursor, "start": cursor, "end": end}
        cursor = end + 1 + BLOCK_GAP
        first = False
    return plan

def resolve_block_order(order, enabled):
    """Enabled blocks, left to right, with Yield never left-most.

    Yield has no reference box of its own to put in the top-left corner and
    the report reads the output area from column 19, so the first slot belongs
    to a data block.
    """
    seen = []
    for name in list(order or ()) + list(DEFAULT_BLOCK_ORDER):
        if name in DEFAULT_BLOCK_ORDER and name not in seen:
            seen.append(name)
    live = [name for name in seen if enabled.get(name)]
    if live and live[0] not in DATA_BLOCKS:
        for i, name in enumerate(live):
            if name in DATA_BLOCKS:
                live.insert(0, live.pop(i))
                break
    return live

def _cell_list(col_let, rows):
    """Explicit cell list (A1,A3,A7) so plain AVERAGE/STDEV/COUNT can take it."""
    return ",".join(f"{col_let}{r}" for r in sorted(set(rows)))

def write_summary_output_labels(ws, r, num_groups, is_arag, output_offset, font,
                                do_c=True, do_o=True, outlier_marker=False):
    """Average/Stdev/Count captions over the normalized slope-group blocks.

    With outlier_marker, every block also gets its own "Outlier Excl." tag in
    the spacer to its left, so the row stays readable however far right you
    scroll instead of relying on one marker at the far edge.
    """
    if num_groups <= 0:
        return
    layout = get_output_layout(num_groups, output_offset)
    for start in out_section_starts(layout, is_arag, do_c, do_o):
        for i in range(num_groups):
            base = out_col(start, i)
            for k, txt in enumerate(("Average", "Stdev", "Count")):
                cell = ws.cell(row=r, column=base + k, value=txt)
                cell.font = font
                cell.alignment = Alignment(horizontal="right")
            if outlier_marker:
                marker = ws.cell(row=r, column=base - 1, value="Outlier Excl.")
                marker.font = Font(bold=True)
                marker.alignment = Alignment(horizontal="right")

def strike_output_cells(ws, r, num_groups, is_arag, output_offset, do_c, do_o, color="00FF0000"):
    """Strikes a run's normalized cells wherever its raw value was excluded.

    Normalization is linear (slope*x + intercept, and VSMOW likewise), so it
    leaves every run's distance from the mean in standard deviations untouched:
    re-running the sigma test on the normalized column flags exactly the rows
    the raw column already flagged. color=None keeps each cell's own colour
    (used for peak-count-flagged runs).
    """
    if num_groups <= 0 or not (do_c or do_o):
        return
    layout = get_output_layout(num_groups, output_offset)
    for start in out_section_starts(layout, is_arag, do_c, do_o):
        for i in range(num_groups):
            cell = ws.cell(row=r, column=out_col(start, i))
            struck = copy(cell.font)
            struck.strike = True
            if color:
                struck.color = Color(rgb=color)
            cell.font = struck

def strike_flagged_cells(ws, r, columns):
    """Strikes a peak-count-flagged run's value cells, keeping their colour."""
    for col in columns:
        cell = ws.cell(row=r, column=col)
        struck = copy(cell.font)
        struck.strike = True
        cell.font = struck

def write_summary_output_cells(ws, r, num_groups, is_arag, rows_c, rows_o,
                               output_format, output_offset, font):
    """Stats over the per-run normalized values above, not a renormalized average.

    Normalizing an average would only re-apply the slope/intercept to a number
    that is already an average, which says nothing about the spread; averaging
    the normalized runs themselves gives a real Stdev and Count. Passing None
    for rows_c or rows_o leaves that section untouched.
    """
    if num_groups <= 0:
        return
    layout = get_output_layout(num_groups, output_offset)
    c_start, o_start, vsmow_start = out_section_starts(layout, is_arag)
    for start, rows in ((c_start, rows_c), (o_start, rows_o), (vsmow_start, rows_o)):
        if not rows:
            continue
        for i in range(num_groups):
            base = out_col(start, i)
            rng = _cell_list(get_column_letter(base), rows)
            if not rng:
                continue
            avg = ws.cell(row=r, column=base, value=f'=IFERROR(AVERAGE({rng}),"")')
            avg.number_format = output_format
            sd = ws.cell(row=r, column=base + 1, value=f'=IFERROR(STDEV({rng}),"")')
            sd.number_format = output_format
            # A slope group with no usable slope leaves the column blank; show
            # nothing there rather than a lone 0.
            cnt = ws.cell(row=r, column=base + 2,
                          value=f'=IF(COUNT({rng})=0,"",COUNT({rng}))')
            cnt.number_format = '0'
            for c in (base, base + 1, base + 2):
                ws.cell(row=r, column=c).font = font

def get_summary_num_format(base_name):
    return '0.000'

def draw_blue_box_structure(ws, offset_col=0, setting_key="Carbonate", box_fill=None):
    if not box_fill:
        box_fill = PatternFill(start_color="DAE9F8", end_color="DAE9F8", fill_type="solid")
        
    if offset_col > 0:
        for r in range(1, 25):
            for c in range(3 + offset_col, 16 + offset_col):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()
                cell.fill = PatternFill()

    materials = settings.get_setting("REFERENCE_MATERIALS", sub_key=setting_key)
    if not materials: materials = []
    
    slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key=setting_key)
    if not slope_groups: slope_groups = []
    
    num_materials = len(materials)
    
    data_start_row = 4
    data_end_row = 4 + num_materials - 1 if num_materials > 0 else 4

    thick = Side(border_style="thick", color="000000")
    medium = Side(border_style="medium", color="000000")
    black_bold = Font(color="000000", bold=True)
    center = Alignment(horizontal="center", vertical="center")
    
    def get_style_font(color_name, is_bold):
        return Font(color=material_color_hex(color_name), bold=is_bold)

    today_str = datetime.today().strftime("%Y-%m-%d")
    
    if offset_col == 0:
        ws.cell(row=1, column=1, value=today_str).alignment = Alignment(horizontal="left", vertical="center")

    header_title = f"Normalization ({setting_key})"
    ws.cell(row=1, column=3 + offset_col, value=header_title).font = black_bold
    ws.cell(row=1, column=3 + offset_col).alignment = center
    
    col_start = 3 + offset_col; col_end = 8 + offset_col
    ws.merge_cells(start_row=2, start_column=3 + offset_col, end_row=3, end_column=3 + offset_col)
    ws.cell(row=2, column=3 + offset_col, value="Reference Materials").font = black_bold
    ws.cell(row=2, column=3 + offset_col).alignment = center
    ws.merge_cells(start_row=2, start_column=6 + offset_col, end_row=2, end_column=7 + offset_col)
    ws.cell(row=2, column=6 + offset_col, value="Published (vs. VPDB)").font = black_bold
    ws.cell(row=2, column=6 + offset_col).alignment = center

    # Setup Subscripts for Carbonate vs CO2
    if setting_key == "Carbonate":
        ws.cell(row=3, column=6 + offset_col, value="δ¹³C")
        ws.cell(row=3, column=7 + offset_col).value = get_subscript_label("δ¹⁸O", "Carbonate")
    else:
        ws.cell(row=3, column=6 + offset_col, value="δ¹³C")
        ws.cell(row=3, column=7 + offset_col).value = get_subscript_label("δ¹⁸O", "CO2")

    ws.cell(row=3, column=6 + offset_col).alignment = center
    ws.cell(row=3, column=7 + offset_col).alignment = center

    for r in range(2, 4):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = box_fill
            cell.border = Border(top=medium if r==2 else None, bottom=medium if r==3 else None, 
                                 left=medium if c==col_start else None, right=medium if c==col_end else None)

    # CO2 runs are temperature dependent, so tag the published δ¹⁸O column with the
    # acid/equilibration temperature instead of burying it in the box title.
    if setting_key != "Carbonate":
        write_co2_temp_tag(ws, offset_col)

    active_slope_materials = []
    for group in slope_groups:
        for item in group:
            active_slope_materials.append(str(item).strip().lower())

    for idx, mat in enumerate(materials):
        r = data_start_row + idx
        mat_name = str(mat.get("col_c", "")).strip().lower()
        mat_clean = re.sub(r'[\s\-_]+', '', mat_name.upper())
        mat_no_std = mat_clean.replace("STD", "")
        
        is_used_in_slope = False
        for active_mat in active_slope_materials:
            active_clean = re.sub(r'[\s\-_]+', '', active_mat.upper())
            active_no_std = active_clean.replace("STD", "")
            if active_clean in mat_clean or (len(active_no_std) >= 4 and active_no_std in mat_no_std):
                is_used_in_slope = True
                break
                
        c_val = mat.get("col_c")
        d_val = mat.get("col_d")
        
        e_raw = mat.get("col_e")
        f_raw = mat.get("col_f")
        g_raw = mat.get("col_g")
        h_raw = mat.get("col_h")
        
        d13c_val = f_raw if (f_raw is not None and str(f_raw).strip() != "") else e_raw
        d18o_val = g_raw if (g_raw is not None and str(g_raw).strip() != "") else h_raw
        
        if is_used_in_slope:
            e_val = None
            f_val = d13c_val
            g_val = d18o_val
            h_val = None
        else:
            e_val = d13c_val
            f_val = None
            g_val = None
            h_val = d18o_val

        vals = [c_val, d_val, e_val, f_val, g_val, h_val]
        font_style = get_style_font(mat.get("color", "black"), mat.get("bold", False))
        
        for i, val in enumerate(vals):
            c = col_start + i
            if i in [2, 3, 4, 5]: 
                try: val = float(val)
                except: pass
                
            target_cell = ws.cell(row=r, column=c, value=val)
            target_cell.font = font_style
            target_cell.alignment = center
            
            if i in [2, 3, 4, 5] and isinstance(val, (int, float)):
                target_cell.number_format = '0.00'
            
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = box_fill
            cell.border = Border(bottom=medium if r==data_end_row else None, 
                                 left=medium if c==col_start else None, right=medium if c==col_end else None)

    ws.merge_cells(start_row=2, start_column=10 + offset_col, end_row=2, end_column=14 + offset_col)
    ws.cell(row=2, column=10 + offset_col, value="Measured (vs. Working Standard)").font = black_bold
    ws.cell(row=2, column=10 + offset_col).alignment = center

    if setting_key == "Carbonate":
        ws.cell(row=3, column=11 + offset_col, value="δ¹³C")
        ws.cell(row=3, column=14 + offset_col).value = get_subscript_label("δ¹⁸O", "Carbonate")
    else:
        ws.cell(row=3, column=11 + offset_col, value="δ¹³C")
        ws.cell(row=3, column=14 + offset_col).value = get_subscript_label("δ¹⁸O", "CO2")

    ws.cell(row=3, column=11 + offset_col).alignment = center
    ws.cell(row=3, column=14 + offset_col).alignment = center
    
    for r in range(2, data_end_row + 1):
        for c in range(10 + offset_col, 15 + offset_col):
            cell = ws.cell(row=r, column=c)
            cell.fill = box_fill
            cell.alignment = center
            top = thick if r == 2 else None
            bottom = thick if r == data_end_row else None
            left = thick if c == 10 + offset_col else None
            right = thick if c == 14 + offset_col else None
            if r == 3: bottom = thick
            if r == 4: top = None
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    red_if = InlineFont(color='00FF0000', b=True)
    blue_if = InlineFont(color='000000FF', b=True)
    green_if = InlineFont(color='008000', b=True)
    black_if = InlineFont(color='000000', b=True)

    def get_group_rich_text(group_list):
        parts = []
        for item in group_list:
            clean = str(item).strip().upper()
            font = black_if
            if "18" in clean: font = red_if
            elif "19" in clean: font = blue_if
            elif "603" in clean: font = green_if
            parts.append((font, clean.replace("NBS","").replace("IAEA","").strip() + " "))
        return create_rich_text(parts)

    current_row = data_end_row + 2 
    slope_info = []

    for group_list in slope_groups:
        ws.cell(row=current_row, column=10 + offset_col, value="Slope").font = black_bold
        ws.cell(row=current_row + 1, column=10 + offset_col, value="Intercept").font = black_bold
        
        ws.cell(row=current_row, column=9 + offset_col).value = get_group_rich_text(group_list)
        ws.cell(row=current_row, column=9 + offset_col).alignment = center
        
        for r_off in [0, 1]:
            ws.cell(row=current_row+r_off, column=11 + offset_col).alignment = center
            ws.cell(row=current_row+r_off, column=14 + offset_col).alignment = center

        slope_info.append({"slope_row": current_row, "intercept_row": current_row + 1})
        current_row += 3 

    mat_row_map = {}
    for idx, mat in enumerate(materials):
        if mat.get("col_c"): mat_row_map[str(mat.get("col_c")).strip().lower()] = data_start_row + idx

    return current_row - 1, current_row + 3, slope_info, mat_row_map

def populate_blue_box_math(ws, slope_info, mat_row_map, offset_col=0, setting_key="Carbonate"):
    slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key=setting_key)
    if not slope_groups: slope_groups = []
    
    calc_mode = settings.get_setting("CALC_MODE_STEP7")
    use_outlier_excluded = (calc_mode == "Outliers Excluded")

    materials = settings.get_setting("REFERENCE_MATERIALS", sub_key=setting_key)
    if not materials: materials = []

    def get_style_font(color_name, is_bold):
        return Font(color=material_color_hex(color_name), bold=is_bold)
    
    identifier_col = 3
    c_avg_col_ref = 11 
    o_avg_col_ref = 14 
    
    found_map = {} 
    
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=c_avg_col_ref).value
        if val and str(val).strip().lower() == "average":
            label_val = str(ws.cell(row=r, column=10).value or "").strip()
            is_excluded_block = ("outlier" in label_val.lower())
            
            if use_outlier_excluded and not is_excluded_block: continue
            if not use_outlier_excluded and is_excluded_block: continue
                
            avg_row = r + 1 
            id_row = r - 1
            ident = ""
            for t in range(id_row, max(1, id_row - 25), -1):
                cellv = ws.cell(row=t, column=identifier_col).value
                if cellv:
                    ident = str(cellv).strip().lower()
                    break
            
            ident_clean = re.sub(r'[\s\-_]+', '', ident.upper())
            ident_no_std = ident_clean.replace("STD", "")
            
            for mat_name, target_row in mat_row_map.items():
                orig_mat = next((m for m in materials if str(m.get("col_c","")).strip().lower() == mat_name), None)
                if not orig_mat: continue
                
                std_clean = re.sub(r'[\s\-_]+', '', str(orig_mat.get("col_c", "")).upper())
                std_no_std = std_clean.replace("STD", "")
                
                is_match = False
                if std_clean in ident_clean:
                    is_match = True
                elif len(std_no_std) >= 4 and std_no_std in ident_no_std:
                    is_match = True
                    
                if is_match:
                    color = orig_mat.get("color", "black")
                    bold = orig_mat.get("bold", False)
                    font_style = get_style_font(color, bold)

                    c_cell = ws.cell(row=target_row, column=11 + offset_col, value=f'=IFERROR({get_column_letter(c_avg_col_ref)}{avg_row},"")')
                    c_cell.font = font_style
                    c_cell.number_format = FMT_3_DEC
                    
                    o_cell = ws.cell(row=target_row, column=14 + offset_col, value=f'=IFERROR({get_column_letter(o_avg_col_ref)}{avg_row},"")')
                    o_cell.font = font_style
                    o_cell.number_format = FMT_3_DEC
                    
                    found_map[mat_name] = target_row
                    break

    for idx, group_list in enumerate(slope_groups):
        if idx >= len(slope_info): break 
        
        current_slope_row = slope_info[idx]["slope_row"]
        
        rows = []
        for n in group_list:
            key = str(n).strip().lower()
            if key in mat_row_map and key in found_map:
                rows.append(mat_row_map[key])
        
        if len(rows) >= 2:
            range_y_pub = get_excel_range(get_column_letter(6 + offset_col), rows)
            range_x_meas = get_excel_range(get_column_letter(11 + offset_col), rows)
            
            range_o_pub = get_excel_range(get_column_letter(7 + offset_col), rows)
            range_o_meas = get_excel_range(get_column_letter(14 + offset_col), rows)
            
            # Carbonate/CO2 Array Formulas (Ctrl+Shift+Enter automation)
            col_11_let = get_column_letter(11 + offset_col)
            col_14_let = get_column_letter(14 + offset_col)

            for cell, formula in (
                (ws.cell(row=current_slope_row, column=11 + offset_col),
                 ArrayFormula(f"{col_11_let}{current_slope_row}", f'=SLOPE({range_y_pub},{range_x_meas})')),
                (ws.cell(row=current_slope_row+1, column=11 + offset_col),
                 ArrayFormula(f"{col_11_let}{current_slope_row+1}", f'=INTERCEPT({range_y_pub},{range_x_meas})')),
                (ws.cell(row=current_slope_row, column=14 + offset_col),
                 ArrayFormula(f"{col_14_let}{current_slope_row}", f'=SLOPE({range_o_pub},{range_o_meas})')),
                (ws.cell(row=current_slope_row+1, column=14 + offset_col),
                 ArrayFormula(f"{col_14_let}{current_slope_row+1}", f'=INTERCEPT({range_o_pub},{range_o_meas})')),
            ):
                cell.value = formula
                cell.number_format = FMT_3_DEC

def yield_table_bottom_row(num_yield_groups):
    """Last row the top Yield table occupies."""
    return 19 + (num_yield_groups * 3)

def draw_yield_table(ws, start_col, box_fill, num_yield_groups):
    max_yield_row = yield_table_bottom_row(num_yield_groups)
    for r in range(1, max_yield_row + 1):
        for c in range(start_col, start_col + 13):
            ws.cell(row=r, column=c).fill = box_fill

    black_bold = Font(color="000000", bold=True)
    thick = Side(border_style="medium", color="000000") 
    
    ws.cell(row=1, column=start_col, value="By Sang-Tae Kim (2026-0726)").font = black_bold
    ws.cell(row=2, column=start_col + 4, value="Yield Calculation (Sample weight vs. Sum area all)").font = black_bold
    
    ah_col = start_col + 1
    ai_col = start_col + 2
    
    col_AI = get_column_letter(start_col + 2)
    col_AL = get_column_letter(start_col + 5)
    col_AN = get_column_letter(start_col + 7)
    col_AP = get_column_letter(start_col + 9)
    col_AQ = get_column_letter(start_col + 10)
    
    ws.cell(row=2, column=ah_col, value="Atomic weight").font = black_bold
    ws.cell(row=2, column=ah_col).alignment = Alignment(horizontal="center")
    
    atomic_data = [
        ("Mn", 54.93805),
        ("Fe", 55.845),
        ("Li", 6.941),
        ("Ca", 40.078),
        ("Mg", 24.3050),
        ("C", 12.0107),
        ("O", 15.9994)
    ]
    
    for i, (elem, weight) in enumerate(atomic_data, start=3):
        ws.cell(row=i, column=ah_col, value=elem)
        cell_wt = ws.cell(row=i, column=ai_col, value=weight)
        cell_wt.number_format = '0.00E+00'
        
    ws.cell(row=10, column=ah_col, value="CO3")
    ws.cell(row=10, column=ai_col, value=f"={col_AI}8+({col_AI}9*3)").number_format = '0.00E+00'
    ws.cell(row=11, column=ah_col, value="CO2")
    ws.cell(row=11, column=ai_col, value=f"={col_AI}8+({col_AI}9*2)").number_format = '0.00E+00'
    
    ws.cell(row=12, column=ai_col, value="Mol. Weight").font = black_bold
    ws.cell(row=12, column=start_col + 3, value="CO2 weight %").font = black_bold
    ws.cell(row=12, column=start_col + 5, value="Yields (umol/mg)").font = black_bold
    ws.cell(row=12, column=start_col + 7, value="Yields (mmol/mg)").font = black_bold
    
    ws.cell(row=11, column=start_col + 9, value="Target").font = black_bold
    ws.cell(row=11, column=start_col + 9).alignment = Alignment(horizontal="center")
    ws.cell(row=12, column=start_col + 9, value="mg").font = black_bold
    ws.cell(row=12, column=start_col + 9).alignment = Alignment(horizontal="center")
    ws.cell(row=12, column=start_col + 10, value="mmol").font = black_bold
    ws.cell(row=12, column=start_col + 10).alignment = Alignment(horizontal="center")
    ws.cell(row=12, column=start_col + 11, value="umol").font = black_bold
    ws.cell(row=12, column=start_col + 11).alignment = Alignment(horizontal="center")
    
    # --- Fetch Dynamic Settings for Yield ---
    yield_compounds = settings.get_setting("YIELD_COMPOUNDS") or {"ref": "CaCO3", "samp": "MnCO3"}
    ref_comp = yield_compounds.get("ref", "CaCO3")
    samp_comp = yield_compounds.get("samp", "MnCO3")
    
    def get_fonts(comp_name):
        """Green for the reference-material compound, pink for the sample
        compound, black for every other compound in the table."""
        if comp_name == ref_comp:
            color = YIELD_REF_COLOR
        elif comp_name == samp_comp:
            color = YIELD_SAMP_COLOR
        else:
            color = "000000"
        return Font(color=color, bold=True), Font(color=color)

    f_name_ca, f_val_ca = get_fonts("CaCO3")
    f_name_mg, f_val_mg = get_fonts("MgCO3")
    f_name_camg, f_val_camg = get_fonts("CaMg(CO3)2")
    f_name_li, f_val_li = get_fonts("Li2CO3")
    f_name_mn, f_val_mn = get_fonts("MnCO3")
    f_name_fe, f_val_fe = get_fonts("FeCO3")
    
    compounds = [
        (13, "CaCO3", f_name_ca, f_val_ca, f"={col_AI}6+{col_AI}10", f"=(${col_AI}$11/{col_AI}13)*100", f"=1000/{col_AI}13", f"={col_AL}13/1000", 0.150),
        (14, "MgCO3", f_name_mg, f_val_mg, f"={col_AI}7+{col_AI}10", f"=(${col_AI}$11/{col_AI}14)*100", f"=1000/{col_AI}14", f"={col_AL}14/1000", 0.126),
        (15, "CaMg(CO3)2", f_name_camg, f_val_camg, f"={col_AI}6+{col_AI}7+(2*{col_AI}10)", f"=(${col_AI}$11)*2/{col_AI}15*100", f"=(1000/{col_AI}15)*2", f"={col_AL}15/1000", 0.138),
        (16, "Li2CO3", f_name_li, f_val_li, f"=({col_AI}5*2)+{col_AI}10", f"=(${col_AI}$11/{col_AI}16)*100", f"=1000/{col_AI}16", f"={col_AL}16/1000", None),
        (17, "MnCO3", f_name_mn, f_val_mn, f"={col_AI}3+{col_AI}10", f"=(${col_AI}$11/{col_AI}17)*100", f"=1000/{col_AI}17", f"={col_AL}17/1000", 0.172),
        (18, "FeCO3", f_name_fe, f_val_fe, f"={col_AI}4+{col_AI}10", f"=(${col_AI}$11/{col_AI}18)*100", f"=1000/{col_AI}18", f"={col_AL}18/1000", 0.174),
    ]
    
    sci_format = '0.00E+00'
    for r, name, name_font, val_font, mol_wt, co2_wt, y_umol, y_mmol, target_mg in compounds:
        ws.cell(row=r, column=ah_col, value=name).font = name_font
        ws.cell(row=r, column=ai_col, value=mol_wt).font = val_font; ws.cell(row=r, column=ai_col).number_format = sci_format
        ws.cell(row=r, column=start_col + 3, value=co2_wt).font = val_font; ws.cell(row=r, column=start_col + 3).number_format = sci_format
        ws.cell(row=r, column=start_col + 5, value=y_umol).font = val_font; ws.cell(row=r, column=start_col + 5).number_format = sci_format
        ws.cell(row=r, column=start_col + 7, value=y_mmol).font = val_font; ws.cell(row=r, column=start_col + 7).number_format = sci_format
        
        if target_mg is not None:
            ws.cell(row=r, column=start_col + 9, value=target_mg).number_format = sci_format
            ws.cell(row=r, column=start_col + 10, value=f"={col_AP}{r}*{col_AN}{r}").number_format = sci_format
            ws.cell(row=r, column=start_col + 11, value=f"={col_AQ}{r}*1000").number_format = sci_format
            
    for r in range(2, 12):
        for c in range(ah_col, ai_col + 1):
            border = Border(
                top=thick if r == 2 else None,
                bottom=thick if r == 11 or r == 2 else None,
                left=thick if c == ah_col else None,
                right=thick if c == ai_col else None
            )
            ws.cell(row=r, column=c).border = border
            
    for r in range(12, 19):
        for c in range(ah_col, start_col + 8):
            border = Border(
                top=thick if r == 12 else None,
                bottom=thick if r == 18 or r == 12 else None,
                left=thick if c == ah_col else None,
                right=thick if c == start_col + 7 else None
            )
            ws.cell(row=r, column=c).border = border

def draw_upper_boxes(ws, h_row, box_fill, black_bold, green_bold, output_offset=0, setting_key="Carbonate"):
    slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key=setting_key) or []
    num_groups = len(slope_groups)
    if num_groups == 0:
        return
    
    layout = get_output_layout(num_groups, output_offset)
    c_start = layout["c_start"]
    o_calc_start = layout["o_calc_start"]
    o_arag_start = layout["o_arag_start"]
    span = layout["span"]
    
    box1_start = layout["box1_start"]
    box1_end = layout["box1_end"]
    
    box2_start = layout["box2_start"]
    vsmow_calc_start = layout["vsmow_calc_start"]
    vsmow_arag_start = layout["vsmow_arag_start"]
    box2_end = layout["box2_end"]
    include_arag = layout["include_arag"]

    thick = Side(style="thick")
    
    red_font = InlineFont(color='00FF0000', b=True)
    blue_font = InlineFont(color='000000FF', b=True)
    green_font = InlineFont(color='008000', b=True)
    black_font = InlineFont(color='000000', b=True)
    
    def get_rich_text_for_group(group_list):
        parts = []
        for item in group_list:
            clean = str(item).strip().upper()
            font = black_font
            if "18" in clean: font = red_font
            elif "19" in clean: font = blue_font
            elif "603" in clean: font = green_font
            parts.append((font, clean.replace("NBS","").replace("IAEA","").strip() + " "))
        return create_rich_text(parts)
        
    for r in range(h_row - 3, h_row + 1):
        for c in range(box1_start, box1_end + 1):
            cell = ws.cell(r, c)
            cell.fill = box_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            top = thick if r == h_row - 3 else None
            bottom = thick if r == h_row else None 
            left = thick if c == box1_start else None
            right = thick if c == box1_end else None
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    ws.merge_cells(start_row=h_row - 3, start_column=box1_start, end_row=h_row - 3, end_column=box1_end)
    ws.cell(h_row - 3, box1_start, f"Normalized VPDB ({setting_key})").font = black_bold

    if num_groups > 1:
        ws.merge_cells(start_row=h_row - 2, start_column=o_calc_start, end_row=h_row - 2, end_column=o_calc_start + span - 1)
        if include_arag:
            ws.merge_cells(start_row=h_row - 2, start_column=o_arag_start, end_row=h_row - 2, end_column=box1_end)
        
    ws.cell(h_row - 2, o_calc_start, "Calcite").font = black_bold
    if include_arag:
        ws.cell(h_row - 2, o_arag_start, "Aragonite").font = green_bold

    o_label = "Carbonate" if setting_key == "Carbonate" else "CO2"

    for i, grp in enumerate(slope_groups):
        rt = get_rich_text_for_group(grp)
        header_sections = [(c_start, "C"), (o_calc_start, "O")]
        if include_arag:
            header_sections.append((o_arag_start, "O"))
        for start, label in header_sections:
            base = out_col(start, i)
            if label == "C":
                ws.cell(h_row - 1, base, "δ¹³C")
            else:
                ws.cell(h_row - 1, base).value = get_subscript_label("δ¹⁸O", o_label)
            ws.cell(h_row, base).value = rt

    for r in range(h_row - 3, h_row + 1):
        for c in range(box2_start, box2_end + 1):
            cell = ws.cell(r, c)
            cell.fill = box_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            top = thick if r == h_row - 3 else None
            bottom = thick if r == h_row else None
            left = thick if c == box2_start else None
            right = thick if c == box2_end else None
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    ws.merge_cells(start_row=h_row - 3, start_column=box2_start, end_row=h_row - 3, end_column=box2_end)
    ws.cell(h_row - 3, box2_start, "VSMOW").font = black_bold

    if num_groups > 1:
        ws.merge_cells(start_row=h_row - 2, start_column=vsmow_calc_start, end_row=h_row - 2, end_column=vsmow_calc_start + span - 1)
        if include_arag:
            ws.merge_cells(start_row=h_row - 2, start_column=vsmow_arag_start, end_row=h_row - 2, end_column=box2_end)
        
    ws.cell(h_row - 2, vsmow_calc_start, "Calcite").font = black_bold
    if include_arag:
        ws.cell(h_row - 2, vsmow_arag_start, "Aragonite").font = green_bold
    
    for i, grp in enumerate(slope_groups):
        rt = get_rich_text_for_group(grp)
        for start in ((vsmow_calc_start, vsmow_arag_start) if include_arag else (vsmow_calc_start,)):
            base = out_col(start, i)
            ws.cell(h_row - 1, base).value = get_subscript_label("δ¹⁸O", o_label)
            ws.cell(h_row, base).value = rt

def draw_lower_boxes(ws, divider_top_row, box_fill, black_bold, green_bold, output_offset=0, setting_key="Carbonate"):
    slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key=setting_key) or []
    num_groups = len(slope_groups)
    if num_groups == 0:
        return
    
    layout = get_output_layout(num_groups, output_offset)
    c_start = layout["c_start"]
    o_calc_start = layout["o_calc_start"]
    o_arag_start = layout["o_arag_start"]
    span = layout["span"]
    
    box1_start = layout["box1_start"]
    box1_end = layout["box1_end"]
    
    box2_start = layout["box2_start"]
    vsmow_calc_start = layout["vsmow_calc_start"]
    vsmow_arag_start = layout["vsmow_arag_start"]
    box2_end = layout["box2_end"]
    include_arag = layout["include_arag"]

    thick = Side(style="thick")
    
    for r in range(divider_top_row - 3, divider_top_row + 2):
        for c in range(box1_start, box1_end + 1):
            cell = ws.cell(r, c)
            cell.fill = box_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            top = thick if r == divider_top_row - 3 else None
            bottom = thick if r == divider_top_row + 1 else None 
            left = thick if c == box1_start else None
            right = thick if c == box1_end else None
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    ws.cell(divider_top_row - 3, box1_start, f"Normalized ({setting_key})").font = black_bold
    ws.cell(divider_top_row - 2, box1_start, "VPDB").font = black_bold
    
    if num_groups > 1:
        ws.merge_cells(start_row=divider_top_row - 1, start_column=o_calc_start, end_row=divider_top_row - 1, end_column=o_calc_start + span - 1)
        if include_arag:
            ws.merge_cells(start_row=divider_top_row - 1, start_column=o_arag_start, end_row=divider_top_row - 1, end_column=box1_end)
        
    ws.cell(divider_top_row - 1, o_calc_start, "Calcite").font = black_bold
    if include_arag:
        ws.cell(divider_top_row - 1, o_arag_start, "Aragonite").font = green_bold
    
    red_font = InlineFont(color='00FF0000', b=True)
    blue_font = InlineFont(color='000000FF', b=True)
    green_font = InlineFont(color='008000', b=True)
    black_font = InlineFont(color='000000', b=True)
    
    def get_rich_text_for_group(group_list):
        parts = []
        for item in group_list:
            clean = str(item).strip().upper()
            font = black_font
            if "18" in clean: font = red_font
            elif "19" in clean: font = blue_font
            elif "603" in clean: font = green_font
            parts.append((font, clean.replace("NBS","").replace("IAEA","").strip() + " "))
        return create_rich_text(parts)

    o_label = "Carbonate" if setting_key == "Carbonate" else "CO2"
        
    for i, grp in enumerate(slope_groups):
        rt = get_rich_text_for_group(grp)
        header_sections = [(c_start, "C"), (o_calc_start, "O")]
        if include_arag:
            header_sections.append((o_arag_start, "O"))
        for start, label in header_sections:
            base = out_col(start, i)
            if label == "C":
                ws.cell(divider_top_row, base, "δ¹³C")
            else:
                ws.cell(divider_top_row, base).value = get_subscript_label("δ¹⁸O", o_label)
            ws.cell(divider_top_row + 1, base).value = rt

    for r in range(divider_top_row - 2, divider_top_row + 2):
        for c in range(box2_start, box2_end + 1):
            cell = ws.cell(r, c)
            cell.fill = box_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            top = thick if r == divider_top_row - 2 else None
            bottom = thick if r == divider_top_row + 1 else None
            left = thick if c == box2_start else None
            right = thick if c == box2_end else None
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    ws.cell(divider_top_row - 2, box2_start, "VSMOW").font = black_bold
    
    if num_groups > 1:
        ws.merge_cells(start_row=divider_top_row - 1, start_column=vsmow_calc_start, end_row=divider_top_row - 1, end_column=vsmow_calc_start + span - 1)
        if include_arag:
            ws.merge_cells(start_row=divider_top_row - 1, start_column=vsmow_arag_start, end_row=divider_top_row - 1, end_column=box2_end)
        
    ws.cell(divider_top_row - 1, vsmow_calc_start, "Calcite").font = black_bold
    if include_arag:
        ws.cell(divider_top_row - 1, vsmow_arag_start, "Aragonite").font = green_bold
    
    for i, grp in enumerate(slope_groups):
        rt = get_rich_text_for_group(grp)
        for start in ((vsmow_calc_start, vsmow_arag_start) if include_arag else (vsmow_calc_start,)):
            base = out_col(start, i)
            ws.cell(divider_top_row, base).value = get_subscript_label("δ¹⁸O", o_label)
            ws.cell(divider_top_row + 1, base).value = rt

def _detect_decimal_places_from_format(fmt: str):
    if not fmt or not isinstance(fmt, str):
        return None
    first_section = fmt.split(';', 1)[0]
    if '.' not in first_section:
        return None
    after_dot = first_section.split('.', 1)[1]
    count = 0
    for ch in after_dot:
        if ch in ('0', '#'):
            count += 1
        else:
            break
    return count if count > 0 else None


# --- Main Function ---
def step6_normalization_carbonate(file_path):
    
    yield_slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key="Yield") or []
    num_yield_groups = len(yield_slope_groups)
    
    EXTRA_VERT_SHIFT = max(0, num_yield_groups - 2) * 3
    
    # --- Fetch UI settings toggles ---
    calc_yield = settings.get_setting("CALC_YIELD")
    if calc_yield is None: calc_yield = True  # Default fallback
    
    calc_co2 = settings.get_setting("CALC_CO2")
    if calc_co2 is None: calc_co2 = True      # Default fallback
    
    calc_carbonate = settings.get_setting("CALC_CARBONATE")
    if calc_carbonate is None: calc_carbonate = True
    
    # Which blocks to draw, left to right. Everything downstream reads this
    # rather than assuming Carbonate is present or sits first.
    block_enabled = {
        "Carbonate": bool(calc_carbonate),
        "Yield": bool(calc_yield),
        "CO2": bool(calc_co2),
    }
    if not (block_enabled["Carbonate"] or block_enabled["CO2"]):
        raise ValueError(
            "Nothing to normalize: switch on the Carbonate or the CO2 table before running Step 6."
        )
    live_order = resolve_block_order(settings.get_setting("BLOCK_ORDER"), block_enabled)
    calc_yield = "Yield" in live_order

    HEADER_ROW_SHIFT = 12 

    # --- Fetch Yield Compounds and Map to Excel Rows ---
    yield_compounds = settings.get_setting("YIELD_COMPOUNDS") or {"ref": "CaCO3", "samp": "MnCO3"}
    compound_row_map = {
        "CaCO3": 13, 
        "MgCO3": 14, 
        "CaMg(CO3)2": 15, 
        "Li2CO3": 16, 
        "MnCO3": 17, 
        "FeCO3": 18
    }
    ref_compound_row = compound_row_map.get(yield_compounds.get("ref", "CaCO3"), 13)
    samp_compound_row = compound_row_map.get(yield_compounds.get("samp", "MnCO3"), 17)

    stdev_is_enabled = settings.get_setting("STDEV_THRESHOLD_ENABLED")
    
    if stdev_is_enabled:
        stdev_threshold = settings.get_setting("STDEV_THRESHOLD")
    else:
        stdev_threshold = None
    outlier_sigma = settings.get_setting("OUTLIER_SIGMA") or 2
    exclusion_mode = settings.get_setting("OUTLIER_EXCLUSION_MODE") or "Individual"

    strike_font = Font(strike=True, color="FF0000")

    EXCLUDED_COLS = {8, 9, 11, 12, 13, 26, 27, 14, 15}
    MAX_SOURCE_COL = 24
    
    col_map = {}
    dest_col = 1
    for src_col in range(1, MAX_SOURCE_COL + 1):
        if src_col not in EXCLUDED_COLS:
            col_map[src_col] = dest_col
            dest_col += 1
    
    col_identifier1 = col_map.get(3, 0)
    if col_identifier1 != 3:
        raise Exception("Identifier column (3) was moved! Logic error in col_map.")
    
    wb = openpyxl.load_workbook(file_path, data_only=False) 
    if "Last 6_DNT" not in wb.sheetnames:
        raise ValueError("Sheet 'Last 6_DNT' not found!")
    ws_last6 = wb["Last 6_DNT"]

    # Yield groups are global defaults, so keep only the ones this run can
    # actually support. Layout metrics that depend on the count follow.
    present_yield_bases = set()
    arag_re = re.compile(r"\bn\.?\s*arag\b", flags=re.IGNORECASE)
    has_arag = False
    for row_cells in ws_last6.iter_rows(min_row=2, min_col=col_identifier1, max_col=col_identifier1):
        ident = row_cells[0].value
        base = _normalize_text(extract_sample_base(ident))
        if base:
            present_yield_bases.add(base)
        if ident and arag_re.search(str(ident)):
            has_arag = True

    yield_slope_groups = filter_yield_groups_by_presence(yield_slope_groups, present_yield_bases)
    num_yield_groups = len(yield_slope_groups)
    EXTRA_VERT_SHIFT = max(0, num_yield_groups - 2) * 3

    # No aragonite in the run means no Aragonite columns anywhere, which makes
    # the output narrower, so the layout has to settle before anything is drawn.
    global OUT_INCLUDE_ARAG
    OUT_INCLUDE_ARAG = has_arag

    # --- SHEET LAYOUT ---
    # Each block is measured and placed in turn, so the order the user chose
    # and the group counts alone decide the columns.
    block_group_counts = {
        name: len(settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key=name) or [])
        for name in DATA_BLOCKS
    }
    layout_plan = plan_block_layout(live_order, block_group_counts, num_yield_groups)
    
    YIELD_START_COL = (layout_plan["Yield"]["start_col"] if calc_yield
                       else max(p["end"] for p in layout_plan.values()) + 1 + BLOCK_GAP)
    LAST_LAYOUT_COL = max(p["end"] for p in layout_plan.values())
    # ------------------------------

    if "Group_DNT" not in wb.sheetnames:
        raise ValueError("Sheet 'Group_DNT' not found. Run Step 5 first.")

    if "Normalization_DNT" in wb.sheetnames:
        wb.remove(wb["Normalization_DNT"])
    pre_group_index = wb.sheetnames.index("Group_DNT")
    ws_group = wb.create_sheet("Normalization_DNT", pre_group_index)

    for s in wb.worksheets:
        try:
            s.sheet_view.tabSelected = False
        except Exception:
            pass
        
    # --- DISTINCT FILL COLORS ---
    carb_fill = _make_fill("DAE9F8")   # Classic Blue
    co2_fill = _make_fill("E2EFDA")    # Soft Mint Green
    yield_fill = _make_fill("FFE0BF")  # Orange White
    
    dark_fill = _make_fill("808080")
    gray_fill = _make_fill("E7E7E7")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    # Identifier colours for the run rows come from the reference tables, so a
    # material added in the UI is coloured here too. Carbonate wins over CO2
    # when both list the same name.
    color_fonts = {}
    for _key in ("CO2", "Carbonate"):
        for _mat in settings.get_setting("REFERENCE_MATERIALS", sub_key=_key) or []:
            _name = _normalize_text(_mat.get("col_c", ""))
            if _name and str(_mat.get("color", "black")).lower() != "black":
                color_fonts[_name] = Font(color=material_color_hex(_mat.get("color")))
    
    MAX_GROUP_COL = dest_col - 1
    
    # TOP BLOCKS: one reference box per enabled data block, in the user's order
    block_fills = {"Carbonate": carb_fill, "CO2": co2_fill}
    data_blocks = []
    for name in live_order:
        if name not in DATA_BLOCKS:
            continue
        placed = layout_plan[name]
        fill = block_fills[name]
        bottom, _, slope_info, mat_map = draw_blue_box_structure(
            ws_group, placed["top_shift"], name, fill)
        data_blocks.append({
            "key": name,
            "fill": fill,
            "top_shift": placed["top_shift"],
            "out_shift": placed["out_shift"],
            "slope_info": slope_info,
            "mat_map": mat_map,
            "bottom": bottom,
        })
    
    blue_box_bottom = max(b["bottom"] for b in data_blocks)
    
    # ---- INJECT YIELD CALCULATIONS TABLE INTO GAP DYNAMICALLY ----
    if calc_yield:
        draw_yield_table(ws_group, start_col=YIELD_START_COL, box_fill=yield_fill, num_yield_groups=num_yield_groups)
    
    header_row = blue_box_bottom + 4 + HEADER_ROW_SHIFT + EXTRA_VERT_SHIFT
    
    # The Yield table's height comes from its own group count, which has
    # nothing to do with the blue boxes'. Few slope groups and many Yield
    # groups leaves the blue boxes ending well above the table, and the header
    # row below them lands on top of it. Clear the table by the same one row
    # the blue boxes get.
    if calc_yield:
        header_row = max(header_row, yield_table_bottom_row(num_yield_groups) + 5)
    
    for block in data_blocks:
        draw_upper_boxes(ws_group, header_row, block["fill"], Font(bold=True, color="000000"),
                         Font(bold=True, color="008000"), block["out_shift"], block["key"])
    
    # The left-most block owns the top-left corner, above the copied data; the
    # rest shade the columns their own box sits over.
    for block in data_blocks:
        if block["top_shift"] == 0:
            min_col, max_col_fill = 1, MAX_GROUP_COL
        else:
            min_col, max_col_fill = 1 + block["top_shift"], 16 + block["top_shift"]
        for row in ws_group.iter_rows(min_row=1, max_row=blue_box_bottom,
                                      min_col=min_col, max_col=max_col_fill):
            for cell in row:
                cell.fill = block["fill"]
        # Re-applied after the shading, which would otherwise bury it.
        if block["key"] == "CO2":
            write_co2_temp_tag(ws_group, block["top_shift"])
            
    headers = []
    first_row_cells = list(ws_last6[1]) if ws_last6.max_row >= 1 else []
    
    for src_col, dest_col_idx in col_map.items():
        src_cell = first_row_cells[src_col - 1] if src_col <= len(first_row_cells) else None
        header_val = src_cell.value if src_cell else None
        headers.append(header_val)
        dest_cell = ws_group.cell(row=header_row, column=dest_col_idx, value=header_val)
        if src_cell:
            dest_cell.number_format = copy(src_cell.number_format)
            dest_cell.font = copy(src_cell.font)
            dest_cell.alignment = copy(src_cell.alignment)
            dest_cell.border = copy(src_cell.border)
            dest_cell.fill = copy(src_cell.fill)

    # --- INLINE HELPER FOR YIELD HEADERS ---
    red_if = InlineFont(color='00FF0000', b=True)
    blue_if = InlineFont(color='000000FF', b=True)
    green_if = InlineFont(color='008000', b=True)
    black_if = InlineFont(color='000000', b=True)
    
    def get_local_yield_rich_text(group_list):
        parts = []
        for item in group_list:
            clean = str(item).strip().upper()
            font = black_if
            if "18" in clean: font = red_if
            elif "19" in clean: font = blue_if
            elif "603" in clean: font = green_if
            parts.append((font, clean.replace("NBS","").replace("IAEA","").strip() + " "))
        return create_rich_text(parts)

    # --- THEORETICAL & YIELD HEADERS CREATION (MULTI-TIERED) ---
    theo_col = YIELD_START_COL + 1
    
    # Right-most column anything is drawn in. Row-wide shading stops here, so
    # it covers every block without running on to the edge of the sheet.
    last_layout_col = max(MAX_GROUP_COL, LAST_LAYOUT_COL)
    
    if calc_yield:
        if num_yield_groups > 0:
            yield_calc_end = theo_col + (num_yield_groups * 4) 
            thick = Side(style="thick")
            
            # 1. Draw the main box including Theoretical (increased height by 1)
            for r in range(header_row - 3, header_row + 1):
                for c in range(theo_col, yield_calc_end + 1):
                    cell = ws_group.cell(row=r, column=c)
                    cell.fill = yield_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    # Outer thick borders for the whole box
                    top = thick if r == header_row - 3 else None
                    bottom = thick if r == header_row else None
                    left = thick if c == theo_col else None
                    right = thick if c == yield_calc_end else None
                    cell.border = Border(top=top, bottom=bottom, left=left, right=right)

            # 2. "Normalized Yield" Title (merged horizontally)
            ws_group.merge_cells(start_row=header_row - 3, start_column=theo_col, end_row=header_row - 3, end_column=yield_calc_end)
            ny_cell = ws_group.cell(row=header_row - 3, column=theo_col, value="Normalized Yield")
            ny_cell.font = Font(bold=True)
            
            # 3. "Theoretical" Header, on the same row as the other headers
            theo_header = ws_group.cell(row=header_row - 1, column=theo_col, value="Theoretical")
            theo_header.font = Font(bold=True, color="000000")
            
            # 4. Populate each Yield Group
            for i in range(num_yield_groups):
                calc_col = theo_col + 2 + (i * 4)
                yield_pct_col = calc_col + 2
                
                # "Calculated" and "Yield (%)" headers
                c_head = ws_group.cell(row=header_row - 1, column=calc_col, value="Calculated")
                c_head.font = Font(bold=True)
                
                y_head = ws_group.cell(row=header_row - 1, column=yield_pct_col, value="Yield (%)")
                y_head.font = Font(bold=True)
                
                # Reference materials underneath, written once under
                # "Calculated" and once under "Yield (%)". Not merged across
                # the three columns, because merges break the formatting when
                # this sheet is copied and pasted downstream.
                rt = get_local_yield_rich_text(yield_slope_groups[i])
                ws_group.cell(row=header_row, column=calc_col).value = rt
                ws_group.cell(row=header_row, column=yield_pct_col).value = get_local_yield_rich_text(yield_slope_groups[i])
                
                # Ensure column widths are spacious
                ws_group.column_dimensions[get_column_letter(calc_col)].width = 15
                ws_group.column_dimensions[get_column_letter(yield_pct_col)].width = 15
        else:
            # Fallback if no yield groups are configured
            theo_header = ws_group.cell(row=header_row - 1, column=theo_col, value="Theoretical")
            theo_header.fill = yield_fill
            theo_header.font = Font(bold=True, color="000000")
            theo_header.alignment = Alignment(horizontal="center", vertical="center")

    data_rows_with_index = []
    for row_idx, row_cells in enumerate(ws_last6.iter_rows(min_row=2, max_col=MAX_SOURCE_COL), start=2):
        row_values_filtered = []
        source_cell_formats = {}
        
        for src_col, cell in enumerate(row_cells, start=1):
            if src_col in col_map:
                row_values_filtered.append(cell.value)
                source_cell_formats[src_col] = {
                    'format': copy(cell.number_format),
                    'font': copy(cell.font), 
                    'fill': copy(cell.fill), 
                    'alignment': copy(cell.alignment),
                    'border': copy(cell.border)
                }
        
        if any(row_values_filtered):
            row_values_filtered = row_values_filtered + [None] * (MAX_GROUP_COL - len(row_values_filtered))
            data_rows_with_index.append((row_idx, tuple(row_values_filtered), source_cell_formats))

    groups = {}

    for original_row_idx, r_values, r_formats in data_rows_with_index:
        ident = r_values[col_identifier1 - 1]
        base = extract_sample_base(ident)
        norm = _normalize_text(base)
        
        if norm not in groups:
            groups[norm] = {"base": base, "rows": []}
        groups[norm]["rows"].append((original_row_idx, r_values, r_formats)) 
        
    carb_materials = settings.get_setting("REFERENCE_MATERIALS", sub_key="Carbonate") or []

    # --- Fetch UI settings for manual Drag & Drop mapping ---
    active_refs = settings.get_setting("ACTIVE_REFERENCES")
    active_samples = settings.get_setting("ACTIVE_SAMPLES")

    def is_reference_group(base_name):
        if not base_name: return False
        raw_text = str(base_name).strip()
        text_clean = re.sub(r'[\s\-_]+', '', raw_text.upper())
        text_no_std = text_clean.replace("STD", "")
        
        # 1. Trust the manual UI Drag & Drop lists completely
        if active_refs is not None and raw_text in active_refs:
            return True
        if active_samples is not None and raw_text in active_samples:
            return False
            
        # 2. Hardcoded exceptions (CO2/HeCO2)
        if text_clean.startswith("CO2") or text_clean.startswith("HECO2") or text_clean.startswith("C02") or text_clean.startswith("HEC02"):
            return True
            
        # 3. Check the global Advanced Settings table (carb_materials)
        for mat in carb_materials:
            std_name = mat.get("col_c")
            if not std_name: continue
            
            std_clean = re.sub(r'[\s\-_]+', '', str(std_name).upper())
            std_no_std = std_clean.replace("STD", "")
            
            if std_clean and std_clean in text_clean:
                return True
            if std_no_std and len(std_no_std) >= 4 and std_no_std in text_no_std:
                return True
                
        return False

    ref_groups = []
    other_groups = []
    col_time_code = col_map.get(2, 0)
    for norm, g in groups.items():
        if is_reference_group(g["base"]):
            # Every run of a standard belongs in its slope: one group, by run number.
            g["rows"].sort(key=lambda item: extract_run_number(item[1][col_identifier1 - 1]))
            ref_groups.append((norm, g))
        else:
            # A sample re-measured under the same name (A r1-r3 on DAY 12,
            # again on DAY 9, ...) gets one group per session, already in
            # measurement order. Split before any sort: sorting by run number
            # first is what interleaved the sessions as r1, r1, ..., r2, r2.
            for session_rows in split_into_sessions(
                    g["rows"],
                    ident_of=lambda item: item[1][col_identifier1 - 1],
                    time_of=lambda item: item[1][col_time_code - 1] if col_time_code else None):
                other_groups.append((norm, {"base": g["base"], "rows": session_rows}))

    present_refs_norm = {norm for norm, _ in ref_groups}
    has_iaea_603 = 'iaea603' in present_refs_norm
            
    current_row = header_row + 1
    n_arag_re = re.compile(r"\bn\.?\s*arag\b", flags=re.IGNORECASE)
    
    yield_mat_row_map = {}

    def _write_sample_output_cells(ws, r, is_arag, has_iaea_603, new_R_col, new_U_col, output_format, slope_info, box_offset=0, output_offset=0):
        num_groups = len(slope_info)
        if num_groups == 0:
            return 27 + output_offset
        
        layout = get_output_layout(num_groups, output_offset)
        is_arag = is_arag and layout["include_arag"]
        c_start = layout["c_start"]
        o_calc_start = layout["o_calc_start"]
        o_arag_start = layout["o_arag_start"]
        vsmow_calc_start = layout["vsmow_calc_start"]
        vsmow_arag_start = layout["vsmow_arag_start"]
        
        col_AA_approx = layout["box2_end"]
        
        green_bold = Font(color="008000", bold=True)
        
        c_slope_col = get_column_letter(11 + box_offset)
        o_slope_col = get_column_letter(14 + box_offset)
        
        for i in range(num_groups):
            s_row = slope_info[i]["slope_row"]
            i_row = slope_info[i]["intercept_row"]
            
            c_formula = f'=IFERROR(${c_slope_col}${s_row}*{get_column_letter(new_R_col)}{r}+${c_slope_col}${i_row},"")'
            o_formula = f'=IFERROR(${o_slope_col}${s_row}*{get_column_letter(new_U_col)}{r}+${o_slope_col}${i_row},"")'
            
            c_col = out_col(c_start, i)
            o_calc_col = out_col(o_calc_start, i)
            o_arag_col = out_col(o_arag_start, i) if o_arag_start else None
            vsmow_calc_col = out_col(vsmow_calc_start, i)
            vsmow_arag_col = out_col(vsmow_arag_start, i) if vsmow_arag_start else None
            
            cell_c = ws.cell(row=r, column=c_col, value=c_formula)
            cell_c.number_format = output_format
            
            if is_arag:
                ws.cell(row=r, column=o_calc_col, value=None)
                ws.cell(row=r, column=vsmow_calc_col, value=None)
                
                cell_o_arag = ws.cell(row=r, column=o_arag_col, value=o_formula)
                cell_o_arag.font = Font(bold=True)
                cell_o_arag.number_format = output_format
                
                vsmow_formula = f'=IFERROR((1.03092*{get_column_letter(o_arag_col)}{r})+30.92,"")'
                cell_vsmow_arag = ws.cell(row=r, column=vsmow_arag_col, value=vsmow_formula)
                cell_vsmow_arag.font = Font(bold=True)
                cell_vsmow_arag.number_format = output_format
                
                for col_idx in (c_col, o_arag_col, vsmow_arag_col):
                    ws.cell(row=r, column=col_idx).font = green_bold
            else:
                if o_arag_col:
                    ws.cell(row=r, column=o_arag_col, value=None)
                    ws.cell(row=r, column=vsmow_arag_col, value=None)
                
                cell_o_calc = ws.cell(row=r, column=o_calc_col, value=o_formula)
                cell_o_calc.number_format = output_format
                
                vsmow_formula = f'=IFERROR((1.03092*{get_column_letter(o_calc_col)}{r})+30.92,"")'
                cell_vsmow_calc = ws.cell(row=r, column=vsmow_calc_col, value=vsmow_formula)
                cell_vsmow_calc.font = Font(bold=True)
                cell_vsmow_calc.number_format = output_format
                
                for col_idx in (c_col, o_calc_col, vsmow_calc_col):
                    original_font = copy(ws.cell(row=r, column=col_idx).font)
                    original_font.bold = True
                    ws.cell(row=r, column=col_idx).font = original_font
                    
        return col_AA_approx

    def write_group(norm, g, is_reference=True, has_iaea_603=False):
        nonlocal current_row
        base_name = _normalize_text(g["base"])
        rows_data = g["rows"]
        font_color = color_fonts.get(base_name)
        
        row_values_list = [item[1] for item in rows_data]
        valid_indices = []
        if base_name in ("co2", "heco2"):
            valid_indices = _get_valid_co2_rows(row_values_list, col_identifier1)
            
        row_map = [] 
        
        c_vals = []; o_vals = [] 
        valid_run_row_indices = [] 

        for i, (source_row_idx, row_values_filtered, source_cell_formats) in enumerate(rows_data):
            excel_row = current_row
            
            for dest_col_idx in range(1, MAX_GROUP_COL + 1):
                src_col = next(s for s, d in col_map.items() if d == dest_col_idx)
                val = row_values_filtered[dest_col_idx - 1]
                dest_cell = ws_group.cell(row=excel_row, column=dest_col_idx, value=val)
                
                original_format = source_cell_formats.get(src_col, {})
                if original_format.get('format'): dest_cell.number_format = original_format['format']
                if original_format.get('font'): dest_cell.font = original_format['font']
                if original_format.get('fill'): dest_cell.fill = original_format['fill']
                if original_format.get('alignment'): dest_cell.alignment = original_format['alignment']
                if original_format.get('border'): dest_cell.border = original_format['border']

                if font_color and dest_col_idx == col_identifier1:
                    new_font = copy(dest_cell.font)
                    new_font.color = font_color.color
                    dest_cell.font = new_font
                    
            if base_name in ("co2", "heco2") and i in valid_indices:
                # Every cell gets the fill explicitly: a cell that already has
                # its own font or number format (the normalized output cells do)
                # carries its own empty fill too, which paints white over any
                # row-level shading.
                for col in range(1, last_layout_col + 1):
                    ws_group.cell(row=excel_row, column=col).fill = gray_fill
                    
            ident_val = str(ws_group.cell(row=excel_row, column=col_identifier1).value or "")
            is_arag = n_arag_re.search(ident_val)
            
            for block in data_blocks:
                _write_sample_output_cells(ws_group, excel_row, bool(is_arag), has_iaea_603, 11, 14,
                                           '0.00', block["slope_info"], block["top_shift"], block["out_shift"])
            
            # --- THEORETICAL & YIELD FORMULA INJECTION ---
            if calc_yield and base_name not in ("co2", "heco2"):
                col_an = get_column_letter(YIELD_START_COL + 7)
                
                target_row_val = ref_compound_row if is_reference else samp_compound_row
                ref_cell = f"${col_an}${target_row_val}"
                
                theo_cell = ws_group.cell(row=excel_row, column=theo_col, value=f'=IF(ISNUMBER(G{excel_row}), G{excel_row}*{ref_cell}, "")')
                theo_cell.number_format = '0.00E+00'
                # Same colour code as the Yield table's compounds
                theo_cell.font = Font(color=YIELD_REF_COLOR if is_reference else YIELD_SAMP_COLOR)
                
                for grp_idx in range(num_yield_groups):
                    calc_col = theo_col + 2 + (grp_idx * 4)
                    yield_pct_col = calc_col + 2
                    
                    slope_cell = f"${get_column_letter(YIELD_START_COL + 5)}${20 + grp_idx*3}"
                    int_cell = f"${get_column_letter(YIELD_START_COL + 5)}${21 + grp_idx*3}"
                    
                    calc_cell = ws_group.cell(row=excel_row, column=calc_col, value=f'=IF(ISNUMBER({slope_cell}),IFERROR(({slope_cell}*Q{excel_row})+{int_cell},""),"")')
                    calc_cell.number_format = '0.00E+00'
                    
                    yp_cell = ws_group.cell(row=excel_row, column=yield_pct_col, value=f'=IFERROR(({get_column_letter(calc_col)}{excel_row}/{get_column_letter(theo_col)}{excel_row})*100,"")')
                    yp_cell.number_format = '0.00'

            # A run with the wrong peak count is unreliable for any material,
            # so it stays on the sheet but out of every average.
            is_flagged = has_peak_count_flag(row_values_filtered[CO2_PEAK_FLAG_COL - 1]
                                             if len(row_values_filtered) >= CO2_PEAK_FLAG_COL else None)

            should_process = True
            if (base_name in ("co2", "heco2") and i not in valid_indices) or is_flagged:
                should_process = False
                c_vals.append(None); o_vals.append(None)

            if is_flagged:
                # Struck in the cell's own colour, so a flagged run reads
                # differently from a red sigma outlier.
                strike_flagged_cells(ws_group, excel_row, (11, 14))
                for block in data_blocks:
                    strike_output_cells(ws_group, excel_row, len(block["slope_info"]), bool(is_arag),
                                        block["out_shift"], True, True, color=None)
            
            if should_process:
                valid_run_row_indices.append(i)
                cv = row_values_filtered[11-1] 
                ov = row_values_filtered[14-1] 
                try: cv = float(cv)
                except: cv = None
                try: ov = float(ov)
                except: ov = None
                c_vals.append(cv); o_vals.append(ov)
                
                if base_name not in yield_mat_row_map:
                    yield_mat_row_map[base_name] = []
                yield_mat_row_map[base_name].append(excel_row)

            row_map.append(excel_row)
            current_row += 1
            
        if is_reference or len(rows_data) > 1:
            valid_c_nums = [v for v in c_vals if v is not None]
            valid_o_nums = [v for v in o_vals if v is not None]

            mean_c = statistics.mean(valid_c_nums) if len(valid_c_nums) > 1 else 0
            stdev_c = statistics.stdev(valid_c_nums) if len(valid_c_nums) > 1 else 0
            mean_o = statistics.mean(valid_o_nums) if len(valid_o_nums) > 1 else 0
            stdev_o = statistics.stdev(valid_o_nums) if len(valid_o_nums) > 1 else 0

            c_up, c_low = mean_c + (outlier_sigma * stdev_c), mean_c - (outlier_sigma * stdev_c)
            o_up, o_low = mean_o + (outlier_sigma * stdev_o), mean_o - (outlier_sigma * stdev_o)

            all_runs_c = []; all_runs_o = []
            # Carbon and Oxygen are tracked in SEPARATE lists. A single shared
            # list of surviving rows cannot express "Individual" mode, where a
            # run can lose its C value to an outlier but keep its O value.
            final_runs_c = []; final_runs_o = []

            # Same split, as bare row numbers: the normalized output blocks need
            # to average their own columns over exactly these rows.
            all_rows_c = []; all_rows_o = []
            final_rows_c = []; final_rows_o = []

            c_col_let = "K"; o_col_let = "N"

            for i in valid_run_row_indices:
                r_num = row_map[i]
                vc = c_vals[i]; vo = o_vals[i]

                if vc is not None: all_runs_c.append(f"{c_col_let}{r_num}"); all_rows_c.append(r_num)
                if vo is not None: all_runs_o.append(f"{o_col_let}{r_num}"); all_rows_o.append(r_num)

                is_c_out = (vc > c_up or vc < c_low) if (vc is not None and len(valid_c_nums) > 2) else False
                is_o_out = (vo > o_up or vo < o_low) if (vo is not None and len(valid_o_nums) > 2) else False

                exclude_c = False; exclude_o = False
                if exclusion_mode == "Exclude Row":
                    if is_c_out or is_o_out: exclude_c = True; exclude_o = True
                else:
                    if is_c_out: exclude_c = True
                    if is_o_out: exclude_o = True
                
                # A value joins the filtered average only if it exists and was
                # not excluded for its own column (same rule as step5_group.py).
                if not exclude_c and vc is not None: final_runs_c.append(f"{c_col_let}{r_num}"); final_rows_c.append(r_num)
                if not exclude_o and vo is not None: final_runs_o.append(f"{o_col_let}{r_num}"); final_rows_o.append(r_num)
                
                if exclude_c and vc is not None: ws_group.cell(row=r_num, column=11).font = strike_font
                if exclude_o and vo is not None: ws_group.cell(row=r_num, column=14).font = strike_font
                
                strike_c = exclude_c and vc is not None
                strike_o = exclude_o and vo is not None
                if strike_c or strike_o:
                    row_is_arag = n_arag_re.search(str(ws_group.cell(row=r_num, column=col_identifier1).value or ""))
                    for block in data_blocks:
                        strike_output_cells(ws_group, r_num, len(block["slope_info"]), bool(row_is_arag),
                                            block["out_shift"], strike_c, strike_o)
        
            summary_font = font_color or Font(bold=True)
            if is_reference and font_color:
                summary_font = copy(font_color); summary_font.bold = True
            
            is_arag_group = n_arag_re.search(g["base"])
            
            # Normalized stats carry the material's colour; an uncoloured
            # aragonite sample follows its column and goes green.
            norm_font = copy(summary_font)
            norm_font.bold = True
            if is_arag_group and not font_color:
                norm_font = Font(color="008000", bold=True)
            
            
            # The normalized blocks only repeat their stats when an outlier
            # actually dropped a run; otherwise the second set would just copy
            # the first. The raw K-P summary still always shows both.
            c_has_outlier = set(all_rows_c) != set(final_rows_c)
            o_has_outlier = set(all_rows_o) != set(final_rows_o)
            any_outlier = c_has_outlier or o_has_outlier
            
            data_fmt = get_summary_num_format(base_name); count_fmt = '0'

            row_all = current_row
            ws_group.cell(row=row_all, column=10, value="--").font = Font(bold=True) 
            for col, txt in zip([11, 12, 13, 14, 15, 16], ["Average", "Stdev", "Count", "Average", "Stdev", "Count"]):
                c = ws_group.cell(row=row_all, column=col, value=txt)
                c.font = summary_font; c.alignment = Alignment(horizontal="right")
            
            for block in data_blocks:
                write_summary_output_labels(ws_group, row_all, len(block["slope_info"]),
                                            bool(is_arag_group), block["out_shift"], norm_font)
                
            if calc_yield and base_name not in ("co2", "heco2"):
                for grp_idx in range(num_yield_groups):
                    yield_pct_col = theo_col + 4 + (grp_idx * 4)
                    c = ws_group.cell(row=row_all, column=yield_pct_col, value="Average")
                    c.font = summary_font; c.alignment = Alignment(horizontal="right")

            row_all_calc = row_all + 1
            if all_runs_c:
                rng_c = ",".join(all_runs_c)
                ws_group.cell(row=row_all_calc, column=11, value=f"=AVERAGE({rng_c})").number_format = data_fmt
                ws_group.cell(row=row_all_calc, column=12, value=f"=STDEV({rng_c})").number_format = data_fmt
                ws_group.cell(row=row_all_calc, column=13, value=f"=COUNT({rng_c})").number_format = count_fmt
            if all_runs_o:
                rng_o = ",".join(all_runs_o)
                ws_group.cell(row=row_all_calc, column=14, value=f"=AVERAGE({rng_o})").number_format = data_fmt
                ws_group.cell(row=row_all_calc, column=15, value=f"=STDEV({rng_o})").number_format = data_fmt
                ws_group.cell(row=row_all_calc, column=16, value=f"=COUNT({rng_o})").number_format = count_fmt
            
            for c in range(11, 17): ws_group.cell(row=row_all_calc, column=c).font = summary_font
            
            if calc_yield and base_name not in ("co2", "heco2"):
                all_runs_rows = [row_map[i] for i in valid_run_row_indices]
                for grp_idx in range(num_yield_groups):
                    yield_pct_col = theo_col + 4 + (grp_idx * 4)
                    y_col_let = get_column_letter(yield_pct_col)
                    y_range_str = get_excel_range(y_col_let, all_runs_rows)
                    if y_range_str:
                        ws_group.cell(row=row_all_calc, column=yield_pct_col, value=f'=IFERROR(AVERAGE({y_range_str}),"")').number_format = '0.00'
                        ws_group.cell(row=row_all_calc, column=yield_pct_col).font = summary_font

            for block in data_blocks:
                write_summary_output_cells(ws_group, row_all_calc, len(block["slope_info"]),
                                           bool(is_arag_group), all_rows_c, all_rows_o, '0.00',
                                           block["out_shift"], norm_font)

            row_filt = row_all_calc + 1
            ws_group.cell(row=row_filt, column=10, value="Outlier Excl.").font = Font(bold=True) 
            for col, txt in zip([11, 12, 13, 14, 15, 16], ["Average", "Stdev", "Count", "Average", "Stdev", "Count"]):
                c = ws_group.cell(row=row_filt, column=col, value=txt)
                c.font = summary_font; c.alignment = Alignment(horizontal="right")
            
            if any_outlier:
                for block in data_blocks:
                    write_summary_output_labels(ws_group, row_filt, len(block["slope_info"]),
                                                bool(is_arag_group), block["out_shift"], norm_font,
                                                c_has_outlier, o_has_outlier, outlier_marker=True)

            row_filt_calc = row_filt + 1
            
            # Filtered ranges, built per column while detecting the outliers above.
            rng_c_filt = final_runs_c
            rng_o_filt = final_runs_o

            if rng_c_filt:
                rng_c_str = ",".join(rng_c_filt)
                ws_group.cell(row=row_filt_calc, column=11, value=f"=AVERAGE({rng_c_str})").number_format = data_fmt
                ws_group.cell(row=row_filt_calc, column=12, value=f"=STDEV({rng_c_str})").number_format = data_fmt
                ws_group.cell(row=row_filt_calc, column=13, value=f"=COUNT({rng_c_str})").number_format = count_fmt
            if rng_o_filt:
                rng_o_str = ",".join(rng_o_filt)
                ws_group.cell(row=row_filt_calc, column=14, value=f"=AVERAGE({rng_o_str})").number_format = data_fmt
                ws_group.cell(row=row_filt_calc, column=15, value=f"=STDEV({rng_o_str})").number_format = data_fmt
                ws_group.cell(row=row_filt_calc, column=16, value=f"=COUNT({rng_o_str})").number_format = count_fmt

            for c in range(11, 17): ws_group.cell(row=row_filt_calc, column=c).font = summary_font

            if any_outlier:
                filt_c = final_rows_c if c_has_outlier else None
                filt_o = final_rows_o if o_has_outlier else None
                for block in data_blocks:
                    write_summary_output_cells(ws_group, row_filt_calc, len(block["slope_info"]),
                                               bool(is_arag_group), filt_c, filt_o, '0.00',
                                               block["out_shift"], norm_font)

            current_row = row_filt_calc + 2 
        else:
            current_row += 3
            
    for norm, g in ref_groups:
        write_group(norm, g, is_reference=True)
        
    if ref_groups:
        current_row += 8
        divider_top_row = current_row 
        # The divider has to reach past the right-most block, which now grows
        # with the slope-group count.
        divider_end_col = max(701, LAST_LAYOUT_COL)
        for _ in range(2):
            for col in range(1, divider_end_col + 1):
                ws_group.cell(row=current_row, column=col).fill = dark_fill
            current_row += 1
            
        current_header_col = 1
        for src_col, dest_col_idx in col_map.items():
            src_cell = first_row_cells[src_col - 1] if src_col <= len(first_row_cells) else None
            
            dest_cell = ws_group.cell(row=current_row, column=current_header_col, value=headers[current_header_col - 1])
            
            if src_cell:
                dest_cell.number_format = copy(src_cell.number_format)
                dest_cell.font = copy(src_cell.font)
                dest_cell.alignment = copy(src_cell.alignment)
                dest_cell.border = copy(src_cell.border)
                dest_cell.fill = copy(src_cell.fill)

        # BOTTOM YIELD BOX START
        if calc_yield:
            if num_yield_groups > 0:
                yield_calc_end = theo_col + (num_yield_groups * 4)
                thick = Side(style="thick")
                
                # Align perfectly with the Carbonate and CO2 lower boxes (which span from -3 to +1)
                bot_y_top = divider_top_row - 3
                bot_y_bottom = divider_top_row + 1
                
                # 1. Draw the main box
                for r in range(bot_y_top, bot_y_bottom + 1):
                    for c in range(theo_col, yield_calc_end + 1):
                        cell = ws_group.cell(row=r, column=c)
                        cell.fill = yield_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        
                        top_border = thick if r == bot_y_top else None
                        bottom_border = thick if r == bot_y_bottom else None
                        left_border = thick if c == theo_col else None
                        right_border = thick if c == yield_calc_end else None
                        cell.border = Border(top=top_border, bottom=bottom_border, left=left_border, right=right_border)
                        
                # 2. "Normalized Yield" Title (merged horizontally)
                ws_group.merge_cells(start_row=bot_y_top, start_column=theo_col, end_row=bot_y_top, end_column=yield_calc_end)
                ny_cell = ws_group.cell(row=bot_y_top, column=theo_col, value="Normalized Yield")
                ny_cell.font = Font(bold=True)
                
                # 3. "Theoretical" Header (placed near the top, no vertical merge)
                theo_header_bottom = ws_group.cell(row=bot_y_top + 4, column=theo_col, value="Theoretical")
                theo_header_bottom.font = Font(bold=True, color="000000")
                
                # 4. Populate each Yield Group
                for i in range(num_yield_groups):
                    calc_col = theo_col + 2 + (i * 4)
                    yield_pct_col = calc_col + 2
                    
                    # "Calculated" and "Yield (%)" headers
                    c_head = ws_group.cell(row=bot_y_bottom - 1, column=calc_col, value="Calculated")
                    c_head.font = Font(bold=True)
                    
                    y_head = ws_group.cell(row=bot_y_bottom - 1, column=yield_pct_col, value="Yield (%)")
                    y_head.font = Font(bold=True)
                    
                    # Reference Materials underneath (merged across the 3 columns)
                    rt = get_local_yield_rich_text(yield_slope_groups[i])
                    ws_group.merge_cells(start_row=bot_y_bottom, start_column=calc_col, end_row=bot_y_bottom, end_column=yield_pct_col)
                    rt_cell = ws_group.cell(row=bot_y_bottom, column=calc_col)
                    rt_cell.value = rt
            else:
                theo_header_bottom = ws_group.cell(row=current_row, column=theo_col, value="Theoretical")
                theo_header_bottom.fill = yield_fill
                theo_header_bottom.font = Font(bold=True, color="000000")
                theo_header_bottom.alignment = Alignment(horizontal="center", vertical="center")
        # BOTTOM YIELD BOX END

        current_header_col += 1

        current_row += 1
        
        for block in data_blocks:
            draw_lower_boxes(ws_group, divider_top_row, block["fill"], Font(bold=True, color="000000"),
                             Font(bold=True, color="008000"), block["out_shift"], block["key"])
        
    for norm, g in other_groups:
        write_group(norm, g, is_reference=False, has_iaea_603=has_iaea_603)
        
    for block in data_blocks:
        populate_blue_box_math(ws_group, block["slope_info"], block["mat_map"],
                               block["top_shift"], block["key"])
    
    # --- YIELD SLOPE & INTERCEPT DYNAMIC POPULATION ---
    if calc_yield:
        def get_group_rich_text(group_list):
            parts = [(black_if, "Used ")]
            for item in group_list:
                clean = str(item).strip().upper()
                font = black_if
                if "18" in clean: font = red_if
                elif "19" in clean: font = blue_if
                elif "603" in clean: font = green_if
                parts.append((font, clean.replace("NBS","").replace("IAEA","").strip() + " "))
            return create_rich_text(parts)
            
        yield_slope_row = 20 
        col_AI = YIELD_START_COL + 2
        col_AJ = YIELD_START_COL + 3
        col_AL = YIELD_START_COL + 5
        
        for grp_idx, group in enumerate(yield_slope_groups):
            yield_slope_row = 20 + (grp_idx * 3)
            group_rows = []
            for mat in group:
                for key, rows in yield_mat_row_map.items():
                    if _yield_name_matches(mat, key):
                        group_rows.extend(rows)
                        
            if len(group_rows) >= 2:
                ws_group.cell(row=yield_slope_row, column=col_AI).value = get_group_rich_text(group)
                
                ws_group.cell(row=yield_slope_row, column=col_AJ, value="slope").font = Font(bold=True)
                ws_group.cell(row=yield_slope_row+1, column=col_AJ, value="intercept").font = Font(bold=True)
                
                range_y = get_excel_range(get_column_letter(theo_col), group_rows)
                range_x = get_excel_range('Q', group_rows)
                
                col_AL_let = get_column_letter(col_AL)

                # Yield Slope Array Formula
                cell_slope = ws_group.cell(row=yield_slope_row, column=col_AL)
                cell_slope.value = ArrayFormula(f"{col_AL_let}{yield_slope_row}", f'=SLOPE({range_y},{range_x})')
                cell_slope.number_format = '0.00E+00'
                cell_slope.font = Font(bold=True)
                
                # Yield Intercept Array Formula
                cell_int = ws_group.cell(row=yield_slope_row+1, column=col_AL)
                cell_int.value = ArrayFormula(f"{col_AL_let}{yield_slope_row+1}", f'=INTERCEPT({range_y},{range_x})')
                cell_int.number_format = '0.00E+00'
                cell_int.font = Font(bold=True)
            
    start_gray_row = header_row + 1 
    max_sheet_row = ws_group.max_row + 50
    
    for row in range(start_gray_row, max_sheet_row + 1):
        for col in (11, 14): 
            ws_group.cell(row=row, column=col).fill = gray_fill
            
    ws_group.column_dimensions["A"].width = 13 
    ws_group.column_dimensions["C"].width = 22 
    ws_group.column_dimensions["H"].width = 15 
    ws_group.column_dimensions["J"].width = 16 

    for offset in [b["top_shift"] for b in data_blocks]:
        ws_group.column_dimensions[get_column_letter(3 + offset)].width = 22
        ws_group.column_dimensions[get_column_letter(6 + offset)].width = 12
        ws_group.column_dimensions[get_column_letter(7 + offset)].width = 12
        ws_group.column_dimensions[get_column_letter(10 + offset)].width = 16
        ws_group.column_dimensions[get_column_letter(11 + offset)].width = 12
        ws_group.column_dimensions[get_column_letter(14 + offset)].width = 12
        
    for output_offset, n_out_groups in [(b["out_shift"], len(b["slope_info"])) for b in data_blocks]:
        if n_out_groups <= 0:
            continue
        out_layout = get_output_layout(n_out_groups, output_offset)
        # Spacers carry the "Outlier Excl." marker, so they set the pace; the
        # value column holds "Average", Stdev and Count need far less.
        for output_col in range(out_layout["box1_start"] - 1, out_layout["box2_end"] + 1):
            ws_group.column_dimensions[get_column_letter(output_col)].width = 14
        for section in ("c_start", "o_calc_start", "o_arag_start", "vsmow_calc_start", "vsmow_arag_start"):
            start = out_layout[section]
            if not start:
                continue
            for i in range(n_out_groups):
                base = out_col(start, i)
                ws_group.column_dimensions[get_column_letter(base)].width = 12
                ws_group.column_dimensions[get_column_letter(base + 1)].width = 9
                ws_group.column_dimensions[get_column_letter(base + 2)].width = 7
            
    if calc_yield:
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 1)].width = 15 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 2)].width = 15 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 3)].width = 15 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 5)].width = 15 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 7)].width = 15 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 9)].width = 10 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 10)].width = 10 
        ws_group.column_dimensions[get_column_letter(YIELD_START_COL + 11)].width = 10 

    max_data_row = ws_group.max_row
    threshold_str = str(stdev_threshold)
    
    ws_group.conditional_formatting.add(
        f"L{start_gray_row}:L{max_data_row}",
        FormulaRule(formula=[f'=AND(ISNUMBER(L{start_gray_row}), L{start_gray_row} > {threshold_str})'], fill=red_fill)
    )

    ws_group.conditional_formatting.add(
        f"O{start_gray_row}:O{max_data_row}",
        FormulaRule(formula=[f'=AND(ISNUMBER(O{start_gray_row}), O{start_gray_row} > {threshold_str})'], fill=red_fill)
    )

    # Dynamic Freeze Pane (Conditional)
    if settings.get_setting("ENABLE_FREEZE_PANE") is not False:
        ws_group.freeze_panes = f'B{start_gray_row}'

    embed_settings_popup(ws_group, "A2")

    # Open on Normalization_DNT with a single tab selected. Both halves matter:
    # tabSelected marks the tab, wb.active sets the workbook's activeTab index.
    # If they point at different sheets, Excel opens the workbook with two tabs
    # selected as a group.
    for s in wb.worksheets:
        try:
            s.sheet_view.tabSelected = False
        except Exception:
            pass
    ws_group.sheet_view.tabSelected = True
    wb.active = wb.index(ws_group)
    ws_group.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    save_workbook_atomic(wb, file_path)
    print(f"✅ Step 6: Normalization completed on {file_path}")