import pandas as pd
import string
from openpyxl import load_workbook
from openpyxl.worksheet.views import Selection
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from copy import copy
import numpy as np
import re
from openpyxl.styles import Protection, Alignment, Border
from openpyxl.formatting.rule import CellIsRule, FormulaRule 
import settings 

# Define the error fill
fill_error = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
# Define Gray fill for HeCO2/CO2
heco2_gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

def extract_run_number(identifier):
    """
    Parses strings like 'r1', 'r2.1', 'R3' into (major, minor).
    Returns (None, None) if parsing fails.
    """
    if not identifier:
        return None, None
    
    match = re.search(r'(?i)[rR](\d+)(?:\.(\d+))?', str(identifier))
    if match:
        major = int(match.group(1))
        minor = int(match.group(2)) if match.group(2) else 0
        return major, minor
    return None, None

def get_valid_heco2_source_rows(df):
    """
    Scans the DataFrame to find valid HeCO2/CO2 rows based on Source_Row.
    Logic: Skip r1, keep lowest minor for r2, r3, etc.
    Returns a set of Source_Rows.
    """
    valid_source_rows = set()
    seen_majors = {} # Map major -> (minor, source_row)
    
    for _, row in df.iterrows():
        ident = str(row['Identifier 1']).strip().lower()
        
        if 'heco2' in ident or 'co2' in ident:
            major, minor = extract_run_number(ident)
            
            # Skip parsing failures or Major Run 1
            if major is None or major == 1: 
                continue
            
            # Logic: Keep lowest minor for each major
            if major not in seen_majors:
                seen_majors[major] = (minor, row['Source_Row'])
            else:
                prev_minor, prev_row = seen_majors[major]
                if minor < prev_minor:
                    seen_majors[major] = (minor, row['Source_Row'])
                    
    # Collect the valid Source Rows
    for _, s_row in seen_majors.values():
        valid_source_rows.add(s_row)
        
    return valid_source_rows

def step5_group_water(file_path: str):
    """
    Fixed Step 4: Grouping by Identifier
    - Groups into HeCO2, Reference, and Others
    - Adds HeCO2 Gray highlighting for specific rows.
    - UPDATED: He/CO2 Calculations ONLY use the gray highlighted rows.
    - Adds a gray divider section after the last reference group.
    - Preserves formatting, column widths, row heights.
    - Applies Conditional Formatting and 3-decimal number formats.
    """

    new_sheet_name = "Group_DNT"
    source_sheet_name = "Last 6_DNT"
    
    # Check for the required setting
    stdev_threshold = settings.get_setting("STDEV_THRESHOLD")
    if stdev_threshold is None:
        print("⚠️ Warning: 'stdev_threshold' not found in settings. Conditional formatting will be skipped.")

    # Define the number format for 3 decimal places
    THREE_DECIMAL_FORMAT = "0.000"

    # --- 1. Load Workbook ---
    wb_values = load_workbook(file_path, data_only=True)
    wb = load_workbook(file_path)

    matched_source = next((s for s in wb.sheetnames if s.lower() == source_sheet_name.lower()), None)
    if matched_source is None:
        raise ValueError(f"Source sheet matching '{source_sheet_name}' not found. Ensure Step 3 was run with 'Last 6'.")

    source_ws_vals = wb_values[matched_source]
    source_ws = wb[matched_source]

    if new_sheet_name in wb.sheetnames:
        del wb[new_sheet_name]

    source_cols_to_read = [1, 2, 3, 4, 5, 6, 7, 10, 14, 15, 16, 17, 18, 19, 20, 21, 22]

    header_row = [source_ws_vals.cell(row=1, column=c).value for c in source_cols_to_read]

    data = []
    for row_idx in range(2, source_ws.max_row + 1):
        if not source_ws.row_dimensions[row_idx].hidden:
            row_values = [source_ws_vals.cell(row=row_idx, column=c).value for c in source_cols_to_read]
            row_values.append(row_idx)
            data.append(row_values)

    df_columns = header_row + ['Source_Row']
    df = pd.DataFrame(data, columns=df_columns)

    new_headers = [
        'Line', 'Time Code', 'Identifier 1', 'Comment', 'Identifier 2', 'Analysis', 'Preparation',
        'Ampl 44', '', '', 'C avg', 'C stdev', '', 'O avg', 'O stdev', '', 'Sum area all'
    ]
    df.columns = new_headers + ['Source_Row']

    # --- 2. Clean and Group Data ---
    def create_group_key(identifier):
        if pd.isna(identifier) or not isinstance(identifier, str):
            return identifier
        return re.sub(r'\s+[rR]\d+(?:\.\d+)*(?:[a-zA-Z]*)?$', '', identifier).strip()

    df['Group_Key'] = df['Identifier 1'].apply(create_group_key)
    df['Line_num'] = pd.to_numeric(df['Line'], errors='coerce')

    group_min_line = df.groupby('Group_Key', sort=False)['Line_num'].min().reset_index().rename(columns={'Line_num': 'min_line'})

    # --- Categorize groups ---
    def is_heco2_key(k):
        if k is None:
            return False
        return bool(re.search(r'(?i)\b(heco2|co2)\b', str(k)))

    def is_reference_key(k):
        if k is None:
            return False
        text = str(k).strip().upper()
        patterns = [
            r'\bMRSI\b', r'\bMRSI[- ]?\d+\b', r'\bMRSI[- ]?W?\d+\b',
            r'\bMRSI[- ]?STD[- ]?W?\d+\b', r'\bUSGS[- ]?W[- ]?\d+\b',
            r'\bMRSIW\d+\b', r'\bUSGSW\d+\b'
        ]
        return any(re.search(p, text) for p in patterns)

    group_min_line['is_heco2'] = group_min_line['Group_Key'].apply(is_heco2_key)
    group_min_line['is_ref'] = group_min_line['Group_Key'].apply(is_reference_key)

    # --- Sort order ---
    heco2 = group_min_line[group_min_line['is_heco2']].sort_values(by=['min_line', 'Group_Key'])
    refs = group_min_line[~group_min_line['is_heco2'] & group_min_line['is_ref']].sort_values(by=['min_line', 'Group_Key'])
    others = group_min_line[~group_min_line['is_heco2'] & ~group_min_line['is_ref']].sort_values(by=['min_line', 'Group_Key'])

    ordered_groups = pd.concat([heco2, refs, others])['Group_Key'].tolist()

    df['Group_Key'] = pd.Categorical(df['Group_Key'], categories=ordered_groups, ordered=True)
    df_sorted = df.sort_values(by=['Group_Key', 'Line_num', 'Source_Row'], na_position='last').reset_index(drop=True)

    # --- 3. Write to New Sheet ---
    # Try to place to the left of 'Pre-Group_DNT'
    target_sheet_name = "Pre-Group_DNT"
    
    if target_sheet_name in wb.sheetnames:
        idx = wb.index(wb[target_sheet_name])
    else:
        # Fallback to source sheet if Pre-Group doesn't exist
        idx = wb.index(source_ws)

    new_ws = wb.create_sheet(new_sheet_name, idx)

    header_fill = PatternFill(start_color="8ED973", end_color="8ED973", fill_type="solid")
    gray_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")

    for col_idx, header in enumerate(new_headers, start=1):
        cell = new_ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font()
        cell.fill = header_fill

    new_ws.freeze_panes = "A2"

    for t_col_idx, src_col_idx in enumerate(source_cols_to_read, start=1):
        src_letter = get_column_letter(src_col_idx)
        tgt_letter = get_column_letter(t_col_idx)
        try:
            width = source_ws.column_dimensions[src_letter].width
            if width is not None:
                new_ws.column_dimensions[tgt_letter].width = width
        except Exception:
            pass

    def copy_cell_style(src_cell, tgt_cell):
        if src_cell is None or tgt_cell is None:
            return
        try:
            if src_cell.has_style:
                tgt_cell.font = copy(src_cell.font)
                tgt_cell.border = copy(src_cell.border)
                tgt_cell.fill = copy(src_cell.fill)
                tgt_cell.number_format = src_cell.number_format 
                tgt_cell.alignment = copy(src_cell.alignment)
                tgt_cell.protection = copy(src_cell.protection)
        except Exception:
            pass

    # --- Identify Valid HeCO2 Rows for Gray Highlighting ---
    valid_heco2_source_rows = get_valid_heco2_source_rows(df)

    cur_row = 3
    grouped = df_sorted.groupby('Group_Key', sort=False)

    # --- Track for divider insertion ---
    last_ref_group = refs['Group_Key'].iloc[-1] if not refs.empty else None

    for group_key, group in grouped:
        data_start_row = cur_row
        
        # Track which destination rows are "valid" for calculation (highlighted gray)
        # This is reset for every group
        valid_calculation_rows = []

        # Write Group Data
        for _, row_series in group.iterrows():
            source_row_idx = int(row_series['Source_Row']) if not pd.isna(row_series['Source_Row']) else None
            
            # Check if this row needs Gray Highlighting (It is a valid He/CO2 row)
            is_gray_row = (source_row_idx in valid_heco2_source_rows)
            
            if is_gray_row:
                valid_calculation_rows.append(cur_row)

            for col_offset, header in enumerate(new_headers, start=1):
                value = None
                if source_row_idx is not None:
                    src_col_idx = source_cols_to_read[col_offset - 1]
                    value = source_ws_vals.cell(row=source_row_idx, column=src_col_idx).value
                else:
                    value = row_series[header]
                
                tgt_cell = new_ws.cell(row=cur_row, column=col_offset, value=value)
                
                if source_row_idx is not None:
                    src_cell = source_ws.cell(row=source_row_idx, column=source_cols_to_read[col_offset - 1])
                    copy_cell_style(src_cell, tgt_cell)
                
                # Apply Gray Fill Override if it's a valid HeCO2 row
                if is_gray_row:
                    tgt_cell.fill = heco2_gray_fill

            try:
                if source_row_idx is not None:
                    rh = source_ws.row_dimensions[source_row_idx].height
                    if rh is not None:
                        new_ws.row_dimensions[cur_row].height = rh
            except Exception:
                pass
            cur_row += 1

        data_end_row = cur_row - 1
        calc_row_mid = cur_row + 1 

        # Set headers for calculation rows
        new_ws.cell(row=calc_row_mid, column=10, value="--").font = Font(bold=True)
        new_ws.cell(row=cur_row, column=11, value="Average").font = Font(bold=True)
        new_ws.cell(row=cur_row, column=12, value="Stdev").font = Font(bold=True)
        new_ws.cell(row=cur_row, column=13, value="Count").font = Font(bold=True)
        new_ws.cell(row=cur_row, column=14, value="Average").font = Font(bold=True)
        new_ws.cell(row=cur_row, column=15, value="Stdev").font = Font(bold=True)
        new_ws.cell(row=cur_row, column=16, value="Count").font = Font(bold=True)

        if data_end_row >= data_start_row:
            col_K_letter = get_column_letter(11)
            col_L_letter = get_column_letter(12) 
            col_N_letter = get_column_letter(14)
            col_O_letter = get_column_letter(15) 
            col_Q_letter = get_column_letter(17)
            bold_font = Font(bold=True)
            
            formulas = {}
            
            # === FORMULA GENERATION LOGIC ===
            # If we have valid_calculation_rows, it means this is a He/CO2 group with specific valid rows.
            # We must use specific cells (comma separated) for the calculation.
            # If valid_calculation_rows is empty, it's a standard group -> use the full range.
            
            if len(valid_calculation_rows) > 0:
                # Construct comma-separated references (e.g., "K5,K8,K9")
                refs_K = ",".join([f"{col_K_letter}{r}" for r in valid_calculation_rows])
                refs_L = ",".join([f"{col_L_letter}{r}" for r in valid_calculation_rows]) # For CF range if needed
                refs_N = ",".join([f"{col_N_letter}{r}" for r in valid_calculation_rows])
                refs_Q = ",".join([f"{col_Q_letter}{r}" for r in valid_calculation_rows])
                
                formulas = {
                    11: f"=AVERAGE({refs_K})",
                    12: f"=STDEV({refs_K})",
                    13: f"=COUNT({refs_K})",
                    14: f"=AVERAGE({refs_N})",
                    15: f"=STDEV({refs_N})",
                    16: f"=COUNT({refs_N})",
                    17: f"=AVERAGE({refs_Q})"
                }
            else:
                # Standard Group: Use continuous Range
                formulas = {
                    11: f"=AVERAGE({col_K_letter}{data_start_row}:{col_K_letter}{data_end_row})",
                    12: f"=STDEV({col_K_letter}{data_start_row}:{col_K_letter}{data_end_row})",
                    13: f"=COUNT({col_K_letter}{data_start_row}:{col_K_letter}{data_end_row})",
                    14: f"=AVERAGE({col_N_letter}{data_start_row}:{col_N_letter}{data_end_row})",
                    15: f"=STDEV({col_N_letter}{data_start_row}:{col_N_letter}{data_end_row})",
                    16: f"=COUNT({col_N_letter}{data_start_row}:{col_N_letter}{data_end_row})",
                    17: f"=AVERAGE({col_Q_letter}{data_start_row}:{col_Q_letter}{data_end_row})"
                }
            
            for c, f in formulas.items():
                cell = new_ws.cell(row=calc_row_mid, column=c, value=f)
                cell.font = bold_font
                
                # Apply number format for 3 decimals
                if c in [11, 12, 14, 15, 17]:
                    cell.number_format = THREE_DECIMAL_FORMAT
            
            # -------------------- CONDITIONAL FORMATTING FOR STDEV COLUMNS (L and O) --------------------
            if stdev_threshold is not None:
                threshold_str = str(stdev_threshold)
                
                # 1. CF for the Group Data Rows (Numeric values)
                # Apply to the whole group range regardless of He/CO2 logic (visualizing high values is useful everywhere)
                data_range_L = f"{col_L_letter}{data_start_row}:{col_L_letter}{data_end_row}"
                data_range_O = f"{col_O_letter}{data_start_row}:{col_O_letter}{data_end_row}"
                
                new_ws.conditional_formatting.add(data_range_L, CellIsRule(operator="greaterThan", formula=[f"{threshold_str}"], fill=fill_error))
                new_ws.conditional_formatting.add(data_range_O, CellIsRule(operator="greaterThan", formula=[f"{threshold_str}"], fill=fill_error))

                # 2. CF for the Calculation Row
                new_ws.conditional_formatting.add(f"{col_L_letter}{calc_row_mid}", CellIsRule(operator="greaterThan", formula=[threshold_str], fill=fill_error))
                new_ws.conditional_formatting.add(f"{col_O_letter}{calc_row_mid}", CellIsRule(operator="greaterThan", formula=[threshold_str], fill=fill_error))
            # -------------------- END CF --------------------

        cur_row = calc_row_mid + 2

        # --- Insert gray divider after last reference group ---
        if group_key == last_ref_group:
            cur_row += 2  # 2 blank rows above divider
            for i in range(1, len(new_headers) + 1):
                divider_cell = new_ws.cell(row=cur_row, column=i, value=None)
                divider_cell.fill = gray_fill
            cur_row += 1
            for i in range(1, len(new_headers) + 1):
                divider_cell = new_ws.cell(row=cur_row, column=i, value=None)
                divider_cell.fill = gray_fill
            cur_row += 3  # 2 blank rows below divider

    # --- Cleanup ---
    try:
        for r_idx in list(new_ws.row_dimensions.keys()):
            new_ws.row_dimensions[r_idx].outlineLevel = 0
    except Exception:
        pass

    for s in wb.worksheets:
        s.sheet_view.tabSelected = False
    new_ws.sheet_view.tabSelected = True
    wb.active = wb.index(new_ws)
    new_ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    wb.save(file_path)
    print(f"✅ Step 5: Group sheet '{new_sheet_name}'")