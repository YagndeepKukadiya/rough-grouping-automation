import streamlit as st
import pandas as pd
import re
import math
from numbers import Number
from io import BytesIO
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formula.translate import Translator
from openpyxl.formatting.rule import CellIsRule, FormulaRule
import base64

st.set_page_config(
    page_title="Rough Grouping Automation",
    page_icon="RGA_App_Icon.ico",
    layout="wide"
)

# ------------------------------------------------------
# Helper Functions
# ------------------------------------------------------
def get_group_key(rn):
    match = re.match(r"^([A-Za-z]+)", str(rn))

    if match:
        return match.group(1)

    return str(rn)


def validate_main_file(df):
    required_columns = ['Date', 'R N', 'R D', 'PC', 'Ct.', 'R1', 'R2', 'P. PC', 'P.Ct.', 'PS', 'P1', 'P2']

    # Convert both to uppercase for a perfect, case-insensitive match
    uploaded_cols_upper = {str(col).upper() for col in df.columns}

    # Find missing columns by stripping and lowercasing the required list too
    missing = [req_col for req_col in required_columns
               if str(req_col).strip().upper() not in uploaded_cols_upper]

    if missing:
        raise ValueError(f"Main Data file is missing required column(s): {', '.join(missing)}")

    # Validate Date
    converted_dates = pd.to_datetime(df["Date"], errors="coerce")

    invalid_dates = converted_dates.isna().sum()

    if invalid_dates > 0:
        raise ValueError(f"Found {invalid_dates} invalid date value(s) in Date column.")

    # Replace with parsed dates
    df["Date"] = converted_dates

    # Blank R N check
    blank_rn = (df["R N"].astype(str).str.strip().eq("").sum())

    if blank_rn > 0:
        raise ValueError(
            f"Found {blank_rn} blank R N value(s)."
        )

    # Blank R D check
    blank_rd = (df["R D"].astype(str).str.strip().eq("").sum())

    if blank_rd > 0:
        raise ValueError(f"Found {blank_rd} blank R D value(s).")

    return df


def validate_keyword_file(df):
    if "Keyword" not in df.columns:
        raise ValueError("Keyword file must contain a column named 'Keyword'.")

    keywords = (df["Keyword"].dropna().astype(str).str.strip())

    keywords = keywords[keywords != ""]

    keywords = (keywords.drop_duplicates().tolist())

    if len(keywords) == 0:
        raise ValueError("No keywords found in the Keyword file.")

    return keywords


def keyword_filter(df, keyword):
    escaped = re.escape(keyword)
    # Match keyword ONLY when it's standalone (surrounded by whitespace or string boundaries)
    # NOT preceded or followed by ANY character (including non-word characters like /, -, etc.)
    pattern = rf"(?:^|\s){escaped}(?:\s|$)"
    
    return df[df["R D"].astype(str).str.contains(pattern, regex=True, case=False, na=False)].copy()


def autofit_columns(ws):
    for column in ws.columns:
        max_length = 0
        # Safe way to get the column letter without relying on column[0]
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            if cell.value is not None:
                # 1. Safely check for formulas if value is a string
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                # 2. If it's a float, format it to 2 decimal places for length calculation
                if isinstance(cell.value, float):
                    cell_str = f"{cell.value:,.2f}"
                else:
                    cell_str = str(cell.value)
                
                # 3. Calculate length based on the formatted string
                max_length = max(max_length, len(cell_str))

        # 3. Add padding, ensuring the column doesn't shrink below a minimum width (e.g., 12)
        ws.column_dimensions[column_letter].width = max_length + 5


# ============================================================
# SUMMARY ROW FORMATTING
# ============================================================
def format_summary_block(ws, summary_row, start_col="D", end_col="R", fill_color="92D050"):
    """
    fill_color: LIGHT_GREEN = "92D050" | AMBER = "FFC000" | LIGHT_PINK = "FFCCCC"
    """
    bold_font = Font(name="Arial", size=11, bold=True)

    thin = Side(style="thin", color="000000")
    thick = Side(style="thick", color="000000")
    
    start_col = column_index_from_string(start_col)
    end_col = column_index_from_string(end_col)

    start_row = summary_row
    end_row = summary_row + 2

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row, col)

            # Determine outer border thickness
            left = thick if col == start_col else thin
            right = thick if col == end_col else thin
            top = thick if row == start_row else thin
            bottom = thick if row == end_row else thin

            # Always apply font and borders
            cell.font = bold_font
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

            # Conditional: Only apply background if specified
            if fill_color is not None:
                custom_fill = PatternFill(fill_type="solid", fgColor=fill_color)
                cell.fill = custom_fill
            else:
                cell.fill = PatternFill(fill_type=None)


# ============================================================
# R N SUMMARY BLOCK
# ============================================================
def add_rn_summary(ws, start_row, end_row, summary_row, fill_color=None):
    avg_row = summary_row + 2
    # ========================================================
    # ROW 1 OF SUMMARY BLOCK
    # ========================================================

    ws[f'D{summary_row}'] = f'=SUM(D{start_row}:D{end_row})'
    #ws[f'D{summary_row}'].number_format = '0.00'

    ws[f'E{summary_row}'] = f'=SUM(E{start_row}:E{end_row})'
    ws[f'E{summary_row}'].number_format = '#,##0.00'

    ws[f"F{summary_row}"] = f"=J{summary_row}/E{summary_row}"
    ws[f"F{summary_row}"].number_format = '#,##0.00'

    ws[f"F{summary_row}"] = f"=J{summary_row}/E{summary_row}"
    ws[f"F{summary_row}"].number_format = '#,##0.00'

    ws[f"G{summary_row}"] = f"=J{avg_row}/E{summary_row}"
    ws[f"G{summary_row}"].number_format = '#,##0.00'

    ws[f"H{summary_row}"] = f"=G{summary_row}-F{summary_row}"
    ws[f"H{summary_row}"].number_format = '#,##0.00'

    ws[f'I{summary_row}'] = f'=SUM(I{start_row}:I{end_row})'
    ws[f'I{summary_row}'].number_format = '#,##0.00'

    ws[f'J{summary_row}'] = f'=SUM(J{start_row}:J{end_row})'
    ws[f'J{summary_row}'].number_format = '#,##0.00'

    ws[f'K{summary_row}'] = f'=SUM(K{start_row}:K{end_row})'
    #ws[f'K{summary_row}'].number_format = '#,##0.00'

    ws[f'L{summary_row}'] = f'=SUM(L{start_row}:L{end_row})'
    ws[f'L{summary_row}'].number_format = '#,##0.00'

    ws[f"M{summary_row}"] = f"=L{summary_row}/K{summary_row}"
    ws[f"M{summary_row}"].number_format = '#,##0.00'

    ws[f"N{summary_row}"] = f"=R{summary_row}/L{summary_row}"
    ws[f"N{summary_row}"].number_format = '#,##0.00'

    ws[f"O{summary_row}"] = f"=R{avg_row}/L{summary_row}"
    ws[f"O{summary_row}"].number_format = '#,##0.00'

    ws[f"P{summary_row}"] = f"=O{summary_row}-N{summary_row}"
    ws[f"P{summary_row}"].number_format = '#,##0.00'

    ws[f'Q{summary_row}'] = f'=SUM(Q{start_row}:Q{end_row})'
    ws[f'Q{summary_row}'].number_format = '#,##0.00'

    ws[f'R{summary_row}'] = f'=SUM(R{start_row}:R{end_row})'
    ws[f'R{summary_row}'].number_format = '#,##0.00'

    # ========================================================
    # ROW 3 OF SUMMARY BLOCK
    # ========================================================

    ws[f'D{avg_row}'] = 'AVG.'
    ws[f'D{avg_row}'].alignment = Alignment(horizontal="right")
    ws[f'E{avg_row}'] = f'=E{summary_row}/D{summary_row}'
    ws[f'E{avg_row}'].number_format = '#,##0.00'

    #ws[f'G{avg_row}'] = '%'
    ws[f'H{avg_row}'] = f'=I{summary_row}/J{summary_row}'
    ws[f'H{avg_row}'].number_format = '0.00%'

    ws[f'J{avg_row}'] = f'=J{summary_row}+I{summary_row}'
    ws[f'J{avg_row}'].number_format = '#,##0.00'

    #ws[f'H{avg_row}'] = '%'
    ws[f'L{avg_row}'] = f'=L{summary_row}/E{summary_row}'
    ws[f'L{avg_row}'].number_format = '0.00%'

    #ws[f'O{avg_row}'] = '%'
    ws[f'P{avg_row}'] = f'=Q{summary_row}/R{summary_row}'
    ws[f'P{avg_row}'].number_format = '0.00%'

    ws[f'R{avg_row}'] = f'=R{summary_row}+Q{summary_row}'
    ws[f'R{avg_row}'].number_format = '#,##0.00'

    # ========================================================
    # FORMAT SUMMARY BLOCK
    # ========================================================
    format_summary_block(ws, summary_row, start_col="D", end_col="R", fill_color=fill_color)
    

# ========================================================
# GRAND SUMMARY BLOCK
# ========================================================
def excel_formula(formula_name, column, rows):
    return f"={formula_name}({','.join(f'{column}{r}' for r in rows)})"


def add_grand_summary(ws, rn_summary_rows, summary_row, fill_color="92D050"):

    avg_row = summary_row + 2
    # ========================================================
    # ROW 1 OF GRAND SUMMARY BLOCK
    # ========================================================

    ws[f"D{summary_row}"] = excel_formula("SUM", "D", rn_summary_rows)
    #ws[f'D{summary_row}'].number_format = '0.00'

    ws[f"E{summary_row}"] = excel_formula("SUM", "E", rn_summary_rows)
    ws[f'E{summary_row}'].number_format = '#,##0.00'

    ws[f"F{summary_row}"] = f"=J{summary_row}/E{summary_row}"
    ws[f"F{summary_row}"].number_format = '#,##0.00'

    ws[f"F{summary_row}"] = f"=J{summary_row}/E{summary_row}"
    ws[f"F{summary_row}"].number_format = '#,##0.00'

    ws[f"G{summary_row}"] = f"=J{avg_row}/E{summary_row}"
    ws[f"G{summary_row}"].number_format = '#,##0.00'

    ws[f"H{summary_row}"] = f"=G{summary_row}-F{summary_row}"
    ws[f"H{summary_row}"].number_format = '#,##0.00'

    ws[f"I{summary_row}"] = excel_formula("SUM", "I", rn_summary_rows)
    ws[f'I{summary_row}'].number_format = '#,##0.00'

    ws[f"J{summary_row}"] = excel_formula("SUM", "J", rn_summary_rows)
    ws[f'J{summary_row}'].number_format = '#,##0.00'

    ws[f"K{summary_row}"] = excel_formula("SUM", "K", rn_summary_rows)
    #ws[f'K{summary_row}'].number_format = '#,##0.00'

    ws[f"L{summary_row}"] = excel_formula("SUM", "L", rn_summary_rows)
    ws[f'L{summary_row}'].number_format = '#,##0.00'

    ws[f"M{summary_row}"] = f"=L{summary_row}/K{summary_row}"
    ws[f"M{summary_row}"].number_format = '#,##0.00'

    ws[f"N{summary_row}"] = f"=R{summary_row}/L{summary_row}"
    ws[f"N{summary_row}"].number_format = '#,##0.00'

    ws[f"O{summary_row}"] = f"=R{avg_row}/L{summary_row}"
    ws[f"O{summary_row}"].number_format = '#,##0.00'

    ws[f"P{summary_row}"] = f"=O{summary_row}-N{summary_row}"
    ws[f"P{summary_row}"].number_format = '#,##0.00'

    ws[f"Q{summary_row}"] = excel_formula("SUM", "Q", rn_summary_rows)
    ws[f'Q{summary_row}'].number_format = '#,##0.00'

    ws[f"R{summary_row}"] = excel_formula("SUM", "R", rn_summary_rows)
    ws[f'R{summary_row}'].number_format = '#,##0.00'

    # ========================================================
    # ROW 3 OF GRAND SUMMARY BLOCK
    # ========================================================

    ws[f'D{avg_row}'] = 'AVG.'
    ws[f'D{avg_row}'].alignment = Alignment(horizontal="right")
    ws[f'E{avg_row}'] = f'=E{summary_row}/D{summary_row}'
    ws[f'E{avg_row}'].number_format = '#,##0.00'

    #ws[f'G{avg_row}'] = '%'
    ws[f'H{avg_row}'] = f'=I{summary_row}/J{summary_row}'
    ws[f'H{avg_row}'].number_format = '0.00%'

    ws[f'J{avg_row}'] = f'=J{summary_row}+I{summary_row}'
    ws[f'J{avg_row}'].number_format = '#,##0.00'

    #ws[f'H{avg_row}'] = '%'
    ws[f'L{avg_row}'] = f'=L{summary_row}/E{summary_row}'
    ws[f'L{avg_row}'].number_format = '0.00%'

    #ws[f'O{avg_row}'] = '%'
    ws[f'P{avg_row}'] = f'=Q{summary_row}/R{summary_row}'
    ws[f'P{avg_row}'].number_format = '0.00%'

    ws[f'R{avg_row}'] = f'=R{summary_row}+Q{summary_row}'
    ws[f'R{avg_row}'].number_format = '#,##0.00'

    # ========================================================
    # FORMAT GRAND SUMMARY BLOCK
    # ========================================================
    format_summary_block(ws, summary_row, start_col="D", end_col="R", fill_color=fill_color)

    
def summary_ws_add_grand_summary(ws, rn_summary_rows, summary_row):
    avg_row = summary_row + 2

    # ========================================================
    # ROW 1 OF GRAND SUMMARY BLOCK
    # ========================================================

    ws[f"B{summary_row}"] = excel_formula("SUM", "B", rn_summary_rows)
    #ws[f'B{summary_row}'].number_format = '0.00'

    ws[f"C{summary_row}"] = excel_formula("SUM", "C", rn_summary_rows)
    ws[f'C{summary_row}'].number_format = '#,##0.00'

    ws[f"D{summary_row}"] = f"=H{summary_row}/C{summary_row}"
    ws[f"D{summary_row}"].number_format = '#,##0.00'

    ws[f"E{summary_row}"] = f"=H{avg_row}/C{summary_row}"
    ws[f"E{summary_row}"].number_format = '#,##0.00'

    ws[f"F{summary_row}"] = f"=E{summary_row}-D{summary_row}"
    ws[f"F{summary_row}"].number_format = '#,##0.00'

    ws[f"G{summary_row}"] = excel_formula("SUM", "G", rn_summary_rows)
    ws[f'G{summary_row}'].number_format = '#,##0.00'

    ws[f"H{summary_row}"] = excel_formula("SUM", "H", rn_summary_rows)
    ws[f'H{summary_row}'].number_format = '#,##0.00'

    ws[f"I{summary_row}"] = excel_formula("SUM", "I", rn_summary_rows)
    #ws[f'I{summary_row}'].number_format = '#,##0.00'

    ws[f"J{summary_row}"] = excel_formula("SUM", "J", rn_summary_rows)
    ws[f'J{summary_row}'].number_format = '#,##0.00'

    ws[f"K{summary_row}"] = f"=J{summary_row}/I{summary_row}"
    ws[f"K{summary_row}"].number_format = '#,##0.00'

    ws[f"L{summary_row}"] = f"=P{summary_row}/J{summary_row}"
    ws[f"L{summary_row}"].number_format = '#,##0.00'

    ws[f"M{summary_row}"] = f"=P{avg_row}/J{summary_row}"
    ws[f"M{summary_row}"].number_format = '#,##0.00'

    ws[f"N{summary_row}"] = f"=M{summary_row}-L{summary_row}"
    ws[f"N{summary_row}"].number_format = '#,##0.00'

    ws[f"O{summary_row}"] = excel_formula("SUM", "O", rn_summary_rows)
    ws[f'O{summary_row}'].number_format = '#,##0.00'

    ws[f"P{summary_row}"] = excel_formula("SUM", "P", rn_summary_rows)
    ws[f'P{summary_row}'].number_format = '#,##0.00'

    # ========================================================
    # ROW 3 OF GRAND SUMMARY BLOCK
    # ========================================================

    ws[f'B{avg_row}'] = 'AVG.'
    ws[f'B{avg_row}'].alignment = Alignment(horizontal="right")
    ws[f'C{avg_row}'] = f'=C{summary_row}/B{summary_row}'
    ws[f'C{avg_row}'].number_format = '#,##0.00'

    #ws[f'E{avg_row}'] = '%'
    ws[f'F{avg_row}'] = f'=G{summary_row}/H{summary_row}'
    ws[f'F{avg_row}'].number_format = '0.00%'

    ws[f'H{avg_row}'] = f'=H{summary_row}+G{summary_row}'
    ws[f'H{avg_row}'].number_format = '#,##0.00'

    #ws[f'I{avg_row}'] = '%'
    ws[f'J{avg_row}'] = f'=J{summary_row}/C{summary_row}'
    ws[f'J{avg_row}'].number_format = '0.00%'

    #ws[f'M{avg_row}'] = '%'
    ws[f'N{avg_row}'] = f'=O{summary_row}/P{summary_row}'
    ws[f'N{avg_row}'].number_format = '0.00%'

    ws[f'P{avg_row}'] = f'=P{summary_row}+O{summary_row}'
    ws[f'P{avg_row}'].number_format = '#,##0.00'

    # ========================================================
    # FORMAT GRAND SUMMARY BLOCK
    # ========================================================
    format_summary_block(ws, summary_row, start_col="B", end_col="P", fill_color="FFC000")


def copy_cell_format(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)

    if source_cell.hyperlink:
        target_cell.hyperlink = copy(source_cell.hyperlink)

    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)


def auto_shift_all_formulas(ws, inserted_at_col_idx, amount=3):
    """
    Scans the entire sheet and automatically adjusts cell references inside 
    any formula that was shifted due to column insertion.
    """
    # Loop through every cell that contains data
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            
            # Check if the cell contains a formula string
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                # Only shift formulas that are AT or to the RIGHT of the insertion point
                if col >= inserted_at_col_idx:
                    # Calculate where this cell used to be before insertion
                    old_col_letter = get_column_letter(col - amount)
                    new_col_letter = get_column_letter(col)
                    
                    old_coordinate = f"{old_col_letter}{row}"
                    new_coordinate = f"{new_col_letter}{row}"
                    
                    # Translate the formula references by shifting them to the right
                    try:
                        translated_formula = Translator(cell.value, origin=old_coordinate).translate_formula(new_coordinate)
                        cell.value = translated_formula
                    except Exception:
                        # Skip if there's a highly complex/unsupported formula structure
                        continue


def sanitize_sheet_name(name):
    invalid_chars = r'[:\\/*?\[\]]'
    cleaned = re.sub(invalid_chars, "_", str(name))

    return cleaned[:31]


def copy_range_format(source_ws, target_ws,
                      source_start_row, source_end_row,
                      target_start_row,
                      source_start_col, source_end_col,
                      target_start_col=None):

    if target_start_col is None:
        target_start_col = source_start_col

    for row_offset in range(source_end_row - source_start_row + 1):
        for col_offset in range(source_end_col - source_start_col + 1):
            source_cell = source_ws.cell(source_start_row + row_offset, source_start_col + col_offset)
            target_cell = target_ws.cell(target_start_row + row_offset, target_start_col + col_offset)

            copy_cell_format(source_cell, target_cell)


def create_summary_sheet(wb, source_ws, r_d_col, repeat_header=False):
    if "Summary" in wb.sheetnames:
        del wb["Summary"]

    summary_ws = wb.create_sheet("Summary", 0)

    # ========================================================
    # HEADER
    # ========================================================
    for source_col in range(r_d_col, source_ws.max_column + 1):
        source_cell = source_ws.cell(1, source_col)
        target_cell = summary_ws.cell(1, source_col)
        target_cell.value = source_cell.value

        copy_cell_format(source_cell, target_cell)
        
    current_row = 2

    summary_grand_rows = []

    # ========================================================
    # PROCESS STATE SHEETS
    # ========================================================
    for ws in wb.worksheets:
        if ws.title not in keywords:
            continue

        # ====================================================
        # OPTIONAL HEADER REPEAT
        # ====================================================
        if repeat_header and current_row > 2:
            for col in range(r_d_col, source_ws.max_column + 1):
                source_cell = source_ws.cell(1, col)
                target_cell = summary_ws.cell(current_row, col)

                target_cell.value = source_cell.value

                copy_cell_format(source_cell, target_cell)

            current_row += 1

        avg_row = ws.max_row
        grand_row = avg_row - 2

        state_cell = summary_ws.cell(current_row, r_d_col)

        state_cell.value = ws.title
        state_cell.alignment = Alignment(horizontal="center", vertical="center")
        state_cell.font = Font(bold=True)
        state_cell.fill = PatternFill(fill_type="solid", fgColor="92D050")

        thick = Side(style="thick", color="000000")
        for row in range(current_row, current_row + 3):
            cell = summary_ws.cell(row, r_d_col)

            cell.border = Border(
                left=thick,
                right=thick,
                top=thick if row == current_row else None,
                bottom=thick if row == current_row + 2 else None
            )

        # ====================================================
        # GRAND TOTAL ROW
        # ====================================================
        for source_col in range(r_d_col+1, source_ws.max_column + 1):
            source_value = ws.cell(grand_row, source_col).value

            if source_value not in (None, ""):
                summary_ws.cell(current_row, source_col).value = f"='{ws.title}'!{get_column_letter(source_col)}{grand_row}"

        # ====================================================
        # AVG ROW
        # ====================================================
        for source_col in range(r_d_col+1, source_ws.max_column + 1):
            source_value = ws.cell(avg_row, source_col).value

            if source_value not in (None, ""):
                summary_ws.cell(current_row + 2, source_col).value = f"='{ws.title}'!{get_column_letter(source_col)}{avg_row}"

        # ====================================================
        # APPLY FORMAT PAINTER FROM GRAND SUMMARY
        # ====================================================
        copy_range_format(ws, summary_ws, grand_row, avg_row, current_row, r_d_col+1, source_ws.max_column+1, r_d_col+1)

        summary_grand_rows.append(current_row)

        current_row += 4

    grand_summary_row = current_row + 1

    summary_ws.move_range(f"C1:{get_column_letter(summary_ws.max_column)}{summary_ws.max_row}", rows=0, cols=-2)
    
    summary_ws_add_grand_summary(summary_ws, summary_grand_rows, grand_summary_row)

    # ====================================================  
    # MERGE STATE CELL (COLUMN A)
    # ====================================================
    for summary_grand_row in summary_grand_rows:
        summary_ws.merge_cells(start_row=summary_grand_row, start_column=1, end_row=summary_grand_row + 2, end_column=1)

    # ====================================================
    # Summary Sheet GRAND Summary Rows Verifcation
    # ====================================================
    validation_row = summary_ws.max_row + 3
    source_ws_gs_row = source_ws.max_row - 2 # source_ws_grand_summary_row

    for i in range(3):
        for summary_col, processed_col in zip(range(2, summary_ws.max_column + 1), range(4, source_ws.max_column + 1)):
            summary_letter = get_column_letter(summary_col)
            processed_letter = get_column_letter(processed_col)
            
            cell = summary_ws.cell(grand_summary_row+i, summary_col)
            
            if isinstance(cell.value, str) and cell.value.startswith("="):
                summary_ws.cell(validation_row+i, summary_col).value = (
                    f'=ROUNDDOWN({summary_letter}{grand_summary_row+i}, 2)'
                    f'=ROUNDDOWN(Processed_Data!{processed_letter}{source_ws_gs_row+i}, 2)')
            else:
                summary_ws.cell(validation_row+i, summary_col).value = (
                    f'={summary_letter}{grand_summary_row+i}'
                    f'=Processed_Data!{processed_letter}{source_ws_gs_row+i}')
            
    light_green = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    light_pink = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    validation_range = (f"B{validation_row}:{get_column_letter(summary_ws.max_column)}{validation_row + 2}")

    summary_ws.conditional_formatting.add(validation_range, FormulaRule(formula=[f'ISBLANK({validation_range.split(":")[0]})'], fill=PatternFill(fill_type=None), stopIfTrue=True))
    summary_ws.conditional_formatting.add(validation_range, CellIsRule(operator="equal", formula=["TRUE"], fill=light_green))
    summary_ws.conditional_formatting.add(validation_range, CellIsRule(operator="equal", formula=["FALSE"], fill=light_pink))

    autofit_columns(summary_ws)

# ========================================================================================================
# Formula Evaluation Functions Start

def split_cell_reference(ref):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", ref.upper())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref}")
    return m.group(1), int(m.group(2))

def expand_range(range_ref):
    start, end = range_ref.split(":")
    start_col, start_row = split_cell_reference(start)
    end_col, end_row = split_cell_reference(end)

    from openpyxl.utils import column_index_from_string, get_column_letter

    cells = []

    for row in range(start_row, end_row + 1):
        for col in range(column_index_from_string(start_col), column_index_from_string(end_col) + 1):
            cells.append(f"{get_column_letter(col)}{row}")

    return cells

def extract_references(formula):
    formula = formula.lstrip("=").strip()
    refs = []

    range_pattern = (r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)")

    for c1, r1, c2, r2 in re.findall(range_pattern, formula, re.I):
        refs.extend(expand_range(f"{c1}{r1}:{c2}{r2}"))

    formula = re.sub(range_pattern, "", formula, flags=re.I)
    cell_pattern = (r"(?<![A-Z0-9_])\$?([A-Z]{1,3})\$?(\d+)(?![A-Z0-9_])")
    refs.extend(f"{col.upper()}{row}" for col, row in re.findall(cell_pattern, formula, re.I))

    return list(dict.fromkeys(refs))

def is_number(value):
    return (isinstance(value, Number) and not isinstance(value, bool) and not (isinstance(value, float) and math.isnan(value)))

def split_arguments(text):
    args, current, depth = [], [], 0

    for char in text:
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1

        if char == "," and depth == 0:
            args.append("".join(current).strip())
            current = []
        else:
            current.append(char)

    if current:
        args.append("".join(current).strip())

    return args

def evaluate_formula(formula, values):
    expression = formula.lstrip("=").strip()

    def function_values(content):
        result = []

        for arg in split_arguments(content):
            if ":" in arg:
                refs = expand_range(arg)
                result.extend(values.get(ref.upper()) for ref in refs)
            else:
                result.append(values.get(arg.upper()))

        return [value for value in result if is_number(value)]

    def replace_sum(match):
        return str(sum(function_values(match.group(1))))

    def replace_average(match):
        nums = function_values(match.group(1))
        return str(sum(nums) / len(nums)) if nums else "0"

    expression = re.sub(r"\bSUM\s*\((.*?)\)", replace_sum, expression, flags=re.I)
    expression = re.sub(r"\bAVERAGE\s*\((.*?)\)", replace_average, expression, flags=re.I)

    def replace_cell(match):
        ref = f"{match.group(1).upper()}{match.group(2)}"
        value = values.get(ref, 0)
        return str(value) if is_number(value) else "0"

    expression = re.sub(r"(?<![A-Z0-9_])\$?([A-Z]{1,3})\$?(\d+)(?![A-Z0-9_])", replace_cell, expression, flags=re.I)

    if not re.fullmatch(r"[0-9eE+\-*/().\s]+", expression):
        raise ValueError(f"Unsupported formula: {formula}")

    try:
        return eval(expression, {"__builtins__": {}})
    except ZeroDivisionError:
        return 0

def evaluate_cell(ws, cell_ref, cache=None, evaluating=None):
    if cache is None:
        cache = {}

    if evaluating is None:
        evaluating = set()

    cell_ref = cell_ref.upper()

    if cell_ref in cache:
        return cache[cell_ref]

    if cell_ref in evaluating:
        raise ValueError(f"Circular reference detected: {cell_ref}")

    evaluating.add(cell_ref)

    try:
        value = ws[cell_ref].value

        if not (isinstance(value, str) and value.startswith("=")):
            cache[cell_ref] = value
            return value

        refs = extract_references(value)
        values = {ref: evaluate_cell(ws, ref, cache, evaluating) for ref in refs}
        result = evaluate_formula(value, values)
        cache[cell_ref] = result
        return result

    finally:
        evaluating.remove(cell_ref)
        
# Formula Evaluation Functions End
# ========================================================================================================


def create_filtered_rn_sheet(wb, rn_group_metadata, filter_threshold, repeat_header=False, need_grand_summary=False):
    if "R N < -10%" in wb.sheetnames:
        del wb["R N < -10%"]

    filtered_ws = wb.create_sheet("R N < -10%", 0)

    current_row = 1
    header_copied = False

    filtered_rn_summary_rows = []
    for item in rn_group_metadata:
        r_d_ws = wb[item["sheet"]]
        avg_row = item["avg_row"]

        h_value = evaluate_cell(r_d_ws, f"H{avg_row}")
        p_value = evaluate_cell(r_d_ws, f"P{avg_row}")

        # print(
        #     f"[Sheet={r_d_ws.title}]"
        #     f"[R N={item['r_n']}]"
        #     f"[Cell=H{avg_row}]"
        #     f"[Type={type(h_value).__name__}]"
        #     f"[Value={h_value}]"
        #     )
        # print(
        #     f"[Sheet={r_d_ws.title}]"
        #     f"[R N={item['r_n']}]"
        #     f"[Cell=P{avg_row}]"
        #     f"[Type={type(p_value).__name__}]"
        #     f"[Value={p_value}]"
        #     )

        if h_value is None or p_value is None:
            continue

        if not (h_value < filter_threshold or p_value < filter_threshold):
            continue

        # ====================================================
        # HEADER
        # ====================================================
        if repeat_header or not header_copied:
            for col in range(1, r_d_ws.max_column + 1):
                source_cell = r_d_ws.cell(1, col)
                target_cell = filtered_ws.cell(current_row, col)

                target_cell.value = source_cell.value

                copy_cell_format(source_cell, target_cell)

            current_row += 1

            header_copied = True

        # ====================================================
        # COPY R N DATA + SUMMARY BLOCK
        # ====================================================
        start_row = item["group_start_row"]
        end_row = item["summary_row"] + 2

        for source_row in range(start_row, end_row + 1):
            for col in range(1, r_d_ws.max_column + 1):
                source_cell = r_d_ws.cell(source_row, col)
                target_cell = filtered_ws.cell(current_row, col)

                if (isinstance(source_cell.value, str) and source_cell.value.startswith("=")):
                    target_cell.value = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                else:
                    target_cell.value = source_cell.value

                copy_cell_format(source_cell, target_cell)

            current_row += 1

        filtered_rn_summary_rows.append(current_row-3)
        # ====================================================
        # BLANK ROW BETWEEN GROUPS
        # ====================================================
        current_row += 1
        
        if need_grand_summary:
            add_grand_summary(filtered_ws, filtered_rn_summary_rows, current_row+1, fill_color="FFCCCC")

    autofit_columns(filtered_ws)


# =======================================================================================================================
# Final Output Function
# =======================================================================================================================
def generate_output(main_file, main_df, keywords):
    wb = load_workbook(main_file)

    original_ws = None

    for ws in wb.worksheets:
        headers = [str(cell.value).strip().upper()
                   if cell.value is not None
                   else ""
                   for cell in ws[1]]

        if ("Date".upper() in headers and "R N" in headers and "R D" in headers):
            original_ws = ws
            break

    if original_ws is None:
        raise ValueError("Could not find a worksheet containing Date, R N and R D columns.")

    # ============================================================
    # CREATE PROCESSED DATA SHEET
    # ============================================================
    source_ws = wb.copy_worksheet(original_ws)
    source_ws.title = "Processed_Data"

    # ============================================================
    # INSERT DIFF., D.AMT., P.AMT. COLUMNS after Header R2 & P2
    # ============================================================
    col_num = column_index_from_string('H')
    source_ws.insert_cols(idx=col_num, amount=3)

    # Fix every formula the owner wrote instantly:
    auto_shift_all_formulas(source_ws, inserted_at_col_idx=col_num, amount=3)

    # 1. Loop and copy formatting first
    for row in range(1, source_ws.max_row + 1):
        copy_cell_format(source_ws[f"G{row}"], source_ws[f"H{row}"])
        copy_cell_format(source_ws[f"G{row}"], source_ws[f"I{row}"])
        copy_cell_format(source_ws[f"G{row}"], source_ws[f"J{row}"])

    # 2. Assign text values last
    source_ws["H1"].value = "DIFF. R"
    source_ws["I1"].value = "D.AMT. R"
    source_ws["J1"].value = "P.AMT. R"

    # Insert 3 column before column 'P'
    col_num = column_index_from_string('P')
    source_ws.insert_cols(idx=col_num, amount=3)

    # Fix every formula the owner wrote instantly:
    auto_shift_all_formulas(source_ws, inserted_at_col_idx=col_num, amount=3)

    # 1. Loop and copy formatting first
    for row in range(1, source_ws.max_row + 1):
        copy_cell_format(source_ws[f"O{row}"], source_ws[f"P{row}"])
        copy_cell_format(source_ws[f"O{row}"], source_ws[f"Q{row}"])
        copy_cell_format(source_ws[f"O{row}"], source_ws[f"R{row}"])

    # 2. Assign text values last
    source_ws["P1"].value = "DIFF. P"
    source_ws["Q1"].value = "D.AMT. P"
    source_ws["R1"].value = "P.AMT. P"

    autofit_columns(source_ws)

    header_mapping = {}

    for col_num in range(1, source_ws.max_column + 1):
        header_mapping[str(source_ws.cell(row=1, column=col_num).value).strip()] = col_num

    r_d_col = header_mapping['R D']

    r_diff = header_mapping["DIFF. R"]
    r_d_amt = header_mapping["D.AMT. R"]
    r_p_amt = header_mapping["P.AMT. R"]

    p_diff = header_mapping["DIFF. P"]
    p_d_amt = header_mapping["D.AMT. P"]
    p_p_amt = header_mapping["P.AMT. P"]
    
    # ----------------------------------------------------
    # Add FORMULAS FOR New Columns
    # ----------------------------------------------------
    for row_num in range(2, source_ws.max_row + 1):
        source_ws.cell(row=row_num, column=r_diff).value = f"=G{row_num}-F{row_num}"
        source_ws.cell(row=row_num, column=r_d_amt).value = f"=H{row_num}*E{row_num}"
        source_ws.cell(row=row_num, column=r_p_amt).value = f"=F{row_num}*E{row_num}"
        
        source_ws.cell(row=row_num, column=p_diff).value = f"=O{row_num}-N{row_num}"
        source_ws.cell(row=row_num, column=p_d_amt).value = f"=P{row_num}*L{row_num}"
        source_ws.cell(row=row_num, column=p_p_amt).value = f"=N{row_num}*L{row_num}"
    
    rn_group_metadata = []
    matched_indexes = set()
    for keyword in keywords:
        # Only match records NOT already matched by a previous keyword
        unmatched_df = main_df.loc[~main_df.index.isin(matched_indexes)]
    
        matched = keyword_filter(unmatched_df, keyword)

        matched_indexes.update(matched.index.tolist())

        if matched.empty:
            continue

        sheet_name = sanitize_sheet_name(keyword)

        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

        ws = wb.create_sheet(sheet_name)

        # Copy column widths
        for col_letter, dim in source_ws.column_dimensions.items():
            ws.column_dimensions[col_letter].width = dim.width

        current_row = 1

        # ========================================================
        # R N GROUPS
        # ========================================================
        rn_summary_rows = []

        grouped_rns = []

        matched["Group_Key"] = (matched["R N"].apply(get_group_key))
        
        for group_key, group in matched.groupby("Group_Key"):
            group = group.sort_values("Date")
            grouped_rns.append(
                (group["Date"].min(), group_key, group))

        grouped_rns.sort(key=lambda x: x[0])

        for col_num in range(1, source_ws.max_column + 1):
            source_cell = source_ws.cell(row=1, column=col_num)
            target_cell = ws.cell(row=current_row, column=col_num, value=source_cell.value)
            copy_cell_format(source_cell, target_cell)

        current_row += 1

        for _, group_key, group in grouped_rns:
            if current_row > 2:
                for col_num in range(1, source_ws.max_column + 1):
                    source_cell = source_ws.cell(row=1, column=col_num)
                    target_cell = ws.cell(row=current_row, column=col_num)
                    copy_cell_format(source_cell, target_cell)

                    # Check if the source cell contains a formula
                    if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                        # Translate the formula from row 1 to current_row
                        translated_formula = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                        target_cell.value = translated_formula
                    else:
                        # If it's regular text, just copy the value normally
                        target_cell.value = source_cell.value

                current_row += 1

            group_start_row = current_row

            for _, row_data in group.iterrows():
                source_row_number = row_data.name + 2

                for col_num in range(1, source_ws.max_column + 1):
                    source_cell = source_ws.cell(row=source_row_number, column=col_num)
                    target_cell = ws.cell(row=current_row, column=col_num)
                    copy_cell_format(source_cell, target_cell)

                    # Check if the source cell contains a formula
                    if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                        # Translate the formula from row 1 to current_row
                        translated_formula = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                        target_cell.value = translated_formula
                    else:
                        # If it's regular text, just copy the value normally
                        target_cell.value = source_cell.value

                if source_row_number in source_ws.row_dimensions:
                    ws.row_dimensions[current_row].height = source_ws.row_dimensions[source_row_number].height

                current_row += 1

            group_end_row = current_row - 1

            # ----------------------------------------------------
            # SUMMARY BLOCK
            # ----------------------------------------------------
            summary_start_row = current_row

            add_rn_summary(ws, group_start_row, group_end_row, summary_start_row)

            rn_summary_rows.append(summary_start_row)
            
            rn_group_metadata.append({
                "sheet": ws.title,
                "r_n": group_key,
                "group_start_row": group_start_row,
                "group_end_row": group_end_row,
                "summary_row": summary_start_row,
                "avg_row": summary_start_row + 2
                })

            current_row += 4

        grand_summary_row = current_row + 1
        
        add_grand_summary(ws, rn_summary_rows, grand_summary_row, fill_color="92D050")

        autofit_columns(ws)

    remaining_df = main_df.loc[~main_df.index.isin(matched_indexes)].copy()

    if not remaining_df.empty:
        sheet_name = "Remaining"
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

        ws = wb.create_sheet(sheet_name)

        # Copy column widths
        for col_letter, dim in source_ws.column_dimensions.items():
            ws.column_dimensions[col_letter].width = dim.width

        # Copy header row
        for col_num in range(1, source_ws.max_column + 1):
            source_cell = source_ws.cell(row=1, column=col_num)
            target_cell = ws.cell(row=1, column=col_num)
            copy_cell_format(source_cell, target_cell)

            # Check if the source cell contains a formula
            if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                # Translate the formula from row 1 to current_row
                translated_formula = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                target_cell.value = translated_formula
            else:
                # If it's regular text, just copy the value normally
                target_cell.value = source_cell.value            

        target_row = 2

        # Copy records exactly as original layout
        for _, row_data in remaining_df.iterrows():
            source_row_number = row_data.name + 2

            for col_num in range(1,source_ws.max_column + 1):
                source_cell = source_ws.cell(row=source_row_number, column=col_num)
                target_cell = ws.cell( row=target_row, column=col_num)
                copy_cell_format(source_cell, target_cell)

                # Check if the source cell contains a formula
                if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                    # Translate the formula from row 1 to current_row
                    translated_formula = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                    target_cell.value = translated_formula
                else:
                    # If it's regular text, just copy the value normally
                    target_cell.value = source_cell.value 

            ws.row_dimensions[target_row].height = (source_ws.row_dimensions[source_row_number].height)

            target_row += 1

        autofit_columns(ws)


    # ============================================================
    # Add Grand Summary in Processed Data Sheet
    # ============================================================
    add_rn_summary(source_ws, 2, source_ws.max_row, source_ws.max_row+2, fill_color="FFC000")

    # ============================================================
    # Create Summary Sheet
    # ============================================================
    create_summary_sheet(wb, source_ws, r_d_col, repeat_header=False)

    # ============================================================
    # Create Filtered R N Sheet for Specific % < -0.10
    # ============================================================
    create_filtered_rn_sheet(wb, rn_group_metadata, -0.10, repeat_header=True, need_grand_summary=False)
    
    # ============================================================
    # SAVE FILE
    # ============================================================
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False

    wb.worksheets[0].sheet_view.tabSelected = True
    wb.active = 0

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


# ------------------------------------------------------
# Streamlit UI
# ------------------------------------------------------
@st.cache_data
def get_base64_image(image_path):
    with open(image_path, "rb") as img:
        return base64.b64encode(img.read()).decode()
    
bg_image = get_base64_image("Mine.webp")

st.markdown(
    f"""
    <style>

    .stApp {{
        background:
            linear-gradient(
                135deg,
                rgba(0,0,0,0.80),
                rgba(0,0,0,0.75)
            ),
            url("data:image/jpeg;base64,{bg_image}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}

    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("""
<style>

/* Main Container */
.block-container {
    max-width: 1100px;
    padding-top: 2rem;
    padding-bottom: 2rem;
}

/* Hero Card */
.hero-card {
    background: white;
    border: 1px solid #E5E7EB;
    border-radius: 16px;
    padding: 15px;
    margin-bottom: 20px;
}

.hero-title {
    font-size: 34px;
    font-weight: 700;
    color: #111827;
}

.hero-subtitle {
    font-size: 15px;
    color: #6B7280;
    margin-top: 8px;
}

/* Footer */
.footer-text {
    text-align: center;
    color: #9CA3AF;
    font-size: 12px;
    margin-top: 30px;
}

/* Upload section label */
.section-title {
    font-size: 18px;
    font-weight: 600;
    margin-bottom: 10px;
    color: white;
}

/* ============================================================
   CLEAN WHITE BORDER FILE UPLOADER
   ============================================================ */

/* 1. Add a solid white border to the main container outer box */
[data-testid="stFileUploader"] {
    border: 2px solid #FFFFFF !important;
    border-radius: 16px !important;
    padding: 16px !important;
    background-color: #1F2937 !important; /* Elegant dark slate background to keep default text fully visible */
}

/* 2. Style the inner dropzone section to match smoothly */
[data-testid="stFileUploader"] section {
    background-color: #111827 !important; /* Deeper dark background for contrast */
    border-radius: 12px !important;
    border: 1px dashed rgba(255, 255, 255, 0.2) !important; /* Subtle inner white dash */
}

/* Download button */
.stDownloadButton button {
    width: 100%;
    height: 3rem;
    font-size: 16px;
    font-weight: 600;
}

/* Hide Streamlit native decoration */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    visibility: hidden;
}

</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------
# Header
# ------------------------------------------------------
st.markdown("""
<div class="hero-card">
    <div class="hero-title">
        💎 Rough Grouping Automation
    </div>
    <div class="hero-subtitle">
        Upload the source workbook and keyword file to automatically create grouped R D worksheets while preserving the original formatting.
    </div>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------
# Upload Section
# ------------------------------------------------------
st.markdown('<div class="section-title">Upload Files</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    main_file = st.file_uploader(
        "Main Data Workbook",
        type=["xlsx", "xls"],
        help="Workbook containing Date, R N and R D columns."
    )

with col2:
    keyword_file = st.file_uploader(
        "Keyword File",
        type=["xlsx", "xls"],
        help="Excel file containing keywords in the Keyword column."
    )

# ------------------------------------------------------
# Processing
# ------------------------------------------------------
if main_file and keyword_file:
    try:
        with st.spinner("Reading files..."):
            main_df = pd.read_excel(main_file)
            main_df.columns = main_df.columns.str.strip()
            
            keyword_df = pd.read_excel(keyword_file)
            keyword_df.columns = keyword_df.columns.str.strip()
            
            main_df = validate_main_file(main_df)
            keywords = validate_keyword_file(keyword_df)

        st.success(f"Validation completed successfully • {len(keywords)} keyword(s) found")

        # --------------------------------------------------
        # Summary Cards
        # --------------------------------------------------
        metric1, metric2 = st.columns(2)

        with metric1:
            st.metric(label="Total Records", value=f"{len(main_df):,}")

        with metric2:
            st.metric(label="Keywords Found", value=f"{len(keywords):,}")

        # --------------------------------------------------
        # Generate Workbook
        # --------------------------------------------------
        with st.spinner("Generating workbook..."):
            output_file = generate_output(main_file, main_df, keywords)

        st.success("Workbook generated successfully.")

        # --------------------------------------------------
        # Download
        # --------------------------------------------------
        st.download_button(
            label="⬇ Download Processed Workbook",
            data=output_file,
            file_name="R_D_Grouped_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")

# ------------------------------------------------------
# Footer
# ------------------------------------------------------
st.markdown("""
<div class="footer-text">
    Rough Grouping Automation
</div>
""", unsafe_allow_html=True)
