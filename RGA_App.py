import streamlit as st
import pandas as pd
import re
from io import BytesIO
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Rough Grouping Automation",
    page_icon="📊",
    layout="wide"
)

# ------------------------------------------------------
# Helper Functions
# ------------------------------------------------------

def validate_main_file(df):
    required_columns = ["Date", "R N", "R D"]

    missing = [
        col for col in required_columns
        if col not in df.columns
    ]

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
        raise ValueError("Keyword file must contain a column named 'R D'.")

    keywords = (df["Keyword"].dropna().astype(str).str.strip())

    keywords = keywords[keywords != ""]

    keywords = (keywords.drop_duplicates().tolist())

    if len(keywords) == 0:
        raise ValueError("No keywords found in the Keyword file.")

    return keywords


def keyword_filter(df, keyword):
    pattern = rf"\b{re.escape(keyword)}\b"

    return df[df["R D"].astype(str).str.contains(pattern, regex=True, case=False, na=False)].copy()


def autofit_columns(ws):
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            try:
                value_length = len(str(cell.value))
                if value_length > max_length:
                    max_length = value_length
            except Exception:
                pass

        ws.column_dimensions[column_letter].width = min(max_length + 3, 60)


def format_workbook(workbook):

    bold_font = Font(bold=True)

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    center_alignment = Alignment(horizontal="center", vertical="center")

    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            is_header = ((str(row[0].value).strip() == "Date"
                          and str(row[1].value).strip() == "R N"
                          and str(row[2].value).strip() == "R D"))

            for cell in row:
                # Center + Middle Align everything
                cell.alignment = center_alignment

                # Format dates
                if (cell.value is not None
                    and hasattr(cell.value, "strftime")):
                    cell.number_format = "Short Date" #"MM/DD/YYYY"

                # Bold headers
                if is_header:
                    cell.font = bold_font

            # Apply borders
            if ws.title == "Original_Data":
                # Border on every populated cell
                for cell in row:
                    if cell.value not in [None, ""]:
                        cell.border = thin_border

            else:
                # Keyword sheets:
                # Border only on group headers and data rows
                # Skip the 4 blank rows between groups

                if not (
                    row[0].value in [None, ""]
                    and row[1].value in [None, ""]
                    and row[2].value in [None, ""]
                ):
                    for cell in row:
                        cell.border = thin_border

        autofit_columns(ws)

    return workbook

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

def sanitize_sheet_name(name):
    invalid_chars = r'[:\\/*?\[\]]'
    cleaned = re.sub(invalid_chars, "_", str(name))

    return cleaned[:31]

def generate_output(main_file, main_df, keywords):

    wb = load_workbook(main_file)

    source_ws = None

    for ws in wb.worksheets:
        headers = [
            str(cell.value).strip()
            if cell.value is not None
            else ""
            for cell in ws[1]
        ]

        if (
            "Date" in headers
            and "R N" in headers
            and "R D" in headers
        ):
            source_ws = ws
            break

    if source_ws is None:
        raise ValueError(
            "Could not find a worksheet containing Date, R N and R D columns."
        )
    
    matched_indexes = set()
    for keyword in keywords:

        matched = keyword_filter(main_df, keyword)

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

        grouped_units = []

        for unit, group in matched.groupby("R N"):
            group = group.sort_values("Date")
            grouped_units.append(
                (group["Date"].min(), unit, group))

        grouped_units.sort(key=lambda x: x[0])

        header_mapping = {}

        for col_num in range(1, source_ws.max_column + 1):
            source_cell = source_ws.cell(row=1, column=col_num)
            target_cell = ws.cell(row=current_row, column=col_num, value=source_cell.value)
            copy_cell_format(source_cell, target_cell)
            header_mapping[str(source_cell.value).strip()] = col_num

        current_row += 1

        for _, unit, group in grouped_units:
            if current_row > 2:
                for col_num in range(1, source_ws.max_column + 1):
                    source_cell = source_ws.cell(row=1, column=col_num)
                    target_cell = ws.cell( row=current_row, column=col_num, value=source_cell.value)
                    copy_cell_format(source_cell, target_cell)

                current_row += 1

            for _, row_data in group.iterrows():
                source_row_number = row_data.name + 2

                for col_num in range(1, source_ws.max_column + 1):
                    source_cell = source_ws.cell(row=source_row_number, column=col_num)
                    target_cell = ws.cell(row=current_row, column=col_num, value=source_cell.value)
                    copy_cell_format(source_cell, target_cell)

                if source_row_number in source_ws.row_dimensions:
                    ws.row_dimensions[current_row].height = source_ws.row_dimensions[source_row_number].height

                current_row += 1

            current_row += 4

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
            target_cell = ws.cell(row=1, column=col_num, value=source_cell.value)

            copy_cell_format(source_cell, target_cell)

        target_row = 2

        # Copy records exactly as original layout
        for _, row_data in remaining_df.iterrows():
            source_row_number = row_data.name + 2

            for col_num in range(1,source_ws.max_column + 1):
                source_cell = source_ws.cell(row=source_row_number, column=col_num)
                target_cell = ws.cell( row=target_row, column=col_num, value=source_cell.value)

                copy_cell_format(source_cell, target_cell)

            ws.row_dimensions[target_row].height = (source_ws.row_dimensions[source_row_number].height)

            target_row += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


# ------------------------------------------------------
# Streamlit UI
# ------------------------------------------------------
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
    padding: 24px;
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
    color: #111827;
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
        📊 Rough Grouping Automation
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
            keyword_df = pd.read_excel(keyword_file)
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
