from helpers.constants import (
    FIELD_COL_FILE,
    FIELD_COL_COLUMN,
    FIELD_COL_INFERRED_TYPE,
    FIELD_COL_DISTINCT_COUNT,
    FIELD_COL_VALUE_COUNT,
    FIELD_COL_EMPTY_VALUES,
    FIELD_COL_TOP5_DISTINCT,
    FIELD_COL_MAX_CHAR_LENGTH,
    FIELD_COL_IMPORT,
    CAPTION_SELECT_FIELD_SOURCE,
    TITLE_NO_FILE_SELECTED,
    TITLE_PROCESSING_ERROR,
    TITLE_FIELD_REPORT,
    FIELD_COL_FLAG
)

from PySide6.QtWidgets import QWidget, QMessageBox
from pathlib import Path
import pandas as pd

from helpers.file_io import (
    read_data_file,
    ask_for_multiple_files,
    show_info,
    show_warning,
    show_error,
)

from helpers.excel_formatting import save_formatted_excel

def get_next_available_filename(base_path: Path) -> Path:
    """
    If file exists, auto-append _2, _3, ... until available.
    """
    if not base_path.exists():
        return base_path

    stem = base_path.stem  # e.g. "_FieldReport"
    suffix = base_path.suffix  # e.g. ".xlsx"
    directory = base_path.parent

    counter = 2
    while True:
        new_name = f"{stem}_{counter}{suffix}"
        new_path = directory / new_name
        if not new_path.exists():
            return new_path
        counter += 1

def profile_column(series: pd.Series) -> dict:
    """Generate metrics and inferred type for a column."""
    
    total = len(series)
    non_null_series = series.dropna()  # used multiple times
    null_count = total - len(non_null_series)

    # ---------- Top 5 Distinct Values ----------
    top_values = (
        non_null_series.value_counts()
        .head(5)
        .index.astype(str)
        .tolist()
    )
    top5 = "; ".join(top_values)

    # ---------- Inferred Type ----------
    normalized = (
        non_null_series.astype(str)
        .str.strip()
        .str.lower()
    )

    boolean_sets = [
        {"true", "false"},
        {"1", "0"},
        {"yes", "no"}
    ]

    if any(set(normalized.unique()).issubset(valid) for valid in boolean_sets):
        inferred_type = "Boolean"
    elif pd.api.types.is_numeric_dtype(series):
        inferred_type = "Numeric"
    elif pd.api.types.is_datetime64_any_dtype(series):
        inferred_type = "Date"
    else:
        inferred_type = "Text"

    # ---------- Max Character Length ----------
    # Measure length after converting each non-null value to string
    if len(non_null_series) > 0:
        max_length = non_null_series.astype(str).str.len().max()
    else:
        max_length = 0

    # ---------- Output Metrics ----------
    return {
        FIELD_COL_DISTINCT_COUNT: non_null_series.nunique(),
        FIELD_COL_TOP5_DISTINCT: top5,
        FIELD_COL_VALUE_COUNT: len(non_null_series),
        FIELD_COL_EMPTY_VALUES: round((null_count / total), 2) if total > 0 else 0,
        FIELD_COL_INFERRED_TYPE: inferred_type,
        FIELD_COL_MAX_CHAR_LENGTH: max_length,
    }

def field_report_generator(parent: QWidget, input_paths : list[str]):
    first_file_dir = Path(input_paths[0]).parent

    output_path = first_file_dir / "_FieldReport.xlsx"

    # If file exists, ask user whether to overwrite
    if output_path.exists():
        reply = QMessageBox.question(
            parent,
            "Overwrite File?",
            f"{output_path.name} already exists. \n\n Do you want to replace it?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.No:
            output_path = get_next_available_filename(output_path)

    """Reads files one at a time and returns a combined field profile report."""
    results = []

    for file_path in input_paths:
        path = Path(file_path)
        
        if path.stem.lower().startswith("_fieldreport"):
            continue
        try:
            df = read_data_file(path)           # Read ONE file at a time
            if df is None:
                continue
        except Exception as e:
            print(f"Error reading {path}: {e}")

        for col in df.columns:
            metrics = profile_column(df[col])
            metrics.update({
                FIELD_COL_FILE: path.name,
                FIELD_COL_COLUMN: col
            })
            results.append(metrics)

    report_df = pd.DataFrame(results)[[
        FIELD_COL_FILE,
        FIELD_COL_COLUMN,
        FIELD_COL_INFERRED_TYPE,
        FIELD_COL_TOP5_DISTINCT,
        FIELD_COL_DISTINCT_COUNT,
        FIELD_COL_VALUE_COUNT,
        FIELD_COL_EMPTY_VALUES,
        FIELD_COL_MAX_CHAR_LENGTH
    ]]
    
    report_df.insert(2, FIELD_COL_IMPORT, None)
    # Add FLAG column that toggles per-file groups; guard against empty dataframes
    if not report_df.empty and FIELD_COL_FILE in report_df.columns:
        try:
            first_fill = report_df[FIELD_COL_FILE].iloc[0]
        except Exception:
            first_fill = None
        report_df[FIELD_COL_FLAG] = (report_df[FIELD_COL_FILE] != report_df[FIELD_COL_FILE].shift(fill_value=first_fill)).cumsum() % 2
    else:
        # Create an integer column with default 0 when missing data
        report_df[FIELD_COL_FLAG] = 0

    column_widths = {
        FIELD_COL_FILE: 30,
        FIELD_COL_TOP5_DISTINCT: 50
    }

    # helper: convert zero-based column index to Excel letter
    def colnum_to_letter(n):
        result = ""
        while n >= 0:
            result = chr(n % 26 + ord("A")) + result
            n = n // 26 - 1
        return result
    
    row_count = len(report_df)
    col_count = len(report_df.columns)
    last_col_letter = colnum_to_letter(col_count - 1)
    first_col_letter = colnum_to_letter(0)

    # Column containing the flag
    # Use report_df's columns and constant for Flag
    try:
        flag_column_letter = colnum_to_letter(report_df.columns.get_loc(FIELD_COL_FLAG))
    except Exception:
        flag_column_letter = None

    # Build conditional formatting for Import column and Flag-based row coloring
    conditional_formats = [
        {
            'columns': FIELD_COL_IMPORT,
            'type': 'cell',
            'criteria': '==',
            'value': 'Yes',
            'format': {'bg_color': "#7CDA8E", 'font_color': '#006100'},
            'first_row': 1,
            'last_row': max(1, report_df.shape[0]),
        },
        # placeholder for flag-based full-row rule; the formula will be derived below when the Flag column is present
        {
            'columns': FIELD_COL_IMPORT,
            'type': 'cell',
            'criteria': '==',
            'value': 'No',
            'format': {'bg_color': "#DD8989", 'font_color': '#9C0006'},
            'first_row': 1,
            'last_row': max(1, report_df.shape[0]),
        },
        {
            'columns': ['*'],
            'type': 'formula',
            'criteria': '',  # populated below once we have the Flag column index
            'format': {'bg_color': "#C8E3EC"},
            'first_row': 1,
            'last_row': max(1, report_df.shape[0]),
        }
    ]

    # Use formatted Excel writer to create a more readable file
    # Example: hide Import column in the final workbook
    hide_cols = [FIELD_COL_FLAG]

    # Populate the formula for the Flag-based full-row rule (apply across all columns)
    try:
        if flag_column_letter:
            first_data_row = 2  # Excel row index (1-based) for first data row, since header is row 1
            conditional_formats[-1]['criteria'] = f'=${flag_column_letter}{first_data_row}=1'
        else:
            # remove wildcard rule if flag column absent
            conditional_formats = [r for r in conditional_formats if not (r.get('type') == 'formula' and r.get('columns') == ['*'])]
    except Exception:
        # remove wildcard rule if flag column absent
        conditional_formats = [r for r in conditional_formats if not (r.get('type') == 'formula' and r.get('columns') == ['*'])]

    save_formatted_excel(
        report_df,
        output_path,
        sheet_name="FieldReport",
        header_style={
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "header_alignment": "center",
            "bg_color": "#0B2763",
            "font_color": "#FFFFFF",
        },
        freeze_header=True,
        autofilter=True,
        column_widths=column_widths,
            all_borders=True,
            header_column_colors={FIELD_COL_IMPORT: {'bg_color': "#524B4B", 'font_color': '#FFFFFF'}},
            validation_columns={FIELD_COL_IMPORT: ['Yes', 'No']},
            hide_columns=hide_cols,
            conditional_format_rules=conditional_formats,
            # Format percentage-like columns as percentages for readability
            value_formats={FIELD_COL_EMPTY_VALUES: 'percent0'},
    )
    return output_path


def run_field_report_workflow(parent: QWidget) -> None:
    input_paths = ask_for_multiple_files(parent, CAPTION_SELECT_FIELD_SOURCE)
    if len(input_paths) == 0:
        show_warning(parent, TITLE_NO_FILE_SELECTED, "Please choose files.")
        return

    
    try:
        output_path = field_report_generator(parent, input_paths)
    except Exception as e:
        show_error(parent, TITLE_PROCESSING_ERROR, f"An error occurred:\n{e}")
        return

    show_info(
        parent,
        TITLE_FIELD_REPORT,
        f"Field report saved to:\n{output_path}"
    )
