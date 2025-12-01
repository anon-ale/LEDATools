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
        FIELD_COL_EMPTY_VALUES: round((null_count / total) * 100, 1) if total > 0 else 0,
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
            continue

        for col in df.columns:
            metrics = profile_column(df[col])
            metrics.update({
                FIELD_COL_FILE: path.name,
                FIELD_COL_COLUMN: col
            })
            results.append(metrics)

        del df  # frees memory

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

    column_widths = {
        FIELD_COL_FILE: 30,
        FIELD_COL_TOP5_DISTINCT: 50
    }

    # Build conditional formatting for Import column
    conditional_formats = [
        {
            'columns': FIELD_COL_IMPORT,
            'type': 'cell',
            'criteria': '==',
            'value': 'Yes',
            'format': {'bg_color': '#C6EFCE', 'font_color': '#006100'},
            'first_row': 1,
            'last_row': max(1, report_df.shape[0]),
        },
        {
            'columns': FIELD_COL_IMPORT,
            'type': 'cell',
            'criteria': '==',
            'value': 'No',
            'format': {'bg_color': '#FFC7CE', 'font_color': '#9C0006'},
            'first_row': 1,
            'last_row': max(1, report_df.shape[0]),
        }
    ]

    # Use formatted Excel writer to create a more readable file
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
            conditional_format_rules=conditional_formats,
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
