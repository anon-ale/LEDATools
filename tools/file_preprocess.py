
from typing import Any, Dict
from pathlib import Path

import pandas as pd
from PySide6.QtWidgets import QWidget

from helpers.file_io import (
    ask_for_multiple_files,
    ask_for_save_excel,
    show_info,
    show_error,
    show_warning,
    read_data_file
)
from helpers.config import load_settings, save_settings

def process_df():
    return None

def file_preprocessing(input_paths: list[str], output_path: str) -> None:


    file_name_to_df: Dict[str, pd.DataFrame] = {}

    for file_path in input_paths:
        result = read_data_file(file_path, read_all_sheets=True)
        if result is None:
            continue
        # Get filename (stem) from the path for friendly keys and logging
        file_name = Path(file_path).stem
        if isinstance(result, dict):
            # Multiple sheets: store each sheet keyed by "filename::sheetname"
            for sheet_name, df in result.items():
                if len(df) == 0:
                    continue
                key = f"{file_name.replace(".xlsx","").replace()}__{sheet_name}"
                file_name_to_df[key] = df
            continue
        else:
            df = result
            # single-sheet file: store by filename
            file_name_to_df[file_name] = df



    df = df.dropna(axis=1, how="all")

    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].astype(str).str.strip()

    df.to_excel(output_path, index=False)


def run_file_preprocessing_workflow(parent: QWidget) -> None:
    settings: Dict[str, Any] = load_settings()

    input_paths = ask_for_multiple_files(parent, "Select Files to Process")
    if len(input_paths) == 0:
        show_warning(parent, "No File Selected", "Please choose an input file.")
        return

    output_path = ask_for_save_excel(parent, "Save Cleaned CRM File As")
    if not output_path:
        show_warning(parent, "No Output Selected", "Please choose where to save the output file.")
        return

    try:
        file_preprocessing(input_paths, output_path)
    except Exception as e:
        show_error(parent, "Processing Error", f"An error occurred:\n{e}")
        return

    save_settings(settings)

    show_info(parent, "Success", f"Cleaned CRM file saved to:\n{output_path}")
