
from typing import Any, Dict

import pandas as pd
from PySide6.QtWidgets import QWidget

from helpers.file_io import (
    ask_for_multiple_files,
    ask_for_save_excel,
    show_info,
    show_error,
    show_warning,
)
from helpers.config import load_settings, save_settings


def clean_crm_excel(input_path: str, output_path: str) -> None:
    df = pd.read_excel(input_path)

    df = df.dropna(axis=1, how="all")

    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].astype(str).str.strip()

    df.to_excel(output_path, index=False)


def run_file_preprocessing_workflow(parent: QWidget) -> None:
    settings: Dict[str, Any] = load_settings()

    input_path = ask_for_multiple_files(parent, "Select Files to Process")
    if not input_path:
        show_warning(parent, "No File Selected", "Please choose an input file.")
        return

    output_path = ask_for_save_excel(parent, "Save Cleaned CRM File As")
    if not output_path:
        show_warning(parent, "No Output Selected", "Please choose where to save the output file.")
        return

    try:
        clean_crm_excel(input_path, output_path)
    except Exception as e:
        show_error(parent, "Processing Error", f"An error occurred:\n{e}")
        return

    save_settings(settings)

    show_info(parent, "Success", f"Cleaned CRM file saved to:\n{output_path}")
