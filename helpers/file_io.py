
from pathlib import Path
from typing import Optional, List
import pandas as pd
from openpyxl import load_workbook

from PySide6.QtWidgets import QFileDialog, QMessageBox, QWidget

def read_data_file(path):
    p = Path(path)

    # CSV handling goes here...
    if p.suffix.lower() == ".csv":
        for enc in ["utf-8", "latin1", "windows-1252"]:
            try:
                return pd.read_csv(p, encoding=enc)
            except Exception:
                continue
        return pd.read_csv(p, engine="python", on_bad_lines="skip")

    # Excel handling with fallback
    try:
        return pd.read_excel(p)   # normal read
    except Exception:
        # Fallback: read values only, ignore broken XML styles
        try:
            wb = load_workbook(filename=p, data_only=True)
            sheet = wb[wb.sheetnames[0]]
            data = sheet.values
            cols = next(data)  # first row as header
            return pd.DataFrame(data, columns=cols)
        except Exception as e:
            print(f"Skipped unreadable Excel file: {p.name} ({e})")
            return None

def ask_for_file(parent: QWidget, caption: str = "Select File") -> Optional[str]:
    file_path, _ = QFileDialog.getOpenFileName(
        parent,
        caption,
        "",
        "Data Files (*.xlsx *.xls *.csv)"
    )
    return file_path or None


def ask_for_multiple_files(parent: QWidget, caption: str = "Select File(s)") -> List[str]:
    files, _ = QFileDialog.getOpenFileNames(
        parent,
        caption,
        "",
        "Data Files (*.xlsx *.xls *.csv)"
    )
    return files or []


def ask_for_save_excel(parent: QWidget, caption: str = "Save Output As") -> Optional[str]:
    file_path, _ = QFileDialog.getSaveFileName(
        parent,
        caption,
        "",
        "Excel Files (*.xlsx)"
    )
    if not file_path:
        return None

    path = Path(file_path)
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    return str(path)


def show_info(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.information(parent, title, message)


def show_error(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.critical(parent, title, message)


def show_warning(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.warning(parent, title, message)
