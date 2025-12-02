"""
Excel formatting helpers used by LEDATools (XlsxWriter)

This module implements utilities to save pandas DataFrames to Excel with simple formatting
using the XlsxWriter engine via pandas' ExcelWriter. Functions are minimal and provide sensible
defaults for header formatting and column autosizing.

Functions:
- save_formatted_excel(df: pd.DataFrame, path: Path, sheet_name: str = "Sheet1", header_style: dict = None)
- autosize_columns(worksheet, df)

Note: This file is intentionally free of PySide6 or GUI references - it provides pure
I/O helpers for writing stylized Excel files for consumption by the user's workflows.
"""

from __future__ import annotations

from pathlib import Path
from helpers.constants import FIELD_COL_IMPORT
from typing import Optional, Dict, Any, Iterable, List

import pandas as pd


def autosize_columns(
    worksheet,
    df: pd.DataFrame,
    column_widths: Optional[dict] = None,
    default_max: Optional[int] = 20,
    add_autofilter_padding: bool = False,
    autofilter_padding: int = 3,
) -> None:
    """Set column widths on the given XlsxWriter worksheet based on df contents.

    - worksheet: XlsxWriter worksheet (obtained from ExcelWriter.sheets[sheet_name])
    - df: pandas DataFrame written to the worksheet
    - add_autofilter_padding: if True, adds `autofilter_padding` to each column width to account for the filter dropdown arrow
    - autofilter_padding: number of extra characters to add when add_autofilter_padding is True
    """
    column_widths = column_widths or {}
    for idx, col in enumerate(df.columns):
        # Start with header length
        max_len = len(str(col))
        # Check values
        # Use try/except to tolerate Series with complex types
        try:
            for v in df[col].fillna(""):
                l = len(str(v))
                if l > max_len:
                    max_len = l
        except Exception:
            # If something goes wrong computing lengths, ignore and leave header width only
            pass
        # Add padding
        width = max_len + 2
        # If caller passed a specific width for this column name, use it
        if col in column_widths:
            try:
                custom_width = int(column_widths[col])
                # Apply optional autosize padding for autofilter arrow
                if add_autofilter_padding:
                    custom_width += int(autofilter_padding)
                worksheet.set_column(idx, idx, custom_width)
                continue
            except Exception:
                # ignore invalid width values, continue with autosize
                pass
        # If autofilter arrow should be accounted for, add a small padding
        if add_autofilter_padding:
            width += int(autofilter_padding)
        # Otherwise, cap width at default_max if provided (apply capping after padding)
        if default_max is not None:
            width = min(width, int(default_max))
        # XlsxWriter uses zero-based column indices
        worksheet.set_column(idx, idx, width)


def save_formatted_excel(
    df: pd.DataFrame,
    path: Path,
    sheet_name: str = "Sheet1",
    header_style: Optional[dict] = None,
    freeze_header: bool = False,
    autofilter: bool = False,
    column_widths: Optional[dict[str, int]] = None,
    default_width_max: Optional[int] = None,
        autofilter_padding: int = 4,
    all_borders: bool = False,
    header_column_colors: Optional[Dict[str, Any]] = None,
    # Removed validate_import_column & import_validation_options — use `validation_columns` instead
    validation_columns: Optional[Dict[str, Iterable[str]]] = None,
    conditional_format_rules: Optional[List[Dict[str, Any]]] = None,
    hide_columns: Optional[List[str]] = None,
) -> Path:
    """Save a DataFrame as an Excel file with basic header formatting and autosized columns.

    Parameters:
    - df: DataFrame to save
    - path: path of the file to write
    - sheet_name: worksheet name (default: "Sheet1")
    - header_style: dictionary with optional style keys for the header row. Supported keys:
        - font_name: font family for header (default: 'Calibri')
        - font_size: integer for header font size (default: 11)
        - bold: bool header bold (default: True)
        - header_alignment: 'left'|'center'|'right' (default: 'center')
        - column_widths: Optional dict mapping column name (str) to integer width to set explicitly
        - default_width_max: Optional max width (int) applied to any column not listed in column_widths. If None, no cap (autosize).
        - autofilter_padding: padding (int) applied to each column's width when `autofilter=True` to account for the autofilter dropdown arrow.
        - header_column_colors: Optional mapping of column name to either a hex color string or a dict like `{'bg_color': '#RRGGBB', 'font_color': '#FFFFFF'}` to override header color per column.
        - validation_columns: Optional mapping of column name -> iterable of allowed values (e.g., {'Import': ['yes', 'no'], 'ColumnB': ['A', 'B']}).
        - conditional_format_rules: Optional list of conditional format rule dicts. Each dict should include:
            - 'columns': column name or list of column names to apply the rule to
            - 'type': 'cell' (others are passed through), default 'cell'
            - 'criteria': criteria string (e.g., '==')
            - 'value': value for the criteria (string or number)
            - 'format': a dict with format keys (bg_color, font_color, bold, border) to be passed to XlsxWriter.add_format
            - Optional 'first_row' and 'last_row' integers (0-based); default applies to data rows only (first row = 1)
        - hide_columns: Optional list of column names to hide in the saved Excel workbook. Case-insensitive matching is used; missing columns are ignored.
        - validation_columns: Optional mapping of column name -> iterable of allowed values (e.g., {'Import': ['yes', 'no'], 'ColumnB': ['A', 'B']}). These take precedence over validate_import_column.

    Returns the path that was written.
    """
    if not isinstance(df, pd.DataFrame):
        raise TypeError("df must be a pandas DataFrame")

    header_style = header_style or {}
    font_name = header_style.get("font_name", "Calibri")
    font_size = header_style.get("font_size", 11)
    bold = header_style.get("bold", True)
    alignment = header_style.get("header_alignment", "center")
    bg_color = header_style.get("bg_color") or header_style.get("cell_color")
    font_color = header_style.get("font_color") or header_style.get("text_color")

    out_path = Path(path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Use pandas ExcelWriter with XlsxWriter engine
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Header format
        header_format = create_header_format(
            workbook,
            bold=bool(bold),
            font_name=font_name,
            font_size=int(font_size),
            bg_color=bg_color,
            font_color=font_color,
            alignment=alignment,
            border=1 if all_borders else None,
        )

        # Apply header format to the first row (default header style)
        for col_num, _ in enumerate(df.columns):
            worksheet.write(0, col_num, df.columns[col_num], header_format)

        # Optional: apply per-column header color overrides
        if header_column_colors:
            ci_map = {str(k).lower(): v for k, v in header_column_colors.items()}
            for c_idx, col_name in enumerate(df.columns):
                override = ci_map.get(str(col_name).lower())
                if not override:
                    continue
                if isinstance(override, str):
                    override = {'bg_color': override}
                if not isinstance(override, dict):
                    continue
                o_bg = override.get('bg_color') or override.get('cell_color')
                o_font = override.get('font_color') or override.get('text_color')
                o_bold = override.get('bold', bold)
                col_format = create_header_format(
                    workbook,
                    bold=bool(o_bold),
                    font_name=font_name,
                    font_size=int(font_size),
                    bg_color=o_bg,
                    font_color=o_font,
                    alignment=alignment,
                    border=1 if all_borders else None,
                )
                worksheet.write(0, c_idx, str(col_name), col_format)

        # Autosize columns using the helper — allows column_widths and default max
        autosize_columns(
            worksheet,
            df,
            column_widths=column_widths,
            default_max=default_width_max,
            add_autofilter_padding=bool(autofilter),
            autofilter_padding=int(autofilter_padding),
        )

        # Optional: freeze the header row
        if freeze_header:
            freeze_top_row(worksheet)

        # Optional: enable autofilter for the written range
        if autofilter:
            apply_autofilter(worksheet, df)

        # Optional: add borders to all cells in the written DataFrame range
        if all_borders:
            # cell format with border
            cell_border_fmt = workbook.add_format({'border': 1})
            # iterate over DataFrame rows and columns and write values with border format
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    value = df.iat[r, c]
                    excel_r = r + 1  # header is row 0
                    try:
                        if pd.isna(value):
                            # write blank value, preserve type
                            worksheet.write_blank(excel_r, c, None, cell_border_fmt)
                        else:
                            worksheet.write(excel_r, c, value, cell_border_fmt)
                    except Exception:
                        # fallback: write empty string with border
                        worksheet.write(excel_r, c, "", cell_border_fmt)

        # Optional: add data validation per supplied mapping
        mapping: Dict[str, Iterable[str]] = {}
        if validation_columns:
            mapping.update(validation_columns)
        # No longer using validate_import_column / import_validation_options; only honor the explicit mapping

        if mapping:
            # Case-insensitive lookup map for actual DataFrame columns
            col_lookup = {str(c).lower(): i for i, c in enumerate(df.columns)}
            for key, opts in mapping.items():
                if not key:
                    continue
                ci = str(key).lower()
                if ci not in col_lookup:
                    continue
                col_idx = col_lookup[ci]
                opts_list = list(opts) if isinstance(opts, Iterable) and not isinstance(opts, str) else [opts]
                if len(opts_list) == 0:
                    continue
                try:
                    first_row = 1
                    last_row = max(1, df.shape[0])
                    worksheet.data_validation(first_row, col_idx, last_row, col_idx, {
                        'validate': 'list',
                        'source': opts_list,
                    })
                except Exception:
                    # ignore errors and continue
                    continue

        # Optional: hide columns if requested
        if hide_columns:
            try:
                # case-insensitive lookup map
                col_lookup = {str(c).lower(): i for i, c in enumerate(df.columns)}
                for col_name in hide_columns:
                    if not col_name:
                        continue
                    ci = str(col_name).lower()
                    idx = col_lookup.get(ci)
                    if idx is None:
                        continue
                    try:
                        # set column hidden using XlsxWriter options
                        worksheet.set_column(idx, idx, None, None, {'hidden': True})
                    except Exception:
                        # ignore issues with hiding individual columns
                        continue
            except Exception:
                pass

        # Optional: apply conditional formatting rules passed from caller
        if conditional_format_rules:
            for rule in conditional_format_rules:
                try:
                    cols = rule.get('columns') or rule.get('col')
                    if not cols:
                        continue
                    if isinstance(cols, str):
                        cols = [cols]
                    cols = list(cols)
                    rule_type = rule.get('type', 'cell')
                    first_row = rule.get('first_row', 1)
                    last_row = rule.get('last_row', max(1, df.shape[0]))
                    criteria = rule.get('criteria')
                    value = rule.get('value')
                    format_spec = rule.get('format', {}) or {}
                    # Build XlsxWriter format
                    # Normalize format_spec values and convert hex colors
                    spec = {}
                    for k, v in format_spec.items():
                        if v is None:
                            continue
                        if k in ('bg_color', 'fg_color', 'fgcolor', 'font_color', 'fontcolour', 'fontcolour'):
                            norm = _normalize_hex_color(v)
                            if norm:
                                # XlsxWriter expects 'fg_color' or 'font_color'
                                if k in ('bg_color', 'fg_color'):
                                    spec['fg_color'] = norm
                                    spec['bg_color'] = norm
                                else:
                                    spec['font_color'] = norm
                            continue
                        spec[k] = v
                    # If a fill color (fg_color / bg_color) is present, set a solid pattern for visibility
                    if 'fg_color' in spec or 'bg_color' in spec:
                        spec.setdefault('pattern', 1)
                    fmt = workbook.add_format(spec)
                    # Prepare kwargs for conditional_format (works for wildcard and per-column rules)
                    kwargs = {'type': rule_type, 'format': fmt}
                    if rule_type == 'cell':
                        kwargs['criteria'] = criteria
                        if isinstance(value, str) and not (value.startswith('"') and value.endswith('"')):
                            kwargs['value'] = f'"{value}"'
                        else:
                            kwargs['value'] = value
                    elif rule_type == 'formula':
                        # For formula-based conditional formatting, XlsxWriter expects a formula string
                        # under the 'criteria' key (e.g. '=A2=1'). Accept 'value' or 'criteria' from the rule.
                        formula = None
                        if isinstance(criteria, str) and criteria.startswith('='):
                            formula = criteria
                        elif isinstance(value, str) and value.startswith('='):
                            formula = value
                        elif isinstance(value, str) and not value.startswith('='):
                            formula = '=' + value
                        if formula is None:
                            # nothing we can do
                            continue
                        kwargs['criteria'] = formula
                    # If wildcard '*' or 'ALL' requested, apply to full width of the df
                    if len(cols) == 1 and (cols[0] == '*' or cols[0] == 'ALL'):
                        first_col = 0
                        last_col = max(0, df.shape[1] - 1)
                        worksheet.conditional_format(first_row, first_col, last_row, last_col, kwargs)
                        continue
                    for col_name in cols:
                        try:
                            col_idx = df.columns.get_loc(col_name)
                        except Exception:
                            # try case-insensitive match
                            try:
                                lookup = {str(c).lower(): i for i, c in enumerate(df.columns)}
                                col_idx = lookup.get(str(col_name).lower())
                                if col_idx is None:
                                    continue
                            except Exception:
                                continue
                        # kwargs already prepared above
                        # Apply to single column range
                        worksheet.conditional_format(first_row, col_idx, last_row, col_idx, kwargs)
                except Exception:
                    # ignore any malformed rule
                    continue

        # writer will be closed automatically when leaving the context

    return out_path


def _normalize_hex_color(color: Optional[str]) -> Optional[str]:
    """Normalize a hex color string to XlsxWriter-compatible '#RRGGBB' or return None.

    Accepts 'RRGGBB' or '#RRGGBB' and returns '#RRGGBB'.
    """
    if not color:
        return None
    try:
        color = str(color).strip()
        if not color:
            return None
        if color.startswith("#"):
            color = color[1:]
        # Ensure length is 6
        if len(color) != 6:
            return None
        # Validate hex characters
        int(color, 16)
        return f"#{color.upper()}"
    except Exception:
        return None


def create_header_format(
    workbook,
    bold: bool = True,
    font_name: str = "Calibri",
    font_size: int = 11,
    bg_color: Optional[str] = None,
    font_color: Optional[str] = None,
    alignment: str = "center",
    border: Optional[int] = None,
) -> object:
    """Build a XlsxWriter header format.

    Parameters:
    - workbook: XlsxWriter workbook object
    - bold: bool
    - font_name: str
    - font_size: int
    - bg_color: optional hex color like '#RRGGBB' or 'RRGGBB'
    - font_color: optional hex color like '#RRGGBB' or 'RRGGBB'
    - alignment: 'left', 'center', 'right'
    """
    fmt_align = {'left': 'left', 'center': 'center', 'right': 'right'}.get(alignment, 'center')
    fmt = {
        'bold': bool(bold),
        'font_name': font_name,
        'font_size': int(font_size),
        'align': fmt_align,
    }
    bg = _normalize_hex_color(bg_color)
    if bg:
        # Ensure a solid fill pattern and set both fg and bg colors for consistent rendering
        fmt['pattern'] = 1
        fmt['fg_color'] = bg
        fmt['bg_color'] = bg
    fc = _normalize_hex_color(font_color)
    if fc:
        fmt['font_color'] = fc
    if border is not None:
        fmt['border'] = int(border)
    return workbook.add_format(fmt)


def freeze_top_row(worksheet) -> None:
    """Freeze the top row (header) in the given XlsxWriter worksheet.

    This keeps the header visible when scrolling vertically.
    """
    # Freeze panes: freeze top row (row 1) — XlsxWriter expects (row_index, col_index)
    worksheet.freeze_panes(1, 0)


def apply_autofilter(worksheet, df: pd.DataFrame, header_row: int = 0) -> None:
    """Apply an autofilter range to the worksheet based on the DataFrame geometry.

    - worksheet: XlsxWriter worksheet
    - df: pandas DataFrame used to compute the data shape
    - header_row: zero-based header row index
    """
    first_row = header_row
    last_row = header_row + df.shape[0]
    first_col = 0
    last_col = max(0, df.shape[1] - 1)
    worksheet.autofilter(first_row, first_col, last_row, last_col)

# Dropdowns removed — function intentionally deleted to reduce scope. Use custom helpers if needed.

