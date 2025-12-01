# Copilot Instructions for LEDATools

## Project Overview
- **LEDA Group Implementation Tools** is a desktop application built with PySide6, providing a GUI for data processing tasks on Excel files.
- The main entry point is `main_app.py`, which launches a window with three primary tools: File Preprocessing, Field Report, and Template Comparator.
- Each tool is implemented in its own module under `tools/` and is invoked via button clicks in the GUI.

## Major Components & Data Flow
- **GUI Layer**: `main_app.py` defines the main window and wires up buttons to workflows in `tools/`.
- **Tool Modules**: Each tool (`file_preprocess.py`, `field_report.py`, `template_comparator.py`) exposes a `run_*_workflow(parent: QWidget)` function, which is called from the GUI.
- **Helpers**: Common dialogs and config management are in `helpers/`:
  - `excel_io.py`: File dialogs for selecting/saving Excel files, info/warning/error popups.
  - `config.py`: Loads/saves persistent settings (e.g., last used directories) in `settings.json`.

## Developer Workflows
- **Dependencies**: Managed via `requirements.txt` (PySide6, pandas, openpyxl).
- **Run App**: Launch with `python main_app.py`.
- **No explicit build/test scripts**; manual testing via GUI is standard.
- **Debugging**: Use print statements or PySide6 error dialogs for runtime issues.

## Project-Specific Patterns
- **All user interactions** (file selection, info/warning/error) use helper functions from `helpers/excel_io.py`.
- **Settings** are always loaded/saved via `helpers/config.py` and stored in `settings.json` at the project root.
- **Excel file operations** use pandas and openpyxl, with file paths always obtained via GUI dialogs.
- **Error Handling**: Show errors to users via dialog popups, not console output.

## Integration Points
- **No external services**; all processing is local to the user's machine.
- **Assets** (e.g., logo) are loaded from the `assets/` directory.

## Example Patterns
- To add a new tool, create a `run_new_tool_workflow(parent: QWidget)` in `tools/`, wire it to a button in `main_app.py`, and use helpers for dialogs.
- To persist user settings, update the dictionary returned by `load_settings()` and call `save_settings()`.

## Key Files & Directories
- `main_app.py`: Main GUI and entry point
- `tools/`: Individual tool workflows
- `helpers/`: Shared dialogs and config
- `assets/`: Static files (e.g., logo)
- `requirements.txt`: Python dependencies

---

**For AI agents:**
- Always use helper functions for file dialogs and popups.
- Follow the workflow pattern: GUI → tool module → helpers.
- Avoid direct file path input; always prompt via dialogs.
- Persist settings only via `helpers/config.py`.
- Reference `main_app.py` for wiring new tools.
