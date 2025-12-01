
from PySide6.QtWidgets import QWidget

from helpers.file_io import (
    ask_for_multiple_files,
    show_info,
    show_warning,
)


def run_template_workflow(parent: QWidget) -> None:
    files = ask_for_multiple_files(parent, "Select 2 Template Files to Compare")
    if len(files) != 2:
        show_warning(parent, "Need Two Files", "Please select exactly two template files.")
        return

    file1, file2 = files

    show_info(
        parent,
        "Field Report",
        f"This is where you'd compare:\n\n{file1}\n{file2}",
    )
