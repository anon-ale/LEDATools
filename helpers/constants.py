# Field report column names (individual constants)
FIELD_COL_FILE = "File"
FIELD_COL_COLUMN = "FileColumn"
FIELD_COL_INFERRED_TYPE = "InferredType"
FIELD_COL_DISTINCT_COUNT = "UniqueValuesCount"
FIELD_COL_VALUE_COUNT = "ValueCount"
FIELD_COL_EMPTY_VALUES = "EmptyValues%"
FIELD_COL_TOP5_DISTINCT = "Top5UniqueValues"
FIELD_COL_MAX_CHAR_LENGTH = "MaxCharacterLength"
FIELD_COL_IMPORT = "Import"
FIELD_COL_FLAG = "Flag"
# Field report columns

"""
Project-wide constants for LEDATools
"""

# Excel file extensions
EXCEL_EXTENSIONS = [".xlsx"]

# Settings keys
LAST_OPEN_DIR_KEY = "last_open_dir"
LAST_SAVE_DIR_KEY = "last_save_dir"

# Dialog captions
CAPTION_SELECT_EXCEL = "Select Excel File"
CAPTION_SELECT_MULTIPLE_EXCEL = "Select Excel File(s)"
CAPTION_SAVE_OUTPUT = "Save Output As"
CAPTION_SELECT_TEMPLATE = "Select 2 Template Files to Compare"
CAPTION_SELECT_FIELD_SOURCE = "Select Source Excel for Field Generation"

# Info/Warning/Error dialog titles
TITLE_NO_FILE_SELECTED = "No File Selected"
TITLE_NO_OUTPUT_SELECTED = "No Output Selected"
TITLE_NEED_TWO_FILES = "Need Two Files"
TITLE_PROCESSING_ERROR = "Processing Error"
TITLE_FIELD_REPORT = "Field Report"

# Default settings
DEFAULT_SETTINGS = {
    LAST_OPEN_DIR_KEY: "",
    LAST_SAVE_DIR_KEY: "",
}
