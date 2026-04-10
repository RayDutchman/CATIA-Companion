"""
Application-wide constants for CATIA Companion.

All magic strings, column definitions, and configuration values are kept here
so they can be imported by any module without circular-dependency risk.
"""

import re

# ---------------------------------------------------------------------------
# Application info
# ---------------------------------------------------------------------------

APP_NAME    = "CATIA Companion"
APP_VERSION = "1.0.0"
APP_DATE    = "2026-04-03"
APP_AUTHOR  = "CHEN Weibo"
APP_CONTACT = "thucwb@gmail.com"

ABOUT_TEXT = f"""{APP_NAME} v{APP_VERSION}

A CATIA V5 productivity tool for engineering teams.
Automates drawing conversion, part export, and
installation of CATIA resources.

─────────────────────────────────────────
Developer   {APP_AUTHOR}
Contact     {APP_CONTACT}
Released    {APP_DATE}
─────────────────────────────────────────

\u00a9 2026 {APP_AUTHOR}. For internal use only."""

# ---------------------------------------------------------------------------
# Default window geometry
# ---------------------------------------------------------------------------

DEFAULT_WIDTH  = 320
DEFAULT_HEIGHT = 500

# ---------------------------------------------------------------------------
# Part template properties
# ---------------------------------------------------------------------------

PART_TEMPLATE_PROPERTIES: list[str] = [
    "物料编码", "物料名称", "规格型号",
    "物料来源", "数据状态", "存货类别", "质量", "备注",
]

# ---------------------------------------------------------------------------
# BOM standard columns
# ---------------------------------------------------------------------------

BOM_ALL_COLUMNS: list[str] = [
    "Level", "Type", "Part Number", "Nomenclature",
    "Definition", "Revision", "Source", "Quantity",
]

BOM_DEFAULT_COLUMNS: list[str] = [
    "Level", "Type", "Part Number", "Nomenclature",
    "Definition", "Revision", "Source", "Quantity",
]

# Preset user-defined property columns (stored on CATPart/CATProduct)
BOM_PRESET_CUSTOM_COLUMNS: list[str] = [
    "物料编码", "物料名称", "规格型号",
    "物料来源", "数据状态", "存货类别", "质量", "备注",
]

# ---------------------------------------------------------------------------
# BOM edit / display constants
# ---------------------------------------------------------------------------

# Columns that are structural / derived – shown read-only in the edit table
BOM_READONLY_COLUMNS: frozenset[str] = frozenset({"Level", "Type", "Filename", "Quantity"})

# Column order used in the BOM edit dialog (internal names)
BOM_EDIT_COLUMN_ORDER: list[str] = [
    "Level", "Type", "Filename", "Part Number", "Quantity",
    "Nomenclature", "Revision", "Definition", "Source",
]

# Internal column name → Chinese display name
BOM_COLUMN_DISPLAY_NAMES: dict[str, str] = {
    "Level":        "层级",
    "Type":         "类型",
    "Filename":     "文件名",
    "Part Number":  "零件编号",
    "Nomenclature": "术语（中文名称）",
    "Definition":   "定义",
    "Revision":     "版本",
    "Source":       "源",
    "Quantity":     "数量",
}

# Minimum column widths (Excel character units) for standard BOM columns
BOM_COLUMN_MIN_WIDTHS: dict[str, int] = {
    "Level":        6,
    "Type":         10,
    "Filename":     30,
    "Part Number":  20,
    "Nomenclature": 20,
    "Definition":   20,
    "Revision":     10,
    "Source":       8,
    "Quantity":     8,
}

# Source field: CATIA integer string ↔ Chinese display label
SOURCE_TO_DISPLAY: dict[str, str]  = {"0": "未知", "1": "自制", "2": "外购"}
SOURCE_FROM_DISPLAY: dict[str, str] = {"未知": "0", "自制": "1", "外购": "2"}
SOURCE_OPTIONS: list[str]           = ["未知", "自制", "外购"]

# ---------------------------------------------------------------------------
# Part Number validation
# ---------------------------------------------------------------------------

# Rejects control characters, non-ASCII characters, and Windows filename-
# forbidden characters  \ / : * ? " < > |
PART_NUMBER_VALID_PATTERN: re.Pattern = re.compile(
    r'^[^\x00-\x1f\x7f-\U0010ffff\\/:*?"<>|]*$'
)
