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
APP_VERSION = "1.1.0"
APP_DATE    = "2026-04-10"
APP_AUTHOR  = "CHEN Weibo"
APP_CONTACT = "thucwb@gmail.com"

ABOUT_TEXT = f"""{APP_NAME} v{APP_VERSION}

一款面向工程团队的 CATIA V5 效率工具。

主要功能：
  • CATDrawing 批量导出 PDF
  • CATPart / CATProduct 批量导出 STEP
  • CATProduct BOM 导出到 Excel
  • BOM 属性在线编辑与回写 CATIA
  • 新建图纸（从模板生成 CATDrawing）
  • 刷新图纸（同步零件属性到图纸参数）
  • CATIA 宏脚本快捷运行
  • 零件模板刷写（添加标准用户自定义属性）
  • 字体文件 / ISO.xml 标准文件一键部署

─────────────────────────────────────────
开发者    {APP_AUTHOR}
联系方式  {APP_CONTACT}
发布日期  {APP_DATE}
─────────────────────────────────────────

\u00a9 2026 {APP_AUTHOR}. 仅供内部使用。"""

# ---------------------------------------------------------------------------
# Default window geometry
# ---------------------------------------------------------------------------

MAIN_WINDOW_DEFAULT_WIDTH  = 480
MAIN_WINDOW_DEFAULT_HEIGHT = 520

# Relative path to the QSS stylesheet (used by main.py entry point)
STYLESHEET_RELATIVE_PATH = "catia_companion/ui/style.qss"

# ---------------------------------------------------------------------------
# Resource file paths (relative to project root / frozen executable directory)
# ---------------------------------------------------------------------------

FONT_FILE_PATH    = "resources/ChangFangSong.ttf"
ISO_XML_FILE_PATH = "resources/ISO.xml"
CRACK_DIR_PATH    = "crack"
APP_ICON_PATH     = "resources/icon.ico"

# ---------------------------------------------------------------------------
# Preset user-defined reference properties
# (used both for CATPart template stamping and as BOM preset custom columns)
# ---------------------------------------------------------------------------

PRESET_USER_REF_PROPERTIES: list[str] = [
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

# ---------------------------------------------------------------------------
# BOM edit / display constants
# ---------------------------------------------------------------------------

# Sentinel value displayed in the Filename cell when a product's backing file
# cannot be resolved via COM (the product is "not found").
FILENAME_NOT_FOUND: str = "未检索到"

# Columns that are structural / derived – shown read-only in the edit table
BOM_READONLY_COLUMNS: frozenset[str] = frozenset({"Level", "Type", "Filename", "Filepath", "Quantity"})

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
    "Filepath":     "完整路径",
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
