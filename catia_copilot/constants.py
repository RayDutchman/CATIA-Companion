"""
Application-wide constants for CATIA Copilot.

All magic strings, column definitions, and configuration values are kept here
so they can be imported by any module without circular-dependency risk.
"""

import re

# ---------------------------------------------------------------------------
# Application info
# ---------------------------------------------------------------------------

APP_NAME    = "CATIA Copilot"
APP_VERSION = "1.4.1"
APP_DATE    = "2026-04-24"
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
  • 紧固件快速装配（VBA 宏批量装配）
  • 托板螺母快速装配（VBA 宏批量装配）
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
STYLESHEET_RELATIVE_PATH = "catia_copilot/ui/style.qss"

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
# "物料编码", "材料", "重量" 这三个属性在新建图纸和刷新图纸的宏也会用到，修改时请注
# 意保持一致
# ---------------------------------------------------------------------------

PRESET_USER_REF_PROPERTIES: list[str] = [
    "零件类型", "设计状态", "材料", "重量",
    "物料编码", "存货类别", "规格型号", "备注",
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

# Sentinel internal column name for the row-number column (always first, read-only)
BOM_ROW_NUMBER_COLUMN: str = "#"

# Columns that are structural / derived – shown read-only in the edit table
BOM_READONLY_COLUMNS: frozenset[str] = frozenset({"#", "Level", "Type", "Filename", "Filepath", "Quantity"})

# Standard BOM columns that can be hidden in the edit dialog
# These are properties that users might not need to see/edit
BOM_HIDEABLE_COLUMNS: list[str] = ["Nomenclature", "Revision", "Definition", "Source"]

# Column order used in the BOM edit dialog (internal names)
BOM_EDIT_COLUMN_ORDER: list[str] = [
    "Level", "Type", "Filename", "Part Number", "Quantity",
    "Nomenclature", "Revision", "Definition", "Source",
]

# Internal column name → Chinese display name
BOM_COLUMN_DISPLAY_NAMES: dict[str, str] = {
    "#":            "#",
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
# BOM thumbnail display
# ---------------------------------------------------------------------------

# Maximum width and height (pixels) for the thumbnail shown in the BOM
# right-click context menu.  Images larger than this are scaled down
# proportionally; images smaller than this are shown at their original size.
BOM_THUMBNAIL_MAX_SIZE: int = 130

# ---------------------------------------------------------------------------
# User-defined property dropdown options
#
# 在此字典中为任意预设用户自定义属性指定可选值列表，该属性在"BOM属性补全"对话框中
# 将自动渲染为下拉框（QComboBox）而非自由文本输入框。
# 字典键必须是 PRESET_USER_REF_PROPERTIES 或用户添加的自定义列中的属性名称。
# 若不希望某属性使用下拉框，只需不在此处添加该属性（或将其删除）即可。
# ---------------------------------------------------------------------------

PRESET_USER_REF_PROPERTY_OPTIONS: dict[str, list[str]] = {
    "零件类型": ["结构件", "外购件", "标准件", "工装"],
    "设计状态": ["设计中", "已定型", "已废弃"],
    "存货类别": ["自制", "外购", "委外"],
}

# ---------------------------------------------------------------------------
# Part Number validation
# ---------------------------------------------------------------------------

# Rejects control characters, non-ASCII characters, and Windows filename-
# forbidden characters  \ / : * ? " < > |
PART_NUMBER_VALID_PATTERN: re.Pattern = re.compile(
    r'^[^\x00-\x1f\x7f-\U0010ffff\\/:*?"<>|]*$'
)
