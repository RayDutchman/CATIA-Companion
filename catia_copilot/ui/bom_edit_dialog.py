"""
BOM 编辑对话框模块。

提供：
- BomEditDialog – 可编辑表格，用于完成 BOM 属性并通过 COM 写回 CATIA。
"""

import copy
import os
import logging
import subprocess
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTreeWidgetItem, QHeaderView, QAbstractItemView,
    QComboBox, QCheckBox, QGroupBox, QMessageBox, QApplication,
    QFileDialog, QProgressDialog, QRadioButton, QButtonGroup,
    QMenu, QWidgetAction, QLineEdit, QGridLayout
)
from PySide6.QtGui import QPixmap, QColor
from PySide6.QtCore import Qt, QSettings

from catia_copilot.constants import (
    PRESET_USER_REF_PROPERTIES,
    PRESET_USER_REF_PROPERTY_OPTIONS,
    BOM_EDIT_COLUMN_ORDER,
    BOM_COLUMN_DISPLAY_NAMES,
    BOM_READONLY_COLUMNS,
    BOM_HIDEABLE_COLUMNS,
    BOM_ROW_NUMBER_COLUMN,
    SOURCE_TO_DISPLAY,
    SOURCE_OPTIONS,
    PART_NUMBER_VALID_PATTERN,
    FILENAME_NOT_FOUND,
    BOM_THUMBNAIL_MAX_SIZE,
)
from catia_copilot.catia.bom_collect import collect_bom_rows, flatten_bom_to_summary
from catia_copilot.catia.bom_write import write_bom_to_catia
from catia_copilot.utils import read_catia_thumbnail
from catia_copilot.ui.bom_catia_helpers import (
    _is_catia_com_error,
    _find_catia_doc_by_path,
)
from catia_copilot.ui.bom_widgets import _BomTreeDelegate, _BomTreeWidget, _ITEM_LOCKED_ROLE
from catia_copilot.ui.bom_file_rename_dialog import _FileRenameDialog

logger = logging.getLogger(__name__)


class BomEditDialog(QDialog):
    """Editable BOM table for completing and writing back product properties.

    - 文件名 / 层级 / 类型 / 数量 are read-only (structural).
    - 零件编号 is editable with duplicate-detection.
    - 源 (Source) uses a QComboBox (未知 / 自制 / 外购).
    - Rows sharing the same Part Number are linked and update together.
    - "应用" writes changes back without closing; "完成" writes and closes.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("BOM属性补全")
        self.setMinimumSize(900, 600)
        self.resize(1100, 700)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowMinimizeButtonHint
        )

        # ── Settings ─────────────────────────────────────────────────────────
        # Share ExportBomDialog's custom-column config
        self._export_settings = QSettings("CATIACompanion", "ExportBOMDialog")
        self._last_browse_dir = self._export_settings.value("last_browse_dir", "")

        saved_custom = self._export_settings.value("custom_columns", [])
        if isinstance(saved_custom, str):
            saved_custom = [saved_custom]
        self._custom_columns: list[str] = list(saved_custom)

        # BomEditDialog-specific settings
        self._edit_settings  = QSettings("CATIACompanion", "BomEditDialog")
        saved_visible        = self._edit_settings.value("visible_preset_columns", [])
        if isinstance(saved_visible, str):
            saved_visible = [saved_visible]
        self._visible_preset_cols: list[str] = [
            c for c in saved_visible if c in PRESET_USER_REF_PROPERTIES
        ]

        # Visible hideable standard columns (Nomenclature, Revision, Definition, Source)
        saved_hideable = self._edit_settings.value("visible_hideable_columns", BOM_HIDEABLE_COLUMNS)
        if isinstance(saved_hideable, str):
            saved_hideable = [saved_hideable]
        self._visible_hideable_cols: list[str] = [
            c for c in saved_hideable if c in BOM_HIDEABLE_COLUMNS
        ]

        self._summarize: bool = self._edit_settings.value("summarize", False, type=bool)
        self._summary_include_assemblies: bool = self._edit_settings.value(
            "summary_include_assemblies", False, type=bool
        )
        self._summary_sort_column: str = self._edit_settings.value(
            "summary_sort_column", "Part Number"
        )

        # All custom columns (including all presets) so that pre-loading from
        # CATIA covers every column regardless of current visibility.
        self._all_custom_columns: list[str] = list(dict.fromkeys(
            self._custom_columns + list(PRESET_USER_REF_PROPERTIES)
        ))

        self._show_filepath_col: bool = self._edit_settings.value(
            "show_filepath_column", False, type=bool,
        )
        self._show_filename_col: bool = self._edit_settings.value(
            "show_filename_column", True, type=bool,
        )

        self._columns: list[str] = self._build_visible_columns()

        # ── State ─────────────────────────────────────────────────────────────
        # {original_pn: {col_name: value}}  (canonical data, Source as display label)
        self._canonical_data: dict[str, dict[str, str]] = {}
        # Snapshot at last load/apply for dirty-only write-back
        self._snapshot_data: dict[str, dict[str, str]] = {}
        # {original_pn: {col_name, ...}}  – fields changed since last write-back
        self._modified_keys: dict[str, set[str]] = {}
        # All BOM rows in traversal order
        self._rows: list[dict] = []
        # Guard against re-entrant change handling
        self._is_updating: bool = False
        # Parallel list: self._item_by_row[i] is the QTreeWidgetItem for self._rows[i]
        self._item_by_row: list[QTreeWidgetItem] = []
        # True once BOM has been successfully loaded at least once
        self._bom_loaded: bool = False
        # Raw (hierarchical) BOM rows as returned by collect_bom_rows(); used to
        # reconstruct the display rows when the user toggles the BOM type.
        self._raw_rows: list[dict] = []
        # PN→Items index for fast linked updates (Performance optimization)
        # Maps Part Number to list of QTreeWidgetItems with that PN
        self._pn_to_items: dict[str, list[QTreeWidgetItem]] = {}
        # Column-name → pixel width cache; persists across column visibility toggles
        # so user-adjusted widths survive adding/removing columns.
        self._col_widths: dict[str, int] = {}

        # ── Layout ────────────────────────────────────────────────────────────
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # Source selection
        self._use_active_chk = QCheckBox("使用当前CATIA活动文档（不选择文件）")
        self._use_active_chk.toggled.connect(self._toggle_file_row)
        layout.addWidget(self._use_active_chk)

        file_row = QHBoxLayout()
        self._file_edit       = QLineEdit()
        self._file_edit.setPlaceholderText("选择一个CATProduct文件...")
        self._file_edit.setReadOnly(True)
        self._file_browse_btn = QPushButton("浏览...")
        self._file_browse_btn.clicked.connect(self._browse_file)
        self._load_btn        = QPushButton("加载BOM")
        self._load_btn.clicked.connect(self._load_bom)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(self._file_browse_btn)
        file_row.addWidget(self._load_btn)
        layout.addLayout(file_row)

        # ── BOM type + display options (single compact group) ────────────────
        display_group  = QGroupBox("BOM类型与显示选项")
        display_group.setMinimumHeight(60)  # Prevent height jumping when switching BOM types
        display_layout = QVBoxLayout(display_group)
        display_layout.setSpacing(4)
        display_layout.setContentsMargins(8, 6, 8, 6)

        # Row 1: radio buttons + filepath checkbox on the same line
        bom_type_row = QHBoxLayout()
        self._bom_type_btn_group = QButtonGroup(self)
        self._radio_hierarchical = QRadioButton("层级BOM")
        self._radio_summary_bom  = QRadioButton("汇总BOM")
        if self._summarize:
            self._radio_summary_bom.setChecked(True)
        else:
            self._radio_hierarchical.setChecked(True)
        self._bom_type_btn_group.addButton(self._radio_hierarchical)
        self._bom_type_btn_group.addButton(self._radio_summary_bom)
        self._radio_summary_bom.toggled.connect(self._on_bom_type_changed)
        bom_type_row.addWidget(self._radio_hierarchical)
        bom_type_row.addWidget(self._radio_summary_bom)

        self._summary_opts_widget = QWidget()
        summary_opts_layout = QHBoxLayout(self._summary_opts_widget)
        summary_opts_layout.setContentsMargins(0, 0, 0, 0)
        summary_opts_layout.setSpacing(8)

        self._include_assemblies_chk = QCheckBox("包含产品和部件（子装配体）")
        self._include_assemblies_chk.setToolTip(
            "勾选后，汇总BOM中也会列出产品和部件（子装配体），而不仅限于零件。"
        )
        self._include_assemblies_chk.setChecked(self._summary_include_assemblies)
        self._include_assemblies_chk.toggled.connect(self._on_include_assemblies_toggled)
        summary_opts_layout.addWidget(self._include_assemblies_chk)
        summary_opts_layout.addSpacing(8)
        summary_opts_layout.addWidget(QLabel("排序列:"))
        self._sort_col_combo = QComboBox()
        _sort_cols = list(BOM_EDIT_COLUMN_ORDER) + [
            c for c in PRESET_USER_REF_PROPERTIES if c not in BOM_EDIT_COLUMN_ORDER
        ] + [
            c for c in self._custom_columns
            if c not in BOM_EDIT_COLUMN_ORDER and c not in PRESET_USER_REF_PROPERTIES
        ]
        for col in _sort_cols:
            self._sort_col_combo.addItem(BOM_COLUMN_DISPLAY_NAMES.get(col, col), col)
        sort_saved_idx = self._sort_col_combo.findData(self._summary_sort_column)
        if sort_saved_idx >= 0:
            self._sort_col_combo.setCurrentIndex(sort_saved_idx)
        self._sort_col_combo.currentIndexChanged.connect(self._on_sort_col_changed)
        summary_opts_layout.addWidget(self._sort_col_combo)

        self._summary_opts_widget.setVisible(self._summarize)
        bom_type_row.addWidget(self._summary_opts_widget)
        bom_type_row.addStretch()
        display_layout.addLayout(bom_type_row)

        layout.addWidget(display_group)

        hint = QLabel(
            "层级 / 类型 / 数量 为结构属性，不可编辑，"
            "零件编号可编辑但不能与其他行冲突，"
            "文件名/路径可编辑。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # Preset column visibility checkboxes (2 rows layout with grid for alignment)
        preset_group  = QGroupBox("属性列（勾选以显示）")
        preset_main_layout = QVBoxLayout(preset_group)
        preset_main_layout.setSpacing(8)
        preset_main_layout.setContentsMargins(8, 6, 8, 6)

        # Use QGridLayout for proper alignment and even distribution
        grid_layout = QGridLayout()
        grid_layout.setSpacing(12)
        grid_layout.setColumnStretch(100, 1)  # Add stretch at the end

        self._preset_checkboxes: dict[str, QCheckBox] = {}

        # Row 0: Filename checkbox + 显示完整路径 + hideable standard columns
        col = 0

        # "Filename" is a built-in column but can be toggled like a preset
        fn_cb = QCheckBox(BOM_COLUMN_DISPLAY_NAMES.get("Filename", "Filename"))
        fn_cb.setChecked(self._show_filename_col)
        fn_cb.toggled.connect(self._on_preset_col_toggled)
        grid_layout.addWidget(fn_cb, 0, col)
        self._preset_checkboxes["Filename"] = fn_cb
        col += 1

        # "显示完整路径" follows immediately after the Filename checkbox
        self._filepath_chk = QCheckBox("显示完整路径")
        self._filepath_chk.setToolTip("勾选后文件名列将显示文件完整路径（含目录），而非仅文件名")
        self._filepath_chk.setChecked(self._show_filepath_col)
        self._filepath_chk.toggled.connect(self._on_show_filepath_toggled)
        grid_layout.addWidget(self._filepath_chk, 0, col)
        col += 1

        # Hideable standard columns (Nomenclature, Revision, Definition, Source)
        for col_name in BOM_HIDEABLE_COLUMNS:
            cb = QCheckBox(BOM_COLUMN_DISPLAY_NAMES.get(col_name, col_name))
            cb.setChecked(col_name in self._visible_hideable_cols)
            cb.toggled.connect(self._on_hideable_col_toggled)
            grid_layout.addWidget(cb, 0, col)
            self._preset_checkboxes[col_name] = cb
            col += 1

        # Row 1: Preset user-defined properties (物料编码, 物料名称, etc.)
        col = 0
        for col_name in PRESET_USER_REF_PROPERTIES:
            cb = QCheckBox(col_name)
            cb.setChecked(col_name in self._visible_preset_cols)
            cb.toggled.connect(self._on_preset_col_toggled)
            grid_layout.addWidget(cb, 1, col)
            self._preset_checkboxes[col_name] = cb
            col += 1

        preset_main_layout.addLayout(grid_layout)
        layout.addWidget(preset_group)

        # BOM tree widget (replaces QTableWidget; tree handles expand/collapse natively)
        self._table = _BomTreeWidget()
        _init_headers = self._display_headers()
        self._table.setColumnCount(len(_init_headers))
        self._table.setHeaderLabels(_init_headers)
        hdr = self._table.header()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        hdr.setSectionsMovable(True)
        hdr.setFixedHeight(28)
        self._table.setUniformRowHeights(True)
        self._table.setRootIsDecorated(True)
        self._table.setSortingEnabled(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setIndentation(16)
        self._table.setStyleSheet("QTreeWidget::item { min-height: 24px; }")
        self._table.itemChanged.connect(self._on_item_changed)
        hdr.sectionResized.connect(self._on_section_resized)
        _delegate = _BomTreeDelegate(lambda: self._columns, self._table)
        self._table.setItemDelegate(_delegate)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._on_tree_context_menu)
        layout.addWidget(self._table, 1)

        # Bottom buttons
        btn_row = QHBoxLayout()

        autofit_btn = QPushButton("自适应列宽")
        autofit_btn.setToolTip("根据内容自动调整所有列的宽度")
        autofit_btn.clicked.connect(self._autofit_columns)
        btn_row.addWidget(autofit_btn)

        expand_btn = QPushButton("全部展开")
        expand_btn.setToolTip("展开结构树中的所有节点")
        expand_btn.clicked.connect(self._table.expandAll)
        btn_row.addWidget(expand_btn)

        collapse_btn = QPushButton("全部折叠")
        collapse_btn.setToolTip("折叠结构树中的所有节点")
        collapse_btn.clicked.connect(self._table.collapseAll)
        btn_row.addWidget(collapse_btn)

        self._rename_btn = QPushButton("按零件编号修改文件名")
        self._rename_btn.setEnabled(False)
        self._rename_btn.clicked.connect(self._rename_by_part_number)
        btn_row.addWidget(self._rename_btn)

        self._rename_file_btn = QPushButton("编辑选中文件名/路径")
        self._rename_file_btn.setToolTip("为选中行的文件执行重命名或移动（通过CATIA另存为）")
        self._rename_file_btn.setEnabled(False)
        self._rename_file_btn.clicked.connect(self._rename_selected_file)
        btn_row.addWidget(self._rename_file_btn)
        btn_row.addStretch()

        self._export_btn = QPushButton("导出表格")
        self._export_btn.setToolTip("将当前表格导出为 Excel（.xlsx）或 CSV 文件")
        self._export_btn.setEnabled(False)
        self._export_btn.clicked.connect(self._export_table)
        btn_row.addWidget(self._export_btn)

        self._save_btn   = QPushButton("应用")
        self._save_btn.setEnabled(False)
        self._save_btn.clicked.connect(self._apply_changes)

        self._finish_btn = QPushButton("完成")
        self._finish_btn.setDefault(True)
        self._finish_btn.setEnabled(False)
        self._finish_btn.clicked.connect(self._finish_and_close)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)

        btn_row.addWidget(self._save_btn)
        btn_row.addWidget(self._finish_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    # ── Source toggle ─────────────────────────────────────────────────────────

    def _toggle_file_row(self, use_active: bool) -> None:
        self._file_edit.setEnabled(not use_active)
        self._file_browse_btn.setEnabled(not use_active)

    # ── BOM type toggle ───────────────────────────────────────────────────────

    def _on_bom_type_changed(self, summary_checked: bool) -> None:
        self._summarize = summary_checked
        self._edit_settings.setValue("summarize", summary_checked)
        self._summary_opts_widget.setVisible(summary_checked)
        # If BOM is already loaded, re-derive display rows from the raw rows and repopulate
        if self._raw_rows:
            self._rows = (
                flatten_bom_to_summary(
                    self._raw_rows,
                    include_assemblies=self._summary_include_assemblies,
                    sort_column=self._summary_sort_column or None,
                )
                if summary_checked else self._raw_rows
            )
            self._rebuild_columns_and_repopulate()

    def _on_include_assemblies_toggled(self, checked: bool) -> None:
        self._summary_include_assemblies = checked
        self._edit_settings.setValue("summary_include_assemblies", checked)
        # Rebuild summary display if BOM is loaded and summary mode is active
        if self._summarize and self._raw_rows:
            self._rows = flatten_bom_to_summary(
                self._raw_rows,
                include_assemblies=checked,
                sort_column=self._summary_sort_column or None,
            )
            # When assemblies are included show the Type column; otherwise hide it
            self._rebuild_columns_and_repopulate()

    def _on_sort_col_changed(self, _index: int) -> None:
        col = self._sort_col_combo.currentData()
        if col:
            self._summary_sort_column = col
            self._edit_settings.setValue("summary_sort_column", col)
            # Re-sort the currently displayed summary rows if applicable
            if self._summarize and self._raw_rows:
                self._rows = flatten_bom_to_summary(
                    self._raw_rows,
                    include_assemblies=self._summary_include_assemblies,
                    sort_column=col,
                )
                self._populate_table()

    # ── Table helpers ─────────────────────────────────────────────────────────

    def _autofit_columns(self) -> None:
        """Resize all columns to fit their content, with a minimum width."""
        # QTreeWidget has resizeColumnToContents(int) not resizeColumnsToContents()
        min_width = 60
        for col_idx, col_name in enumerate(self._columns):
            self._table.resizeColumnToContents(col_idx)
            if self._table.columnWidth(col_idx) < min_width:
                self._table.setColumnWidth(col_idx, min_width)
            # Update the cache so subsequent column-visibility toggles keep these widths
            self._col_widths[col_name] = self._table.columnWidth(col_idx)

    def _on_section_resized(self, logical_index: int, _old_size: int, new_size: int) -> None:
        """Cache the new width whenever the user (or code) resizes a column."""
        if logical_index < len(self._columns):
            self._col_widths[self._columns[logical_index]] = new_size

    def _rebuild_columns_and_repopulate(self) -> None:
        """Rebuild the visible column list, update headers, and repopulate if rows are loaded."""
        # --- Immutable width snapshot (used for both anchor and restoration) ---
        #
        # Take this snapshot before any Qt tree/header operations.  The snapshot
        # merges the persistent cache (gives widths for previously-hidden columns)
        # with the live columnWidth() values (authoritative for visible columns),
        # so both anchor computation and width restoration read from a stable copy
        # rather than from self._col_widths which may be modified mid-rebuild.
        width_snapshot: dict[str, int] = dict(self._col_widths)
        for col_idx, col_name in enumerate(self._columns):
            w = self._table.columnWidth(col_idx)
            width_snapshot[col_name] = w
            self._col_widths[col_name] = w  # keep persistent cache current

        # --- Anchor-column horizontal scroll bookkeeping ---
        # We want the leftmost visible column to stay on screen after columns are
        # added/removed.  A raw pixel hscroll value is not stable across column
        # count changes because Qt may reset the scroll bar, and because removing
        # columns to the LEFT of the viewport shifts the pixel positions of all
        # remaining visible columns.
        #
        # Strategy: identify which column name sits at the left viewport edge
        # (and how many pixels into that column) BEFORE the rebuild, then
        # recompute the target hscroll from column widths AFTER the rebuild.
        old_columns = list(self._columns)
        old_hscroll = self._table.horizontalScrollBar().value()

        anchor_col_name: str | None = None  # column at the left viewport edge
        anchor_offset: int = 0              # pixels into that column

        x = 0
        for col_name in old_columns:
            w = width_snapshot[col_name]    # use snapshot, not live columnWidth
            if x + w > old_hscroll:
                anchor_col_name = col_name
                anchor_offset = old_hscroll - x
                break
            x += w

        vscroll = self._table.verticalScrollBar().value()

        self._columns = self._build_visible_columns()
        if self._rows:
            self._populate_table()  # sets column count and headers internally
            # Restore widths from the immutable snapshot; auto-fit only new
            # columns that have never been seen before.
            for col_idx, col_name in enumerate(self._columns):
                if col_name in width_snapshot:
                    self._table.setColumnWidth(col_idx, width_snapshot[col_name])
                else:
                    self._table.resizeColumnToContents(col_idx)

            # Compute new hscroll using the anchor column's new pixel position.
            new_hscroll = 0
            if anchor_col_name is not None:
                # Try to find the anchor column in the new layout
                x = 0
                found = False
                for col_idx, col_name in enumerate(self._columns):
                    if col_name == anchor_col_name:
                        new_hscroll = x + anchor_offset
                        found = True
                        break
                    x += self._table.columnWidth(col_idx)

                if not found:
                    # Anchor column was hidden; scroll to the first new column
                    # that followed the anchor in the old layout (the next
                    # surviving column to the right of the removed one).
                    old_col_order = {c: i for i, c in enumerate(old_columns)}
                    anchor_old_idx = old_col_order.get(anchor_col_name, -1)
                    x = 0
                    for col_idx, col_name in enumerate(self._columns):
                        if old_col_order.get(col_name, -1) > anchor_old_idx:
                            new_hscroll = x
                            break
                        x += self._table.columnWidth(col_idx)
                    else:
                        # All remaining columns are to the left of where the
                        # anchor was; scroll to the end.
                        new_hscroll = self._table.horizontalScrollBar().maximum()

            new_hscroll = max(0, min(new_hscroll, self._table.horizontalScrollBar().maximum()))
            self._table.verticalScrollBar().setValue(vscroll)
            self._table.horizontalScrollBar().setValue(new_hscroll)
        else:
            # No rows yet: just update the column count and header labels so the
            # table reflects the new column selection before any BOM is loaded.
            _headers = self._display_headers()
            self._table.setColumnCount(len(_headers))
            self._table.setHeaderLabels(_headers)

    # ── Preset column helpers ─────────────────────────────────────────────────

    def _display_headers(self) -> list[str]:
        """Return display header labels for the current column list.

        When "文件名列显示完整路径" is active the Filename column header is
        shown as "完整路径" so users can tell what they're looking at.
        """
        result = []
        for c in self._columns:
            if c == "Filename" and self._show_filepath_col:
                result.append("完整路径")
            else:
                result.append(BOM_COLUMN_DISPLAY_NAMES.get(c, c))
        return result

    def _build_visible_columns(self) -> list[str]:
        base = list(BOM_EDIT_COLUMN_ORDER)
        # Filter out hidden columns (Filename and hideable columns)
        if not self._show_filename_col:
            base = [c for c in base if c != "Filename"]
        # Filter out hidden hideable columns
        base = [c for c in base if c not in BOM_HIDEABLE_COLUMNS or c in self._visible_hideable_cols]
        if self._summarize:
            # In summary mode Level has no meaning; also hide Type unless assemblies shown
            cols_to_hide = {"Level"}
            if not self._summary_include_assemblies:
                cols_to_hide.add("Type")
            base = [c for c in base if c not in cols_to_hide]
        visible_preset = [
            c for c in PRESET_USER_REF_PROPERTIES if c in self._visible_preset_cols
        ]
        other_custom   = [
            c for c in self._custom_columns
            if c not in BOM_EDIT_COLUMN_ORDER and c not in PRESET_USER_REF_PROPERTIES
        ]
        # Insert "#" immediately after "Level" (column 0 → Level, column 1 → "#")
        # so that the QTreeWidget tree-decoration (branch lines) stays in the
        # Level column (logical index 0).  In summary mode Level is hidden, so
        # "#" falls to the front (column 0) which is fine.
        result = base + visible_preset + other_custom
        if "Level" in result:
            level_idx = result.index("Level")
            result.insert(level_idx + 1, BOM_ROW_NUMBER_COLUMN)
        else:
            result.insert(0, BOM_ROW_NUMBER_COLUMN)
        return result

    def _on_preset_col_toggled(self) -> None:
        # "Filename" checkbox controls the built-in filename column visibility
        if "Filename" in self._preset_checkboxes:
            new_show_fn = self._preset_checkboxes["Filename"].isChecked()
            if new_show_fn != self._show_filename_col:
                self._show_filename_col = new_show_fn
                self._edit_settings.setValue("show_filename_column", self._show_filename_col)
        self._visible_preset_cols = [
            name for name, cb in self._preset_checkboxes.items()
            if name != "Filename" and name not in BOM_HIDEABLE_COLUMNS and cb.isChecked()
        ]
        self._edit_settings.setValue("visible_preset_columns", self._visible_preset_cols)
        self._rebuild_columns_and_repopulate()

    def _on_hideable_col_toggled(self) -> None:
        """Handle hideable column checkbox toggle (Nomenclature, Revision, Definition, Source)."""
        self._visible_hideable_cols = [
            name for name, cb in self._preset_checkboxes.items()
            if name in BOM_HIDEABLE_COLUMNS and cb.isChecked()
        ]
        self._edit_settings.setValue("visible_hideable_columns", self._visible_hideable_cols)
        self._rebuild_columns_and_repopulate()

    def _on_show_filepath_toggled(self, checked: bool) -> None:
        self._show_filepath_col = checked
        self._edit_settings.setValue("show_filepath_column", checked)
        self._rebuild_columns_and_repopulate()

    # ── File picker ───────────────────────────────────────────────────────────

    def _browse_file(self) -> None:
        file, _ = QFileDialog.getOpenFileName(
            self, "选择CATProduct文件",
            self._last_browse_dir,
            "*.CATProduct (*.CATProduct);;All Files (*)",
        )
        if file:
            self._file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._export_settings.setValue("last_browse_dir", self._last_browse_dir)

    # ── Load BOM ──────────────────────────────────────────────────────────────

    def _load_bom(self) -> None:
        if self._use_active_chk.isChecked():
            file_path = None
        else:
            file_path = self._file_edit.text().strip()
            if not file_path:
                QMessageBox.warning(self, "未选择文件", "请先选择一个CATProduct文件。")
                return
            if not Path(file_path).exists():
                QMessageBox.warning(self, "文件不存在", f"文件不存在：\n{file_path}")
                return

        self._load_btn.setEnabled(False)
        self._load_btn.setText("加载中…")
        QApplication.processEvents()

        progress = QProgressDialog("正在加载BOM，请稍候…", None, 0, 0, self)
        progress.setWindowTitle("加载BOM")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(300)
        progress.setValue(0)

        def _on_row_collected(count: int) -> None:
            progress.setLabelText(f"正在加载BOM，请稍候… 已读取 {count} 个节点")
            progress.repaint()
            QApplication.processEvents()

        try:
            all_read_cols = list(dict.fromkeys(
                BOM_EDIT_COLUMN_ORDER
                + [c for c in self._all_custom_columns if c not in BOM_EDIT_COLUMN_ORDER]
            ))
            rows = collect_bom_rows(
                file_path, all_read_cols, self._all_custom_columns,
                progress_callback=_on_row_collected,
            )
        except Exception as e:
            progress.close()
            logger.error(f"Failed to load BOM for edit: {e}")
            QMessageBox.critical(
                self, "加载失败",
                f"加载BOM时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            self._load_btn.setEnabled(True)
            self._load_btn.setText("加载BOM")
            return
        finally:
            progress.close()

        self._load_btn.setEnabled(True)
        self._load_btn.setText("重新加载BOM")

        # Always save the raw hierarchical rows so we can switch modes later
        self._raw_rows = rows

        # In summary mode collapse the hierarchy into unique parts
        display_rows = (
            flatten_bom_to_summary(
                rows,
                include_assemblies=self._summary_include_assemblies,
                sort_column=self._summary_sort_column or None,
            )
            if self._summarize else rows
        )

        self._rows = display_rows

        # Build PN-keyed canonical data from the raw rows (first occurrence wins).
        # Using raw rows ensures all parts are indexed regardless of current mode.
        all_data_cols = list(dict.fromkeys(
            BOM_EDIT_COLUMN_ORDER
            + [c for c in self._all_custom_columns if c not in BOM_EDIT_COLUMN_ORDER]
        ))
        self._canonical_data = {}
        for row in rows:
            pn = str(row.get("Part Number", ""))
            if pn and pn not in self._canonical_data:
                data: dict[str, str] = {}
                for col in all_data_cols:
                    val = str(row.get(col, ""))
                    if col == "Source":
                        val = SOURCE_TO_DISPLAY.get(val, val)
                    data[col] = val
                self._canonical_data[pn] = data

        self._snapshot_data  = copy.deepcopy(self._canonical_data)
        self._modified_keys.clear()

        # Save current widths by column name before repopulating
        if self._bom_loaded:
            for col_idx, col_name in enumerate(self._columns):
                self._col_widths[col_name] = self._table.columnWidth(col_idx)

        self._populate_table()
        if not self._bom_loaded:
            # First load: auto-fit all columns and seed the cache; '#' gets a
            # fixed default of 40 px (resizable afterwards like any other column)
            for _c, col_name in enumerate(self._columns):
                if col_name == BOM_ROW_NUMBER_COLUMN:
                    self._table.setColumnWidth(_c, 40)
                    self._col_widths[col_name] = 40
                else:
                    self._table.resizeColumnToContents(_c)
                    self._col_widths[col_name] = self._table.columnWidth(_c)
            self._bom_loaded = True
        else:
            # Subsequent reloads: restore saved widths by column name
            for col_idx, col_name in enumerate(self._columns):
                if col_name in self._col_widths:
                    self._table.setColumnWidth(col_idx, self._col_widths[col_name])

        self._save_btn.setEnabled(True)
        self._finish_btn.setEnabled(True)
        self._rename_btn.setEnabled(True)
        self._rename_file_btn.setEnabled(True)
        self._export_btn.setEnabled(True)

    def _populate_table(self) -> None:
        self._is_updating = True
        self._table.blockSignals(True)

        # Summary mode: all rows are flat top-level items with no children.
        # Keeping setRootIsDecorated(True) reserves space for the expand arrow on
        # every row, which pushes column-0 content to the right.  Disable it in
        # summary mode; re-enable it in hierarchical mode so expand arrows show.
        self._table.setRootIsDecorated(not self._summarize)

        self._table.clear()                          # removes all items; headers persist
        headers = self._display_headers()
        self._table.setColumnCount(len(headers))     # Qt never shrinks column count on its own
        self._table.setHeaderLabels(headers)
        self._item_by_row = []
        self._pn_to_items.clear()  # Reset PN→Item index

        # parent_stack: list of (level, item_or_None)
        # The sentinel at position 0 represents the invisible root (level −1).
        parent_stack: list[tuple[int, QTreeWidgetItem | None]] = [(-1, None)]

        for row_idx, row_data in enumerate(self._rows):
            level = 0 if self._summarize else int(row_data.get("Level", 0))

            # Pop until the top of the stack has a level strictly below ours
            while len(parent_stack) > 1 and parent_stack[-1][0] >= level:
                parent_stack.pop()

            parent_item = parent_stack[-1][1]
            item = QTreeWidgetItem()
            # Store row_idx in UserRole of column 0 for reverse lookup
            item.setData(0, Qt.ItemDataRole.UserRole, row_idx)

            if parent_item is None:
                self._table.addTopLevelItem(item)
            else:
                parent_item.addChild(item)

            parent_stack.append((level, item))
            self._item_by_row.append(item)

            pn         = str(row_data.get("Part Number", ""))
            not_found  = bool(row_data.get("_not_found"))
            unreadable = bool(row_data.get("_unreadable"))
            row_locked = unreadable or not_found

            # Build PN→Item index for fast linked updates
            if pn:
                self._pn_to_items.setdefault(pn, []).append(item)

            for col_idx, col_name in enumerate(self._columns):

                # Source → QComboBox (overlay widget; not stored as item text)
                if col_name == "Source":
                    raw    = str(row_data.get("Source", ""))
                    pn_val = self._canonical_data.get(pn, {}).get(
                        "Source", SOURCE_TO_DISPLAY.get(raw, raw)
                    )
                    if pn_val not in SOURCE_OPTIONS:
                        pn_val = SOURCE_TO_DISPLAY.get(pn_val, SOURCE_OPTIONS[0])
                    combo = QComboBox()
                    combo.blockSignals(True)
                    combo.addItems(SOURCE_OPTIONS)
                    combo.setCurrentText(pn_val)
                    combo.blockSignals(False)
                    if row_locked:
                        combo.setEnabled(False)
                    else:
                        combo.currentTextChanged.connect(
                            lambda text, r=row_idx: self._on_source_changed(r, text)
                        )
                    self._table.setItemWidget(item, col_idx, combo)
                    continue

                # User-defined property with constrained options → QComboBox
                opts = PRESET_USER_REF_PROPERTY_OPTIONS.get(col_name)
                if opts is not None:
                    pn_val = self._canonical_data.get(pn, {}).get(
                        col_name, str(row_data.get(col_name, ""))
                    )
                    # Build the effective item list:
                    #   • always prepend "" so an unset property shows as blank
                    #   • if the stored value is not in the allowed list AND is
                    #     non-empty, append it so the real value remains visible
                    display_opts = [""] + list(opts)
                    if pn_val and pn_val not in opts:
                        logger.debug(
                            "属性 '%s' 的值 '%s' 不在可选列表中，将以原始值显示（零件编号: %s）",
                            col_name, pn_val, pn,
                        )
                        display_opts.append(pn_val)
                    combo = QComboBox()
                    combo.blockSignals(True)
                    combo.addItems(display_opts)
                    combo.setCurrentText(pn_val)
                    combo.blockSignals(False)
                    if row_locked:
                        combo.setEnabled(False)
                    else:
                        combo.currentTextChanged.connect(
                            lambda text, r=row_idx, c=col_name: self._on_option_col_changed(r, c, text)
                        )
                    self._table.setItemWidget(item, col_idx, combo)
                    continue

                # All other columns → item text
                if col_name == BOM_ROW_NUMBER_COLUMN:
                    value = str(row_idx + 1)
                elif col_name == "Quantity":
                    value = str(row_data.get("Quantity", "1"))
                elif col_name == "Filename":
                    fp = str(row_data.get("_filepath", ""))
                    fn = str(row_data.get("Filename", ""))
                    if self._show_filepath_col:
                        value = fp if fp else fn
                    else:
                        # Always show filename with extension when the backing
                        # path is known; fall back to the stored stem (which may
                        # equal FILENAME_NOT_FOUND) when it is not.
                        value = Path(fp).name if fp else fn
                elif col_name == "Filepath":
                    value = str(row_data.get("_filepath", ""))
                elif col_name in BOM_READONLY_COLUMNS:
                    value = str(row_data.get(col_name, ""))
                else:
                    value = str(
                        self._canonical_data.get(pn, {}).get(
                            col_name, row_data.get(col_name, "")
                        )
                    )
                item.setText(col_idx, value)

                if col_name == "Filename":
                    fp = str(row_data.get("_filepath", ""))
                    fn = str(row_data.get("Filename", ""))
                    if fp:
                        if self._show_filepath_col:
                            # Column shows full path; tooltip shows just name+ext
                            name_with_ext = Path(fp).name
                            if name_with_ext and name_with_ext != FILENAME_NOT_FOUND:
                                item.setToolTip(col_idx, name_with_ext)
                        else:
                            # Column shows name+ext; tooltip shows full path
                            item.setToolTip(col_idx, fp)

            # Non-locked rows: allow in-place editing (delegate blocks read-only columns)
            if not row_locked:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                item.setData(0, _ITEM_LOCKED_ROLE, False)
            else:
                grey = QColor(160, 160, 160)
                bg   = QColor(255, 205, 205) if not_found else QColor(245, 245, 245)
                item.setData(0, _ITEM_LOCKED_ROLE, True)
                for ci in range(len(self._columns)):
                    item.setForeground(ci, grey)
                    item.setBackground(ci, bg)
                fn_col = self._columns.index("Filename") if "Filename" in self._columns else -1
                if fn_col >= 0:
                    tip = (
                        "该零件/装配体的文件未被CATIA检索到，行内容不可编辑。"
                        if not_found else
                        "该零件/装配体处于轻量化模式，无法读取属性。"
                    )
                    item.setToolTip(fn_col, tip)

        self._table.expandAll()
        self._table.blockSignals(False)
        self._is_updating = False

    # ── Tree helpers ──────────────────────────────────────────────────────────

    def _iter_all_items(self):
        """Yield every QTreeWidgetItem in DFS (pre-order) traversal."""
        def _walk(parent: QTreeWidgetItem):
            yield parent
            for i in range(parent.childCount()):
                yield from _walk(parent.child(i))
        for i in range(self._table.topLevelItemCount()):
            yield from _walk(self._table.topLevelItem(i))

    # ── Source combo change ───────────────────────────────────────────────────

    def _on_source_changed(self, row_idx: int, text: str) -> None:
        if self._is_updating:
            return
        if "Source" not in self._columns:
            return
        src_col_idx = self._columns.index("Source")

        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        direct_rows = selected_row_indices if row_idx in selected_row_indices else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            pn = str(self._rows[r].get("Part Number", ""))
            if pn:
                pns_to_update.add(pn)

        for pn in pns_to_update:
            if pn in self._canonical_data:
                self._canonical_data[pn]["Source"] = text
                self._modified_keys.setdefault(pn, set()).add("Source")

        # Performance optimization: use PN→Item index instead of full tree traversal
        self._is_updating = True
        for pn in pns_to_update:
            if pn in self._pn_to_items:
                for other_item in self._pn_to_items[pn]:
                    combo = self._table.itemWidget(other_item, src_col_idx)
                    if isinstance(combo, QComboBox) and combo.currentText() != text:
                        combo.blockSignals(True)
                        combo.setCurrentText(text)
                        combo.blockSignals(False)
        self._is_updating = False

    # ── User-defined option column combo change ───────────────────────────────

    def _on_option_col_changed(self, row_idx: int, col_name: str, text: str) -> None:
        if self._is_updating:
            return
        if col_name not in self._columns:
            return
        col_idx = self._columns.index(col_name)

        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        direct_rows = selected_row_indices if row_idx in selected_row_indices else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            pn = str(self._rows[r].get("Part Number", ""))
            if pn:
                pns_to_update.add(pn)

        for pn in pns_to_update:
            if pn in self._canonical_data:
                self._canonical_data[pn][col_name] = text
                self._modified_keys.setdefault(pn, set()).add(col_name)

        # Performance optimization: use PN→Item index instead of full tree traversal
        self._is_updating = True
        for pn in pns_to_update:
            if pn in self._pn_to_items:
                for other_item in self._pn_to_items[pn]:
                    combo = self._table.itemWidget(other_item, col_idx)
                    if isinstance(combo, QComboBox) and combo.currentText() != text:
                        combo.blockSignals(True)
                        combo.setCurrentText(text)
                        combo.blockSignals(False)
        self._is_updating = False

    # ── Regular cell edit ─────────────────────────────────────────────────────

    def _on_item_changed(self, item: QTreeWidgetItem, col_idx: int) -> None:
        if self._is_updating:
            return
        row_idx = item.data(0, Qt.ItemDataRole.UserRole)
        if row_idx is None:
            return
        col_name = self._columns[col_idx]

        if col_name in BOM_READONLY_COLUMNS or col_name == "Source" or col_name in PRESET_USER_REF_PROPERTY_OPTIONS:
            return

        new_value = item.text(col_idx)
        pn        = str(self._rows[row_idx].get("Part Number", ""))

        if col_name == "Part Number":
            # ── Empty / whitespace-only PN ────────────────────────────────────
            if not new_value.strip():
                QMessageBox.warning(
                    self, "零件编号不能为空",
                    "零件编号不能为空或仅含空格，请输入有效的零件编号。",
                )
                self._is_updating = True
                item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                self._is_updating = False
                return

            # ── Strip leading/trailing whitespace silently ────────────────────
            if new_value != new_value.strip():
                new_value = new_value.strip()
                self._is_updating = True
                item.setText(col_idx, new_value)
                self._is_updating = False

            # ── Character validity ────────────────────────────────────────────
            if not PART_NUMBER_VALID_PATTERN.fullmatch(new_value):
                QMessageBox.warning(
                    self, "零件编号含非法字符",
                    f"零件编号 \"{new_value}\" 含有非法字符。\n"
                    "不允许：控制字符、非ASCII字符，以及Windows文件名禁用字符"
                    "（\\ / : * ? \" < > |）。",
                )
                self._is_updating = True
                item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                self._is_updating = False
                return

            # ── Conflict with current canonical values ────────────────────────
            for other_pn, data in self._canonical_data.items():
                if other_pn == pn:
                    continue
                if data.get("Part Number", other_pn) == new_value:
                    QMessageBox.warning(
                        self, "零件编号冲突",
                        f"零件编号 \"{new_value}\" 与 \"{other_pn}\" "
                        f"的当前零件编号冲突，不允许修改。",
                    )
                    self._is_updating = True
                    item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                    self._is_updating = False
                    return

            # ── Conflict with snapshot (what CATIA currently holds) ───────────
            for other_pn, data in self._snapshot_data.items():
                if other_pn == pn:
                    continue
                if data.get("Part Number", other_pn) == new_value:
                    QMessageBox.warning(
                        self, "零件编号冲突",
                        f"零件编号 \"{new_value}\" 与 \"{other_pn}\" "
                        f"的原始零件编号冲突，不允许修改。",
                    )
                    self._is_updating = True
                    item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                    self._is_updating = False
                    return

        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        direct_rows = selected_row_indices if row_idx in selected_row_indices else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            r_pn = str(self._rows[r].get("Part Number", ""))
            if r_pn:
                pns_to_update.add(r_pn)
                if r_pn in self._canonical_data:
                    self._canonical_data[r_pn][col_name] = new_value
                    self._modified_keys.setdefault(r_pn, set()).add(col_name)

        # Performance optimization: use PN→Item index instead of full tree traversal
        self._is_updating = True
        for pn in pns_to_update:
            if pn in self._pn_to_items:
                for other_item in self._pn_to_items[pn]:
                    if other_item.text(col_idx) != new_value:
                        other_item.setText(col_idx, new_value)
        self._is_updating = False

    # ── Write-back ────────────────────────────────────────────────────────────

    def _rename_by_part_number(self) -> None:
        """SaveAs each CATIA file using its Part Number as the filename."""
        if self._modified_keys:
            ret = QMessageBox.question(
                self, "存在未回传的修改",
                "检测到BOM属性尚未写回CATIA。\n\n"
                "必须先将修改写回CATIA，才能确保零件编号与CATIA文件一致。\n\n"
                "是否立即执行写回？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
            self._write_back(close_on_success=False)
            # If write-back failed or was only partial, modified_keys is still
            # non-empty – stop here so the user can fix the issue first.
            if self._modified_keys:
                return
            # Write-back cleared all modifications; fall through to rename.

        to_rename: list[tuple[str, str]] = []
        seen_fps:  set[str] = set()
        for row in self._rows:
            fp = str(row.get("_filepath", ""))
            if not fp or fp in seen_fps:
                continue
            seen_fps.add(fp)
            orig_pn = str(row.get("Part Number", ""))
            pn      = str(self._canonical_data.get(orig_pn, {}).get("Part Number", orig_pn))
            if pn and Path(fp).stem != pn:
                to_rename.append((fp, pn))

        if not to_rename:
            QMessageBox.information(self, "无需改名", "所有文件名已与零件编号一致。")
            return

        delete_old = (
            QMessageBox.question(
                self, "是否删除旧文件",
                "另存为完成后，是否删除旧文件？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            ) == QMessageBox.StandardButton.Yes
        )

        QMessageBox.information(self, "请在CATIA中继续操作", "准备就绪，请在CATIA中确认后续操作。")

        renamed_count = 0

        # Performance optimization: Build document cache once to avoid repeated scans
        from pycatia import catia as _pycatia
        caa         = _pycatia()
        application = caa.application
        application.visible = True
        documents   = application.documents

        # Cache: filepath → document
        doc_cache: dict[Path, object] = {}
        for i in range(1, documents.count + 1):
            try:
                doc = documents.item(i)
                doc_path = Path(doc.full_name).resolve()
                doc_cache[doc_path] = doc
            except Exception:
                pass

        for fp, pn in reversed(to_rename):
            if not PART_NUMBER_VALID_PATTERN.fullmatch(pn):
                QMessageBox.warning(
                    self, "零件编号含非法字符",
                    f"零件编号 「{pn}」 含有非法字符。\n"
                    "不允许：控制字符、非ASCII字符，以及Windows文件名禁用字符"
                    "（\\ / : * ? \" < > |）。\n请在表格中修改此零件编号后重试。",
                )
                continue

            if not Path(fp).exists():
                continue

            ext    = Path(fp).suffix
            new_fp = str(Path(fp).parent / (pn + ext))
            target_existed_before = Path(new_fp).exists()

            try:
                src = Path(fp).resolve()

                # Use cache for fast lookup
                target_doc = doc_cache.get(src)
                if target_doc is None:
                    documents.open(str(src))
                    target_doc = _find_catia_doc_by_path(documents, src)
                    if target_doc:
                        doc_cache[src] = target_doc

                if target_doc is None:
                    QMessageBox.warning(
                        self, "无法找到文档",
                        f"无法在CATIA中找到或打开文档：\n{fp}",
                    )
                    continue

                target_doc.com_object.SaveAs(new_fp)

                if delete_old and Path(fp).resolve() != Path(new_fp).resolve():
                    try:
                        os.remove(fp)
                    except Exception as del_err:
                        logger.warning(f"Failed to delete old file {fp}: {del_err}")

                for row in self._rows:
                    if str(row.get("_filepath", "")) == fp:
                        row["_filepath"] = new_fp
                        row["Filename"]  = pn
                for row in self._raw_rows:
                    if str(row.get("_filepath", "")) == fp:
                        row["_filepath"] = new_fp
                        row["Filename"]  = pn
                renamed_count += 1

                # Update cache with new path
                if Path(new_fp).resolve() != src:
                    doc_cache[Path(new_fp).resolve()] = target_doc
                    if src in doc_cache:
                        del doc_cache[src]

            except Exception as e:
                # Only treat the exception as a user-initiated cancel when:
                #   1. The exception came from the CATIA COM layer (pywintypes.com_error),
                #      which is what CATIA raises when the user clicks Cancel or No in
                #      its own SaveAs dialog – NOT for OS-level failures.
                #   2. The source file is still intact.
                #   3. The target file was either pre-existing or was never created.
                # Any other exception (OSError, PermissionError, non-COM errors like
                # disk full) must be surfaced to the user as a real failure.
                if _is_catia_com_error(e) and Path(fp).exists() and (
                    target_existed_before or not Path(new_fp).exists()
                ):
                    logger.info(
                        f"SaveAs skipped for {Path(fp).name} "
                        "(user cancelled or declined overwrite in CATIA)"
                    )
                    continue
                QMessageBox.warning(
                    self, "另存为失败", f"文件「{Path(fp).name}」另存为失败：\n{e}"
                )

        if renamed_count > 0:
            QMessageBox.information(
                self, "改名完成",
                f"已成功将 {renamed_count} 个文件通过CATIA另存为功能改名。",
            )
            self._populate_table()

    def _rename_selected_file(self) -> None:
        """Rename or move the file for a single selected BOM row via CATIA SaveAs."""
        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        if len(selected_row_indices) != 1:
            QMessageBox.warning(
                self, "请选择单行",
                "请在表格中选中恰好一行，再执行此操作。",
            )
            return

        row_idx  = next(iter(selected_row_indices))
        row_data = self._rows[row_idx]
        fp       = str(row_data.get("_filepath", ""))

        if not fp or row_data.get("_not_found"):
            QMessageBox.warning(self, "无有效路径", "该行没有可用的文件路径，无法执行重命名/移动。")
            return
        if not Path(fp).exists():
            QMessageBox.warning(self, "文件不存在", f"文件不存在：\n{fp}")
            return

        # Require attribute write-back before renaming to keep file content consistent.
        orig_pn = str(row_data.get("Part Number", ""))
        if orig_pn in self._modified_keys:
            ret = QMessageBox.question(
                self, "存在未写回的属性修改",
                f"零件「{orig_pn}」的属性尚未写回CATIA。\n\n"
                "必须先将修改写回CATIA，才能确保文件内容与表格一致。\n\n"
                "是否立即执行写回？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
            self._write_back(close_on_success=False)
            # Only proceed if write-back actually cleared the modifications;
            # if it failed (error dialog shown), modified_keys still has the entry.
            if orig_pn in self._modified_keys:
                return

        dlg = _FileRenameDialog(fp, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_fp                = dlg.new_path
        target_existed_before = Path(new_fp).exists()

        delete_old = (
            QMessageBox.question(
                self, "是否删除旧文件",
                f"另存为完成后，是否删除旧文件？\n\n旧文件：{fp}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            ) == QMessageBox.StandardButton.Yes
        )

        QMessageBox.information(self, "请在CATIA中继续操作", "准备就绪，请在CATIA中确认后续操作。")

        try:
            from pycatia import catia as _pycatia
            caa         = _pycatia()
            application = caa.application
            application.visible = True
            documents   = application.documents
            src         = Path(fp).resolve()

            target_doc = _find_catia_doc_by_path(documents, src)
            if target_doc is None:
                documents.open(str(src))
                target_doc = _find_catia_doc_by_path(documents, src)

            if target_doc is None:
                QMessageBox.warning(
                    self, "无法找到文档",
                    f"无法在CATIA中找到或打开文档：\n{fp}",
                )
                return

            target_doc.com_object.SaveAs(new_fp)

            if delete_old and Path(fp).resolve() != Path(new_fp).resolve():
                try:
                    os.remove(fp)
                except Exception as del_err:
                    logger.warning(f"Failed to delete old file {fp}: {del_err}")

            new_stem = Path(new_fp).stem
            for row in self._rows:
                if str(row.get("_filepath", "")) == fp:
                    row["_filepath"] = new_fp
                    row["Filename"]  = new_stem
            for row in self._raw_rows:
                if str(row.get("_filepath", "")) == fp:
                    row["_filepath"] = new_fp
                    row["Filename"]  = new_stem
            self._populate_table()
            QMessageBox.information(
                self, "操作成功",
                f"文件已成功另存为：\n{new_fp}",
            )

        except Exception as e:
            # Only treat the exception as a user-initiated cancel when:
            #   1. The exception came from the CATIA COM layer (pywintypes.com_error),
            #      which is what CATIA raises when the user clicks Cancel or No in
            #      its own SaveAs dialog – NOT for OS-level failures.
            #   2. The source file is still intact.
            #   3. The target file was either pre-existing or was never created.
            # Any other exception (OSError, PermissionError, non-COM errors like
            # disk full) must be surfaced to the user as a real failure.
            if _is_catia_com_error(e) and Path(fp).exists() and (
                target_existed_before or not Path(new_fp).exists()
            ):
                # Most likely the user clicked Cancel or No in CATIA's own SaveAs
                # prompt – treat this as a deliberate skip with no error dialog.
                logger.info(
                    f"SaveAs skipped for {Path(fp).name} "
                    f"(user cancelled or declined overwrite in CATIA; exception: {e})"
                )
                return
            QMessageBox.warning(self, "另存为失败", f"文件操作失败：\n{e}")

    def _write_back(self, *, close_on_success: bool) -> None:
        """Write only the changed fields back to CATIA."""
        if self._use_active_chk.isChecked():
            file_path = None
        else:
            file_path = self._file_edit.text().strip()
            if not file_path:
                QMessageBox.warning(self, "未选择文件", "请选择一个CATProduct文件。")
                return

        # dirty_data must be keyed by the *current* CATIA PN, which may differ
        # from the internal canonical key (orig_pn) when a PN rename was already
        # written back in a previous write-back operation.  We keep pn_remap to
        # go back from current_pn → orig_pn for the post-write snapshot update.
        dirty_data: dict[str, dict[str, str]] = {}
        pn_remap:   dict[str, str]            = {}  # current_pn → orig_pn
        for pn, dirty_cols in self._modified_keys.items():
            if pn not in self._canonical_data:
                continue
            changed = {
                col: self._canonical_data[pn][col]
                for col in dirty_cols if col in self._canonical_data[pn]
            }
            if changed:
                # Use the snapshot PN (= what CATIA currently holds for this
                # node) as the lookup key for the traversal.  If the PN has
                # never been written back, the snapshot value equals orig_pn.
                current_pn = self._snapshot_data.get(pn, {}).get(
                    "Part Number", pn
                )
                dirty_data[current_pn] = changed
                pn_remap[current_pn]   = pn

        if not dirty_data:
            if close_on_success:
                self.accept()
            else:
                QMessageBox.information(self, "无更改", "没有检测到任何修改，无需写回。")
            return

        self._save_btn.setEnabled(False)
        self._finish_btn.setEnabled(False)
        QApplication.processEvents()

        progress = QProgressDialog("正在写回CATIA，请稍候…", None, 0, 0, self)
        progress.setWindowTitle("写回CATIA")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(300)
        progress.setValue(0)

        def _on_node_written(count: int) -> None:
            progress.setLabelText(f"正在写回CATIA，请稍候… 已处理 {count} 个节点")
            progress.repaint()
            QApplication.processEvents()

        try:
            write_bom_to_catia(file_path, dirty_data, self._all_custom_columns,
                               _on_node_written)
        except Exception as e:
            progress.close()
            logger.error(f"Failed to write BOM back to CATIA: {e}")
            self._save_btn.setEnabled(True)
            self._finish_btn.setEnabled(True)
            QMessageBox.critical(
                self, "写回失败",
                f"写回CATIA时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            return
        finally:
            progress.close()

        for current_pn, changed in dirty_data.items():
            pn = pn_remap.get(current_pn, current_pn)
            if pn in self._snapshot_data:
                self._snapshot_data[pn].update(changed)
            if pn in self._modified_keys:
                self._modified_keys[pn] -= set(changed.keys())
                if not self._modified_keys[pn]:
                    del self._modified_keys[pn]

        self._save_btn.setEnabled(True)
        self._finish_btn.setEnabled(True)

        if close_on_success:
            QMessageBox.information(
                self, "完成",
                "BOM属性已成功写回CATIA，请在CATIA中手动保存文件。",
            )
            self.accept()
        else:
            QMessageBox.information(
                self, "应用成功",
                "BOM属性已成功写回CATIA，请在CATIA中手动保存文件。",
            )

    # ── Export table ──────────────────────────────────────────────────────────

    def _export_table(self) -> None:
        """Export the currently displayed BOM table to an Excel or CSV file."""
        if not self._bom_loaded or not self._rows:
            QMessageBox.warning(self, "无数据", "请先加载BOM。")
            return

        # Suggest a default filename derived from the source file
        initial_name = ""
        if not self._use_active_chk.isChecked():
            fp_src = self._file_edit.text().strip()
            if fp_src:
                suffix_hint = "_BOM汇总" if self._summarize else "_BOM"
                initial_name = str(Path(fp_src).with_name(Path(fp_src).stem + suffix_hint))

        dest, selected_filter = QFileDialog.getSaveFileName(
            self,
            "导出BOM表格",
            initial_name,
            "Excel工作簿 (*.xlsx);;CSV文件 (*.csv)",
        )
        if not dest:
            return

        dest_path = Path(dest)
        # Infer format from extension; fall back to xlsx when ambiguous
        suffix = dest_path.suffix.lower()
        if suffix not in (".xlsx", ".csv"):
            dest_path = dest_path.with_suffix(".xlsx")
            suffix = ".xlsx"

        # Columns to export: current visible columns, excluding the row-number "#" column
        export_cols = [c for c in self._columns if c != BOM_ROW_NUMBER_COLUMN]

        # Snapshot current column widths from the live table (pixels)
        col_px_widths: dict[str, int] = {}
        for col_idx, col_name in enumerate(self._columns):
            if col_name != BOM_ROW_NUMBER_COLUMN:
                col_px_widths[col_name] = self._table.columnWidth(col_idx)

        # Collect row data using the same value-resolution logic as _populate_table
        rows_data: list[dict] = []
        for row_data in self._rows:
            pn = str(row_data.get("Part Number", ""))
            row_out: dict = {}
            for col_name in export_cols:
                if col_name == "Source":
                    raw = str(row_data.get("Source", ""))
                    val = self._canonical_data.get(pn, {}).get(
                        "Source", SOURCE_TO_DISPLAY.get(raw, raw)
                    )
                elif col_name in PRESET_USER_REF_PROPERTY_OPTIONS:
                    val = self._canonical_data.get(pn, {}).get(
                        col_name, str(row_data.get(col_name, ""))
                    )
                elif col_name == "Filename":
                    fp_val = str(row_data.get("_filepath", ""))
                    fn_val = str(row_data.get("Filename", ""))
                    if self._show_filepath_col:
                        val = fp_val if fp_val else fn_val
                    else:
                        val = Path(fp_val).name if fp_val else fn_val
                elif col_name in BOM_READONLY_COLUMNS:
                    val = str(row_data.get(col_name, ""))
                else:
                    val = str(
                        self._canonical_data.get(pn, {}).get(
                            col_name, row_data.get(col_name, "")
                        )
                    )
                row_out[col_name] = val
            rows_data.append(row_out)

        try:
            if suffix == ".xlsx":
                self._write_xlsx(dest_path, export_cols, col_px_widths, rows_data)
            else:
                self._write_csv(dest_path, export_cols, rows_data)
        except PermissionError:
            QMessageBox.critical(
                self, "导出失败",
                f"无法写入文件（文件可能已在其他程序中打开）：\n{dest_path}",
            )
            return
        except Exception as e:
            logger.error(f"BOM table export failed: {e}")
            QMessageBox.critical(self, "导出失败", f"导出时出错：\n{e}")
            return

        QMessageBox.information(self, "导出成功", f"BOM已成功导出：\n{dest_path}")

    def _export_header(self, col_name: str) -> str:
        """Return the display header string for a column, matching the live table."""
        if col_name == "Filename" and self._show_filepath_col:
            return "完整路径"
        return BOM_COLUMN_DISPLAY_NAMES.get(col_name, col_name)

    def _write_xlsx(
        self,
        dest: Path,
        cols: list[str],
        px_widths: dict[str, int],
        rows: list[dict],
    ) -> None:
        """Write *rows* to an .xlsx workbook at *dest*."""
        import openpyxl
        from openpyxl.styles import Font, Alignment

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BOM汇总" if self._summarize else "BOM"

        center = Alignment(horizontal="center")

        # Header row
        for col_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(row=1, column=col_idx, value=self._export_header(col_name))
            cell.font = Font(bold=True)

        # Data rows
        for row_idx, row in enumerate(rows, start=2):
            for col_idx, col_name in enumerate(cols, start=1):
                raw_val = row.get(col_name, "")
                # Store numbers as integers so Excel can sort/filter them
                if col_name in ("Level", "Quantity"):
                    try:
                        value = int(raw_val)
                    except (ValueError, TypeError):
                        logger.debug(
                            "Could not convert %r to int for column '%s'", raw_val, col_name
                        )
                        value = raw_val
                else:
                    value = raw_val
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if col_name in ("Level", "Quantity", "Type"):
                    cell.alignment = center

        # Column widths: convert pixel widths → Excel character units
        # Calibri 11pt default: ~7 px per character unit is a reasonable approximation
        PX_PER_CHAR = 7.0
        for col_idx, col_name in enumerate(cols, start=1):
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            px = px_widths.get(col_name, 80)
            char_width = max(8.0, px / PX_PER_CHAR)
            ws.column_dimensions[col_letter].width = round(char_width, 1)

        wb.save(str(dest))
        logger.info(f"BOM table exported (xlsx) -> {dest}")

    def _write_csv(
        self,
        dest: Path,
        cols: list[str],
        rows: list[dict],
    ) -> None:
        """Write *rows* to a UTF-8 CSV file at *dest*."""
        import csv

        with open(dest, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([self._export_header(c) for c in cols])
            for row in rows:
                writer.writerow([row.get(c, "") for c in cols])
        logger.info(f"BOM table exported (csv) -> {dest}")

    def _apply_changes(self) -> None:
        """Write changes back to CATIA and keep the dialog open."""
        self._write_back(close_on_success=False)
    def _finish_and_close(self) -> None:
        """Write changes back to CATIA and close the dialog."""
        self._write_back(close_on_success=True)

    # ── Right-click context menu ──────────────────────────────────────────────

    def _on_tree_context_menu(self, pos) -> None:
        """Show a context menu for the right-clicked BOM row.

        If a thumbnail is embedded in the backing file it is shown at the top
        of the menu as a non-interactive image widget.
        """
        item = self._table.itemAt(pos)
        if item is None:
            return
        row_idx = item.data(0, Qt.ItemDataRole.UserRole)
        if row_idx is None:
            return

        row_data     = self._rows[row_idx]
        fp           = str(row_data.get("_filepath", ""))
        fp_path      = Path(fp) if fp else None
        is_component = row_data.get("Type") == "部件"
        not_found    = bool(row_data.get("_not_found"))
        unreadable   = bool(row_data.get("_unreadable"))
        pn           = str(row_data.get("Part Number", ""))

        # Ensure the right-clicked row is selected so that the downstream
        # helpers (_rename_selected_file etc.) can find it.
        if not item.isSelected():
            self._table.clearSelection()
            item.setSelected(True)

        menu = QMenu(self)

        # ── Embedded thumbnail (shown inline at the top when available) ───────
        # Conditions where we skip thumbnail extraction:
        #   • 部件: filepath belongs to the parent product, not this component
        #   • not_found: CATIA couldn't resolve the file
        #   • file doesn't exist on disk (unsaved or missing)
        if fp and not is_component and not not_found and fp_path is not None and fp_path.exists():
            img_bytes = read_catia_thumbnail(fp)
            if img_bytes:
                pixmap = QPixmap()
                loaded = pixmap.loadFromData(img_bytes)
                if loaded and not pixmap.isNull():
                    if (pixmap.width() > BOM_THUMBNAIL_MAX_SIZE
                            or pixmap.height() > BOM_THUMBNAIL_MAX_SIZE):
                        pixmap = pixmap.scaled(
                            BOM_THUMBNAIL_MAX_SIZE,
                            BOM_THUMBNAIL_MAX_SIZE,
                            Qt.AspectRatioMode.KeepAspectRatio,
                            Qt.TransformationMode.SmoothTransformation,
                        )
                    thumb_label = QLabel()
                    thumb_label.setPixmap(pixmap)
                    thumb_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    thumb_label.setContentsMargins(6, 4, 6, 4)
                    thumb_action = QWidgetAction(menu)
                    thumb_action.setDefaultWidget(thumb_label)
                    menu.addAction(thumb_action)
                    menu.addSeparator()

        # ── 打开路径 ──────────────────────────────────────────────────────────
        act_open_path = menu.addAction("打开路径")
        path_available = bool(fp) and fp_path is not None and (
            fp_path.exists() or fp_path.parent.exists()
        )
        act_open_path.setEnabled(path_available)

        # ── 复制路径 ──────────────────────────────────────────────────────────
        act_copy_path = menu.addAction("复制路径")
        act_copy_path.setEnabled(bool(fp))

        # ── 在CATIA中打开 ─────────────────────────────────────────────────────
        # Enabled only when the file exists on disk and is not a broken/unreadable
        # reference.  Component rows share the parent product's filepath so are
        # excluded as well.
        act_open_catia = menu.addAction("在CATIA中打开")
        catia_available = (
            not is_component and not not_found and not unreadable
            and fp_path is not None and fp_path.exists()
        )
        act_open_catia.setEnabled(catia_available)

        menu.addSeparator()

        # ── 编辑文件名/路径 ───────────────────────────────────────────────────
        act_edit_path = menu.addAction("编辑文件名/路径")
        act_edit_path.setEnabled(
            bool(fp) and not not_found
            and fp_path is not None and fp_path.exists()
        )

        action = menu.exec(self._table.viewport().mapToGlobal(pos))

        if action == act_open_path:
            self._open_path(fp)
        elif action == act_copy_path:
            QApplication.clipboard().setText(fp)
        elif action == act_open_catia:
            self._open_in_catia(fp)
        elif action == act_edit_path:
            self._rename_selected_file()

    def _open_path(self, fp: str) -> None:
        """Open the folder containing *fp* in Windows Explorer, highlighting the file."""
        p = Path(fp).resolve()
        try:
            if p.exists():
                # Quote the path so Explorer handles spaces in directory names.
                subprocess.Popen(f'explorer /select,"{p}"', shell=True)
            elif p.parent.exists():
                subprocess.Popen(f'explorer "{p.parent}"', shell=True)
        except Exception as exc:
            logger.warning(f"Failed to open path in Explorer: {exc}")

    def _open_in_catia(self, fp: str) -> None:
        """Open the CATIA document at file path *fp* via ``documents.open``.

        After opening, the CATIA V5 main window is brought to the Windows
        foreground via ``win32gui`` when available.
        """
        try:
            from pycatia import catia as _pycatia  # noqa: PLC0415
            caa         = _pycatia()
            application = caa.application
            application.visible = True
            documents   = application.documents

            fp_resolved = Path(fp).resolve()
            documents.open(str(fp_resolved))

            # ── Bring the CATIA V5 main window to the Windows foreground ──────
            try:
                import win32gui  # noqa: PLC0415
                import win32con  # noqa: PLC0415

                def _raise_catia_window(hwnd, _extra):
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    title = win32gui.GetWindowText(hwnd)
                    # Match only the CATIA V5 application window, not other
                    # windows that happen to contain the word "CATIA".
                    if title.startswith("CATIA V5"):
                        try:
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            win32gui.SetForegroundWindow(hwnd)
                        except Exception:
                            pass
                        # Stop enumeration after the first CATIA V5 window.
                        return False

                win32gui.EnumWindows(_raise_catia_window, None)
            except ImportError:
                pass
            except Exception:
                pass

        except Exception as e:
            QMessageBox.warning(self, "在CATIA中打开失败", f"无法在CATIA中打开文件：\n{e}")
