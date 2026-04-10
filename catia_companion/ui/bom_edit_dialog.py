"""
BOM edit dialog.

Provides:
- BomEditDialog – editable table for completing BOM properties and writing them
                  back to CATIA via COM.
"""

import copy
import os
import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QComboBox, QCheckBox, QGroupBox, QMessageBox, QApplication,
    QFileDialog,
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QSettings

from catia_companion.constants import (
    BOM_PRESET_CUSTOM_COLUMNS,
    BOM_EDIT_COLUMN_ORDER,
    BOM_COLUMN_DISPLAY_NAMES,
    BOM_READONLY_COLUMNS,
    SOURCE_TO_DISPLAY,
    SOURCE_OPTIONS,
    PART_NUMBER_VALID_PATTERN,
    FILENAME_NOT_FOUND,
)
from catia_companion.catia.bom_collect import collect_bom_rows
from catia_companion.catia.bom_write import write_bom_to_catia

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
            c for c in saved_visible if c in BOM_PRESET_CUSTOM_COLUMNS
        ]

        # All custom columns (including all presets) so that pre-loading from
        # CATIA covers every column regardless of current visibility.
        self._all_custom_columns: list[str] = list(dict.fromkeys(
            self._custom_columns + list(BOM_PRESET_CUSTOM_COLUMNS)
        ))

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
        # Row indices of collapsed assembly rows
        self._collapsed_rows: set[int] = set()
        # True once BOM has been successfully loaded at least once
        self._bom_loaded: bool = False

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

        hint = QLabel(
            "文件名 / 层级 / 类型 / 数量 为结构属性，不可编辑。"
            "零件编号可编辑但不能与其他行冲突。"
            "相同零件编号的行会联动更新。请确保 CATIA 已启动。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # Preset column visibility checkboxes
        preset_group  = QGroupBox("自定义属性列（勾选以显示）")
        preset_layout = QHBoxLayout(preset_group)
        preset_layout.setSpacing(12)
        self._preset_checkboxes: dict[str, QCheckBox] = {}
        for col_name in BOM_PRESET_CUSTOM_COLUMNS:
            cb = QCheckBox(col_name)
            cb.setChecked(col_name in self._visible_preset_cols)
            cb.toggled.connect(self._on_preset_col_toggled)
            preset_layout.addWidget(cb)
            self._preset_checkboxes[col_name] = cb
        layout.addWidget(preset_group)

        # Editable table
        display_headers = [BOM_COLUMN_DISPLAY_NAMES.get(c, c) for c in self._columns]
        self._table = QTableWidget(0, len(self._columns))
        self._table.setHorizontalHeaderLabels(display_headers)
        hdr = self._table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        hdr.setSectionsMovable(True)
        hdr.setFixedHeight(28)
        self._table.verticalHeader().setDefaultSectionSize(24)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._table.setAlternatingRowColors(True)
        self._table.itemChanged.connect(self._on_item_changed)
        self._table.cellClicked.connect(self._on_cell_clicked)
        layout.addWidget(self._table)

        # Bottom buttons
        btn_row = QHBoxLayout()

        autofit_btn = QPushButton("自适应列宽")
        autofit_btn.setToolTip("根据内容自动调整所有列的宽度")
        autofit_btn.clicked.connect(self._autofit_columns)
        btn_row.addWidget(autofit_btn)

        self._rename_btn = QPushButton("按零件编号将文件改名")
        self._rename_btn.setEnabled(False)
        self._rename_btn.clicked.connect(self._rename_by_part_number)
        btn_row.addWidget(self._rename_btn)
        btn_row.addStretch()

        self._save_btn   = QPushButton("应用（写回CATIA）")
        self._save_btn.setEnabled(False)
        self._save_btn.clicked.connect(self._apply_changes)

        self._finish_btn = QPushButton("完成（写回CATIA）")
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

    # ── Table helpers ─────────────────────────────────────────────────────────

    def _autofit_columns(self) -> None:
        """Resize all columns to fit their content, with a minimum width."""
        self._table.resizeColumnsToContents()
        # Enforce a reasonable minimum width
        min_width = 60
        for col in range(self._table.columnCount()):
            if self._table.columnWidth(col) < min_width:
                self._table.setColumnWidth(col, min_width)

    # ── Preset column helpers ─────────────────────────────────────────────────

    def _build_visible_columns(self) -> list[str]:
        visible_preset = [
            c for c in BOM_PRESET_CUSTOM_COLUMNS if c in self._visible_preset_cols
        ]
        other_custom   = [
            c for c in self._custom_columns
            if c not in BOM_EDIT_COLUMN_ORDER and c not in BOM_PRESET_CUSTOM_COLUMNS
        ]
        return BOM_EDIT_COLUMN_ORDER + visible_preset + other_custom

    def _on_preset_col_toggled(self) -> None:
        self._visible_preset_cols = [
            name for name, cb in self._preset_checkboxes.items() if cb.isChecked()
        ]
        self._edit_settings.setValue("visible_preset_columns", self._visible_preset_cols)
        self._columns = self._build_visible_columns()
        display_headers = [BOM_COLUMN_DISPLAY_NAMES.get(c, c) for c in self._columns]
        self._table.setColumnCount(len(self._columns))
        self._table.setHorizontalHeaderLabels(display_headers)
        if self._rows:
            self._populate_table()
            self._table.resizeColumnsToContents()

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

        try:
            all_read_cols = list(dict.fromkeys(
                BOM_EDIT_COLUMN_ORDER
                + [c for c in self._all_custom_columns if c not in BOM_EDIT_COLUMN_ORDER]
            ))
            rows = collect_bom_rows(file_path, all_read_cols, self._all_custom_columns)
        except Exception as e:
            logger.error(f"Failed to load BOM for edit: {e}")
            QMessageBox.critical(
                self, "加载失败",
                f"加载BOM时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            self._load_btn.setEnabled(True)
            self._load_btn.setText("加载BOM")
            return

        self._load_btn.setEnabled(True)
        self._load_btn.setText("重新加载BOM")
        self._rows = rows
        self._collapsed_rows.clear()

        # Build PN-keyed canonical data (first occurrence wins)
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

        saved_widths = (
            [self._table.columnWidth(i) for i in range(self._table.columnCount())]
            if self._bom_loaded else []
        )

        self._populate_table()
        if not self._bom_loaded:
            self._table.resizeColumnsToContents()
            self._bom_loaded = True
        else:
            for i, w in enumerate(saved_widths):
                if i < self._table.columnCount():
                    self._table.setColumnWidth(i, w)

        self._save_btn.setEnabled(True)
        self._finish_btn.setEnabled(True)
        self._rename_btn.setEnabled(True)

    def _populate_table(self) -> None:
        self._is_updating = True

        display_headers = [BOM_COLUMN_DISPLAY_NAMES.get(c, c) for c in self._columns]
        self._table.setColumnCount(len(self._columns))
        self._table.setHorizontalHeaderLabels(display_headers)
        self._table.setRowCount(0)
        self._table.setRowCount(len(self._rows))

        for row_idx, row_data in enumerate(self._rows):
            pn         = str(row_data.get("Part Number", ""))
            not_found  = bool(row_data.get("_not_found"))
            unreadable = bool(row_data.get("_unreadable"))

            for col_idx, col_name in enumerate(self._columns):

                # Source → QComboBox
                if col_name == "Source":
                    raw       = str(row_data.get("Source", ""))
                    pn_val    = self._canonical_data.get(pn, {}).get(
                        "Source", SOURCE_TO_DISPLAY.get(raw, raw)
                    )
                    if pn_val not in SOURCE_OPTIONS:
                        pn_val = SOURCE_TO_DISPLAY.get(pn_val, SOURCE_OPTIONS[0])
                    combo = QComboBox()
                    combo.addItems(SOURCE_OPTIONS)
                    combo.setCurrentText(pn_val)
                    combo.currentTextChanged.connect(
                        lambda text, r=row_idx: self._on_source_changed(r, text)
                    )
                    self._table.setCellWidget(row_idx, col_idx, combo)
                    continue

                # All other columns → QTableWidgetItem
                if col_name == "Level":
                    value = self._level_cell_text(row_idx)
                elif col_name == "Quantity":
                    value = str(row_data.get("Quantity", "1"))
                elif col_name in BOM_READONLY_COLUMNS:
                    value = str(row_data.get(col_name, ""))
                else:
                    value = str(
                        self._canonical_data.get(pn, {}).get(
                            col_name, row_data.get(col_name, "")
                        )
                    )

                item = QTableWidgetItem(value)
                if col_name in BOM_READONLY_COLUMNS:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if col_name == "Filename":
                    fp = str(row_data.get("_filepath", ""))
                    if fp:
                        item.setToolTip(fp)
                self._table.setItem(row_idx, col_idx, item)

            # Lock rows that cannot be accessed or whose backing file is missing
            row_locked = unreadable or not_found
            if row_locked:
                grey = QColor(160, 160, 160)
                bg   = QColor(250, 245, 245) if not_found else QColor(245, 245, 245)
                for ci in range(len(self._columns)):
                    it = self._table.item(row_idx, ci)
                    if it:
                        it.setForeground(grey)
                        it.setBackground(bg)
                        it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    w = self._table.cellWidget(row_idx, ci)
                    if isinstance(w, QComboBox):
                        w.setEnabled(False)
                # Show a tooltip explaining why the row is locked
                fn_col = self._columns.index("Filename") if "Filename" in self._columns else -1
                if fn_col >= 0:
                    it = self._table.item(row_idx, fn_col)
                    if it:
                        if not_found:
                            it.setToolTip("该零件/装配体的文件未被CATIA检索到，行内容不可编辑。")
                        else:
                            it.setToolTip("该零件/装配体处于轻量化模式，无法读取属性。")

        self._is_updating = False

    # ── Collapse / expand helpers ─────────────────────────────────────────────

    def _row_has_children(self, row_idx: int) -> bool:
        if row_idx + 1 >= len(self._rows):
            return False
        return (
            self._rows[row_idx + 1].get("Level", 0)
            > self._rows[row_idx].get("Level", 0)
        )

    def _level_cell_text(self, row_idx: int) -> str:
        level = self._rows[row_idx].get("Level", 0)
        if self._row_has_children(row_idx):
            indicator = "▶ " if row_idx in self._collapsed_rows else "▼ "
        else:
            indicator = "  "
        return "  " * level + indicator + str(level)

    def _update_row_visibility(self) -> None:
        hide_depth_stack: list[int] = []
        for r, row_data in enumerate(self._rows):
            level = row_data.get("Level", 0)
            while hide_depth_stack and hide_depth_stack[-1] >= level:
                hide_depth_stack.pop()
            should_hide = bool(hide_depth_stack)
            self._table.setRowHidden(r, should_hide)
            if not should_hide and r in self._collapsed_rows:
                hide_depth_stack.append(level)

    def _on_cell_clicked(self, row: int, col: int) -> None:
        if "Level" not in self._columns:
            return
        level_col = self._columns.index("Level")
        if col != level_col or not self._row_has_children(row):
            return
        if row in self._collapsed_rows:
            self._collapsed_rows.discard(row)
        else:
            self._collapsed_rows.add(row)
        item = self._table.item(row, level_col)
        if item:
            item.setText(self._level_cell_text(row))
        self._update_row_visibility()

    # ── Source combo change ───────────────────────────────────────────────────

    def _on_source_changed(self, row_idx: int, text: str) -> None:
        if self._is_updating:
            return
        if "Source" not in self._columns:
            return
        src_col_idx = self._columns.index("Source")

        selected_rows = {idx.row() for idx in self._table.selectedIndexes()}
        direct_rows   = selected_rows if row_idx in selected_rows else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            pn = str(self._rows[r].get("Part Number", ""))
            if pn:
                pns_to_update.add(pn)

        for pn in pns_to_update:
            if pn in self._canonical_data:
                self._canonical_data[pn]["Source"] = text
                self._modified_keys.setdefault(pn, set()).add("Source")

        self._is_updating = True
        for r in range(self._table.rowCount()):
            if r == row_idx:
                continue
            other_pn = str(self._rows[r].get("Part Number", ""))
            if other_pn in pns_to_update:
                combo = self._table.cellWidget(r, src_col_idx)
                if isinstance(combo, QComboBox) and combo.currentText() != text:
                    combo.blockSignals(True)
                    combo.setCurrentText(text)
                    combo.blockSignals(False)
        self._is_updating = False

    # ── Regular cell edit ─────────────────────────────────────────────────────

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._is_updating:
            return
        col_idx  = item.column()
        row_idx  = item.row()
        col_name = self._columns[col_idx]

        if col_name in BOM_READONLY_COLUMNS or col_name == "Source":
            return

        new_value = item.text()
        pn        = str(self._rows[row_idx].get("Part Number", ""))

        # Part Number conflict checking
        if col_name == "Part Number":
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
                    item.setText(self._canonical_data.get(pn, {}).get("Part Number", pn))
                    self._is_updating = False
                    return
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
                    item.setText(self._canonical_data.get(pn, {}).get("Part Number", pn))
                    self._is_updating = False
                    return

        # Part Number character validity
        if col_name == "Part Number" and new_value:
            if not PART_NUMBER_VALID_PATTERN.fullmatch(new_value):
                QMessageBox.warning(
                    self, "零件编号含非法字符",
                    f"零件编号 \"{new_value}\" 含有非法字符。\n"
                    "不允许：控制字符、非英文字符，以及Windows文件名禁用字符"
                    "（\\ / : * ? \" < > |）。",
                )
                self._is_updating = True
                item.setText(self._canonical_data.get(pn, {}).get("Part Number", pn))
                self._is_updating = False
                return

        selected_rows = {idx.row() for idx in self._table.selectedIndexes()}
        direct_rows   = selected_rows if row_idx in selected_rows else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            r_pn = str(self._rows[r].get("Part Number", ""))
            if r_pn:
                pns_to_update.add(r_pn)
                if r_pn in self._canonical_data:
                    self._canonical_data[r_pn][col_name] = new_value
                    self._modified_keys.setdefault(r_pn, set()).add(col_name)

        self._is_updating = True
        for r in range(self._table.rowCount()):
            if r == row_idx:
                continue
            other_pn = str(self._rows[r].get("Part Number", ""))
            if other_pn in pns_to_update:
                other_item = self._table.item(r, col_idx)
                if other_item and other_item.text() != new_value:
                    other_item.setText(new_value)
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
            return

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

        for fp, pn in reversed(to_rename):
            if not PART_NUMBER_VALID_PATTERN.fullmatch(pn):
                QMessageBox.warning(
                    self, "零件编号含非法字符",
                    f"零件编号 「{pn}」 含有非法字符。\n"
                    "不允许：控制字符、非英文字符，以及Windows文件名禁用字符"
                    "（\\ / : * ? \" < > |）。\n请在表格中修改此零件编号后重试。",
                )
                continue

            if not Path(fp).exists():
                continue

            ext    = Path(fp).suffix
            new_fp = str(Path(fp).parent / (pn + ext))
            target_existed_before = Path(new_fp).exists()

            try:
                from catia_companion.catia.connection import connect_to_catia
                caa         = connect_to_catia()
                application = caa.application
                application.visible = True
                documents   = application.documents
                src         = Path(fp).resolve()

                def _find_doc(docs, path: Path):
                    for i in range(1, docs.count + 1):
                        try:
                            d = docs.item(i)
                            if Path(d.full_name).resolve() == path:
                                return d
                        except Exception:
                            pass
                    return None

                target_doc = _find_doc(documents, src)
                if target_doc is None:
                    documents.open(str(src))
                    target_doc = _find_doc(documents, src)

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
                renamed_count += 1

            except Exception as e:
                if target_existed_before and Path(fp).exists():
                    # The target file already existed before SaveAs and the source
                    # file is still intact.  This most likely means the user clicked
                    # "No" when CATIA asked whether to overwrite – treat as a
                    # user-initiated skip and move on silently.
                    logger.info(
                        f"SaveAs skipped for {Path(fp).name} "
                        "(target already existed; user likely declined overwrite)"
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

    def _write_back(self, *, close_on_success: bool) -> None:
        """Write only the changed fields back to CATIA."""
        if self._use_active_chk.isChecked():
            file_path = None
        else:
            file_path = self._file_edit.text().strip()
            if not file_path:
                QMessageBox.warning(self, "未选择文件", "请选择一个CATProduct文件。")
                return

        dirty_data: dict[str, dict[str, str]] = {}
        for pn, dirty_cols in self._modified_keys.items():
            if pn not in self._canonical_data:
                continue
            changed = {
                col: self._canonical_data[pn][col]
                for col in dirty_cols if col in self._canonical_data[pn]
            }
            if changed:
                dirty_data[pn] = changed

        if not dirty_data:
            if close_on_success:
                self.accept()
            else:
                QMessageBox.information(self, "无更改", "没有检测到任何修改，无需写回。")
            return

        self._save_btn.setEnabled(False)
        self._finish_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            write_bom_to_catia(file_path, dirty_data, self._all_custom_columns)
        except Exception as e:
            logger.error(f"Failed to write BOM back to CATIA: {e}")
            self._save_btn.setEnabled(True)
            self._finish_btn.setEnabled(True)
            QMessageBox.critical(
                self, "写回失败",
                f"写回CATIA时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            return

        for pn, changed in dirty_data.items():
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

    def _apply_changes(self) -> None:
        """Write changes back to CATIA and keep the dialog open."""
        self._write_back(close_on_success=False)

    def _finish_and_close(self) -> None:
        """Write changes back to CATIA and close the dialog."""
        self._write_back(close_on_success=True)
