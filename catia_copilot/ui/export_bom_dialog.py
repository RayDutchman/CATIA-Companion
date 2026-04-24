"""
BOM 导出对话框。

提供：
- ExportBomDialog – 用于选择 CATProduct、选择列并将 BOM 导出到 Excel 的对话框。
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QFileDialog, QAbstractItemView, QRadioButton, QButtonGroup, QLineEdit,
    QGroupBox, QPushButton, QMessageBox, QProgressDialog, QApplication,
    QCheckBox, QComboBox,
)
from PySide6.QtCore import Qt, QSettings

from catia_copilot.constants import (
    BOM_ALL_COLUMNS,
    BOM_DEFAULT_COLUMNS,
    PRESET_USER_REF_PROPERTIES,
    BOM_COLUMN_DISPLAY_NAMES,
)
from catia_copilot.catia.bom_export import export_bom_to_excel

logger = logging.getLogger(__name__)


class ExportBomDialog(QDialog):
    """将 CATProduct 的 BOM 导出到 Excel 文件的对话框。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("从CATProduct导出BOM")
        self.setMinimumSize(560, 580)

        self._settings        = QSettings("CATIACompanion", "ExportBOMDialog")
        self._last_browse_dir = self._settings.value("last_browse_dir", "")
        self._last_output_dir = self._settings.value("last_output_dir", "")

        saved_custom = self._settings.value("custom_columns", [])
        if isinstance(saved_custom, str):
            saved_custom = [saved_custom]
        self._custom_columns: list[str] = list(saved_custom)

        self._summarize: bool = self._settings.value("summarize", False, type=bool)
        self._summary_include_assemblies: bool = self._settings.value(
            "summary_include_assemblies", False, type=bool
        )
        self._summary_sort_column: str = self._settings.value(
            "summary_sort_column", "Part Number"
        )

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Source selection ────────────────────────────────────────────────
        src_group  = QGroupBox("数据来源")
        src_layout = QVBoxLayout(src_group)
        self._src_btn_group = QButtonGroup(self)
        self._radio_active  = QRadioButton("使用当前CATIA活动文档")
        self._radio_file    = QRadioButton("选择文件:")
        self._radio_file.setChecked(True)
        self._src_btn_group.addButton(self._radio_active)
        self._src_btn_group.addButton(self._radio_file)
        src_layout.addWidget(self._radio_active)

        file_row = QHBoxLayout()
        file_row.addWidget(self._radio_file)
        self._file_edit       = QLineEdit()
        self._file_edit.setPlaceholderText("选择一个CATProduct文件...")
        self._file_edit.setReadOnly(True)
        self._file_browse_btn = QPushButton("浏览...")
        self._file_browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(self._file_browse_btn)
        src_layout.addLayout(file_row)
        self._radio_active.toggled.connect(self._toggle_source_row)
        layout.addWidget(src_group)

        # ── Output folder ───────────────────────────────────────────────────
        output_group  = QGroupBox("输出文件夹")
        output_layout = QVBoxLayout(output_group)
        self._radio_same   = QRadioButton("与源文件相同目录")
        self._radio_custom = QRadioButton("自定义目录:")
        self._radio_same.setChecked(True)
        _btn_group = QButtonGroup(self)
        _btn_group.addButton(self._radio_same)
        _btn_group.addButton(self._radio_custom)
        output_layout.addWidget(self._radio_same)
        output_layout.addWidget(self._radio_custom)

        folder_row = QHBoxLayout()
        self._folder_edit = QLineEdit()
        self._folder_edit.setPlaceholderText("选择输出文件夹...")
        self._folder_edit.setReadOnly(True)
        self._folder_edit.setEnabled(False)
        self._folder_browse_btn = QPushButton("浏览...")
        self._folder_browse_btn.setEnabled(False)
        self._folder_browse_btn.clicked.connect(self._browse_output_folder)
        folder_row.addWidget(self._folder_edit)
        folder_row.addWidget(self._folder_browse_btn)
        output_layout.addLayout(folder_row)
        self._radio_custom.toggled.connect(self._toggle_folder_row)
        layout.addWidget(output_group)

        if self._last_output_dir:
            self._radio_custom.setChecked(True)
            self._folder_edit.setText(self._last_output_dir)

        # ── BOM type + summary options (combined group) ─────────────────────
        bom_opts_group  = QGroupBox("BOM类型与汇总选项")
        bom_opts_layout = QVBoxLayout(bom_opts_group)

        # Radio buttons (层级 vs 汇总)
        type_row = QHBoxLayout()
        self._bom_type_btn_group = QButtonGroup(self)
        self._radio_hierarchical = QRadioButton("层级BOM")
        self._radio_summary      = QRadioButton("汇总BOM")
        if self._summarize:
            self._radio_summary.setChecked(True)
        else:
            self._radio_hierarchical.setChecked(True)
        self._bom_type_btn_group.addButton(self._radio_hierarchical)
        self._bom_type_btn_group.addButton(self._radio_summary)
        type_row.addWidget(self._radio_hierarchical)
        type_row.addWidget(self._radio_summary)
        type_row.addStretch()
        bom_opts_layout.addLayout(type_row)
        self._radio_summary.toggled.connect(self._on_bom_type_changed)

        # Summary-only options (shown only in summary mode, no sub-groupbox)
        self._include_assemblies_chk = QCheckBox("包含产品和部件（子装配体）")
        self._include_assemblies_chk.setToolTip(
            "勾选后，汇总BOM中也会列出产品和部件（子装配体），而不仅限于零件。"
        )
        self._include_assemblies_chk.setChecked(self._summary_include_assemblies)
        self._include_assemblies_chk.toggled.connect(self._on_include_assemblies_toggled)
        bom_opts_layout.addWidget(self._include_assemblies_chk)

        sort_row = QHBoxLayout()
        self._sort_row_label = QLabel("排序列:")
        self._sort_col_combo = QComboBox()
        sort_row.addWidget(self._sort_row_label)
        sort_row.addWidget(self._sort_col_combo)
        sort_row.addStretch()
        bom_opts_layout.addLayout(sort_row)

        self._include_assemblies_chk.setVisible(self._summarize)
        self._sort_row_label.setVisible(self._summarize)
        self._sort_col_combo.setVisible(self._summarize)
        layout.addWidget(bom_opts_group)
        col_group  = QGroupBox("导出列（拖动以排序）")
        col_outer  = QVBoxLayout(col_group)
        col_layout = QHBoxLayout()

        avail_layout = QVBoxLayout()
        avail_layout.addWidget(QLabel("可用列:"))
        self._avail_list = QListWidget()
        self._avail_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self._avail_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        avail_layout.addWidget(self._avail_list)
        col_layout.addLayout(avail_layout)

        arrow_layout = QVBoxLayout()
        arrow_layout.addStretch()
        add_btn    = QPushButton("→")
        remove_btn = QPushButton("←")
        up_btn     = QPushButton("↑")
        down_btn   = QPushButton("↓")
        for btn in (add_btn, remove_btn, up_btn, down_btn):
            btn.setFixedWidth(36)
        add_btn.clicked.connect(self._add_column)
        remove_btn.clicked.connect(self._remove_column)
        up_btn.clicked.connect(self._move_up)
        down_btn.clicked.connect(self._move_down)
        arrow_layout.addWidget(add_btn)
        arrow_layout.addWidget(remove_btn)
        arrow_layout.addSpacing(10)
        arrow_layout.addWidget(up_btn)
        arrow_layout.addWidget(down_btn)
        arrow_layout.addStretch()
        col_layout.addLayout(arrow_layout)

        selected_layout = QVBoxLayout()
        selected_layout.addWidget(QLabel("已选列:"))
        self._selected_list = QListWidget()
        self._selected_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self._selected_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        selected_layout.addWidget(self._selected_list)
        col_layout.addLayout(selected_layout)
        col_outer.addLayout(col_layout)
        layout.addWidget(col_group, 1)

        # Populate column lists
        saved = self._settings.value("selected_columns", BOM_DEFAULT_COLUMNS)
        if isinstance(saved, str):
            saved = [saved]
        all_known = BOM_ALL_COLUMNS + [
            c for c in PRESET_USER_REF_PROPERTIES if c not in BOM_ALL_COLUMNS
        ] + [
            c for c in self._custom_columns
            if c not in BOM_ALL_COLUMNS and c not in PRESET_USER_REF_PROPERTIES
        ]
        for col in saved:
            if col in all_known:
                self._selected_list.addItem(self._make_col_item(col))
        for col in all_known:
            if col not in saved:
                self._avail_list.addItem(self._make_col_item(col))

        # Populate sort column combo (after all_known is built)
        for col in all_known:
            self._sort_col_combo.addItem(
                BOM_COLUMN_DISPLAY_NAMES.get(col, col), col
            )
        saved_sort_idx = self._sort_col_combo.findData(self._summary_sort_column)
        if saved_sort_idx >= 0:
            self._sort_col_combo.setCurrentIndex(saved_sort_idx)
        self._sort_col_combo.currentIndexChanged.connect(self._on_sort_col_changed)

        # ── Action buttons ──────────────────────────────────────────────────
        action_row  = QHBoxLayout()
        confirm_btn = QPushButton("导出")
        confirm_btn.setDefault(True)
        cancel_btn  = QPushButton("取消")
        confirm_btn.clicked.connect(self._confirm)
        cancel_btn.clicked.connect(self.reject)
        action_row.addStretch()
        action_row.addWidget(confirm_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    # ── Helpers ──────────────────────────────────────────────────────────────

    @staticmethod
    def _make_col_item(internal_name: str) -> QListWidgetItem:
        item = QListWidgetItem(
            BOM_COLUMN_DISPLAY_NAMES.get(internal_name, internal_name)
        )
        item.setData(Qt.ItemDataRole.UserRole, internal_name)
        return item

    @staticmethod
    def _item_internal(item: QListWidgetItem) -> str:
        data = item.data(Qt.ItemDataRole.UserRole)
        return data if data else item.text()

    def _toggle_folder_row(self, checked: bool) -> None:
        self._folder_edit.setEnabled(checked)
        self._folder_browse_btn.setEnabled(checked)

    def _toggle_source_row(self, active_checked: bool) -> None:
        self._file_edit.setEnabled(not active_checked)
        self._file_browse_btn.setEnabled(not active_checked)
        if active_checked and self._radio_same.isChecked():
            self._radio_custom.setChecked(True)

    def _browse_file(self) -> None:
        file, _ = QFileDialog.getOpenFileName(
            self, "选择CATProduct文件",
            self._last_browse_dir,
            "*.CATProduct (*.CATProduct);;All Files (*)",
        )
        if file:
            self._file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)

    def _browse_output_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "选择输出文件夹", self._last_output_dir
        )
        if folder:
            self._folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    def _add_column(self) -> None:
        for item in self._avail_list.selectedItems():
            internal = self._item_internal(item)
            self._avail_list.takeItem(self._avail_list.row(item))
            self._selected_list.addItem(self._make_col_item(internal))

    def _remove_column(self) -> None:
        for item in self._selected_list.selectedItems():
            internal = self._item_internal(item)
            self._selected_list.takeItem(self._selected_list.row(item))
            self._avail_list.addItem(self._make_col_item(internal))

    def _move_up(self) -> None:
        row = self._selected_list.currentRow()
        if row > 0:
            item = self._selected_list.takeItem(row)
            self._selected_list.insertItem(row - 1, item)
            self._selected_list.setCurrentRow(row - 1)

    def _move_down(self) -> None:
        row = self._selected_list.currentRow()
        if row < self._selected_list.count() - 1:
            item = self._selected_list.takeItem(row)
            self._selected_list.insertItem(row + 1, item)
            self._selected_list.setCurrentRow(row + 1)

    def _on_bom_type_changed(self, summary_checked: bool) -> None:
        """When BOM type switches, move the 'Level' column between the lists."""
        self._summarize = summary_checked
        self._settings.setValue("summarize", summary_checked)

        # Show/hide summary options
        self._include_assemblies_chk.setVisible(summary_checked)
        self._sort_row_label.setVisible(summary_checked)
        self._sort_col_combo.setVisible(summary_checked)

        if summary_checked:
            # Move all "Level" items from selected to available
            # (iterate in reverse so takeItem indices stay valid)
            for i in range(self._selected_list.count() - 1, -1, -1):
                item = self._selected_list.item(i)
                if self._item_internal(item) == "Level":
                    self._selected_list.takeItem(i)
                    self._avail_list.addItem(self._make_col_item("Level"))

    def _on_include_assemblies_toggled(self, checked: bool) -> None:
        self._summary_include_assemblies = checked
        self._settings.setValue("summary_include_assemblies", checked)

    def _on_sort_col_changed(self, _index: int) -> None:
        col = self._sort_col_combo.currentData()
        if col:
            self._summary_sort_column = col
            self._settings.setValue("summary_sort_column", col)

    def _confirm(self) -> None:
        use_active = self._radio_active.isChecked()
        if use_active:
            file_path = None
        else:
            file_path = self._file_edit.text().strip()
            if not file_path:
                QMessageBox.warning(self, "未选择文件", "请选择一个CATProduct文件。")
                return

        selected_cols = [
            self._item_internal(self._selected_list.item(i))
            for i in range(self._selected_list.count())
        ]
        if not selected_cols:
            QMessageBox.warning(self, "未选择列", "请至少选择一列进行导出。")
            return
        self._settings.setValue("selected_columns", selected_cols)

        if self._radio_same.isChecked() and not use_active:
            output_folder = None
        else:
            output_folder = self._folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(
                    self, "未选择输出文件夹",
                    "请选择一个输出文件夹（使用活动文档时需指定）。",
                )
                return

        summarize = self._radio_summary.isChecked()
        label_text = "正在导出汇总BOM，请稍候…" if summarize else "正在导出BOM，请稍候…"
        progress = QProgressDialog(label_text, None, 0, 0, self)
        progress.setWindowTitle("导出BOM汇总" if summarize else "导出BOM")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(300)
        progress.setValue(0)

        def _on_row_collected(count: int) -> None:
            base = "正在导出汇总BOM，请稍候…" if summarize else "正在导出BOM，请稍候…"
            progress.setLabelText(f"{base} 已读取 {count} 个节点")
            progress.repaint()
            QApplication.processEvents()

        try:
            export_bom_to_excel(
                [file_path], output_folder,
                columns=selected_cols,
                custom_columns=self._custom_columns,
                row_progress_callback=_on_row_collected,
                summarize=summarize,
                summary_include_assemblies=self._summary_include_assemblies,
                summary_sort_column=self._summary_sort_column or None,
            )
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "导出失败", f"导出BOM时出错：\n{e}")
            return
        finally:
            progress.close()

        QMessageBox.information(self, "导出成功", "BOM已成功导出为Excel文件。")
        self.accept()
