"""
Generic file-conversion dialog.

Provides:
- FileConvertDialog – a reusable dialog for "pick files → run conversion function"
  workflows.  Used for drawing-to-PDF, part-to-STEP, and template stamping.

Supports drag & drop of files onto the file list and shows a QProgressBar
during conversion.
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QFileDialog,
    QAbstractItemView, QRadioButton, QButtonGroup, QLineEdit, QGroupBox,
    QCheckBox, QPushButton, QMessageBox, QProgressBar, QApplication, QWidget,
)
from PySide6.QtCore import Qt, QSettings

logger = logging.getLogger(__name__)


class FileConvertDialog(QDialog):
    """Dialog for selecting files and running a batch conversion function.

    Parameters
    ----------
    title:
        Window title.
    file_label:
        Label shown above the file list.
    file_filter:
        File filter string for the open-file dialog (Qt format).
    no_files_msg:
        Warning shown when the user confirms without selecting any files.
    conversion_fn:
        Callable invoked as ``conversion_fn(files, output_folder)`` or
        ``conversion_fn(files, output_folder, prefix=..., suffix=...)``
        when *show_prefix_option* is ``True``.
    settings_key:
        Key suffix used to persist settings in QSettings.
    show_prefix_option:
        Whether to show the prefix / suffix input rows.
    prefix:
        Default prefix value.
    note:
        Optional grey hint text displayed below the controls.
    """

    def __init__(
        self,
        parent=None,
        title: str = "Convert",
        file_label: str = "已选文件:",
        file_filter: str = "All Files (*)",
        no_files_msg: str = "请至少选择一个文件。",
        conversion_fn=None,
        settings_key: str = "default",
        show_prefix_option: bool = False,
        prefix: str = "",
        note: str = "",
        show_update_option: bool = False,
        show_active_doc_option: bool = False,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(520, 450)

        self._file_filter           = file_filter
        self._no_files_msg          = no_files_msg
        self._conversion_fn         = conversion_fn
        self._show_prefix_option    = show_prefix_option
        self._show_update_option    = show_update_option
        self._show_active_doc_option = show_active_doc_option

        self._settings         = QSettings("CATIACompanion", f"ConvertDialog_{settings_key}")
        self._last_browse_dir  = self._settings.value("last_browse_dir", "")
        self._last_output_dir  = self._settings.value("last_output_dir", "")
        self._is_stamp_dialog  = settings_key == "StampPartTemplate"

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── "Use active document" option ─────────────────────────────────────
        if show_active_doc_option:
            self._use_active_chk = QCheckBox("使用当前CATIA活动文档（不选择文件）")
            self._use_active_chk.toggled.connect(self._toggle_file_section)
            layout.addWidget(self._use_active_chk)
        else:
            self._use_active_chk = None

        # ── File list section (can be hidden in active-doc mode) ─────────────
        self._file_section = QWidget()
        file_section_layout = QVBoxLayout(self._file_section)
        file_section_layout.setContentsMargins(0, 0, 0, 0)
        file_section_layout.setSpacing(6)

        file_section_layout.addWidget(QLabel(file_label))

        self._file_list = QListWidget()
        self._file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        # Enable drag & drop – users can drop files directly onto the list
        self.setAcceptDrops(True)
        # Restore previously saved file list
        saved_files: list = self._settings.value("saved_files", []) or []
        if isinstance(saved_files, str):
            saved_files = [saved_files]
        for f in saved_files:
            if Path(f).exists():
                self._file_list.addItem(f)
        file_section_layout.addWidget(self._file_list)

        btn_row = QHBoxLayout()
        browse_btn     = QPushButton("浏览...")
        remove_btn     = QPushButton("移除所选")
        remove_all_btn = QPushButton("全部移除")
        browse_btn.clicked.connect(self._browse_files)
        remove_btn.clicked.connect(self._remove_selected)
        remove_all_btn.clicked.connect(self._remove_all)
        btn_row.addWidget(browse_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addWidget(remove_all_btn)
        btn_row.addStretch()
        file_section_layout.addLayout(btn_row)

        layout.addWidget(self._file_section)

        # ── Output folder (hidden for stamp dialog) ─────────────────────────
        if not self._is_stamp_dialog:
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
        else:
            self._radio_same  = None
            self._folder_edit = None

        # ── Prefix / suffix rows ────────────────────────────────────────────
        if show_prefix_option:
            saved_add_prefix   = self._settings.value("add_prefix", True)
            saved_prefix_value = self._settings.value("prefix_value", prefix)
            if isinstance(saved_add_prefix, str):
                saved_add_prefix = saved_add_prefix.lower() == "true"

            prefix_row = QHBoxLayout()
            self._prefix_checkbox = QCheckBox("添加前缀:")
            self._prefix_checkbox.setChecked(saved_add_prefix)
            self._prefix_edit = QLineEdit(saved_prefix_value)
            self._prefix_edit.setEnabled(saved_add_prefix)
            self._prefix_checkbox.toggled.connect(self._prefix_edit.setEnabled)
            prefix_row.addWidget(self._prefix_checkbox)
            prefix_row.addWidget(self._prefix_edit)
            layout.addLayout(prefix_row)

            saved_add_suffix   = self._settings.value("add_suffix", False)
            saved_suffix_value = self._settings.value("suffix_value", "")
            if isinstance(saved_add_suffix, str):
                saved_add_suffix = saved_add_suffix.lower() == "true"

            suffix_row = QHBoxLayout()
            self._suffix_checkbox = QCheckBox("添加后缀:")
            self._suffix_checkbox.setChecked(saved_add_suffix)
            self._suffix_edit = QLineEdit(saved_suffix_value)
            self._suffix_edit.setEnabled(saved_add_suffix)
            self._suffix_checkbox.toggled.connect(self._suffix_edit.setEnabled)
            suffix_row.addWidget(self._suffix_checkbox)
            suffix_row.addWidget(self._suffix_edit)
            layout.addLayout(suffix_row)
        else:
            self._prefix_checkbox = None
            self._prefix_edit     = None
            self._suffix_checkbox = None
            self._suffix_edit     = None

        # ── Update-before-export option (e.g. CATDrawing → PDF) ─────────────
        if show_update_option:
            saved_update = self._settings.value("update_before_export", False)
            if isinstance(saved_update, str):
                saved_update = saved_update.lower() == "true"
            self._update_checkbox = QCheckBox("更新图纸后再输出")
            self._update_checkbox.setToolTip(
                "导出PDF前先强制更新CATDrawing中每一页的所有视图，"
                "确保导出结果与最新模型状态一致。"
            )
            self._update_checkbox.setChecked(saved_update)
            layout.addWidget(self._update_checkbox)
        else:
            self._update_checkbox = None

        if note:
            note_label = QLabel(note)
            note_label.setWordWrap(True)
            note_label.setStyleSheet("color: gray; font-size: 11px;")
            layout.addWidget(note_label)

        # ── Progress bar (hidden until conversion starts) ─────────────────
        self._progress_bar = QProgressBar()
        self._progress_bar.setTextVisible(True)
        self._progress_bar.setVisible(False)
        layout.addWidget(self._progress_bar)

        # ── Action buttons ──────────────────────────────────────────────────
        action_row  = QHBoxLayout()
        self._confirm_btn = QPushButton("确认")
        self._confirm_btn.setDefault(True)
        cancel_btn  = QPushButton("取消")
        self._confirm_btn.clicked.connect(self._confirm)
        cancel_btn.clicked.connect(self.reject)
        action_row.addStretch()
        action_row.addWidget(self._confirm_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    # ── Drag & drop support ─────────────────────────────────────────────────

    def _accepted_extensions(self) -> set[str]:
        """Parse the file_filter to extract accepted file extensions."""
        import re
        exts: set[str] = set()
        for m in re.finditer(r"\*(\.\w+)", self._file_filter):
            exts.add(m.group(1).lower())
        return exts

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        accepted = self._accepted_extensions()
        existing = {self._file_list.item(i).text()
                    for i in range(self._file_list.count())}
        added = 0
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if not path:
                continue
            # Accept the file if the filter is a wildcard or the extension matches
            if accepted and Path(path).suffix.lower() not in accepted:
                continue
            if path not in existing:
                self._file_list.addItem(path)
                existing.add(path)
                added += 1
        if added:
            self._persist_file_list()
        event.acceptProposedAction()

    # ── File list management ─────────────────────────────────────────────────

    def _browse_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", self._last_browse_dir, self._file_filter
        )
        if files:
            self._last_browse_dir = str(Path(files[0]).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)
        existing = {self._file_list.item(i).text()
                    for i in range(self._file_list.count())}
        for f in files:
            if f not in existing:
                self._file_list.addItem(f)
        self._persist_file_list()

    def _remove_selected(self) -> None:
        for item in self._file_list.selectedItems():
            self._file_list.takeItem(self._file_list.row(item))
        self._persist_file_list()

    def _remove_all(self) -> None:
        self._file_list.clear()
        self._persist_file_list()

    def _persist_file_list(self) -> None:
        files = [self._file_list.item(i).text()
                 for i in range(self._file_list.count())]
        self._settings.setValue("saved_files", files)

    # ── Output folder ────────────────────────────────────────────────────────

    def _toggle_file_section(self, use_active: bool) -> None:
        """Show/hide the file-list section when active-doc mode is toggled."""
        self._file_section.setVisible(not use_active)

    def _toggle_folder_row(self, checked: bool) -> None:
        self._folder_edit.setEnabled(checked)
        self._folder_browse_btn.setEnabled(checked)

    def _browse_output_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "选择输出文件夹", self._last_output_dir
        )
        if folder:
            self._folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    # ── Confirm ──────────────────────────────────────────────────────────────

    def _confirm(self) -> None:
        use_active = (
            self._use_active_chk is not None and self._use_active_chk.isChecked()
        )

        if use_active:
            # Resolve the path of the current CATIA active document
            try:
                from pycatia import catia as _catia
                _caa = _catia()
                active_path = _caa.application.active_document.full_name
            except Exception as e:
                QMessageBox.warning(
                    self, "无法获取活动文档",
                    f"无法从CATIA获取当前活动文档路径：\n{e}\n\n请确保CATIA已启动且有活动文档。",
                )
                return
            files = [active_path]
            output_folder = None
        else:
            files = [self._file_list.item(i).text()
                     for i in range(self._file_list.count())]
            if not files:
                QMessageBox.warning(self, "未选择文件", self._no_files_msg)
                return

            if self._radio_same is None:
                output_folder = None
            elif self._radio_same.isChecked():
                output_folder = None
            else:
                output_folder = self._folder_edit.text().strip()
                if not output_folder:
                    QMessageBox.warning(self, "未选择输出文件夹", "请选择一个输出文件夹。")
                    return

        # Show progress bar and disable confirm button during conversion
        total = len(files)
        self._progress_bar.setRange(0, total)
        self._progress_bar.setValue(0)
        self._progress_bar.setVisible(True)
        self._confirm_btn.setEnabled(False)

        def _progress(i: int, _total: int) -> None:
            self._progress_bar.setValue(i)
            self._progress_bar.setFormat(
                f"正在转换 ({i + 1}/{_total}): {Path(files[i]).name}"
            )
            QApplication.processEvents()

        success_count = 0
        try:
            if self._prefix_checkbox is not None:
                prefix_value = (
                    self._prefix_edit.text() if self._prefix_checkbox.isChecked() else ""
                )
                suffix_value = (
                    self._suffix_edit.text() if self._suffix_checkbox.isChecked() else ""
                )
                kwargs: dict = dict(
                    prefix=prefix_value, suffix=suffix_value,
                    progress_callback=_progress,
                )
                if self._update_checkbox is not None:
                    kwargs["update_before_export"] = self._update_checkbox.isChecked()
                if use_active:
                    kwargs["keep_open"] = True
                success_count = self._conversion_fn(files, output_folder, **kwargs)
            else:
                kwargs = dict(progress_callback=_progress)
                if self._update_checkbox is not None:
                    kwargs["update_before_export"] = self._update_checkbox.isChecked()
                if use_active:
                    kwargs["keep_open"] = True
                success_count = self._conversion_fn(files, output_folder, **kwargs)
        except Exception as e:
            logger.error("Conversion failed: %s", e)

        # Persist prefix/suffix and update settings after conversion
        if self._prefix_checkbox is not None:
            self._settings.setValue("add_prefix",   self._prefix_checkbox.isChecked())
            self._settings.setValue("prefix_value", self._prefix_edit.text())
            self._settings.setValue("add_suffix",   self._suffix_checkbox.isChecked())
            self._settings.setValue("suffix_value", self._suffix_edit.text())
        if self._update_checkbox is not None:
            self._settings.setValue("update_before_export", self._update_checkbox.isChecked())

        self._progress_bar.setValue(total)
        self._progress_bar.setFormat(f"完成 ({success_count}/{total})")
        self._confirm_btn.setEnabled(True)

        QMessageBox.information(
            self, "导出完成",
            f"已成功导出 {success_count} / {total} 个文件。",
        )
        self.accept()
