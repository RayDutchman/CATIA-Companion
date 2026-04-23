"""文件重命名/移动对话框，供 BOM 编辑器使用。"""

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QWidget, QFormLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox, QFileDialog,
)

from catia_companion.constants import PART_NUMBER_VALID_PATTERN


class _FileRenameDialog(QDialog):
    """Dialog for renaming or moving a single CATIA file via CATIA SaveAs.

    Lets the user change the file stem (name without extension) and/or the
    target directory independently.  Validates the new stem with
    :data:`~catia_companion.constants.PART_NUMBER_VALID_PATTERN` and creates
    the target directory on demand.
    """

    def __init__(self, current_fp: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("编辑文件名/路径")
        self.setMinimumWidth(540)
        self._current_fp = current_fp
        self._p          = Path(current_fp)

        layout = QFormLayout(self)
        layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        layout.setSpacing(8)

        # Current path (read-only)
        cur_label = QLabel(current_fp)
        cur_label.setWordWrap(True)
        cur_label.setStyleSheet("color: #555;")
        layout.addRow("当前路径：", cur_label)

        # New filename (stem only; extension is preserved automatically)
        self._name_edit = QLineEdit(self._p.stem)
        layout.addRow(f"新文件名（不含扩展名 {self._p.suffix}）：", self._name_edit)

        # New directory (with browse button)
        dir_widget = QWidget()
        dir_layout = QHBoxLayout(dir_widget)
        dir_layout.setContentsMargins(0, 0, 0, 0)
        self._dir_edit = QLineEdit(str(self._p.parent))
        dir_btn        = QPushButton("浏览…")
        dir_btn.setFixedWidth(64)
        dir_btn.clicked.connect(self._browse_dir)
        dir_layout.addWidget(self._dir_edit)
        dir_layout.addWidget(dir_btn)
        layout.addRow("新目录：", dir_widget)

        # Path preview
        self._preview_label = QLabel()
        self._preview_label.setWordWrap(True)
        self._preview_label.setStyleSheet("color: #333; font-style: italic;")
        layout.addRow("新路径预览：", self._preview_label)

        self._name_edit.textChanged.connect(self._update_preview)
        self._dir_edit.textChanged.connect(self._update_preview)
        self._update_preview()

        # Buttons
        btn_widget = QWidget()
        btn_row    = QHBoxLayout(btn_widget)
        btn_row.setContentsMargins(0, 0, 0, 0)
        btn_row.addStretch()
        ok_btn     = QPushButton("确认")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self._validate_and_accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addRow(btn_widget)

    # ── Properties ─────────────────────────────────────────────────────────

    @property
    def new_stem(self) -> str:
        return self._name_edit.text().strip()

    @property
    def new_dir(self) -> str:
        return self._dir_edit.text().strip()

    @property
    def new_path(self) -> str:
        stem      = self.new_stem or self._p.stem
        directory = self.new_dir  or str(self._p.parent)
        return str(Path(directory) / (stem + self._p.suffix))

    # ── Slots ───────────────────────────────────────────────────────────────

    def _update_preview(self) -> None:
        self._preview_label.setText(self.new_path)

    def _browse_dir(self) -> None:
        d = QFileDialog.getExistingDirectory(
            self, "选择目标目录",
            self._dir_edit.text() or str(self._p.parent),
        )
        if d:
            self._dir_edit.setText(d)

    def _validate_and_accept(self) -> None:
        stem = self.new_stem or self._p.stem
        if stem != self._p.stem and not PART_NUMBER_VALID_PATTERN.fullmatch(stem):
            QMessageBox.warning(
                self, "文件名含非法字符",
                f"文件名 「{stem}」 含有非法字符。\n"
                "不允许：控制字符、非ASCII字符，以及Windows文件名禁用字符"
                "（\\ / : * ? \" < > |）。",
            )
            return
        new_p = Path(self.new_path)
        if new_p.resolve() == self._p.resolve():
            QMessageBox.warning(self, "路径未改变", "新路径与当前路径相同，无需操作。")
            return
        dest_dir = new_p.parent
        if not dest_dir.exists():
            ret = QMessageBox.question(
                self, "目录不存在",
                f"目标目录不存在：\n{dest_dir}\n\n是否创建该目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                QMessageBox.critical(self, "创建目录失败", f"无法创建目录：\n{exc}")
                return
        self.accept()
