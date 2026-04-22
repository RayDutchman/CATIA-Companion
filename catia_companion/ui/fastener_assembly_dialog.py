"""
Fastener quick-assembly dialog.

Provides:
- FastenerAssemblyDialog – dialog for locating the ``fastener_assembly.catvba``
  VBA macro file and launching it inside the running CATIA instance.
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QApplication,
)
from PySide6.QtCore import QSettings

logger = logging.getLogger(__name__)


class FastenerAssemblyDialog(QDialog):
    """Launcher dialog for the fastener-assembly VBA macro.

    The dialog lets the user locate the ``fastener_assembly.catvba`` VBA macro
    file and launch it inside the running CATIA instance via
    ``SystemService.ExecuteScript``.  The macro itself handles all interactive
    steps (selecting edges, switching documents, placing fasteners).
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("快速装配紧固件")
        self.setMinimumWidth(480)

        self._settings = QSettings("CATIACompanion", "FastenerAssemblyDialog")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Macro file selector ──────────────────────────────────────────
        layout.addWidget(QLabel("VBA 宏文件（.catvba）："))

        file_row = QHBoxLayout()
        self._file_edit = QLineEdit()
        self._file_edit.setPlaceholderText("选择 fastener_assembly.catvba 文件…")
        saved = self._settings.value("last_catvba", "")
        if saved:
            self._file_edit.setText(saved)
        browse_btn = QPushButton("浏览…")
        browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(browse_btn)
        layout.addLayout(file_row)

        # ── Instructions ─────────────────────────────────────────────────
        hint = QLabel(
            "使用说明：\n"
            "1. 确保 CATIA 已启动，并已打开紧固件（CATPart）和目标装配体（CATProduct）文档。\n"
            '2. 选择宏文件后，点击\u201c开始装配\u201d，CATIA 宏将自动引导后续操作。\n'
            "3. 按照 CATIA 浮动面板的提示，依次选择紧固件圆形边和目标产品，"
            "即可连续放置紧固件实例。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # ── Status ───────────────────────────────────────────────────────
        self._status_label = QLabel("状态：就绪")
        self._status_label.setStyleSheet("color: #555;")
        layout.addWidget(self._status_label)

        # ── Action buttons ────────────────────────────────────────────────
        action_row = QHBoxLayout()
        self._start_btn = QPushButton("开始装配")
        self._start_btn.setDefault(True)
        self._start_btn.clicked.connect(self._start_assembly)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.reject)
        action_row.addStretch()
        action_row.addWidget(self._start_btn)
        action_row.addWidget(close_btn)
        layout.addLayout(action_row)

    # ── File browsing ────────────────────────────────────────────────────

    def _browse_file(self) -> None:
        last = self._settings.value("last_catvba", "")
        start_dir = str(Path(last).parent) if last else ""
        file, _ = QFileDialog.getOpenFileName(
            self, "选择 VBA 宏文件", start_dir,
            "CATIA VBA 宏 (*.catvba);;All Files (*)",
        )
        if file:
            self._file_edit.setText(file)
            self._settings.setValue("last_catvba", file)

    # ── Launch macro ──────────────────────────────────────────────────────

    def _start_assembly(self) -> None:
        """Validate the macro path and launch it via CATIA COM."""
        catvba_str = self._file_edit.text().strip()
        if not catvba_str:
            QMessageBox.warning(self, "未选择文件", "请先选择 .catvba 宏文件。")
            return
        catvba_path = Path(catvba_str)
        if not catvba_path.exists():
            QMessageBox.warning(
                self, "文件不存在",
                f"找不到宏文件：\n{catvba_path}\n\n请确认文件路径是否正确。",
            )
            return

        self._status_label.setText("状态：正在启动宏…")
        QApplication.processEvents()

        try:
            from pycatia import catia as _catia
            caa = _catia()
            app = caa.application
            # ExecuteScript signature:
            #   (LibraryName, LibraryType, ProgramName, FunctionName, Parameters)
            # LibraryType = 2 for a .catvba VBA project file.
            # LibraryName = full path to the .catvba file.
            # ProgramName = VBA module name (conventionally the file stem).
            module_name = catvba_path.stem
            app.com_object.SystemService.ExecuteScript(
                str(catvba_path), 2, module_name, "CATMain", [],
            )
            self._status_label.setText("状态：宏已完成")
        except Exception as e:
            logger.error("Fastener assembly macro failed: %s", e)
            self._status_label.setText("状态：运行失败")
            QMessageBox.critical(
                self, "宏执行失败",
                f"无法运行宏：\n{e}\n\n"
                "请确保：\n"
                "  1. CATIA 已启动\n"
                "  2. 所选文件为有效的 .catvba 宏文件\n"
                "  3. 宏模块名与文件名一致，且包含 CATMain 子程序",
            )
