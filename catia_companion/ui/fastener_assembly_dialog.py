"""
Fastener quick-assembly dialog.

Provides:
- FastenerAssemblyDialog – dialog for specifying a fastener CATPart and
  launching the ``fastener_assembly.catvbs`` macro to interactively place
  fastener instances in an assembly.
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QRadioButton, QButtonGroup, QGroupBox,
    QApplication,
)
from PySide6.QtCore import QSettings

logger = logging.getLogger(__name__)


class FastenerAssemblyDialog(QDialog):
    """Dialog for the quick fastener-assembly workflow.

    The dialog lets the user pick a fastener CATPart (either the current
    active CATIA document or a file on disk), displays its Part Number,
    and launches the ``fastener_assembly.catvbs`` macro that drives the
    interactive CATIA assembly loop.
    """

    def __init__(self, parent=None, *, execute_fn=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("快速装配紧固件")
        self.setMinimumSize(480, 300)

        self._settings = QSettings("CATIACompanion", "FastenerAssemblyDialog")
        self._execute_fn = execute_fn
        self._fastener_path: str = ""
        self._fastener_pn: str = ""

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Fastener source ─────────────────────────────────────────────
        src_group = QGroupBox("紧固件来源")
        src_layout = QVBoxLayout(src_group)

        self._btn_group = QButtonGroup(self)
        self._radio_active = QRadioButton("使用当前CATIA活动文档")
        self._radio_file = QRadioButton("选择文件:")
        self._btn_group.addButton(self._radio_active)
        self._btn_group.addButton(self._radio_file)
        self._radio_active.setChecked(True)

        src_layout.addWidget(self._radio_active)

        file_row = QHBoxLayout()
        file_row.addWidget(self._radio_file)
        self._file_edit = QLineEdit()
        self._file_edit.setReadOnly(True)
        self._file_edit.setPlaceholderText("选择一个CATPart文件...")
        saved_path = self._settings.value("last_file", "")
        if saved_path:
            self._file_edit.setText(saved_path)
        self._file_browse_btn = QPushButton("浏览...")
        self._file_browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(self._file_browse_btn)
        src_layout.addLayout(file_row)

        layout.addWidget(src_group)

        # ── Info display ────────────────────────────────────────────────
        info_group = QGroupBox("信息")
        info_layout = QVBoxLayout(info_group)

        pn_row = QHBoxLayout()
        pn_row.addWidget(QLabel("紧固件零件编号："))
        self._pn_label = QLabel("（未读取）")
        self._pn_label.setStyleSheet("font-weight: bold;")
        pn_row.addWidget(self._pn_label)
        pn_row.addStretch()
        info_layout.addLayout(pn_row)

        layout.addWidget(info_group)

        # ── Hint ────────────────────────────────────────────────────────
        hint = QLabel(
            "点击"开始装配"后，请按照 CATIA 弹窗提示依次选取紧固件的圆柱面"
            "和端面，再切换到装配体并选择目标产品节点，之后即可连续装配紧固件。\n"
            "在选取过程中按 ESC 可随时终止。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # ── Status ──────────────────────────────────────────────────────
        self._status_label = QLabel("状态：就绪")
        self._status_label.setStyleSheet("color: #555;")
        layout.addWidget(self._status_label)

        # ── Action buttons ──────────────────────────────────────────────
        action_row = QHBoxLayout()
        self._start_btn = QPushButton("开始装配")
        self._start_btn.setDefault(True)
        self._start_btn.clicked.connect(self._start_assembly)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        action_row.addStretch()
        action_row.addWidget(self._start_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    # ── File browsing ───────────────────────────────────────────────────

    def _browse_file(self) -> None:
        last = self._settings.value("last_file", "")
        start_dir = str(Path(last).parent) if last else ""
        file, _ = QFileDialog.getOpenFileName(
            self, "选择紧固件CATPart文件", start_dir,
            "CATPart (*.CATPart);;All Files (*)",
        )
        if file:
            self._file_edit.setText(file)
            self._settings.setValue("last_file", file)
            self._radio_file.setChecked(True)

    # ── Start assembly ──────────────────────────────────────────────────

    def _start_assembly(self) -> None:
        """Read fastener info, then launch the VBScript macro."""

        use_active = self._radio_active.isChecked()

        # Resolve fastener path
        if use_active:
            fastener_path = ""  # macro will use CATIA.ActiveDocument
        else:
            fastener_path = self._file_edit.text().strip()
            if not fastener_path:
                QMessageBox.warning(
                    self, "未选择文件", "请选择一个紧固件CATPart文件。"
                )
                return
            if not Path(fastener_path).exists():
                QMessageBox.warning(
                    self, "文件不存在",
                    f"文件不存在：\n{fastener_path}",
                )
                return

        # Read fastener part number via COM
        self._status_label.setText("状态：正在读取紧固件信息…")
        QApplication.processEvents()

        try:
            from pycatia import catia as _catia
            caa = _catia()
            app = caa.application

            if use_active:
                pn = app.active_document.product.part_number
                fastener_path = app.active_document.full_name
            else:
                # Leave the document open – the VBScript macro needs it.
                doc = app.documents.open(fastener_path)
                pn = doc.product.part_number
        except Exception as e:
            self._status_label.setText("状态：读取失败")
            QMessageBox.critical(
                self, "读取失败",
                f"无法读取紧固件信息：\n{e}\n\n请确保CATIA已启动。",
            )
            return

        self._fastener_path = fastener_path
        self._fastener_pn = pn
        self._pn_label.setText(pn)

        # Launch the VBScript macro
        self._status_label.setText("状态：正在装配（按ESC终止）…")
        QApplication.processEvents()

        if self._execute_fn is None:
            QMessageBox.warning(self, "内部错误", "未指定宏执行函数。")
            return

        try:
            self._execute_fn(fastener_path)
            self._status_label.setText("状态：装配完成")
        except Exception as e:
            logger.error("Fastener assembly macro failed: %s", e)
            self._status_label.setText("状态：装配中断")
            QMessageBox.critical(
                self, "宏执行失败",
                f"运行宏时出错：\n{e}\n\n请确保CATIA已启动。",
            )
