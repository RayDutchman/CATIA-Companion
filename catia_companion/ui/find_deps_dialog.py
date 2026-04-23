"""
查找依赖项对话框。

提供：
- FindDependenciesDialog – 通过 COM 打开 CATIA 文件并列出它依赖的所有文档。
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QPlainTextEdit, QFileDialog, QMessageBox, QApplication,
)
from PySide6.QtCore import QSettings

from catia_companion.catia.dependencies import find_dependencies

logger = logging.getLogger(__name__)


class FindDependenciesDialog(QDialog):
    """发现 CATIA 文档依赖的所有文件的对话框。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("查找所有依赖项")
        self.setMinimumSize(540, 420)

        self._settings = QSettings("CATIACompanion", "FindDependenciesDialog")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        layout.addWidget(
            QLabel("目标CATIA文件（CATPart / CATProduct / CATDrawing）:")
        )

        target_row = QHBoxLayout()
        self._target_edit = QLineEdit()
        self._target_edit.setReadOnly(True)
        self._target_edit.setPlaceholderText("选择目标CATIA文件...")
        self._target_edit.setText(self._settings.value("last_target", ""))
        target_browse_btn = QPushButton("浏览...")
        target_browse_btn.clicked.connect(self._browse_target)
        target_row.addWidget(self._target_edit)
        target_row.addWidget(target_browse_btn)
        layout.addLayout(target_row)

        hint = QLabel(
            "通过CATIA COM打开文件并自动收集所有引用文档，请确保CATIA已运行。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        action_row = QHBoxLayout()
        self._search_btn = QPushButton("开始搜索")
        self._search_btn.setDefault(True)
        self._search_btn.clicked.connect(self._start_search)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(self._search_btn)
        action_row.addWidget(cancel_btn)
        action_row.addStretch()
        layout.addLayout(action_row)

        layout.addWidget(QLabel("找到的依赖项："))
        self._result_view = QPlainTextEdit()
        self._result_view.setReadOnly(True)
        self._result_view.setMinimumHeight(150)
        layout.addWidget(self._result_view)

        copy_btn = QPushButton("复制结果")
        copy_btn.clicked.connect(self._copy_results)
        layout.addWidget(copy_btn)

    def _browse_target(self) -> None:
        last       = self._settings.value("last_target", "")
        start_dir  = str(Path(last).parent) if last else ""
        file, _    = QFileDialog.getOpenFileName(
            self, "选择目标CATIA文件", start_dir,
            "*.CATPart *.CATProduct *.CATDrawing"
            " (*.CATPart *.CATProduct *.CATDrawing);;All Files (*)",
        )
        if file:
            self._target_edit.setText(file)
            self._settings.setValue("last_target", file)

    def _start_search(self) -> None:
        target = self._target_edit.text().strip()
        if not target:
            QMessageBox.warning(self, "未选择目标文件", "请先选择一个目标CATIA文件。")
            return
        if not Path(target).exists():
            QMessageBox.warning(self, "文件不存在", f"目标文件不存在：\n{target}")
            return

        self._search_btn.setEnabled(False)
        self._result_view.setPlainText("正在通过CATIA COM搜索依赖项，请稍候…")
        QApplication.processEvents()

        def _progress(msg: str) -> None:
            self._result_view.setPlainText(msg)
            QApplication.processEvents()

        try:
            results = find_dependencies(target, progress_callback=_progress)
        except Exception as e:
            self._search_btn.setEnabled(True)
            QMessageBox.critical(
                self, "搜索失败",
                f"通过CATIA COM搜索依赖项时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            self._result_view.setPlainText(f"搜索失败：{e}")
            return

        self._search_btn.setEnabled(True)

        if results:
            summary = f"搜索完成，共找到 {len(results)} 个依赖项：\n\n" + "\n".join(results)
        else:
            summary = "搜索完成，未找到任何依赖项。"
        self._result_view.setPlainText(summary)

    def _copy_results(self) -> None:
        text = self._result_view.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
