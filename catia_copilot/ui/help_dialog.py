"""
帮助对话框 – 在可滚动的富文本窗口中显示用户文档。
HTML 内容存放于同目录下的 help_doc.html，运行时动态填入版本信息。
"""

import pathlib
import string

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QTextBrowser, QPushButton, QHBoxLayout,
)

from catia_copilot.constants import APP_NAME, APP_VERSION, APP_AUTHOR, APP_CONTACT, MAX_INERTIA_INDEX
from catia_copilot.utils import resource_path

_UI_DIR = pathlib.Path(__file__).parent

_HELP_HTML = string.Template(
    (_UI_DIR / "help_doc.html").read_text(encoding="utf-8")
).safe_substitute(
    APP_NAME=APP_NAME,
    APP_VERSION=APP_VERSION,
    APP_AUTHOR=APP_AUTHOR,
    APP_CONTACT=APP_CONTACT,
    MAX_INERTIA_INDEX=MAX_INERTIA_INDEX,
)


class HelpDialog(QDialog):
    """Scrollable help dialog with rich-text documentation."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} — 帮助文档")
        self.resize(700, 560)
        self.setMinimumSize(480, 360)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        # Allow relative <img> paths in the HTML to resolve from the resources folder
        browser.setSearchPaths([str(resource_path("resources"))])
        browser.setHtml(_HELP_HTML)
        layout.addWidget(browser)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        layout.addLayout(btn_layout)
