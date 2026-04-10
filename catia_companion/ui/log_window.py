"""
Log window widget.

Provides:
- LogWindow – a floating, non-closeable QWidget that displays application logs
"""

import sys
import os
import subprocess
import logging

from PySide6.QtWidgets import QWidget, QVBoxLayout, QPlainTextEdit, QPushButton, QLabel, QMessageBox
from PySide6.QtCore import Qt

from catia_companion.logging_setup import LOG_FILE

logger = logging.getLogger(__name__)


class LogWindow(QWidget):
    """Floating log viewer window.

    Subscribes to :data:`~catia_companion.logging_setup.log_signal_emitter`
    to receive formatted log messages in real time.

    Closing the window hides it rather than destroying it so that the log
    history is preserved.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent, Qt.WindowType.Window)
        self.setWindowTitle("CATIA Companion – Log")
        self.resize(660, 420)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        self._log_view = QPlainTextEdit()
        self._log_view.setReadOnly(True)
        self._log_view.setStyleSheet(
            "background-color: #1e1e1e; color: #d4d4d4;"
            " font-family: Consolas, 'Courier New', monospace; font-size: 9pt;"
        )
        layout.addWidget(self._log_view)

        open_log_btn = QPushButton("打开日志文件")
        open_log_btn.clicked.connect(self._open_log_file)
        layout.addWidget(open_log_btn)

        log_path_label = QLabel(f"Log: {LOG_FILE}")
        log_path_label.setStyleSheet("color: gray; font-size: 9pt;")
        log_path_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        layout.addWidget(log_path_label)

    def append_log(self, message: str) -> None:
        """Append *message* to the log view and scroll to the bottom."""
        self._log_view.appendPlainText(message)
        self._log_view.verticalScrollBar().setValue(
            self._log_view.verticalScrollBar().maximum()
        )

    def _open_log_file(self) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(LOG_FILE)
            else:
                subprocess.Popen(
                    ["xdg-open", str(LOG_FILE)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
        except Exception as e:
            QMessageBox.warning(
                self, "无法打开日志文件",
                f"无法打开日志文件：\n{LOG_FILE}\n\n{e}",
            )

    def closeEvent(self, event) -> None:
        # Hide instead of destroying so log history is preserved
        event.ignore()
        self.hide()
        # Uncheck the corresponding menu action in the parent MainWindow
        parent = self.parent()
        if parent and hasattr(parent, "_show_log_action"):
            parent._show_log_action.setChecked(False)
