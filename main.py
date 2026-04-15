"""
CATIA Companion – application entry point.

All application logic lives under the ``catia_companion`` package.
This file only bootstraps the Qt application and launches the main window.
"""

import sys

# Ensure logging and Qt signal emitter are initialised before any widget is created.
import catia_companion.logging_setup  # noqa: F401

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication
from catia_companion.utils import resource_path
from catia_companion.constants import STYLESHEET_RELATIVE_PATH, APP_ICON_PATH
from catia_companion.ui.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("CATIA Companion")

    # Apply the unified QSS stylesheet
    qss_path = resource_path(STYLESHEET_RELATIVE_PATH)
    if qss_path.exists():
        app.setStyleSheet(qss_path.read_text(encoding="utf-8"))

    # Set application icon (resources/icon.ico); silently skipped if not yet provided
    icon_path = resource_path(APP_ICON_PATH)
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
