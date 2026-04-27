"""
CATIA Copilot - 应用程序入口点。

所有应用逻辑都在 ``catia_copilot`` 包中实现。
本文件仅负责启动 Qt 应用程序并显示主窗口。
"""

import sys

# 确保在创建任何控件之前初始化日志系统和 Qt 信号发射器
import catia_copilot.logging_setup  # noqa: F401

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication
from catia_copilot.utils import resource_path
from catia_copilot.constants import STYLESHEET_RELATIVE_PATH, APP_ICON_PATH
from catia_copilot.ui.main_window import MainWindow


def main() -> None:
    """应用程序主入口函数。

    初始化 Qt 应用程序，加载样式表和图标，显示主窗口。
    """
    app = QApplication(sys.argv)
    app.setApplicationName("CATIA Copilot 1.4.1")

    # 应用统一的 QSS 样式表
    qss_path = resource_path(STYLESHEET_RELATIVE_PATH)
    if qss_path.exists():
        app.setStyleSheet(qss_path.read_text(encoding="utf-8"))

    # 设置应用程序图标（resources/icon.ico）；如果文件不存在则静默跳过
    icon_path = resource_path(APP_ICON_PATH)
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
