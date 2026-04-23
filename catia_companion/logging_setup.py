"""
CATIA Companion 日志基础设施模块。

设置轮转文件日志和标准输出日志，以及基于 Qt 信号的处理器，
使日志消息可以转发到应用内的 LogWindow 窗口。
"""

import sys
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from PySide6.QtCore import QObject, Signal

# ---------------------------------------------------------------------------
# 日志目录和文件配置
# ---------------------------------------------------------------------------

LOG_DIR: Path  = Path.home() / "CATIA_Companion" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE: Path = LOG_DIR / "catia_companion.log"

# 日志文件大小和备份配置
LOG_MAX_BYTES = 2 * 1024 * 1024  # 单个日志文件最大 2MB
LOG_BACKUP_COUNT = 3              # 保留 3 个备份文件

# ---------------------------------------------------------------------------
# Qt 信号发射器
# ---------------------------------------------------------------------------

class LogSignalEmitter(QObject):
    """为每条日志记录发射信号，供 GUI 控件订阅。"""
    message_logged = Signal(str)


log_signal_emitter = LogSignalEmitter()


class QtLogHandler(logging.Handler):
    """将格式化的日志记录转发到 *log_signal_emitter*。"""

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        log_signal_emitter.message_logged.emit(msg)


# ---------------------------------------------------------------------------
# 根日志器配置（在导入时调用一次）
# ---------------------------------------------------------------------------

_LOG_FORMAT  = "%(asctime)s [%(levelname)s] %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

logging.basicConfig(
    level=logging.DEBUG,
    format=_LOG_FORMAT,
    datefmt=_DATE_FORMAT,
    handlers=[
        RotatingFileHandler(
            LOG_FILE,
            maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT,
            encoding="utf-8",
        ),
        logging.StreamHandler(sys.stdout),
    ],
)

_qt_handler = QtLogHandler()
_qt_handler.setFormatter(logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT))
logging.getLogger().addHandler(_qt_handler)
