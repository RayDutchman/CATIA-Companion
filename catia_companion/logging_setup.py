"""
Logging infrastructure for CATIA Companion.

Sets up rotating file + stdout logging and a Qt signal-based handler so log
messages can be forwarded to the in-app LogWindow.
"""

import sys
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from PySide6.QtCore import QObject, Signal

# ---------------------------------------------------------------------------
# Log directory and file
# ---------------------------------------------------------------------------

LOG_DIR: Path  = Path.home() / "CATIA_Companion" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE: Path = LOG_DIR / "catia_companion.log"

# ---------------------------------------------------------------------------
# Qt signal emitter
# ---------------------------------------------------------------------------

class LogSignalEmitter(QObject):
    """Emits a Signal for every log record so GUI widgets can subscribe."""
    message_logged = Signal(str)


log_signal_emitter = LogSignalEmitter()


class QtLogHandler(logging.Handler):
    """Forwards formatted log records to *log_signal_emitter*."""

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        log_signal_emitter.message_logged.emit(msg)


# ---------------------------------------------------------------------------
# Root logger configuration (called once at import time)
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
            maxBytes=2 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8",
        ),
        logging.StreamHandler(sys.stdout),
    ],
)

_qt_handler = QtLogHandler()
_qt_handler.setFormatter(logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT))
logging.getLogger().addHandler(_qt_handler)
