"""
Utility helpers for CATIA Companion.

Provides:
- resource_path()        – resolve bundled resource files (PyInstaller-aware)
- detect_catia_root()    – auto-detect CATIA installation directory via registry
- estimate_column_width() – approximate Excel column width for CJK/ASCII text
"""

import sys
import unicodedata
import winreg
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def resource_path(filename: str) -> Path:
    """Return the absolute path to a bundled resource file.

    When running as a PyInstaller-frozen executable ``sys._MEIPASS`` is used
    (the ``_internal/`` directory where PyInstaller 6.x extracts data files);
    otherwise the project root directory is used.
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).parent.parent / filename


def detect_catia_root() -> str | None:
    """Return the CATIA V5 installation root directory, or *None* if not found.

    Searches the Windows registry under HKEY_LOCAL_MACHINE for Dassault Systèmes
    release keys and returns the first path that contains a ``win_b64`` sub-folder.
    """
    registry_paths = [
        r"SOFTWARE\Dassault Systemes",
        r"SOFTWARE\WOW6432Node\Dassault Systemes",
    ]
    for reg_path in registry_paths:
        logger.debug(f"Trying registry path: HKEY_LOCAL_MACHINE\\{reg_path}")
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as ds_key:
                i = 0
                while True:
                    try:
                        release = winreg.EnumKey(ds_key, i)
                        logger.debug(
                            f"  Trying key: HKEY_LOCAL_MACHINE\\{reg_path}\\{release}\\0"
                        )
                        try:
                            with winreg.OpenKey(ds_key, rf"{release}\0") as release_key:
                                try:
                                    install_path, _ = winreg.QueryValueEx(
                                        release_key, "DEST_FOLDER"
                                    )
                                    candidate = Path(install_path)
                                    if (candidate / "win_b64").exists():
                                        logger.debug(
                                            f"    -> Valid CATIA installation found: {candidate}"
                                        )
                                        return str(candidate)
                                except FileNotFoundError:
                                    pass
                        except OSError:
                            pass
                        i += 1
                    except OSError:
                        break
        except OSError:
            pass

    logger.debug("No valid CATIA installation detected.")
    return None


def estimate_column_width(text: str) -> int:
    """Return the approximate display width of *text* in Excel column-width units.

    CJK / wide characters count as 2; all others count as 1.
    """
    return sum(
        2 if unicodedata.east_asian_width(c) in ("W", "F") else 1
        for c in str(text)
    )
