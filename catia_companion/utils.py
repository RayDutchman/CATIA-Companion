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
    """Return the absolute path to a resource file.

    When running as a PyInstaller-frozen executable, bundled files (e.g.
    ``resources/``, ``catia_companion/ui/style.qss``) live inside
    ``sys._MEIPASS`` (which may be an ``_internal/`` sub-folder on
    PyInstaller 6.x).  User-placed files such as ``macros/`` and
    ``drawing_templates/`` are never bundled and always live next to the
    executable.  This function checks ``sys._MEIPASS`` first so that
    bundled resources are found regardless of the PyInstaller version or
    ``contents_directory`` setting, then falls back to the executable's
    directory for user-managed folders.
    """
    if hasattr(sys, "_MEIPASS"):
        meipass_candidate = Path(sys._MEIPASS) / filename
        if meipass_candidate.exists():
            return meipass_candidate
        return Path(sys.executable).parent / filename
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
