"""
CATIA V5 COM connection utility.

Provides connect_to_catia() as a drop-in replacement for pycatia's catia().

Improvements over the stock pycatia catia() call
-------------------------------------------------
* GetActiveObject is tried first so that an already-running CATIA V5 session
  is always reused (pycatia uses only Dispatch, which can spawn a new instance).
* Multiple ProgIDs are attempted in order (CATIA_COM_PROGIDS) so that
  installations that register CATIA V5 under a non-standard name still work.
* Errors carry the list of ProgIDs that were tried, making diagnosis easier.
"""

import logging
from typing import Optional

from catia_companion.constants import CATIA_COM_PROGIDS

logger = logging.getLogger(__name__)


class _CatiaHandle:
    """Thin wrapper that mimics the interface of pycatia's CATIADocHandler.

    Callers access the pycatia Application object via ``handle.application``,
    exactly as they would with the object returned by pycatia's ``catia()``.
    """

    __slots__ = ("_com_obj",)

    def __init__(self, com_obj) -> None:
        self._com_obj = com_obj

    @property
    def application(self):
        from pycatia.in_interfaces.application import Application
        return Application(self._com_obj)


def connect_to_catia(progid: Optional[str] = None) -> _CatiaHandle:
    """Connect to a running CATIA V5 instance (or launch one).

    Strategy
    --------
    For each ProgID in the candidate list:

    1. Try ``win32com.client.GetActiveObject(progid)`` – attaches to the
       CATIA process that is **already running**.  This is non-destructive and
       does not change the currently active document.
    2. If step 1 raises (CATIA is not running under that ProgID), try
       ``win32com.client.Dispatch(progid)`` – launches CATIA or connects to
       the COM class registered for that ProgID.

    Parameters
    ----------
    progid:
        Optional override ProgID tried *before* the defaults in
        CATIA_COM_PROGIDS.  Useful if the local CATIA installation uses an
        unusual registration string.

    Returns
    -------
    _CatiaHandle
        Object with an ``.application`` property that returns the pycatia
        ``Application`` instance.

    Raises
    ------
    pywintypes.com_error
        If none of the candidate ProgIDs result in a successful connection.
    """
    try:
        import win32com.client
        import pywintypes
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for CATIA COM automation but is not installed."
        ) from exc

    if progid:
        candidates = [progid] + [p for p in CATIA_COM_PROGIDS if p != progid]
    else:
        candidates = list(CATIA_COM_PROGIDS)

    last_error: Exception = RuntimeError("No ProgID candidates available.")
    for pid in candidates:
        # ── 1. Attach to running instance ──────────────────────────────────
        try:
            com_obj = win32com.client.GetActiveObject(pid)
            logger.debug("Connected to running CATIA V5 via GetActiveObject(%r)", pid)
            return _CatiaHandle(com_obj)
        except pywintypes.com_error as exc:
            logger.debug("GetActiveObject(%r) failed: %s", pid, exc)
            last_error = exc

        # ── 2. Launch / dispatch ────────────────────────────────────────────
        try:
            com_obj = win32com.client.Dispatch(pid)
            logger.debug("Connected to CATIA V5 via Dispatch(%r)", pid)
            return _CatiaHandle(com_obj)
        except pywintypes.com_error as exc:
            logger.debug("Dispatch(%r) failed: %s", pid, exc)
            last_error = exc

    logger.error(
        "Could not connect to CATIA V5. Tried ProgIDs: %s", candidates
    )
    raise last_error
