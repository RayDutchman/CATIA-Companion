"""
CATIA dependency finder.

Provides:
- find_dependencies() – collect all documents that a target CATIA file depends on
"""

import logging
from collections.abc import Callable
from pathlib import Path

logger = logging.getLogger(__name__)


def find_dependencies(
    target_path: str,
    progress_callback: Callable[[str], None] | None = None,
) -> list[str]:
    """Return the full paths of every file that *target_path* depends on.

    Opens the target file in a running CATIA instance; CATIA automatically
    loads all referenced documents.  The function collects the paths of every
    newly-opened document, then closes all of them before returning.

    Parameters
    ----------
    target_path:
        Absolute path to a ``.CATPart``, ``.CATProduct``, or ``.CATDrawing``.
    progress_callback:
        Optional ``callable(str)`` invoked with a status message while the
        search is running.
    """
    from pycatia import catia

    target      = Path(target_path).resolve()
    caa         = catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    # Snapshot of documents already open before we do anything
    already_open: set[Path] = set()
    for i in range(1, documents.count + 1):
        try:
            already_open.add(Path(documents.item(i).full_name).resolve())
        except Exception:
            pass

    logger.info(f"Opening target for dependency scan: {target}")
    if progress_callback:
        progress_callback("正在打开文件，请稍候…")

    documents.open(str(target))

    results:      list[str]  = []
    newly_opened: set[Path]  = set()

    for i in range(1, documents.count + 1):
        try:
            doc      = documents.item(i)
            doc_path = Path(doc.full_name).resolve()
            if doc_path == target or doc_path in already_open:
                continue
            newly_opened.add(doc_path)
            results.append(str(doc_path))
            logger.info(f"  Dependency: {doc_path}")
        except Exception as e:
            logger.debug(f"  Could not read document {i}: {e}")

    # Close all documents we opened (target last)
    for i in range(documents.count, 0, -1):
        try:
            doc      = documents.item(i)
            doc_path = Path(doc.full_name).resolve()
            if doc_path in newly_opened or doc_path == target:
                doc.close()
        except Exception:
            pass

    logger.info(
        f"Dependency scan complete: {len(results)} found for {target.name}"
    )
    return results
