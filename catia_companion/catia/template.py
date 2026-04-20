"""
CATPart template stamping.

Provides:
- apply_part_template() – add standard user-defined properties to CATPart files
"""

import logging
from pathlib import Path

from catia_companion.constants import PRESET_USER_REF_PROPERTIES

logger = logging.getLogger(__name__)


def _find_open_document(application, resolved_path: Path):
    """Return the CATIA document object if *resolved_path* is already open, else ``None``."""
    try:
        docs = application.documents
        for i in range(1, docs.count + 1):
            try:
                doc_com = docs.item(i).com_object
                full_name = doc_com.FullName
                if Path(full_name).resolve() == resolved_path:
                    return docs.item(i)
            except Exception:
                continue
    except Exception:
        pass
    return None


def apply_part_template(
    file_paths: list[str],
    output_folder: str | None = None,
    *,
    progress_callback=None,
    keep_open: bool = False,
) -> int:
    """Add the standard user-defined properties to each CATPart if they are absent.

    Properties are added as empty strings and the file is saved automatically
    after stamping.  *output_folder* is accepted for API compatibility with the
    generic :class:`~catia_companion.ui.convert_dialog.FileConvertDialog` but
    is otherwise unused (parts are saved in place).

    Parameters
    ----------
    progress_callback:
        Optional ``(index, total)`` callback invoked after each file is
        processed, compatible with
        :class:`~catia_companion.ui.convert_dialog.FileConvertDialog`.
    keep_open:
        When ``True`` the document is **not** closed after stamping.  Use
        this when operating on a document that is already open in CATIA (e.g.
        the current active document selected via "使用当前活动文档").  When
        ``False`` (the default), each file is opened if not already open and
        closed afterwards only if it was not already open before stamping.

    Returns
    -------
    int
        Number of files stamped successfully.
    """
    from pycatia import catia
    from pycatia.mec_mod_interfaces.part_document import PartDocument

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    succeeded: list[str] = []
    failed:    list[str] = []
    total = len(file_paths)

    for idx, path in enumerate(file_paths):
        src = Path(path)
        was_already_open = False
        try:
            if keep_open:
                # The document is already active in CATIA – skip re-opening it.
                # Using documents.open() on an already-open file would trigger
                # a CATIA "reload?" dialog, and would fail for unsaved documents
                # that have no file on disk yet.
                logger.info(f"Stamping active document: {src.name}")
                doc = PartDocument(application.active_document.com_object)
                was_already_open = True
            else:
                src = src.resolve()
                existing_doc = _find_open_document(application, src)
                if existing_doc is not None:
                    logger.info(f"File already open, reusing: {src.name}")
                    was_already_open = True
                    # Make the already-open document the active one so that
                    # active_document refers to it after this branch.
                    existing_doc.com_object.Activate()
                else:
                    logger.info(f"Opening: {src}")
                    documents.open(str(src))
                doc = PartDocument(application.active_document.com_object)

            product    = doc.product
            user_props = product.user_ref_properties

            existing_names: set[str] = set()
            for i in range(1, user_props.count + 1):
                try:
                    # CATIA returns a qualified path such as
                    # "Part6\属性\物料编码" (or using "/" depending on locale).
                    # Only the trailing leaf name is relevant for dedup.
                    raw = user_props.item(i).name
                    leaf = raw.replace("/", "\\").rsplit("\\", 1)[-1]
                    existing_names.add(leaf)
                except Exception:
                    pass

            added: list[str] = []
            for prop_name in PRESET_USER_REF_PROPERTIES:
                if prop_name not in existing_names:
                    user_props.create_string(prop_name, "")
                    added.append(prop_name)
                    logger.info(f"  Added property: '{prop_name}'")
                else:
                    logger.info(f"  Skipped (already exists): '{prop_name}'")

            try:
                doc.save()
                logger.info(f"  Saved: {src.name}")
            except Exception as save_err:
                # Unsaved (new) documents have no disk path yet – log a warning
                # but still count the stamp as successful since properties were
                # written into memory.
                logger.warning(f"  Could not save {src.name} (may be unsaved): {save_err}")

            succeeded.append(f"{src.name} (+{len(added)} added)")

        except Exception as e:
            logger.error(f"  ERROR processing {src.name}: {e}")
            failed.append(f"{src.name}: {e}")
        finally:
            if not was_already_open:
                try:
                    application.active_document.close()
                except Exception:
                    pass

        if progress_callback is not None:
            try:
                progress_callback(idx, total)
            except Exception:
                pass

    return len(succeeded)
