"""
CATPart template stamping.

Provides:
- apply_part_template() – add standard user-defined properties to CATPart files
"""

import logging
from pathlib import Path

from catia_companion.constants import PART_TEMPLATE_PROPERTIES

logger = logging.getLogger(__name__)


def apply_part_template(
    file_paths: list[str],
    output_folder: str | None = None,
) -> None:
    """Add the standard user-defined properties to each CATPart if they are absent.

    Properties are added as empty strings and the file is saved automatically
    after stamping.  *output_folder* is accepted for API compatibility with the
    generic :class:`~catia_companion.ui.convert_dialog.FileConvertDialog` but
    is otherwise unused (parts are saved in place).
    """
    from catia_companion.catia.connection import connect_to_catia
    from pycatia.mec_mod_interfaces.part_document import PartDocument
    from PySide6.QtWidgets import QMessageBox

    caa = connect_to_catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    succeeded: list[str] = []
    failed:    list[str] = []

    for path in file_paths:
        src = Path(path).resolve()
        logger.info(f"Opening: {src}")
        try:
            documents.open(str(src))
            doc        = PartDocument(application.active_document.com_object)
            product    = doc.product
            user_props = product.user_ref_properties

            existing_names: set[str] = set()
            for i in range(1, user_props.count + 1):
                try:
                    existing_names.add(user_props.item(i).name)
                except Exception:
                    pass

            added: list[str] = []
            for prop_name in PART_TEMPLATE_PROPERTIES:
                if prop_name not in existing_names:
                    user_props.create_string(prop_name, "")
                    added.append(prop_name)
                    logger.info(f"  Added property: '{prop_name}'")
                else:
                    logger.info(f"  Skipped (already exists): '{prop_name}'")

            doc.save()
            logger.info(f"  Saved: {src.name}")
            succeeded.append(f"{src.name} (+{len(added)} added)")

        except Exception as e:
            logger.error(f"  ERROR processing {src.name}: {e}")
            failed.append(f"{src.name}: {e}")
        finally:
            try:
                application.active_document.close()
            except Exception:
                pass

    msg = "Stamping complete.\n\n"
    if succeeded:
        msg += "✔ Succeeded:\n" + "\n".join(f"  {s}" for s in succeeded)
    if failed:
        msg += "\n\n✘ Failed:\n" + "\n".join(f"  {f}" for f in failed)

    if failed:
        QMessageBox.warning(None, "刷写零件模板", msg)
    else:
        QMessageBox.information(None, "刷写零件模板", msg)
