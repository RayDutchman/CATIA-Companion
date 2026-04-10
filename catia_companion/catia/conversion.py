"""
CATIA file-conversion helpers.

Provides:
- convert_drawing_to_pdf()  – export CATDrawing files to PDF
- convert_part_to_step()    – export CATPart/CATProduct files to STEP (.stp)
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import QMessageBox

logger = logging.getLogger(__name__)


def _prompt_overwrite(dest: Path) -> str:
    """Show an overwrite-conflict dialog for *dest*.

    Returns one of: ``"skip"``, ``"skip_all"``, ``"overwrite"``,
    ``"overwrite_all"``, or ``"cancel"``.
    """
    msg = QMessageBox()
    msg.setWindowTitle("文件已存在")
    msg.setText(f'"{dest.name}" 已存在于输出文件夹中。')
    msg.setInformativeText(str(dest.parent))
    msg.setIcon(QMessageBox.Icon.Warning)
    skip_btn          = msg.addButton("跳过",     QMessageBox.ButtonRole.RejectRole)
    skip_all_btn      = msg.addButton("全部跳过", QMessageBox.ButtonRole.RejectRole)
    _overwrite_btn    = msg.addButton("覆盖",     QMessageBox.ButtonRole.AcceptRole)
    overwrite_all_btn = msg.addButton("全部覆盖", QMessageBox.ButtonRole.AcceptRole)
    cancel_btn        = msg.addButton("取消",     QMessageBox.ButtonRole.DestructiveRole)
    msg.exec()
    clicked = msg.clickedButton()
    if clicked is cancel_btn:
        return "cancel"
    if clicked is skip_all_btn:
        return "skip_all"
    if clicked is skip_btn:
        return "skip"
    if clicked is overwrite_all_btn:
        return "overwrite_all"
    return "overwrite"


def convert_drawing_to_pdf(
    file_paths: list[str],
    output_folder: str | None = None,
    prefix: str = "DR_",
    suffix: str = "",
) -> None:
    """Convert CATDrawing files to PDF using pyCATIA.

    If *prefix* is non-empty it is prepended to the output filename unless the
    stem already starts with it.  If *suffix* is non-empty it is appended
    unless the stem already ends with it.
    """
    from pycatia import catia
    from pycatia.drafting_interfaces.drawing_document import DrawingDocument

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    bulk_action: str | None = None  # "skip_all", "overwrite_all", or "cancel"

    for path in file_paths:
        if bulk_action == "cancel":
            break

        src      = Path(path).resolve()
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent
        dest_dir.mkdir(parents=True, exist_ok=True)

        stem = src.stem
        if prefix and not stem.startswith(prefix):
            stem = f"{prefix}{stem}"
        if suffix and not stem.endswith(suffix):
            stem = f"{stem}{suffix}"

        dest = dest_dir / f"{stem}.pdf"
        logger.info(f"Opening: {src}")

        if dest.exists():
            if bulk_action == "skip_all":
                logger.info(f"  Skipped (skip all): {dest}")
                continue
            if bulk_action == "overwrite_all":
                dest.unlink()
            else:
                action = _prompt_overwrite(dest)
                if action == "cancel":
                    bulk_action = "cancel"
                    break
                if action == "skip_all":
                    bulk_action = "skip_all"
                    logger.info(f"  Skipped (skip all): {dest}")
                    continue
                if action == "skip":
                    logger.info(f"  Skipped: {dest}")
                    continue
                if action == "overwrite_all":
                    bulk_action = "overwrite_all"
                dest.unlink()

        documents.open(str(src))
        drawing_doc = DrawingDocument(application.active_document.com_object)
        sheet_count = drawing_doc.drawing_root.sheets.count
        drawing_doc.export_data(str(dest), "pdf")

        if not dest.exists():
            logger.warning(f"  WARNING: export_data did not create {dest}")
        else:
            logger.info(f"  Exported {sheet_count} sheet(s) -> {dest}")

        drawing_doc.close()
        logger.info(f"Done: {src.name}\n")


def convert_part_to_step(
    file_paths: list[str],
    output_folder: str | None = None,
    prefix: str = "MD_",
    suffix: str = "",
) -> None:
    """Convert CATPart/CATProduct files to STEP (.stp) using pyCATIA.

    If *prefix* is non-empty it is prepended to the output filename unless the
    stem already starts with it.  If *suffix* is non-empty it is appended
    unless the stem already ends with it.
    """
    from pycatia import catia

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    bulk_action: str | None = None

    for path in file_paths:
        if bulk_action == "cancel":
            break

        src      = Path(path)
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent.resolve()
        dest_dir.mkdir(parents=True, exist_ok=True)

        stem = src.stem
        if prefix and not stem.startswith(prefix):
            stem = f"{prefix}{stem}"
        if suffix and not stem.endswith(suffix):
            stem = f"{stem}{suffix}"

        dest = dest_dir / f"{stem}.stp"
        logger.info(f"Opening: {src}")

        if dest.exists():
            if bulk_action == "skip_all":
                logger.info(f"  Skipped (skip all): {dest}")
                continue
            if bulk_action == "overwrite_all":
                dest.unlink()
            else:
                action = _prompt_overwrite(dest)
                if action == "cancel":
                    bulk_action = "cancel"
                    break
                if action == "skip_all":
                    bulk_action = "skip_all"
                    logger.info(f"  Skipped (skip all): {dest}")
                    continue
                if action == "skip":
                    logger.info(f"  Skipped: {dest}")
                    continue
                if action == "overwrite_all":
                    bulk_action = "overwrite_all"
                dest.unlink()

        documents.open(str(src))
        doc = application.active_document
        doc.export_data(str(dest), "stp")
        logger.info(f"  Exported -> {dest}")
        doc.close()
        logger.info(f"Done: {src.name}\n")
