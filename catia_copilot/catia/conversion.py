"""
CATIA 文件转换辅助模块。

提供：
- convert_drawing_to_pdf()  – 将 CATDrawing 文件导出为 PDF
- convert_part_to_step()    – 将 CATPart/CATProduct 文件导出为 STEP (.stp)
"""

import logging
from collections.abc import Callable
from pathlib import Path

from PySide6.QtWidgets import QMessageBox

logger = logging.getLogger(__name__)


def _prompt_overwrite(dest: Path) -> str:
    """显示 *dest* 的覆盖冲突对话框。

    返回以下之一：``"skip"``、``"skip_all"``、``"overwrite"``、
    ``"overwrite_all"`` 或 ``"cancel"``。

    参数：
        dest: 目标文件路径

    返回：
        用户选择的操作
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


def _resolve_overwrite(
    dest: Path,
    bulk_action: str | None,
) -> tuple[str, str | None]:
    """决定批量转换循环中 *dest* 已存在时的处理方式。

    返回 ``(result, new_bulk_action)``，其中 *result* 为以下之一：

    * ``"proceed"``  – 调用者可以写入目标文件（旧文件已删除）。
    * ``"skip"``     – 跳过此文件并移至下一个。
    * ``"cancel"``   – 中止整个批次。
    """
    if bulk_action == "skip_all":
        logger.info(f"  Skipped (skip all): {dest}")
        return "skip", bulk_action
    if bulk_action == "overwrite_all":
        dest.unlink()
        return "proceed", bulk_action
    action = _prompt_overwrite(dest)
    if action == "cancel":
        return "cancel", "cancel"
    if action == "skip_all":
        logger.info(f"  Skipped (skip all): {dest}")
        return "skip", "skip_all"
    if action == "skip":
        logger.info(f"  Skipped: {dest}")
        return "skip", bulk_action
    if action == "overwrite_all":
        bulk_action = "overwrite_all"
    dest.unlink()
    return "proceed", bulk_action


def convert_drawing_to_pdf(
    file_paths: list[str],
    output_folder: str | None = None,
    prefix: str = "DR_",
    suffix: str = "",
    progress_callback: Callable[[int, int], None] | None = None,
    update_before_export: bool = False,
) -> int:
    """使用 pyCATIA 将 CATDrawing 文件转换为 PDF。

    如果 *prefix* 非空，则在输出文件名前添加前缀（除非文件名已包含该前缀）。
    如果 *suffix* 非空，则在文件名后添加后缀（除非文件名已包含该后缀）。

    参数：
        file_paths: CATDrawing 文件路径列表
        output_folder: 输出文件夹路径，默认为源文件所在目录
        prefix: 输出文件名前缀，默认 "DR_"
        suffix: 输出文件名后缀，默认为空
        progress_callback: 进度回调函数，在处理每个文件前调用 ``progress_callback(i, total)``
                          （0 基索引）
        update_before_export: 为 ``True`` 时，在导出 PDF 前更新图纸文档（刷新所有视图）

    返回：
        成功导出的文件数量
    """
    from pycatia import catia
    from pycatia.drafting_interfaces.drawing_document import DrawingDocument

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    bulk_action: str | None = None  # "skip_all", "overwrite_all", or "cancel"
    success_count = 0
    total = len(file_paths)

    for i, path in enumerate(file_paths):
        if progress_callback:
            progress_callback(i, total)

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
            result, bulk_action = _resolve_overwrite(dest, bulk_action)
            if result == "cancel":
                break
            if result == "skip":
                continue

        try:
            documents.open(str(src))
            drawing_doc = DrawingDocument(application.active_document.com_object)
            sheet_count = drawing_doc.drawing_root.sheets.count

            if update_before_export:
                logger.info(f"  Updating drawing ({sheet_count} sheet(s))…")
                drawing_doc.com_object.Update()

            drawing_doc.export_data(str(dest), "pdf")

            if not dest.exists():
                logger.warning(f"  WARNING: export_data did not create {dest}")
            else:
                logger.info(f"  Exported {sheet_count} sheet(s) -> {dest}")

            drawing_doc.close()
            logger.info(f"Done: {src.name}\n")
            if dest.exists():
                success_count += 1
        except Exception as e:
            logger.error("Failed to convert %s: %s", path, e)

    return success_count


def convert_part_to_step(
    file_paths: list[str],
    output_folder: str | None = None,
    prefix: str = "MD_",
    suffix: str = "",
    progress_callback: Callable[[int, int], None] | None = None,
) -> int:
    """使用 pyCATIA 将 CATPart/CATProduct 文件转换为 STEP (.stp)。

    如果 *prefix* 非空，则在输出文件名前添加前缀（除非文件名已包含该前缀）。
    如果 *suffix* 非空，则在文件名后添加后缀（除非文件名已包含该后缀）。

    参数：
        file_paths: CATPart/CATProduct 文件路径列表
        output_folder: 输出文件夹路径，默认为源文件所在目录
        prefix: 输出文件名前缀，默认 "MD_"
        suffix: 输出文件名后缀，默认为空
        progress_callback: 进度回调函数，在处理每个文件前调用 ``progress_callback(i, total)``
                          （0 基索引）

    返回：
        成功导出的文件数量
    """
    from pycatia import catia

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    bulk_action: str | None = None
    success_count = 0
    total = len(file_paths)

    for i, path in enumerate(file_paths):
        if progress_callback:
            progress_callback(i, total)

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
            result, bulk_action = _resolve_overwrite(dest, bulk_action)
            if result == "cancel":
                break
            if result == "skip":
                continue

        try:
            documents.open(str(src))
            doc = application.active_document
            doc.export_data(str(dest), "stp")
            logger.info(f"  Exported -> {dest}")
            doc.close()
            logger.info(f"Done: {src.name}\n")
            success_count += 1
        except Exception as e:
            logger.error("Failed to convert %s: %s", path, e)

    return success_count
