"""CATIA COM 辅助函数，供 BOM 相关模块共用。"""

import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def _is_catia_com_error(exc: Exception) -> bool:
    """如果 *exc* 是来自 CATIA COM 层的 ``pywintypes.com_error`` 则返回 True。

    用于区分用户主动取消信号（当用户在 CATIA 自己的 SaveAs 对话框中点击取消或否时，
    CATIA 会抛出 COM 错误）与真正的操作系统级错误（如磁盘已满或权限拒绝），
    后者是普通 Python 异常，必须始终报告给用户。
    """
    try:
        import pywintypes  # noqa: PLC0415
        return isinstance(exc, pywintypes.com_error)
    except ImportError:
        return False


def _find_catia_doc_by_path(docs, path: Path) -> object | None:
    """返回解析路径与 *path* 匹配的 CATIA 文档对象，如果未找到则返回 ``None``。

    参数：
        docs: CATIA 文档集合
        path: 要匹配的文件路径

    返回：
        匹配的 CATIA 文档对象，或 None
    """
    logger.debug("_find_catia_doc_by_path: searching for path=%r across %d doc(s)", str(path), docs.count)
    for i in range(1, docs.count + 1):
        try:
            d = docs.item(i)
            try:
                d_resolved = Path(d.full_name).resolve()
            except Exception as path_err:
                logger.debug(
                    "_find_catia_doc_by_path:   doc[%d] full_name resolve error: %s",
                    i, path_err,
                )
                continue
            match = d_resolved == path
            logger.debug(
                "_find_catia_doc_by_path:   doc[%d] path=%r  match=%s",
                i, str(d_resolved), match,
            )
            if match:
                return d
        except Exception as exc:
            logger.debug("_find_catia_doc_by_path:   doc[%d] unexpected error: %s", i, exc)
    logger.debug("_find_catia_doc_by_path: no match found")
    return None


def _find_catia_doc_by_part_number(docs, pn: str) -> object | None:
    """返回根产品零件编号与 *pn* 匹配的第一个已打开 CATIA 文档对象。

    当零件尚未保存到磁盘时（无文件路径），可通过零件编号定位已在 CATIA
    中打开的文档。如果零件编号不可用则回退到按文档名（不含扩展名）匹配。
    未找到时返回 ``None``。

    参数：
        docs: CATIA 文档集合
        pn:   要匹配的零件编号

    返回：
        匹配的 CATIA 文档对象，或 None
    """
    logger.debug("_find_catia_doc_by_part_number: searching for pn=%r across %d doc(s)", pn, docs.count)
    for i in range(1, docs.count + 1):
        try:
            d = docs.item(i)
            try:
                doc_pn = d.com_object.Product.PartNumber
                logger.debug(
                    "_find_catia_doc_by_part_number:   doc[%d] name=%r  PN=%r  match=%s",
                    i, d.name, doc_pn, doc_pn == pn,
                )
            except Exception as pn_err:
                # 非零件/产品类文档（如工程图）无 .Product；回退到文档名茎
                doc_pn = Path(d.name).stem
                logger.debug(
                    "_find_catia_doc_by_part_number:   doc[%d] name=%r  no Product.PartNumber (%s)  -> using name stem=%r  match=%s",
                    i, d.name, pn_err, doc_pn, doc_pn == pn,
                )
            if doc_pn == pn:
                return d
        except Exception as exc:
            logger.debug("_find_catia_doc_by_part_number:   doc[%d] unexpected error: %s", i, exc)
    logger.debug("_find_catia_doc_by_part_number: no match found")
    return None
