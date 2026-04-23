"""CATIA COM 辅助函数，供 BOM 相关模块共用。"""

from pathlib import Path


def _is_catia_com_error(exc: Exception) -> bool:
    """Return True if *exc* is a ``pywintypes.com_error`` from the CATIA COM layer.

    This distinguishes deliberate user-cancel signals (CATIA raises a COM error
    when the user clicks Cancel or No in its own SaveAs dialog) from genuine
    OS-level failures such as disk-full or permission-denied, which are plain
    Python exceptions and must always be reported to the user.
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
    for i in range(1, docs.count + 1):
        try:
            d = docs.item(i)
            if Path(d.full_name).resolve() == path:
                return d
        except Exception:
            pass
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
    for i in range(1, docs.count + 1):
        try:
            d = docs.item(i)
            try:
                doc_pn = d.com_object.Product.PartNumber
            except Exception:
                # 非零件/产品类文档（如工程图）无 .Product；回退到文档名茎
                doc_pn = Path(d.name).stem
            if doc_pn == pn:
                return d
        except Exception:
            pass
    return None
