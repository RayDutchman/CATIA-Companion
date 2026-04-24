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
        path: 要匹配的文件路径（已解析的绝对路径）

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
