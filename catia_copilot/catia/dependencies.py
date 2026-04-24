"""
CATIA 依赖项查找器。

提供：
- find_dependencies() – 收集目标 CATIA 文件依赖的所有文档
"""

import logging
from collections.abc import Callable
from pathlib import Path

logger = logging.getLogger(__name__)


def find_dependencies(
    target_path: str,
    progress_callback: Callable[[str], None] | None = None,
) -> list[str]:
    """返回 *target_path* 依赖的所有文件的完整路径。

    在运行中的 CATIA 实例中打开目标文件；CATIA 会自动加载所有引用的文档。
    该函数收集每个新打开文档的路径，然后在返回前关闭所有这些文档。

    参数
    ----------
    target_path:
        ``.CATPart``、``.CATProduct`` 或 ``.CATDrawing`` 的绝对路径。
    progress_callback:
        可选的 ``callable(str)``，在搜索运行时使用状态消息调用。
    """
    from pycatia import catia

    target      = Path(target_path).resolve()
    caa         = catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    # 在我们执行任何操作之前，已打开文档的快照
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

    # 关闭我们打开的所有文档（目标文件最后关闭）
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
