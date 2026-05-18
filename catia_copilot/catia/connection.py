"""
CATIA V5 R28 连接辅助模块。

提供 get_catia_v5_application()，用于替代 pycatia 原生的 catia()，确保：

1. 即使机器上同时安装了 3DEXPERIENCE，也只连接到 CATIA V5（不会误启动 3DE）。
2. 当 CATIA V5 未启动时，自动通过注册表检测 CATIA V5 安装路径并启动 CNEXT.exe。
3. 如果注册表中 "CATIA.Application" ProgID 不存在（CO_E_CLASSSTRING），
   则依次尝试已知 CLSID 直连和 ROT 枚举来找到 V5 对象。

所有模块内部的 CATIA COM 调用都应使用本模块提供的 get_catia_v5_application()
而非直接调用 pycatia 的 catia()，以避免意外连接到 3DEXPERIENCE 或新建实例。
"""

import logging
import subprocess
import time
from pathlib import Path

logger = logging.getLogger(__name__)

# CATIA V5 启动等待超时（秒）和轮询间隔（秒）
_CATIA_V5_STARTUP_TIMEOUT = 60
_CATIA_V5_POLL_INTERVAL = 2


# ---------------------------------------------------------------------------
# 内部辅助：查找 CNEXT.exe 路径
# ---------------------------------------------------------------------------

def _get_cnext_exe() -> Path | None:
    """返回 CATIA V5 主可执行文件（CNEXT.exe）的路径，未找到返回 None。"""
    from catia_copilot.utils import detect_catia_root
    catia_root = detect_catia_root()
    if catia_root is None:
        return None
    for candidate in (
        Path(catia_root) / "win_b64" / "code" / "bin" / "CNEXT.exe",
        Path(catia_root) / "win_b64" / "code" / "bin" / "CATIA.exe",
    ):
        if candidate.exists():
            return candidate
    return None


# ---------------------------------------------------------------------------
# 内部辅助：获取运行中的 CATIA V5 COM 对象
# ---------------------------------------------------------------------------

def _get_v5_com_object():
    """尝试获取运行中的 CATIA V5 COM dispatch 对象。

    委托给 catia_copilot.utils.get_catia_v5_com_dispatch()，
    该函数综合使用注册表 ProgID、已知 CLSID 和 ROT 枚举三种策略，
    确保在 ProgID 不存在时仍能连接到运行中的 CATIA V5 实例。

    返回 COM dispatch 对象，或 None（CATIA V5 未运行）。
    """
    from catia_copilot.utils import get_catia_v5_com_dispatch
    return get_catia_v5_com_dispatch()


# ---------------------------------------------------------------------------
# 内部辅助：启动 CATIA V5 并等待 COM 就绪
# ---------------------------------------------------------------------------

def _launch_catia_v5() -> bool:
    """启动 CATIA V5（CNEXT.exe）并等待其通过 COM 注册到 ROT。

    返回 True 表示启动成功且 COM 连接已就绪；False 表示失败。
    """
    cnext_exe = _get_cnext_exe()
    if cnext_exe is None:
        logger.error(
            "找不到 CATIA V5 可执行文件（CNEXT.exe）。"
            "请确认 CATIA V5 已正确安装。"
        )
        return False

    logger.info(f"正在启动 CATIA V5：{cnext_exe}")
    try:
        subprocess.Popen([str(cnext_exe)])
    except Exception as exc:
        logger.error(f"启动 CATIA V5 失败：{exc}")
        return False

    # 等待 CATIA V5 注册到 COM ROT
    deadline = time.time() + _CATIA_V5_STARTUP_TIMEOUT
    while time.time() < deadline:
        time.sleep(_CATIA_V5_POLL_INTERVAL)
        if _get_v5_com_object() is not None:
            logger.info("CATIA V5 已启动，COM 连接就绪。")
            return True

    logger.error("CATIA V5 在超时时间内未能注册到 COM。")
    return False


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------

def get_catia_v5_application():
    """获取连接到 CATIA V5 R28 的 pycatia Application 对象。

    本函数替代 ``pycatia.catia()``，以确保在同时安装了 3DEXPERIENCE 的环境中
    只连接到 CATIA V5，而不会误启动 3DEXPERIENCE 或新建 CATIA 实例。

    行为：
    1. 若 CATIA V5 已在运行，直接连接并返回（不新建实例）。
    2. 若 "CATIA.Application" ProgID 不存在，则依次尝试 CLSID 直连和 ROT 枚举。
    3. 若 CATIA V5 未运行（进程不存在），自动通过注册表找到并启动 CNEXT.exe，然后连接。
    4. 若 CATIA V5 进程存在但 COM 连接失败（例如权限不匹配），直接报错，
       不尝试启动新实例。

    返回：
        pycatia Application 对象（已连接到 CATIA V5）。

    抛出：
        RuntimeError：无法连接到 CATIA V5（未安装、启动失败或 COM 连接被拒绝）时。
    """
    from pycatia.in_interfaces.application import Application
    from catia_copilot.utils import _is_catia_process_running

    com_obj = _get_v5_com_object()
    if com_obj is None:
        if _is_catia_process_running():
            # 进程存在但 COM 连不上（通常是权限不匹配），不要新建实例
            raise RuntimeError(
                "检测到 CATIA V5 正在运行，但无法通过 COM 连接。\n\n"
                "最常见原因：CATIA 以管理员权限运行，而本程序以普通用户权限运行。\n"
                "解决方法：将 CATIA 改为普通用户权限运行（取消「以管理员身份运行」），"
                "使两侧权限级别一致。"
            )
        # 进程不存在，尝试自动启动
        logger.info("CATIA V5 未运行，尝试自动启动……")
        success = _launch_catia_v5()
        if success:
            com_obj = _get_v5_com_object()
        if com_obj is None:
            raise RuntimeError(
                "无法连接到 CATIA V5 R28。\n\n"
                "可能原因：\n"
                "  • CATIA V5 未安装或无法找到启动程序（CNEXT.exe）\n"
                "  • CATIA V5 启动超时\n\n"
                "请手动启动 CATIA V5 R28 后重试。"
            )

    return Application(com_obj)

