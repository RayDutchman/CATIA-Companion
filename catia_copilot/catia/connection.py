"""
CATIA V5 R28 连接辅助模块。

提供 get_catia_v5_application()，用于替代 pycatia 原生的 catia()，确保：

1. 即使机器上同时安装了 3DEXPERIENCE，也只连接到 CATIA V5（不会误启动 3DE）。
2. 当 CATIA V5 未启动时，自动通过注册表检测 CATIA V5 安装路径并启动 CNEXT.exe。
3. 如果注册表中 "CATIA.Application" ProgID 已被 3DE 覆盖导致 GetActiveObject 失败，
   则通过枚举 Windows ROT（Running Object Table）来直接找到 V5 对象。

所有模块内部的 CATIA COM 调用都应使用本模块提供的 get_catia_v5_application()
而非直接调用 pycatia 的 catia()，以避免意外连接到 3DEXPERIENCE。
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
# 内部辅助：V5 版本识别
# ---------------------------------------------------------------------------

def _is_catia_v5_dispatch(dispatch) -> bool:
    """返回 True 表示 dispatch 对象属于 CATIA V5，而非 3DEXPERIENCE。

    判据（宽松优先，避免误判）：
    - 版本字符串中含 "3DEXPERIENCE" → 3DE，返回 False
    - 版本字符串中含 "V5" → V5，返回 True
    - 版本号是 < 100 的纯数字（如 "28" 代表 R28）→ V5，返回 True
    - 其余情况默认视为 V5（保持向后兼容）
    - 若 .Version 访问抛出异常（可能由 gen_py 早绑定缓存污染引起），
      则回退至 .Name 检查：Name 为 "CNEXT"/"CATIA" 时视为 V5
    """
    try:
        version = str(dispatch.Version)
        if "3DEXPERIENCE" in version.upper():
            return False
        if "V5" in version.upper():
            return True
        # 版本号可能是纯数字，如 "28" 代表 R28
        try:
            if float(version.strip()) < 100:
                return True
        except (ValueError, AttributeError):
            pass
        return True   # 默认视为 V5
    except Exception:
        # .Version 访问失败（例如 gen_py 早绑定缓存与实际对象接口不兼容）。
        # 回退策略：用 .Name 属性判断——CATIA V5 和 3DEXPERIENCE 均使用 "CNEXT"
        # 作为 Application.Name；如果 .Name 也不可读，则跳过此对象。
        try:
            name = str(dispatch.Name).upper()
            return name in ("CNEXT", "CATIA")
        except Exception:
            return False


# ---------------------------------------------------------------------------
# 内部辅助：通过 ROT 枚举查找运行中的 CATIA V5
# ---------------------------------------------------------------------------

def _find_catia_v5_in_rot():
    """枚举 Windows Running Object Table，返回 CATIA V5 的 COM dispatch 对象。

    当 "CATIA.Application" ProgID 在注册表中已被 3DEXPERIENCE 覆盖时，
    GetActiveObject("CATIA.Application") 无法找到 V5；此函数通过直接枚举 ROT
    来绕过 ProgID→CLSID 映射，确保能找到正在运行的 V5 实例。

    使用 win32com.client.dynamic.Dispatch 而非普通的 Dispatch，强制晚绑定以
    避免 gen_py 早绑定缓存干扰（早绑定缓存可能来自 3DEXPERIENCE，导致用其
    接口描述调用 V5 对象时抛出异常）。

    找到返回 COM dispatch 对象；未找到返回 None。
    """
    try:
        import pythoncom
        import win32com.client as _wcc
        from win32com.client import dynamic as _wcc_dynamic
    except ImportError:
        return None

    try:
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            moniker = monikers[0]
            try:
                obj = rot.GetObject(moniker)
                # 强制使用晚绑定（绕过 gen_py 缓存）
                dispatch = _wcc_dynamic.Dispatch(obj)
                # 必须能访问 Name 属性（CATIA Application 对象的基本属性）
                _ = dispatch.Name
                if _is_catia_v5_dispatch(dispatch):
                    logger.debug("通过 ROT 枚举找到 CATIA V5 COM 对象")
                    return dispatch
            except Exception:
                continue
    except Exception as exc:
        logger.debug(f"ROT 枚举失败：{exc}")
    return None


# ---------------------------------------------------------------------------
# 内部辅助：获取运行中的 CATIA V5 COM 对象
# ---------------------------------------------------------------------------

def _get_v5_com_object():
    """尝试获取运行中的 CATIA V5 COM dispatch 对象。

    按顺序尝试：
    1. GetActiveObject("CATIA.Application") + 版本校验（使用晚绑定，绕过 gen_py）
    2. ROT 枚举（应对 ProgID 被 3DE 覆盖的场景）

    返回 COM dispatch 对象，或 None（CATIA V5 未运行）。
    """
    try:
        import pythoncom
        from win32com.client import dynamic as _wcc_dynamic
    except ImportError:
        return None

    # ── 方式 1：标准 GetActiveObject（强制晚绑定，绕过 gen_py 缓存）────────
    # win32com.client.GetActiveObject 内部自动完成 ProgID→CLSID→QI(IDispatch)
    # 的全套流程，在所有 pywin32 版本中均可靠工作。
    # app._oleobj_ 取出底层 PyIDispatch，再交给 dynamic.Dispatch 包装为晚绑定
    # 代理，避免 gen_py 早绑定缓存（如 3DEXPERIENCE 写入的 typelib 缓存）干扰。
    try:
        import win32com.client as _wcc
        _raw_app = _wcc.GetActiveObject("CATIA.Application")
        app = _wcc_dynamic.Dispatch(_raw_app._oleobj_)
        _ = app.Name        # 功能性测试
        if _is_catia_v5_dispatch(app):
            return app
        # GetActiveObject 返回的是 3DE，继续尝试 ROT 枚举
        logger.debug("GetActiveObject 返回的是 3DEXPERIENCE，尝试 ROT 枚举查找 V5")
    except Exception:
        pass

    # ── 方式 2：ROT 枚举 ──────────────────────────────────────────────────
    return _find_catia_v5_in_rot()


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
    只连接到 CATIA V5，而不会误启动 3DEXPERIENCE。

    行为：
    1. 若 CATIA V5 已在运行，直接连接并返回。
    2. 若 "CATIA.Application" ProgID 在注册表中被 3DE 覆盖，则通过 ROT 枚举找到 V5。
    3. 若 CATIA V5 未运行，自动通过注册表找到并启动 CNEXT.exe，然后连接。

    返回：
        pycatia Application 对象（已连接到 CATIA V5）。

    抛出：
        RuntimeError：无法连接到 CATIA V5（未安装或启动失败）时。
    """
    from pycatia.in_interfaces.application import Application

    com_obj = _get_v5_com_object()
    if com_obj is None:
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
