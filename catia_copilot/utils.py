"""
CATIA Copilot 实用工具函数模块。

提供：
- resource_path()               – 解析打包资源文件路径（支持 PyInstaller）
- detect_catia_root()           – 通过注册表自动检测 CATIA V5 安装目录（优先返回 V5，不返回 3DE）
- is_admin()                    – 检测当前进程是否以管理员身份运行
- check_catia_connection()      – 4 态 COM 连接检测（"connected"/"broken"/"access_denied"/"disconnected"）
- diagnose_catia_connection()   – 详细 COM 诊断，返回含版本、文档数等信息的字典
- ensure_clean_gencache()       – 启动时清理 win32com 早绑定缓存（gen_py 目录）
- estimate_column_width()       – 估算 Excel 列宽度（支持中日韩字符）
"""

import ctypes
import ctypes.wintypes as _wt
import os
import shutil
import struct
import sys
import tempfile
import unicodedata
import winreg
import logging
from pathlib import Path

try:
    import win32com.client as _win32com_client
except ImportError:
    _win32com_client = None  # type: ignore[assignment]

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# CATIA V5 识别辅助
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
        try:
            if float(version.strip()) < 100:
                return True
        except (ValueError, AttributeError):
            pass
        return True
    except Exception:
        # .Version 访问失败（例如 gen_py 早绑定缓存与实际对象接口不兼容）。
        # 回退策略：用 .Name 属性判断——CATIA V5 和 3DEXPERIENCE 均使用 "CNEXT"
        # 作为 Application.Name；如果 .Name 也不可读，则跳过此对象。
        try:
            name = str(dispatch.Name).upper()
            return name in ("CNEXT", "CATIA")
        except Exception:
            return False


def _find_catia_v5_in_rot():
    """枚举 Windows Running Object Table，返回 CATIA V5 的 COM dispatch 对象。

    当 "CATIA.Application" ProgID 在注册表中不存在（CO_E_CLASSSTRING）时，
    GetActiveObject("CATIA.Application") 无法找到 V5；此函数通过直接枚举 ROT
    来绕过 ProgID→CLSID 映射。找到返回 dispatch 对象，否则返回 None。

    修复：rot.GetObject(moniker) 返回的是 PyIUnknown，需要先通过
    QueryInterface(IID_IDispatch) 获取 IDispatch 接口，再用 dynamic.Dispatch
    包装为晚绑定对象（避免 gen_py 缓存干扰）。

    同时将每个 moniker 的显示名和失败原因写入 DEBUG 日志，便于诊断问题。
    """
    if _win32com_client is None:
        return None
    try:
        import pythoncom
        from win32com.client import dynamic as _wcc_dynamic
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        # 用于获取 moniker 显示名的 bind context
        try:
            bind_ctx = pythoncom.CreateBindCtx(0)
        except Exception:
            bind_ctx = None
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            moniker = monikers[0]
            # 尝试获取 moniker 显示名（用于日志/诊断）
            display_name = "<unknown>"
            if bind_ctx is not None:
                try:
                    display_name = moniker.GetDisplayName(bind_ctx, None)
                except Exception as dn_exc:
                    display_name = f"<DisplayName 失败: {dn_exc}>"
            try:
                obj = rot.GetObject(moniker)
                # rot.GetObject 返回 PyIUnknown；需要先 QI 到 IDispatch
                # 再包装为 dynamic.Dispatch 对象（晚绑定，绕过 gen_py 缓存）
                try:
                    idispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                except Exception as qi_exc:
                    logger.debug(
                        f"ROT moniker [{display_name}] QI(IID_IDispatch) 失败：{qi_exc}"
                    )
                    continue
                dispatch = _wcc_dynamic.Dispatch(idispatch)
                _ = dispatch.Name   # 功能性测试
                if _is_catia_v5_dispatch(dispatch):
                    logger.debug(f"通过 ROT 枚举找到 CATIA V5 COM 对象：{display_name}")
                    return dispatch
                else:
                    logger.debug(f"ROT moniker [{display_name}] 不是 CATIA V5，跳过")
            except Exception as exc:
                logger.debug(f"ROT moniker [{display_name}] 处理失败：{exc}")
                continue
    except Exception as exc:
        logger.debug(f"ROT 枚举失败：{exc}")
    return None


def _find_catia_progids_in_rot() -> list[str]:
    """枚举 ROT 中所有条目，返回其显示名列表（用于诊断）。

    与 _find_catia_v5_in_rot() 不同，此函数不尝试 QI/Dispatch，
    仅收集显示名，即使对象不可 Dispatch 也会记录。

    返回：
        显示名字符串列表，枚举失败时返回空列表。
    """
    monikers_list: list[str] = []
    if _win32com_client is None:
        return monikers_list
    try:
        import pythoncom
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        try:
            bind_ctx = pythoncom.CreateBindCtx(0)
        except Exception:
            bind_ctx = None
        while True:
            items = enum.Next(1)
            if not items:
                break
            moniker = items[0]
            display_name = "<unknown>"
            if bind_ctx is not None:
                try:
                    display_name = moniker.GetDisplayName(bind_ctx, None)
                except Exception:
                    pass
            monikers_list.append(display_name)
    except Exception as exc:
        logger.debug(f"_find_catia_progids_in_rot 枚举失败：{exc}")
    return monikers_list


def _find_catia_progids_in_registry() -> list[str]:
    """在 HKEY_CLASSES_ROOT 中搜索所有以 "CATIA" 开头的 ProgID 键。

    CATIA V5 通常注册的 ProgID 为 "CATIA.Application" 等，但当
    3DEXPERIENCE 共存时可能被覆盖，或者 ProgID 的实际名称与预期不同。
    此函数枚举 HKCR 顶级键，收集所有匹配 "CATIA*" 的 ProgID 及其
    对应的 CLSID（如果有），供连接函数动态选择正确的 ProgID。

    返回：
        ProgID 字符串列表（如 ["CATIA.Application", "CATIA.Document", ...]），
        无法访问注册表时返回空列表。
    """
    catia_progids: list[str] = []
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as hkcr:
            i = 0
            while True:
                try:
                    key_name = winreg.EnumKey(hkcr, i)
                    if key_name.upper().startswith("CATIA"):
                        catia_progids.append(key_name)
                        logger.debug(f"发现 CATIA ProgID：{key_name}")
                    i += 1
                except OSError:
                    break
    except Exception as exc:
        logger.debug(f"_find_catia_progids_in_registry 失败：{exc}")
    return catia_progids


# CATIA V5 Application 对象的已知 CLSID（与 ProgID 注册无关）。
# 即使 HKCR 中不存在 "CATIA.Application" ProgID，仍可用 CLSID 直接
# 调用 GetActiveObject，绕过 ProgID→CLSID 查找失败（CO_E_CLASSSTRING）。
_CATIA_V5_KNOWN_CLSIDS = [
    "{87FD6F40-E252-11D5-8040-0001B5FA1031}",  # CATIA V5 Application（R20+ 通用）
]


def _try_get_active_object_by_clsid(clsid: str):
    """尝试用 CLSID 字符串直接调用 GetActiveObject，绕过 ProgID 查找。

    win32com.client.GetActiveObject 支持传入 CLSID 格式（"{...}"）字符串，
    内部通过 CLSIDFromString 解析，不依赖 ProgID 注册。

    返回晚绑定 dispatch 对象（成功时），或 None（失败时）。
    """
    if _win32com_client is None:
        return None
    try:
        import win32com.client as _wcc
        from win32com.client import dynamic as _dyn
        _raw = _wcc.GetActiveObject(clsid)
        _oleobj = getattr(_raw, "_oleobj_", None)
        return _dyn.Dispatch(_oleobj) if _oleobj is not None else _raw
    except Exception as exc:
        logger.debug(f"GetActiveObject({clsid!r}) 失败：{exc}")
        return None


def resource_path(filename: str) -> Path:
    """返回打包资源文件的绝对路径。

    当作为 PyInstaller 冻结的可执行文件运行时，使用 ``sys._MEIPASS``
    （PyInstaller 6.x 解压数据文件的 ``_internal/`` 目录）；
    否则使用项目根目录。

    参数：
        filename: 相对于项目根目录的文件路径

    返回：
        资源文件的绝对路径
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).parent.parent / filename


def detect_catia_root() -> str | None:
    """返回 CATIA V5 安装根目录，如果未找到则返回 *None*。

    在 Windows 注册表的 HKEY_LOCAL_MACHINE 下搜索 Dassault Systèmes
    发布版本键，优先返回 CATIA V5（release key 为 B21–B99 的小版本号）而非
    3DEXPERIENCE（release key 含大版本号如 B421+）。

    判定逻辑：
    - 若 release key 名为 "B<数字>"，数字 < 100 → 视为 CATIA V5
    - 若 release key 含 "V5" → 视为 CATIA V5
    - 其余视为 3DE 或未知版本，降低优先级

    返回：
        CATIA V5 安装根目录路径，或 None（如果未检测到）
    """
    registry_paths = [
        r"SOFTWARE\Dassault Systemes",
        r"SOFTWARE\WOW6432Node\Dassault Systemes",
    ]

    def _release_number(release_key_name: str) -> int | None:
        """Return the numeric part of a 'B<n>' release key, or None if not applicable."""
        name = release_key_name.upper()
        if name.startswith("B"):
            try:
                return int(name[1:])
            except ValueError:
                pass
        return None

    def _release_is_v5(release_key_name: str) -> bool:
        """启发式：判断 release key 名称是否属于 CATIA V5（而非 3DE）。"""
        name = release_key_name.upper()
        if "V5" in name:
            return True
        # "B28", "B29" etc. (B + small number ≤ 99) are CATIA V5 releases;
        # 3DE releases use large numbers like B421, B422 …
        num = _release_number(release_key_name)
        if num is not None:
            return num <= 99
        return False

    v5_candidates: list[tuple[int, str]] = []   # (release_number, path)
    other_candidates: list[str] = []

    for reg_path in registry_paths:
        logger.debug(f"Trying registry path: HKEY_LOCAL_MACHINE\\{reg_path}")
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as ds_key:
                i = 0
                while True:
                    try:
                        release = winreg.EnumKey(ds_key, i)
                        logger.debug(
                            f"  Trying key: HKEY_LOCAL_MACHINE\\{reg_path}\\{release}\\0"
                        )
                        try:
                            with winreg.OpenKey(ds_key, rf"{release}\0") as release_key:
                                try:
                                    install_path, _ = winreg.QueryValueEx(
                                        release_key, "DEST_FOLDER"
                                    )
                                    candidate = Path(install_path)
                                    if (candidate / "win_b64").exists():
                                        logger.debug(
                                            f"    -> Valid installation found: {candidate}"
                                        )
                                        if _release_is_v5(release):
                                            num = _release_number(release) or 0
                                            v5_candidates.append((num, str(candidate)))
                                        else:
                                            other_candidates.append(str(candidate))
                                except FileNotFoundError:
                                    pass
                        except OSError:
                            pass
                        i += 1
                    except OSError:
                        break
        except OSError:
            pass

    if v5_candidates:
        # Return the highest release number among V5 candidates (e.g. R28 over R21)
        v5_candidates.sort(key=lambda t: t[0], reverse=True)
        best = v5_candidates[0][1]
        logger.debug(f"Selected CATIA V5 root: {best}")
        return best

    if other_candidates:
        logger.debug(f"No V5 found; falling back to: {other_candidates[0]}")
        return other_candidates[0]

    logger.debug("No valid CATIA installation detected.")
    return None


def is_admin() -> bool:
    """返回 True 表示当前进程以 Windows 管理员（提升权限）身份运行。

    使用 ``ctypes.windll.shell32.IsUserAnAdmin()`` 检测。
    在非 Windows 平台或调用失败时始终返回 False。
    """
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _is_catia_process_running() -> bool:
    """检测 CNEXT.exe（CATIA V5 主进程）是否在系统中运行。

    通过调用 ``tasklist /FI "IMAGENAME eq CNEXT.exe" /NH`` 检测，
    无需管理员权限。返回 True 表示进程存在，False 表示不存在或检测失败。
    """
    import subprocess
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq CNEXT.exe", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        # tasklist 找到匹配项时输出中会包含 "CNEXT.exe"
        return "CNEXT.exe" in result.stdout
    except Exception as exc:
        logger.debug(f"_is_catia_process_running 检测失败：{exc}")
        return False


def get_catia_v5_com_dispatch():
    """获取运行中的 CATIA V5 COM dispatch 对象（晚绑定，绕过 gen_py 缓存）。

    这是所有业务代码应调用的统一入口，替代直接使用
    ``GetActiveObject("CATIA.Application")`` 或 ``_find_catia_v5_in_rot()``。

    连接策略（按优先级）：
    1. 遍历 HKCR 中发现的所有 "CATIA*" ProgID + 经典备选。
    2. 使用已知 CLSID ``{87FD6F40-E252-11D5-8040-0001B5FA1031}`` 直连
       （解决 HKCR 无 ProgID 导致的 CO_E_CLASSSTRING 问题）。
    3. ROT 枚举 + IID_IDispatch QueryInterface（解决 CATIA 普通用户运行
       但 ProgID 注册缺失的场景）。

    返回：
        CATIA V5 的晚绑定 COM dispatch 对象（成功时），或 None（未找到时）。
    """
    if _win32com_client is None:
        return None

    import win32com.client as _wcc
    from win32com.client import dynamic as _dyn

    def _try_get(key: str):
        try:
            _raw = _wcc.GetActiveObject(key)
            _oleobj = getattr(_raw, "_oleobj_", None)
            candidate = _dyn.Dispatch(_oleobj) if _oleobj is not None else _raw
            _ = candidate.Name  # 可用性测试
            if _is_catia_v5_dispatch(candidate):
                return candidate
        except Exception:
            pass
        return None

    # 阶段 1：注册表 ProgID + 经典备选
    progids = _find_catia_progids_in_registry()
    for classic in ("CATIA.Application", "CNEXT.Application"):
        if classic not in progids:
            progids.append(classic)
    for key in progids:
        obj = _try_get(key)
        if obj is not None:
            logger.debug(f"get_catia_v5_com_dispatch: 通过 ProgID [{key}] 连接成功")
            return obj

    # 阶段 2：已知 CLSID 直连（绕过 CO_E_CLASSSTRING）
    for clsid in _CATIA_V5_KNOWN_CLSIDS:
        obj = _try_get(clsid)
        if obj is not None:
            logger.debug(f"get_catia_v5_com_dispatch: 通过 CLSID [{clsid}] 连接成功")
            return obj

    # 阶段 3：ROT 枚举（加 IID_IDispatch QI，应对 ProgID/CLSID 均失败的场景）
    obj = _find_catia_v5_in_rot()
    if obj is not None:
        logger.debug("get_catia_v5_com_dispatch: 通过 ROT 枚举连接成功")
        return obj

    return None


def check_catia_connection() -> str:
    """检测 CATIA V5 是否正在运行并可通过 COM 访问。

    返回以下四种状态之一：

    - ``"connected"``     — CATIA V5 已运行，COM 对象可获取，且功能性测试通过。
    - ``"broken"``        — COM 对象可获取但访问属性时抛出异常；或 CATIA 进程
                            存在但所有 COM 连接方式均失败（COM 注册异常等）。
    - ``"access_denied"`` — CATIA V5 进程正在运行，但当前进程权限不足（CATIA 以
                            管理员身份运行而本进程未提权），导致 COM ROT 对本进程
                            不可见。需要以管理员身份重启本程序。
    - ``"disconnected"``  — CATIA V5 未运行或 COM 完全不可用。

    连接策略（按优先级）：
    1. 遍历 HKCR 中发现的所有 "CATIA*" ProgID + 经典备选。
    2. 用已知 CLSID 直连（绕过 ProgID 注册，解决 CO_E_CLASSSTRING）。
    3. ROT 枚举（加 IID_IDispatch QI，绕过 ProgID→CLSID 映射）。
    4. 上述均失败时根据进程是否存在及权限判断原因。

    所有 COM 包装均强制使用晚绑定（dynamic.Dispatch），避免 gen_py 早绑定
    缓存污染干扰 .Version 等属性的读取。
    """
    if _win32com_client is None:
        return "disconnected"

    # E_ACCESSDENIED HRESULT（0x80070005），以有符号 int32 表示
    _E_ACCESSDENIED = -2147024891

    import win32com.client as _wcc
    from win32com.client import dynamic as _dyn

    def _try_progid(progid: str):
        """尝试通过 progid/clsid 获取晚绑定 dispatch 对象。
        返回 (app, hresult_error)：app 非 None 表示成功。
        """
        try:
            _raw = _wcc.GetActiveObject(progid)
            _oleobj = getattr(_raw, "_oleobj_", None)
            return (_dyn.Dispatch(_oleobj) if _oleobj is not None else _raw), None
        except Exception as exc:
            hr = getattr(exc, "hresult", None)
            if hr is None:
                try:
                    hr = exc.args[0]
                except Exception:
                    pass
            return None, hr

    # ── 阶段 1：遍历注册表 ProgID + 经典备选 ────────────────────────────
    registry_progids = _find_catia_progids_in_registry()
    for classic in ("CATIA.Application", "CNEXT.Application"):
        if classic not in registry_progids:
            registry_progids.append(classic)

    _got_access_denied = False
    _got_broken = False

    for progid in registry_progids:
        app, hr = _try_progid(progid)
        if hr == _E_ACCESSDENIED:
            _got_access_denied = True
            logger.debug(f"GetActiveObject({progid!r}) → E_ACCESSDENIED")
            continue
        if app is not None:
            try:
                _ = app.Name
                if _is_catia_v5_dispatch(app):
                    logger.debug(f"GetActiveObject({progid!r}) 成功，连接 CATIA V5")
                    return "connected"
                logger.debug(f"GetActiveObject({progid!r}) 返回非 V5 对象，继续")
            except Exception:
                _got_broken = True

    if _got_access_denied and not _got_broken:
        logger.debug("所有 ProgID 均返回 E_ACCESSDENIED → access_denied")
        return "access_denied"

    if _got_broken:
        return "broken"

    # ── 阶段 2：已知 CLSID 直连（绕过 CO_E_CLASSSTRING）────────────────
    for clsid in _CATIA_V5_KNOWN_CLSIDS:
        app, hr = _try_progid(clsid)
        if hr == _E_ACCESSDENIED:
            _got_access_denied = True
            logger.debug(f"GetActiveObject({clsid!r}) → E_ACCESSDENIED")
            continue
        if app is not None:
            try:
                _ = app.Name
                if _is_catia_v5_dispatch(app):
                    logger.debug(f"GetActiveObject({clsid!r}) 成功，连接 CATIA V5")
                    return "connected"
            except Exception:
                return "broken"

    if _got_access_denied:
        logger.debug("CLSID 直连也返回 E_ACCESSDENIED → access_denied")
        return "access_denied"

    # ── 阶段 3：ROT 枚举（加 IID_IDispatch QI，绕过 ProgID 映射）────────
    try:
        app = _find_catia_v5_in_rot()
        if app is not None:
            _ = app.Name
            return "connected"
    except Exception:
        return "broken"

    # ── 所有方式均失败：根据进程是否存在及当前权限判断原因 ──────────────
    if _is_catia_process_running():
        if not is_admin():
            # 进程存在但本程序未提权，且没有收到 E_ACCESSDENIED → COM 其他原因失败
            # （注意：如果 CATIA 也是普通用户，以管理员运行反而看不到 ROT；
            #  此处只要进程存在 + 未提权 + ROT 枚举失败 → 保守判定为 broken）
            logger.debug("CNEXT.exe 存在，未提权，所有 COM 方式均失败 → broken")
            return "broken"
        else:
            logger.debug("CNEXT.exe 存在，已提权，但 COM 连接失败 → broken")
            return "broken"

    return "disconnected"


def diagnose_catia_connection() -> dict:
    """对 CATIA V5 COM 连接进行详细诊断，返回包含各项检测结果的字典。

    返回字典包含以下键：

    - ``status``                   (str)            — "connected" / "broken" / "access_denied" / "disconnected"
    - ``error``                    (str | None)     — 最近一次异常描述（如有）
    - ``get_active_error``         (str | None)     — GetActiveObject 的实际报错（与 error 区分）
    - ``rot_found``                (bool)           — ROT 枚举是否找到 CATIA V5 对象
    - ``app_name``                 (str | None)     — CATIA 应用名称（如 "CATIA"）
    - ``app_version``              (str | None)     — CATIA 版本字符串
    - ``is_v5``                    (bool | None)    — True 表示连接到 CATIA V5；False 表示 3DEXPERIENCE
    - ``active_doc``               (str | None)     — 当前活动文档名称
    - ``doc_count``                (int | None)     — 已打开文档数量
    - ``gen_py_path``              (str)            — win32com gen_py 缓存目录路径
    - ``gen_py_exists``            (bool)           — gen_py 缓存目录是否存在
    - ``is_elevated``              (bool)           — 当前进程是否以管理员身份运行
    - ``catia_process_running``    (bool)           — CNEXT.exe 进程是否正在运行
    - ``registry_catia_progids``   (list[str])      — HKCR 中找到的所有 CATIA* ProgID
    - ``rot_monikers``             (list[str])      — ROT 中所有条目的显示名（不论是否 CATIA）
    """
    result: dict = {
        "status": "disconnected",
        "error": None,
        "get_active_error": None,
        "rot_found": False,
        "app_name": None,
        "app_version": None,
        "is_v5": None,
        "active_doc": None,
        "doc_count": None,
        "gen_py_path": "",
        "gen_py_exists": False,
        "is_elevated": is_admin(),
        "catia_process_running": _is_catia_process_running(),
        "registry_catia_progids": [],
        "rot_monikers": [],
    }

    # ── gen_py 缓存目录 ────────────────────────────────────────────────────
    gen_py_path: Path | None = None
    if _win32com_client is not None:
        try:
            from win32com.client import gencache as _gencache
            gen_py_path = Path(_gencache.GetGeneratePath())
        except Exception:
            pass
    if gen_py_path is None:
        gen_py_path = Path.home() / "AppData" / "Local" / "Temp" / "gen_py"
    result["gen_py_path"] = str(gen_py_path)
    result["gen_py_exists"] = gen_py_path.exists()

    # ── 收集注册表 ProgID 列表（诊断用）─────────────────────────────────────
    result["registry_catia_progids"] = _find_catia_progids_in_registry()

    # ── 收集 ROT 显示名列表（诊断用）──────────────────────────────────────
    result["rot_monikers"] = _find_catia_progids_in_rot()

    # ── COM 连接检测 ──────────────────────────────────────────────────────
    if _win32com_client is None:
        result["error"] = "win32com 未安装，无法进行 COM 调用"
        return result

    # E_ACCESSDENIED HRESULT（0x80070005），以有符号 int32 表示
    _E_ACCESSDENIED = -2147024891

    import win32com.client as _wcc
    from win32com.client import dynamic as _dyn

    def _try_get(key: str):
        """尝试用 ProgID 或 CLSID 获取晚绑定 dispatch，返回 (app, hresult)。"""
        try:
            _raw = _wcc.GetActiveObject(key)
            _oleobj = getattr(_raw, "_oleobj_", None)
            _candidate = _dyn.Dispatch(_oleobj) if _oleobj is not None else _raw
            _ = _candidate.Name  # 可用性测试
            return _candidate, None
        except Exception as exc:
            hr = getattr(exc, "hresult", None)
            if hr is None:
                try:
                    hr = exc.args[0]
                except Exception:
                    pass
            return None, hr

    # 构建要尝试的 ProgID/CLSID 列表：注册表发现的 + 经典备选 + 已知 CLSID
    progids_to_try = list(result["registry_catia_progids"])
    for classic in ("CATIA.Application", "CNEXT.Application"):
        if classic not in progids_to_try:
            progids_to_try.append(classic)
    all_keys_to_try = progids_to_try + [
        c for c in _CATIA_V5_KNOWN_CLSIDS if c not in progids_to_try
    ]

    app = None
    _got_access_denied = False
    _last_error: str | None = None

    for key in all_keys_to_try:
        _candidate, hr = _try_get(key)
        if _candidate is not None:
            app = _candidate
            logger.debug(f"diagnose: GetActiveObject({key!r}) 成功")
            break
        _last_error = f"GetActiveObject({key!r}) 失败：{hr}"
        if hr == _E_ACCESSDENIED:
            _got_access_denied = True
            logger.debug(f"diagnose: GetActiveObject({key!r}) → E_ACCESSDENIED")

    result["get_active_error"] = _last_error

    # 若所有方式均 E_ACCESSDENIED，直接判定
    if app is None and _got_access_denied:
        result["status"] = "access_denied"
        result["error"] = (
            "E_ACCESSDENIED (0x80070005)：CATIA 可能以管理员身份运行，"
            "而本程序未提权。请以管理员身份重启本程序。\n"
            f"尝试过的 ProgID/CLSID：{all_keys_to_try}"
        )
        return result

    if app is None:
        # 方式 3：ROT 枚举（加 IID_IDispatch QI）
        rot_app = _find_catia_v5_in_rot()
        if rot_app is not None:
            result["rot_found"] = True
            app = rot_app
        else:
            # 所有方式均失败，根据进程是否存在及权限判断原因
            if result["catia_process_running"]:
                if not result["is_elevated"]:
                    # 未提权但所有 COM 方式均失败（没有收到 E_ACCESSDENIED）
                    # 可能是 CATIA 也以普通用户运行但 ROT QI 仍失败，
                    # 或 COM 注册表问题 → 保守判定为 broken
                    result["status"] = "broken"
                    result["error"] = (
                        "CNEXT.exe 进程存在，GetActiveObject（ProgID + CLSID）和 ROT 枚举均失败。\n"
                        "可能原因：\n"
                        "  ① HKCR 中无 CATIA ProgID，且已知 CLSID 直连也失败\n"
                        "  ② CATIA 以不同权限级别运行（UAC 隔离）\n"
                        "  ③ CATIA 正在初始化中，尚未注册 COM 对象\n"
                        f"尝试过的 ProgID/CLSID：{all_keys_to_try}\n"
                        f"注册表发现的 CATIA ProgID：{result['registry_catia_progids']}\n"
                        f"ROT 条目（前10条）：{result['rot_monikers'][:10]}\n"
                        f"最近错误：{result['get_active_error']}"
                    )
                else:
                    result["status"] = "broken"
                    result["error"] = (
                        "CNEXT.exe 进程存在，本程序已以管理员身份运行，\n"
                        "但所有 COM 连接方式均失败。\n"
                        "⚠️ 注意：如果 CATIA 是以普通用户（非管理员）运行的，\n"
                        "则管理员进程无法看到普通用户的 ROT 对象。\n"
                        "建议：请改为以普通用户身份运行本程序（与 CATIA 相同权限级别）。\n"
                        "可能原因：\n"
                        "  ① CATIA 以普通用户运行，而本程序以管理员运行（权限级别不匹配）\n"
                        "  ② HKCR 中无 CATIA ProgID，CLSID 直连也失败\n"
                        "  ③ gen_py 早绑定缓存问题或 CATIA 正在初始化\n"
                        f"尝试过的 ProgID/CLSID：{all_keys_to_try}\n"
                        f"注册表发现的 CATIA ProgID：{result['registry_catia_progids']}\n"
                        f"ROT 条目（前10条）：{result['rot_monikers'][:10]}\n"
                        f"最近错误：{result['get_active_error']}"
                    )
            else:
                result["status"] = "disconnected"
            return result

    # GetActiveObject 或 ROT 枚举成功 — 继续功能性测试
    try:
        result["app_name"] = app.Name
    except Exception as exc:
        result["status"] = "broken"
        result["error"] = f"获取 .Name 失败：{exc}"
        return result

    result["status"] = "connected"
    result["is_v5"] = _is_catia_v5_dispatch(app)

    try:
        result["app_version"] = str(app.Version)
    except Exception:
        pass

    try:
        result["doc_count"] = int(app.Documents.Count)
    except Exception:
        pass

    try:
        result["active_doc"] = str(app.ActiveDocument.Name)
    except Exception:
        pass  # 无活动文档时正常

    return result


def ensure_clean_gencache() -> None:
    """启动时清理 win32com 早绑定缓存目录（gen_py）。

    win32com 的 ``EnsureDispatch`` 会在 ``%LOCALAPPDATA%\\Temp\\gen_py\\``
    写入 CATIA 类型库的早绑定缓存。一旦该缓存存在，后续所有晚绑定调用
    （包括本程序使用的 ``GetActiveObject``）都可能受到干扰，导致无法连接 CATIA。

    本程序仅使用晚绑定，因此在每次启动时主动删除该目录可彻底消除上述隐患。

    为确保在 PyInstaller 打包环境中（不论可执行文件放在 C: 还是 D: 等任意盘符下）
    都能找到并清理正确的 gen_py 目录，本函数会同时清理所有已知候选路径：

    1. ``gencache.GetGeneratePath()`` 返回的版本特定路径（如 …/gen_py/3.11）
    2. 上述路径的父目录（gen_py）
    3. ``tempfile.gettempdir()/gen_py``（最通用的临时目录下的 gen_py）
    4. ``%LOCALAPPDATA%/Temp/gen_py``（Windows 标准本地临时目录）
    5. ``Path.home()/AppData/Local/Temp/gen_py``（家目录推算的路径）

    清理操作是幂等的：目录不存在时静默跳过。
    """
    candidates: set[Path] = set()

    # 1. 通过 gencache 模块获取版本特定路径及其父目录
    if _win32com_client is not None:
        try:
            from win32com.client import gencache as _gencache
            gp = Path(_gencache.GetGeneratePath())
            candidates.add(gp)          # e.g., …/gen_py/3.11
            candidates.add(gp.parent)   # e.g., …/gen_py
        except Exception:
            pass

    # 2. tempfile.gettempdir() / gen_py
    try:
        candidates.add(Path(tempfile.gettempdir()) / "gen_py")
    except Exception:
        pass

    # 3. %LOCALAPPDATA% / Temp / gen_py
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.add(Path(local_app_data) / "Temp" / "gen_py")

    # 4. Path.home() / AppData / Local / Temp / gen_py
    candidates.add(Path.home() / "AppData" / "Local" / "Temp" / "gen_py")

    for path in candidates:
        if path.exists():
            try:
                shutil.rmtree(path, ignore_errors=True)
                logger.debug(f"[gencache] 已清理早绑定缓存目录：{path}")
            except Exception as exc:
                logger.warning(f"[gencache] 清理缓存目录失败（{path}）：{exc}")


def estimate_column_width(text: str) -> int:
    """返回 *text* 在 Excel 列宽度单位下的近似显示宽度。

    中日韩全角字符计为 2，其他字符计为 1。

    参数：
        text: 要测量的文本

    返回：
        估算的列宽度（Excel 单位）
    """
    return sum(
        2 if unicodedata.east_asian_width(c) in ("W", "F") else 1
        for c in str(text)
    )


# ---------------------------------------------------------------------------
# CATIA V5 embedded-thumbnail helpers
# ---------------------------------------------------------------------------

def read_catia_thumbnail(filepath: str) -> bytes | None:
    """Extract the thumbnail from a CATIA V5 file and return raw image bytes.

    Uses ``IShellItemImageFactory`` (the same API used by Windows Explorer)
    to read thumbnails from the system thumbnail cache. Repeated calls for
    the same file are nearly instant and no file parsing is needed.

    Returns raw image bytes (BMP) suitable for ``QPixmap.loadFromData``,
    or *None* when no thumbnail is available.
    """
    if sys.platform == "win32":
        return _read_thumbnail_via_windows_shell(filepath)
    return None


def _read_thumbnail_via_windows_shell(filepath: str, size: int = 256) -> bytes | None:
    """Use ``IShellItemImageFactory`` to get the Windows Shell thumbnail for *filepath*.

    Returns raw BMP bytes or *None* on any failure.
    Only works on Windows.
    """
    try:
        return _read_thumbnail_via_windows_shell_inner(filepath, size)
    except Exception:
        return None


def _read_thumbnail_via_windows_shell_inner(filepath: str, size: int) -> bytes | None:
    # ── COM / GDI structure types ────────────────────────────────────────────

    class _GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", _wt.DWORD),
            ("Data2", _wt.WORD),
            ("Data3", _wt.WORD),
            ("Data4", ctypes.c_ubyte * 8),
        ]

    # IID_IShellItemImageFactory  {BCC18B79-BA16-442F-80C4-8A59C30C463B}
    _IID_ISIIF = _GUID(
        0xBCC18B79, 0xBA16, 0x442F,
        (ctypes.c_ubyte * 8)(0x80, 0xC4, 0x8A, 0x59, 0xC3, 0x0C, 0x46, 0x3B),
    )

    class _SIZE(ctypes.Structure):
        _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]

    class _BITMAP(ctypes.Structure):
        _fields_ = [
            ("bmType",       ctypes.c_long),
            ("bmWidth",      ctypes.c_long),
            ("bmHeight",     ctypes.c_long),
            ("bmWidthBytes", ctypes.c_long),
            ("bmPlanes",     ctypes.c_ushort),
            ("bmBitsPixel",  ctypes.c_ushort),
            ("bmBits",       ctypes.c_void_p),
        ]

    class _BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize",          _wt.DWORD),
            ("biWidth",         ctypes.c_long),
            ("biHeight",        ctypes.c_long),
            ("biPlanes",        _wt.WORD),
            ("biBitCount",      _wt.WORD),
            ("biCompression",   _wt.DWORD),
            ("biSizeImage",     _wt.DWORD),
            ("biXPelsPerMeter", ctypes.c_long),
            ("biYPelsPerMeter", ctypes.c_long),
            ("biClrUsed",       _wt.DWORD),
            ("biClrImportant",  _wt.DWORD),
        ]

    _PSZ = ctypes.sizeof(ctypes.c_void_p)

    def _vtcall(this: int, idx: int, restype, argtypes: list, args: list):
        vtbl = ctypes.cast(this, ctypes.POINTER(ctypes.c_void_p))[0]
        fptr = ctypes.cast(vtbl + idx * _PSZ, ctypes.POINTER(ctypes.c_void_p))[0]
        return ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)(fptr)(this, *args)

    shell32 = ctypes.windll.shell32
    gdi32   = ctypes.windll.gdi32
    user32  = ctypes.windll.user32

    shell32.SHCreateItemFromParsingName.argtypes = [
        ctypes.c_wchar_p, ctypes.c_void_p,
        ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p),
    ]
    shell32.SHCreateItemFromParsingName.restype = ctypes.HRESULT

    pFactory = ctypes.c_void_p(0)
    if shell32.SHCreateItemFromParsingName(
        filepath, None, ctypes.byref(_IID_ISIIF), ctypes.byref(pFactory),
    ) != 0 or not pFactory:
        return None

    try:
        # IShellItemImageFactory::GetImage (vtable slot 3)
        # SIIGBF_THUMBNAILONLY = 0x08 – fail if only a generic file-type icon
        # is available; this prevents returning the CATIA icon as a thumbnail.
        hbm = ctypes.c_void_p(0)
        if _vtcall(
            pFactory.value, 3,
            ctypes.HRESULT,
            [_SIZE, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p)],
            [_SIZE(size, size), 0x08, ctypes.byref(hbm)],
        ) != 0 or not hbm:
            return None

        try:
            bm = _BITMAP()
            if not gdi32.GetObjectW(hbm, ctypes.sizeof(_BITMAP), ctypes.byref(bm)):
                return None
            w, h = bm.bmWidth, bm.bmHeight
            if w <= 0 or h <= 0:
                return None

            bih = _BITMAPINFOHEADER()
            bih.biSize        = ctypes.sizeof(_BITMAPINFOHEADER)
            bih.biWidth       = w
            bih.biHeight      = -h     # negative → top-down
            bih.biPlanes      = 1
            bih.biBitCount    = 32
            bih.biCompression = 0      # BI_RGB
            bih.biSizeImage   = w * h * 4

            buf = ctypes.create_string_buffer(w * h * 4)
            hdc = user32.GetDC(None)
            if not hdc:
                return None
            try:
                if not gdi32.GetDIBits(hdc, hbm, 0, h, buf, ctypes.byref(bih), 0):
                    return None
            finally:
                user32.ReleaseDC(None, hdc)

            dib = bytes(bih) + bytes(buf)
            return (b"BM"
                    + struct.pack("<IHHI", 14 + len(dib), 0, 0,
                                  14 + ctypes.sizeof(_BITMAPINFOHEADER))
                    + dib)
        finally:
            gdi32.DeleteObject(hbm)
    finally:
        _vtcall(pFactory.value, 2, ctypes.c_ulong, [], [])   # IUnknown::Release

