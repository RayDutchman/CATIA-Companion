"""
CATIA Copilot 实用工具函数模块。

提供：
- resource_path()               – 解析打包资源文件路径（支持 PyInstaller）
- detect_catia_root()           – 通过注册表自动检测 CATIA V5 安装目录（优先返回 V5，不返回 3DE）
- check_catia_connection()      – 3 态 COM 连接检测（"connected"/"broken"/"disconnected"）
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

    当 "CATIA.Application" ProgID 在注册表中已被 3DEXPERIENCE 覆盖时，
    GetActiveObject("CATIA.Application") 无法找到 V5；此函数通过直接枚举 ROT
    来绕过 ProgID→CLSID 映射。找到返回 dispatch 对象，否则返回 None。

    使用 win32com.client.dynamic.Dispatch 而非普通的 Dispatch，强制晚绑定以
    避免 gen_py 早绑定缓存干扰（早绑定缓存可能来自 3DEXPERIENCE，导致用其
    接口描述调用 V5 对象时抛出异常）。
    """
    if _win32com_client is None:
        return None
    try:
        import pythoncom
        from win32com.client import dynamic as _wcc_dynamic
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
                _ = dispatch.Name   # 功能性测试
                if _is_catia_v5_dispatch(dispatch):
                    logger.debug("通过 ROT 枚举找到 CATIA V5 COM 对象")
                    return dispatch
            except Exception:
                continue
    except Exception as exc:
        logger.debug(f"ROT 枚举失败：{exc}")
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


def check_catia_connection() -> str:
    """检测 CATIA V5 是否正在运行并可通过 COM 访问。

    返回以下三种状态之一：

    - ``"connected"``     — CATIA V5 已运行，COM 对象可获取，且功能性测试通过。
    - ``"broken"``        — COM 对象可获取（GetActiveObject 成功），但访问属性时
                            抛出异常，说明连接存在异常（例如早绑定缓存污染）。
    - ``"disconnected"``  — CATIA V5 未运行或 COM 完全不可用。

    当 "CATIA.Application" ProgID 在注册表中被 3DEXPERIENCE 覆盖时，此函数
    会自动通过 ROT 枚举查找运行中的 CATIA V5 实例，确保状态正确显示。

    所有 COM 包装均强制使用晚绑定（dynamic.Dispatch），避免 gen_py 早绑定
    缓存污染干扰 .Version 等属性的读取。
    """
    if _win32com_client is None:
        return "disconnected"

    # ── 方式 1：标准 GetActiveObject（强制晚绑定，绕过 gen_py 缓存）────────
    # win32com.client.GetActiveObject 内部自动完成 ProgID→CLSID→QI(IDispatch)
    # 的全套流程，在所有 pywin32 版本中均可靠工作。
    # app._oleobj_ 取出底层 PyIDispatch，再交给 dynamic.Dispatch 包装为晚绑定
    # 代理，避免 gen_py 早绑定缓存干扰 .Version 等属性的读取。
    _broken = False
    try:
        import win32com.client as _wcc
        from win32com.client import dynamic as _dyn
        _raw_app = _wcc.GetActiveObject("CATIA.Application")
        app = _dyn.Dispatch(_raw_app._oleobj_)
        try:
            _ = app.Name  # 功能性测试
            if _is_catia_v5_dispatch(app):
                return "connected"
            # GetActiveObject 返回的是 3DE，继续尝试 ROT 枚举
        except Exception:
            _broken = True
    except Exception:
        pass

    if _broken:
        return "broken"

    # ── 方式 2：ROT 枚举（ProgID 被 3DE 覆盖时的降级路径）────────────────
    try:
        app = _find_catia_v5_in_rot()
        if app is not None:
            _ = app.Name
            return "connected"
    except Exception:
        return "broken"

    return "disconnected"


def diagnose_catia_connection() -> dict:
    """对 CATIA V5 COM 连接进行详细诊断，返回包含各项检测结果的字典。

    返回字典包含以下键：

    - ``status``         (str)            — "connected" / "broken" / "disconnected"
    - ``error``          (str | None)     — 最近一次异常描述（如有）
    - ``app_name``       (str | None)     — CATIA 应用名称（如 "CATIA"）
    - ``app_version``    (str | None)     — CATIA 版本字符串
    - ``is_v5``          (bool | None)    — True 表示连接到 CATIA V5；False 表示 3DEXPERIENCE
    - ``active_doc``     (str | None)     — 当前活动文档名称
    - ``doc_count``      (int | None)     — 已打开文档数量
    - ``gen_py_path``    (str)            — win32com gen_py 缓存目录路径
    - ``gen_py_exists``  (bool)           — gen_py 缓存目录是否存在
    """
    result: dict = {
        "status": "disconnected",
        "error": None,
        "app_name": None,
        "app_version": None,
        "is_v5": None,
        "active_doc": None,
        "doc_count": None,
        "gen_py_path": "",
        "gen_py_exists": False,
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

    # ── COM 连接检测 ──────────────────────────────────────────────────────
    if _win32com_client is None:
        result["error"] = "win32com 未安装，无法进行 COM 调用"
        return result

    # 方式 1：标准 GetActiveObject（强制晚绑定，绕过 gen_py 缓存）
    # win32com.client.GetActiveObject 内部自动完成 ProgID→CLSID→QI(IDispatch)
    # 的全套流程，在所有 pywin32 版本中均可靠工作。
    # app._oleobj_ 取出底层 PyIDispatch，再交给 dynamic.Dispatch 包装为晚绑定代理。
    app = None
    try:
        import win32com.client as _wcc
        from win32com.client import dynamic as _dyn
        _raw_app = _wcc.GetActiveObject("CATIA.Application")
        app = _dyn.Dispatch(_raw_app._oleobj_)
    except Exception as exc:
        result["error"] = str(exc)

    if app is None:
        # 方式 2：ROT 枚举（应对 ProgID 被 3DE 覆盖的场景）
        app = _find_catia_v5_in_rot()
        if app is None:
            result["status"] = "disconnected"
            return result
        result["error"] = None  # 通过 ROT 枚举找到，清除之前的错误

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

