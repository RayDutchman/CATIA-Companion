"""
CATIA Copilot 实用工具函数模块。

提供：
- resource_path()               – 解析打包资源文件路径（支持 PyInstaller）
- detect_catia_root()           – 通过注册表自动检测 CATIA 安装目录
- check_catia_connection()      – 3 态 COM 连接检测（"connected"/"broken"/"disconnected"）
- diagnose_catia_connection()   – 详细 COM 诊断，返回含版本、文档数等信息的字典
- ensure_clean_gencache()       – 启动时清理 win32com 早绑定缓存（gen_py 目录）
- estimate_column_width()       – 估算 Excel 列宽度（支持中日韩字符）
"""

import ctypes
import ctypes.wintypes as _wt
import shutil
import struct
import sys
import unicodedata
import winreg
import logging
from pathlib import Path

try:
    import win32com.client as _win32com_client
except ImportError:
    _win32com_client = None  # type: ignore[assignment]

logger = logging.getLogger(__name__)


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
    发布版本键，返回第一个包含 ``win_b64`` 子文件夹的路径。

    返回：
        CATIA 安装根目录路径，或 None（如果未检测到）
    """
    registry_paths = [
        r"SOFTWARE\Dassault Systemes",
        r"SOFTWARE\WOW6432Node\Dassault Systemes",
    ]
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
                                            f"    -> Valid CATIA installation found: {candidate}"
                                        )
                                        return str(candidate)
                                except FileNotFoundError:
                                    pass
                        except OSError:
                            pass
                        i += 1
                    except OSError:
                        break
        except OSError:
            pass

    logger.debug("No valid CATIA installation detected.")
    return None


def check_catia_connection() -> str:
    """检测 CATIA V5 是否正在运行并可通过 COM 访问。

    返回以下三种状态之一：

    - ``"connected"``     — CATIA 已运行，COM 对象可获取，且功能性测试通过。
    - ``"broken"``        — COM 对象可获取（GetActiveObject 成功），但访问属性时
                            抛出异常，说明连接存在异常（例如早绑定缓存污染）。
    - ``"disconnected"``  — CATIA 未运行或 COM 完全不可用。
    """
    if _win32com_client is None:
        return "disconnected"
    try:
        app = _win32com_client.GetActiveObject("CATIA.Application")
    except Exception:
        return "disconnected"
    try:
        _ = app.Name  # 功能性测试：确认 COM 对象实际可用
        return "connected"
    except Exception:
        return "broken"


def diagnose_catia_connection() -> dict:
    """对 CATIA COM 连接进行详细诊断，返回包含各项检测结果的字典。

    返回字典包含以下键：

    - ``status``         (str)            — "connected" / "broken" / "disconnected"
    - ``error``          (str | None)     — 最近一次异常描述（如有）
    - ``app_name``       (str | None)     — CATIA 应用名称（如 "CATIA"）
    - ``app_version``    (str | None)     — CATIA 版本字符串
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

    try:
        app = _win32com_client.GetActiveObject("CATIA.Application")
    except Exception as exc:
        result["status"] = "disconnected"
        result["error"] = str(exc)
        return result

    # GetActiveObject 成功 — 继续功能性测试
    try:
        result["app_name"] = app.Name
    except Exception as exc:
        result["status"] = "broken"
        result["error"] = f"GetActiveObject 成功但无法访问 .Name：{exc}"
        return result

    result["status"] = "connected"

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
    删除操作是幂等的：目录不存在时静默跳过。
    """
    gen_py_path: Path | None = None

    # 优先通过 gencache 模块获取实际路径
    if _win32com_client is not None:
        try:
            from win32com.client import gencache as _gencache
            gen_py_path = Path(_gencache.GetGeneratePath())
        except Exception:
            pass

    # 回退到默认路径
    if gen_py_path is None:
        local_app_data = Path.home() / "AppData" / "Local"
        gen_py_path = local_app_data / "Temp" / "gen_py"

    if gen_py_path.exists():
        try:
            shutil.rmtree(gen_py_path, ignore_errors=True)
            logger.debug(f"[gencache] 已清理早绑定缓存目录：{gen_py_path}")
        except Exception as exc:
            logger.warning(f"[gencache] 清理缓存目录失败（{gen_py_path}）：{exc}")
    else:
        logger.debug(f"[gencache] 缓存目录不存在，无需清理：{gen_py_path}")


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

