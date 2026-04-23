"""
CATIA Companion 实用工具函数模块。

提供：
- resource_path()        – 解析打包资源文件路径（支持 PyInstaller）
- detect_catia_root()    – 通过注册表自动检测 CATIA 安装目录
- estimate_column_width() – 估算 Excel 列宽度（支持中日韩字符）
"""

import ctypes
import ctypes.wintypes as _wt
import struct
import sys
import unicodedata
import winreg
import logging
from pathlib import Path

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

def _dib_to_bmp(dib: bytes) -> bytes | None:
    """Prepend a BITMAPFILEHEADER to a raw DIB blob to form a valid BMP byte string.

    Returns *None* if the DIB header is too short or obviously malformed.
    """
    if len(dib) < 4:
        return None
    header_size: int = struct.unpack_from("<I", dib, 0)[0]
    if header_size < 12 or header_size > len(dib):
        return None
    if header_size == 12:
        # BITMAPCOREHEADER (OS/2 v1): color entries are RGBTRIPLE (3 bytes)
        if len(dib) < 12:
            return None
        bit_count: int = struct.unpack_from("<H", dib, 10)[0]
        colors = 0 if bit_count > 8 else (1 << bit_count)
        color_table_bytes = colors * 3
    else:
        # BITMAPINFOHEADER (40 bytes) and later variants
        if len(dib) < 24:
            return None
        bit_count = struct.unpack_from("<H", dib, 14)[0]
        compression = struct.unpack_from("<I", dib, 16)[0]
        colors_used = struct.unpack_from("<I", dib, 32)[0]
        if bit_count in (1, 4, 8):
            colors = colors_used if colors_used else (1 << bit_count)
            color_table_bytes = colors * 4  # RGBQUAD
        elif compression == 3:          # BI_BITFIELDS: three DWORD color masks
            color_table_bytes = 12
        else:
            color_table_bytes = 0
    pixel_offset = 14 + header_size + color_table_bytes
    file_size    = 14 + len(dib)
    bmp_header   = b"BM" + struct.pack("<IHHI", file_size, 0, 0, pixel_offset)
    return bmp_header + dib


def _parse_pidsi_thumbnail(data: bytes) -> bytes | None:
    """Parse an OLE ``\\x05SummaryInformation`` property stream and return the
    raw thumbnail bytes (JPEG or reconstructed BMP), or *None* if absent.

    The thumbnail is property ID 17 (``PIDSI_THUMBNAIL``), of type ``VT_CF``
    (0x47).  The ``CLIPDATA`` payload is returned as-is when it looks like a
    JPEG (magic ``FF D8``), or wrapped in a ``BITMAPFILEHEADER`` when the
    clipboard format indicates ``CF_DIB`` (8).
    """
    if len(data) < 48:
        return None
    if struct.unpack_from("<H", data, 0)[0] != 0xFFFE:
        return None
    if struct.unpack_from("<I", data, 24)[0] < 1:
        return None
    sec_offset: int = struct.unpack_from("<I", data, 44)[0]
    if sec_offset + 8 > len(data):
        return None
    prop_count: int = struct.unpack_from("<I", data, sec_offset + 4)[0]
    id_base = sec_offset + 8
    for i in range(prop_count):
        entry = id_base + i * 8
        if entry + 8 > len(data):
            break
        prop_id, prop_off = struct.unpack_from("<II", data, entry)
        if prop_id != 17:               # PIDSI_THUMBNAIL
            continue
        abs_off = sec_offset + prop_off
        if abs_off + 12 > len(data):
            break
        if struct.unpack_from("<I", data, abs_off)[0] != 0x0047:   # VT_CF
            break
        cb_size  = struct.unpack_from("<I", data, abs_off + 4)[0]
        clip_fmt = struct.unpack_from("<i", data, abs_off + 8)[0]
        if cb_size < 4:
            break
        img_bytes = data[abs_off + 12: abs_off + 4 + cb_size]
        if not img_bytes:
            break
        if img_bytes[:2] == b"\xff\xd8":        # JPEG
            return img_bytes
        if clip_fmt == 8:                        # CF_DIB → reconstruct BMP
            return _dib_to_bmp(img_bytes)
        return img_bytes                         # unknown – let Qt try
    return None


def read_catia_thumbnail(filepath: str) -> bytes | None:
    """Extract the thumbnail from a CATIA V5 file and return raw image bytes.

    On Windows, ``IShellItemImageFactory`` (the same API used by Windows
    Explorer) is tried first.  It reads from the system thumbnail cache, so
    repeated calls for the same file are nearly instant and no file parsing is
    needed.

    When the Shell API fails, the function falls back to reading the
    ``\\x05SummaryInformation`` OLE stream embedded in classic CATIA V5
    OLE2-format files.

    Returns raw image bytes (BMP or JPEG) suitable for
    ``QPixmap.loadFromData``, or *None* when no thumbnail is available.
    """
    if sys.platform == "win32":
        result = _read_thumbnail_via_windows_shell(filepath)
        if result:
            return result

    # Fallback: OLE2 SummaryInformation stream (classic CATIA V5 files).
    try:
        import olefile
    except ImportError:
        return None
    try:
        if olefile.isOleFile(filepath):
            with olefile.OleFileIO(filepath) as ole:
                stream_name = "\x05SummaryInformation"
                if ole.exists(stream_name):
                    return _parse_pidsi_thumbnail(ole.openstream(stream_name).read())
    except Exception:
        pass
    return None


def _read_thumbnail_via_windows_shell(filepath: str, size: int = 256) -> bytes | None:
    """Use ``IShellItemImageFactory`` to get the Windows Shell thumbnail for *filepath*.

    Returns raw BMP bytes or *None* on any failure.
    Only works on Windows.
    """
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
        # SIIGBF_BIGGERSIZEOK = 0x01
        hbm = ctypes.c_void_p(0)
        if _vtcall(
            pFactory.value, 3,
            ctypes.HRESULT,
            [_SIZE, ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p)],
            [_SIZE(size, size), 0x01, ctypes.byref(hbm)],
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

