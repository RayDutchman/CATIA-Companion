# build.spec — PyInstaller spec file for CATIA Companion
#
# How to build:
#   pyinstaller build.spec
#
# Output: dist/CATIA Companion/
# All files (resources/, macros/, drawing_templates/, etc.) are placed
# directly next to CATIA Companion.exe — no _internal/ subdirectory.
#
# Before building, place the application icon at:
#   resources/icon.ico
# Then uncomment the `icon=` line in the EXE block below.

import PyInstaller
_pyinstaller_ver = tuple(int(x) for x in PyInstaller.__version__.split('.')[:2])

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('resources', 'resources'),
        ('catia_companion/ui/style.qss', 'catia_companion/ui'),
        ('catia_companion', 'catia_companion'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CATIA Companion',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/icon.ico',  # Uncomment after placing icon.ico in resources/
)

# PyInstaller 6.0 introduced the _internal/ sub-folder and the
# contents_directory option to opt out of it.  Older versions do not
# accept (or need) this parameter, so it is passed conditionally.
_coll_kw = {'contents_directory': '.'} if _pyinstaller_ver >= (6, 0) else {}

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CATIA Companion',
    **_coll_kw,
)
