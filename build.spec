# build.spec — PyInstaller spec file for CATIA Copilot
#
# How to build:
#   pyinstaller build.spec
#
# Output: dist/CATIA Copilot/
# The executable is placed in dist/CATIA Copilot/ and all supporting files
# (resources/, macros/, catia_copilot/, etc.) are placed inside the default
# _internal/ subdirectory alongside it.
#
# Before building, place the application icon at:
#   resources/icon.ico
# Then uncomment the `icon=` line in the EXE block below.

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('resources', 'resources'),
        ('macros', 'macros'),
        ('drawing_templates', 'drawing_templates'),
        ('crack', 'crack'),
        ('catia_copilot/ui/style.qss', 'catia_copilot/ui'),
        ('catia_copilot', 'catia_copilot'),
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
    name='CATIA Copilot 1.4.2',
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

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CATIA Copilot 1.4.2',
)
