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

# contents_directory='.' places all bundle files directly next to the .exe,
# eliminating the _internal/ subdirectory introduced in PyInstaller 6.0.
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CATIA Companion',
    contents_directory='.',
)
