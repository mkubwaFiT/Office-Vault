# -*- mode: python ; coding: utf-8 -*-
#
# Optimized build spec for the harmonized Vault Toolkit.
#
# Why this is faster than the old OfficeVault/TextVault one-file builds:
#   * onedir (exclude_binaries=True + COLLECT) instead of --onefile.
#     One-file unpacks the whole ~11 MB runtime into %TEMP% on EVERY launch;
#     onedir runs straight from the folder -> dramatically faster startup.
#   * upx=False everywhere. UPX decompresses on launch AND frequently trips
#     Windows Defender heuristics, which then scans the binary every run.
#   * optimize=2 strips docstrings/asserts; excludes drop heavy libs that
#     PyInstaller would otherwise vacuum up. This is a stdlib + tkinter app.

block_cipher = None

a = Analysis(
    ['vault_toolkit.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['sqlite3'],  # belt-and-suspenders; also NOT in excludes below
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Heavy third-party libs (only excluded if present) — never imported here
        'numpy', 'pandas', 'scipy', 'matplotlib', 'PIL', 'cv2',
        'IPython', 'notebook', 'jupyter',
        # Build/dev tooling that should never ship in the runtime
        'pip', 'setuptools', 'wheel', 'pytest', 'pydoc',
        # Unused stdlib subsystems for a local tkinter file tool.
        # NOTE: do NOT exclude 'sqlite3' — v2 needs it for the FTS5 catalog;
        # PyInstaller's hook bundles _sqlite3.pyd + sqlite3.dll automatically.
        'test', 'unittest', 'lib2to3', 'distutils',
    ],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,          # onedir mode
    name='Vault_Toolkit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                      # no UPX -> faster launch, fewer AV false positives
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Vault_Toolkit',
)
