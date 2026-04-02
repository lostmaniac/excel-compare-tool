# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller spec file for Excel Compare Tool

This configuration file is optimized for Windows builds.
Run: pyinstaller ExcelCompare.spec
"""

import os

a = Analysis(
    ['src/excel_compare/main.py'],
    pathex=[],
    binaries=[],
    datas=[('src/excel_compare', 'excel_compare')],
    hiddenimports=[
        'tkinterdnd2',
        'tkinterdnd2.tkdnd',
        'tkinterdnd2.tkdnd_wrapper',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy.testing',
        'scipy.testing',
        'pandas.testing',
        'pandas.tests',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelCompare',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
