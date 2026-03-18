# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# ── Data files ────────────────────────────────────────────────────────────────
added_files = [
    # ── Your project assets (all land in the root of the bundle) ──
    ('QEDC - Full Logo (Primary Color).png',  '.'),
    ('generated-image.png',                    '.'),
    ('icon_256x256.ico',                       '.'),
    ('Zipcodes-with-Reps-Complete.xlsx',        '.'),

    # ── Runtime dependencies ──
    *collect_data_files('openpyxl'),
    *collect_data_files('certifi'),
]

# ── Hidden imports ────────────────────────────────────────────────────────────
hidden = [
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.writer.excel',
    'openpyxl.reader.excel',
    'bs4',
    'lxml',
    'lxml.etree',
    'requests',
    'certifi',
    'urllib3',
    'charset_normalizer',
    'pandas',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.skiplist',
    'PyQt5',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.sip',
]

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ['main.py'],
    pathex=[os.path.abspath('.')],
    binaries=[],
    datas=[
        ('QEDC - Full Logo (Primary Color).png',  '.'),
        ('generated-image.png',                    '.'),
        ('icon_256x256.ico',                       '.'),
        ('Zipcodes-with-Reps-Complete.xlsx',        '.'),
        *collect_data_files('openpyxl'),
        *collect_data_files('certifi'),
    ],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'PIL',
        'tkinter',
        'test',
        'unittest',
        'email',
        'xmlrpc',
        'pydoc',
        'doctest',
        'difflib',
        'ftplib',
        'getopt',
        'calendar',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ── PYZ archive ───────────────────────────────────────────────────────────────
pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

# ── Single-file EXE ───────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='NYC_Representatives_Lookup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'python3*.dll',
        'Qt5Core.dll',
        'Qt5Gui.dll',
        'Qt5Widgets.dll',
    ],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon_256x256.ico',           # ← matches your exact .ico filename
)
