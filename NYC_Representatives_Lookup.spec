# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('QEDC - Full Logo (Primary Color).png', '.'),
        ('icon_256x256.ico',                      '.'),
        ('Zipcodes-with-Reps-Complete.xlsx',       '.'),
    ],
    hiddenimports=[
        'PyQt5', 'PyQt5.QtWidgets', 'PyQt5.QtCore', 'PyQt5.QtGui',
        'pandas', 'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype', 'pandas._libs.tslibs.timedeltas',
        'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
        'requests', 'bs4', 'lxml', 'lxml.etree',
        'certifi', 'charset_normalizer', 'urllib3',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='NYC_Representatives_Lookup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False, upx=True,
    console=False,
    icon='icon_256x256.ico',
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, upx_exclude=[],
    name='NYC_Representatives_Lookup',
)
