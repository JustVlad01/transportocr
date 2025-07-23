# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['optimoroute_sorter_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('delivery_sequence_data.json', '.'),
        ('supabase_config.py', '.'),
        ('supabase_schema.sql', '.'),
        ('your_order_format.xlsx', '.'),
    ],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui', 
        'PySide6.QtWidgets',
        'pytesseract',
        'PIL._tkinter_finder',
        'requests',
        'pandas',
        'fitz',
        'reportlab',
        'barcode',
        'openpyxl'
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OptimoRouteSorter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # You can add an icon file here
    version_file=None
)
