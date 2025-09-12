# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('app_data', 'app_data'), ('app_icon.ico', '.')]
binaries = []
hiddenimports = [
    'barcode', 'barcode.writer', 'barcode.codex', 'barcode.codex.code128', 'barcode.codex.code39', 
    'barcode.codex.ean13', 'barcode.codex.ean8', 'barcode.codex.upc', 'barcode.codex.isbn10', 
    'barcode.codex.isbn13', 'barcode.codex.issn', 'barcode.codex.jan', 'barcode.codex.pzn', 
    'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont', 'PIL.ImageTk', 
    'pytesseract', 'fitz', 'reportlab', 'reportlab.pdfgen', 'reportlab.lib.pagesizes', 'reportlab.lib.colors', 
    'pandas', 'openpyxl', 'supabase', 'supabase_config', 'requests', 
    'PySide6', 'PySide6.QtWidgets', 'PySide6.QtCore', 'PySide6.QtGui',
    # Printer communication libraries
    'serial', 'pyserial', 'win32print', 'win32api', 'win32con', 'win32gui', 'win32ui',
    'subprocess', 'tempfile', 'shutil', 'hashlib', 'datetime', 'time', 'io', 're',
    # Additional Windows API modules
    'pywintypes', 'pythoncom', 'win32com', 'win32com.client',
    # System modules
    'pathlib', 'json', 'os', 'sys'
]
hiddenimports += collect_submodules('barcode')
hiddenimports += collect_submodules('PIL')
hiddenimports += collect_submodules('reportlab')
tmp_ret = collect_all('barcode')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('reportlab')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('supabase')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PySide6')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('win32print')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('win32api')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('serial')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['dispatch_scanning_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 'matplotlib', 'numpy', 'scipy',  # Exclude heavy unused libraries
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Dispatch Scanning App',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'win32print', 'win32api', 'serial',  # Don't compress printer-related modules
    ],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.ico'],
    version=None,
    uac_admin=False,
    uac_uiaccess=False,
)
