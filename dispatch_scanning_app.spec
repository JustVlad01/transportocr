# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['dispatch_scanning_app.py'],
    pathex=[],
    binaries=[],
    datas=[('app_data', 'app_data')],
    hiddenimports=['serial', 'pyserial', 'win32print', 'win32api', 'win32con', 'win32gui', 'win32ui', 'pywintypes', 'pythoncom', 'win32com', 'win32com.client', 'subprocess', 'tempfile', 'shutil', 'hashlib', 'datetime', 'time', 'io', 're', 'pathlib', 'json', 'os', 'sys', 'barcode', 'barcode.writer', 'barcode.codex', 'barcode.codex.code128', 'barcode.codex.code39', 'barcode.codex.ean13', 'barcode.codex.ean8', 'barcode.codex.upc', 'barcode.codex.isbn10', 'barcode.codex.isbn13', 'barcode.codex.issn', 'barcode.codex.jan', 'barcode.codex.pzn', 'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont', 'PIL.ImageTk', 'pytesseract', 'fitz', 'PyMuPDF', 'reportlab', 'reportlab.pdfgen', 'reportlab.lib.pagesizes', 'reportlab.lib.colors', 'pandas', 'openpyxl', 'supabase', 'supabase_config', 'requests', 'PySide6', 'PySide6.QtWidgets', 'PySide6.QtCore', 'PySide6.QtGui', 'numpy'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy'],
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
    name='dispatch_scanning_app',
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
    icon=['app_icon.ico'],
)
