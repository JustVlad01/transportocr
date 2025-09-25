# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['dispatch_scanning_app.py'],
    pathex=[],
    binaries=[],
    datas=[('app_data', 'app_data'), ('C:\\Users\\vladk\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\barcode\\fonts', 'barcode/fonts')],
    hiddenimports=['sys', 'os', 'json', 'pathlib', 'pandas', 'fitz', 'PyMuPDF', 'pytesseract', 'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont', 'PIL.ImageTk', 'io', 're', 'openpyxl', 'openpyxl.styles', 'openpyxl.workbook', 'openpyxl.worksheet', 'openpyxl.cell', 'openpyxl.utils', 'reportlab', 'reportlab.pdfgen', 'reportlab.pdfgen.canvas', 'reportlab.lib', 'reportlab.lib.pagesizes', 'reportlab.lib.colors', 'reportlab.lib.units', 'reportlab.lib.utils', 'barcode', 'barcode.writer', 'barcode.writer.base', 'barcode.writer.image', 'barcode.writer.svg', 'barcode.writer.pdf', 'barcode.base', 'barcode.errors', 'barcode.codex', 'barcode.codex.base', 'barcode.codex.code128', 'barcode.codex.code39', 'barcode.codex.ean13', 'barcode.codex.ean8', 'barcode.codex.upc', 'barcode.codex.isbn10', 'barcode.codex.isbn13', 'barcode.codex.issn', 'barcode.codex.jan', 'barcode.codex.pzn', 'hashlib', 'requests', 'datetime', 'time', 'serial', 'pyserial', 'subprocess', 'tempfile', 'shutil', 'win32print', 'win32api', 'win32con', 'win32gui', 'win32ui', 'pywintypes', 'pythoncom', 'win32com', 'win32com.client', 'PySide6', 'PySide6.QtWidgets', 'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtOpenGL', 'PySide6.QtPrintSupport', 'PySide6.QtSvg', 'PySide6.QtTest', 'PySide6.QtUiTools', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineCore', 'PySide6.QtWebSockets', 'PySide6.QtXml', 'PySide6.QtNetwork', 'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets', 'PySide6.QtPositioning', 'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtQuickWidgets', 'PySide6.QtSql', 'PySide6.QtSvgWidgets', 'PySide6.QtTest', 'PySide6.QtUiTools', 'PySide6.QtWebChannel', 'PySide6.QtWebEngine', 'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebSockets', 'PySide6.QtXml', 'PySide6.QtXmlPatterns', 'supabase', 'supabase_config', 'numpy', 'uuid', 'typing', 'platform', 'PyQt5', 'PyQt5.QtWidgets', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtPrintSupport', 'PyQt5.QtSvg', 'PyQt5.QtTest', 'PyQt5.QtUiTools', 'PyQt5.QtWebEngineWidgets', 'PyQt5.QtWebEngineCore', 'PyQt5.QtWebSockets', 'PyQt5.QtXml', 'PyQt5.QtNetwork', 'PyQt5.QtMultimedia', 'PyQt5.QtMultimediaWidgets', 'PyQt5.QtPositioning', 'PyQt5.QtQml', 'PyQt5.QtQuick', 'PyQt5.QtQuickWidgets', 'PyQt5.QtSql', 'PyQt5.QtSvgWidgets', 'PyQt5.QtTest', 'PyQt5.QtUiTools', 'PyQt5.QtWebChannel', 'PyQt5.QtWebEngine', 'PyQt5.QtWebEngineCore', 'PyQt5.QtWebEngineWidgets', 'PyQt5.QtWebSockets', 'PyQt5.QtXml', 'PyQt5.QtXmlPatterns'],
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
