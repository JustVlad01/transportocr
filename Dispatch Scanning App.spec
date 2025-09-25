# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all
import os
import barcode

# Get the barcode package directory to find font files
barcode_package_dir = os.path.dirname(barcode.__file__)

datas = [
    ('app_data', 'app_data'), 
    ('app_icon.ico', '.'),
    # Include barcode font files
    (os.path.join(barcode_package_dir, 'fonts'), 'barcode/fonts'),
    # Include any other barcode data files
    (os.path.join(barcode_package_dir, 'data'), 'barcode/data'),
]
binaries = []
hiddenimports = [
    'barcode', 'barcode.writer', 'barcode.codex', 'barcode.codex.code128', 'barcode.codex.code39', 
    'barcode.codex.ean13', 'barcode.codex.ean8', 'barcode.codex.upc', 'barcode.codex.isbn10', 
    'barcode.codex.isbn13', 'barcode.codex.issn', 'barcode.codex.jan', 'barcode.codex.pzn', 
    # Additional barcode writer components
    'barcode.writer.base', 'barcode.writer.svg', 'barcode.writer.image', 'barcode.writer.pdf',
    'barcode.errors', 'barcode.base', 'barcode.codex.base',
    'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont', 'PIL.ImageTk', 
    'pytesseract', 'fitz', 'reportlab', 'reportlab.pdfgen', 'reportlab.lib.pagesizes', 'reportlab.lib.colors', 
    'pandas', 'numpy', 'openpyxl', 'supabase', 'supabase_config', 'requests', 
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
# Collect all barcode resources
tmp_ret = collect_all('barcode')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Additional barcode data collection - ensure fonts are included
try:
    import barcode
    barcode_dir = os.path.dirname(barcode.__file__)
    
    # Add any additional barcode data files
    for root, dirs, files in os.walk(barcode_dir):
        for file in files:
            if file.endswith(('.ttf', '.otf', '.woff', '.woff2', '.json', '.txt')):
                rel_path = os.path.relpath(root, barcode_dir)
                if rel_path == '.':
                    datas.append((os.path.join(root, file), 'barcode'))
                else:
                    datas.append((os.path.join(root, file), f'barcode/{rel_path}'))
except Exception as e:
    print(f"Warning: Could not collect additional barcode resources: {e}")
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('reportlab')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('numpy')
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
    runtime_hooks=['hook-barcode.py'],
    excludes=[
        'tkinter', 'matplotlib', 'scipy',  # Exclude heavy unused libraries (numpy needed for pandas)
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
