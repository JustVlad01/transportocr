# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['dispatch_scanning_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app_data/*.json', 'app_data'),
        ('app_icon.ico', '.'),
    ],
    hiddenimports=[
        'barcode',
        'barcode.writer',
        'barcode.writer.ImageWriter',
        'barcode.codex',
        'barcode.codex.code128',
        'barcode.codex.code39',
        'barcode.codex.ean13',
        'barcode.codex.ean8',
        'barcode.codex.upc',
        'barcode.codex.isbn10',
        'barcode.codex.isbn13',
        'barcode.codex.issn',
        'barcode.codex.jan',
        'barcode.codex.pzn',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        'pytesseract',
        'fitz',
        'PyMuPDF',
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.lib.pagesizes',
        'reportlab.lib.colors',
        'pandas',
        'openpyxl',
        'supabase',
        'supabase_config'
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
    name='Dispatch Scanning App',
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
    icon='app_icon.ico'
)
