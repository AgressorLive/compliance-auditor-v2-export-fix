# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files, collect_dynamic_libs

hiddenimports = [
    # PySide6
    'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets',
    'PySide6.QtNetwork', 'PySide6.QtPrintSupport', 'PySide6.QtSvg',
    # Pillow – import all sub-modules so imaging works inside exe
    'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont',
    'PIL.ImageFilter', 'PIL.ImageOps', 'PIL.PngImagePlugin',
    'PIL.JpegImagePlugin', 'PIL.BmpImagePlugin', 'PIL.GifImagePlugin',
    'PIL.TiffImagePlugin', 'PIL.WebPImagePlugin',
    # python-docx
    'docx', 'docx.oxml', 'docx.oxml.ns', 'docx.shared',
    # fpdf2
    'fpdf',
    # networking / scraping
    'requests', 'urllib3', 'certifi', 'charset_normalizer',
    'bs4', 'bs4.builder', 'bs4.builder._htmlparser',
    # OpenAI
    'openai',
    # local core package
    'core', 'core.site_profiler', 'core.analyzer', 'core.ban_analyzer',
    'core.bulk_scanner', 'core.cloaking_detector', 'core.profile_builder',
    'core.review_engine', 'core.summary',
]
hiddenimports += collect_submodules('PIL')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('fpdf')
hiddenimports += collect_submodules('openai')
hiddenimports += collect_submodules('PySide6')

datas = [
    ('C:\\Users\\alexx\\OneDrive\\Desktop\\Bastliga One.otf', '.'),
    (r'C:\Users\alexx\AppData\Local\Microsoft\Windows\Fonts\DieselpowerPersonalUse-axaY5.ttf', '.'),
    (r'C:\Users\alexx\AppData\Local\Microsoft\Windows\Fonts\Interplanetary Crap.otf', '.'),
    ('..\\core', 'core'),   # include entire local core/ package
    ('AdsMailImporter.gs', '.'),
]
datas += collect_data_files('PIL')
datas += collect_data_files('fpdf')
datas += collect_data_files('certifi')

a = Analysis(
    ['doc_generator.py'],
    pathex=['.', '..'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['..\\pyi_rthook_pil.py'],
    excludes=['tkinter', 'matplotlib', 'numpy', 'pandas', 'scipy'],
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
    name='DocGenerator',
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
    icon=['assets\\docgen.ico'],
)
