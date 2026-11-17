# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, copy_metadata

datas = [('app_icon.ico', '.'), ('app_icon.png', '.')]
binaries = []
hiddenimports = [
    'numpy',
    'onnxruntime',
    'keyring',
    'keyring.backends',
    'keyring.backends.Windows',
    'keyring.backends.fail',
    'jaraco.classes',
    'jaraco.functools',
    'jaraco.context',
    'openai',
    'pydub',
    'pymupdf',
    'fitz',
    'olefile',
    'tkinterdnd2',
    'PIL',
    'PIL.Image',
    'pandas',
    'openpyxl',
    'pdfminer',
    'pdfminer.high_level',
    'pdfplumber',
    # Tab 2: PPT ke PDF
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    # Tab 3: Gabung PDF
    'pypdf',
    'pypdf._reader',
    'pypdf._writer',
    'pypdf.filters',
    'pypdf.generic',
    'pypdf.errors',
    # Tab 4: PDF Optimizer / Compressor
    'pdf_optimizer',
]

for pkg in ['magika', 'markitdown', 'markitdown_ocr', 'tkinterdnd2', 'sv_ttk']:
    d, b, h = collect_all(pkg)
    datas.extend(d)
    binaries.extend(b)
    hiddenimports.extend(h)

datas.extend(copy_metadata('markitdown_ocr'))
datas.extend(copy_metadata('markitdown'))
datas.extend(copy_metadata('magika'))

a = Analysis(
    ['markitdown_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch', 'torchvision', 'torchaudio',
        'scipy', 'matplotlib',
        'IPython', 'jupyter', 'notebook',
        'cv2', 'sklearn', 'tensorflow', 'transformers',
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
    name='DocKit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['_fitz.pyd', 'mupdf.dll'],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
    version_file=None,
)
