# -*- mode: python ; coding: utf-8 -*-

# 排除目前環境中與標籤列印無關的大型套件，避免 PyInstaller 因可選 hook
# 把 AI、科學運算、資料分析、測試框架等依賴一併打包。
EXCLUDES = [
    'IPython',
    'PyQt5',
    'PyQt6',
    'PySide2',
    'PySide6',
    'bcrypt',
    'cryptography',
    'cv2',
    'easyocr',
    'fsspec',
    'imageio',
    'jinja2',
    'matplotlib',
    'numba',
    'numpy',
    'paddle',
    'paddleocr',
    'pandas',
    'psutil',
    'psycopg2',
    'pytest',
    'scipy',
    'skimage',
    'sklearn',
    'sqlalchemy',
    'sympy',
    'tensorboard',
    'tensorflow',
    'torch',
    'torchvision',
]

a = Analysis(
    ['進料標籤列印.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['fitz', 'pymupdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='必榮進料標籤列印',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='必榮進料標籤列印',
)
