# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Usuario\\conversor markdown\\assets\\logo.jpg', 'assets'), ('C:\\Users\\Usuario\\conversor markdown\\LICENSE', '.'), ('C:\\Users\\Usuario\\conversor markdown\\.venv\\Lib\\site-packages\\rapidocr_onnxruntime', 'rapidocr_onnxruntime')],
    hiddenimports=['docx', 'pptx', 'openpyxl', 'pymupdf', 'rapidocr_onnxruntime', 'striprtf', 'tkinterdnd2', 'requests', 'markdownify', 'bs4', 'pyclipper', 'shapely', 'shapely.geometry', 'six'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='preparador_de_archivos_para_ia',
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
    version='C:\\Users\\Usuario\\conversor markdown\\build\\version_info.txt',
    icon=['C:\\Users\\Usuario\\conversor markdown\\icon.ico'],
)
