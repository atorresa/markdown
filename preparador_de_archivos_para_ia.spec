# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/logo.jpg', 'assets')],
    hiddenimports=['docx', 'pptx', 'openpyxl', 'pypdf', 'tkinterdnd2', 'requests', 'markdownify', 'bs4'],
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
    icon=['C:\\Users\\Usuario\\conversor markdown\\icon.ico'],
)
