# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\João Amadeu\\OneDrive\\Documentos\\GitHub\\TirelessBud\\code\\main_exe_safe.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\João Amadeu\\OneDrive\\Documentos\\GitHub\\TirelessBud\\LogoTBud.ico', '.')],
    hiddenimports=['pymupdf', 'PyPDF2', 'openpyxl', 'PIL'],
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
    name='TirelessBud_v1.0.0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='C:\\Users\\João Amadeu\\OneDrive\\Documentos\\GitHub\\TirelessBud\\build\\file_version_info.txt',
    icon=['C:\\Users\\João Amadeu\\OneDrive\\Documentos\\GitHub\\TirelessBud\\LogoTBud.ico'],
    manifest='C:\\Users\\João Amadeu\\OneDrive\\Documentos\\GitHub\\TirelessBud\\build\\app.manifest',
)
