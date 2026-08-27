# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_dynamic_libs

datas = []
binaries = [('C:/Users/mengqingqi/AppData/Local/Programs/Python/Python313/Lib/site-packages/pywin32_system32/pythoncom313.dll', '.'), ('C:/Users/mengqingqi/AppData/Local/Programs/Python/Python313/Lib/site-packages/pywin32_system32/pywintypes313.dll', '.')]
datas += collect_data_files('tkinterdnd2')
binaries += collect_dynamic_libs('tkinterdnd2')


a = Analysis(
    ['decrypt_gui.py'],
    pathex=['C:/Users/mengqingqi/AppData/Local/Programs/Python/Python313/Lib/site-packages'],
    binaries=binaries,
    datas=datas,
    hiddenimports=['win32com', 'win32com.client', 'win32com.gen_py', 'pythoncom', 'tkinterdnd2', 'context_menu', 'win32event', 'win32api'],
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
    name='DLP批量解密工具',
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
)
