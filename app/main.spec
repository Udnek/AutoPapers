# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['C:/Coding/AutoCards_(SS2023)/app/main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Coding/AutoCards_(SS2023)/app/app.py', '.'), ('C:/Coding/AutoCards_(SS2023)/app/generator.py', '.'), ('C:/Coding/AutoCards_(SS2023)/app/get_excel_list.py', '.'), ('C:/Coding/AutoCards_(SS2023)/app/libs_dima.py', '.'), ('C:/Coding/AutoCards_(SS2023)/app/libs_gleb.py', '.'), ('C:/Coding/AutoCards_(SS2023)/app/manager.py', '.')],
    hiddenimports=[],
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
    name='main',
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
    icon=['C:\\Coding\\AutoCards_(SS2023)\\app\\assets\\icon.ico'],
)
