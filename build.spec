# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon.png', '.'),  # 包含图标文件
        ('C:/Python310/Lib/site-packages/scipy/.libs/*', 'scipy/.libs')  # 包含scipy依赖
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'tkinter', 'PIL'],  # 排除不需要的模块
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
    name='BCAAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 启用UPX压缩
    console=False,  # 不显示控制台窗口
    icon='icon.ico',  # 使用单独的图标文件
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
