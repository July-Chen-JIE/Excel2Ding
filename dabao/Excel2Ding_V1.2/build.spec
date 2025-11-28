# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src/Excel2Ding_V1.2.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['babel.numbers'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 排除不必要的模块
    excludes=['matplotlib', 'notebook', 'scipy', 'pandas.io.formats.style',
              'PySide2', 'PyQt6', 'tcl', 'tk', 'wx'],
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
    name='Excel2Ding',
    debug=False,
    strip=True,  # 减小二进制体积
    upx=True,    # 使用UPX压缩
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='src/icon.ico',  # 确保图标文件路径正确
    version='file_version_info.txt',  # 添加版本信息
    # 添加优化选项
    optimize=2,
    bundle_identifier=None,  # 禁用 Mac bundle
    uac_admin=False,        # 禁用 UAC 提权
)