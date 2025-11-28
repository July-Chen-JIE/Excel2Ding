# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src/Excel2Ding_V1.4.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('src/Excel2Ding.ico', '.'),
        ('src/config/column_mapping.json', 'config/'),  # 修改配置文件路径
    ],
    hiddenimports=[
        'babel.numbers',
        'tkinter',
        'tkinter.ttk',
        'tkcalendar',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'notebook', 'scipy',
        'pandas.io.formats.style',
        'PySide2', 'PyQt6', 'wx'
    ],
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
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='src/Excel2Ding.ico',
    version='file_version_info.txt',
    optimize=2,
)