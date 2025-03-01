# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['update_auto_create_overwrite_word.py'],
    pathex=['D:\\project_root'],
    binaries=[],
    datas=[
        ('SyncBackup', 'SyncBackup'),
        ('word_icon.ico', '.')
    ],
    hiddenimports=[
        'win32timezone',
        'win32com.client',
        'pythoncom',
        'psutil._psutil_windows'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter'],
    cipher=block_cipher,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AutoWordSync',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    icon='word_icon.ico',
    version_info={
        'version': '1.0.0',
        'company_name': 'YourCompany',
        'file_description': 'Word实时同步工具'
    },
    onefile=True  # ✅ 关键配置
)

# ❌ 删除或注释掉COLLECT部分
# coll = COLLECT(...)
