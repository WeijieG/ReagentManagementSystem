# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

import os
import sys
import ctypes

# 获取 PyQt5 插件路径
try:
    from PyQt5 import QtCore
    pyqt5_plugin_dir = os.path.join(os.path.dirname(QtCore.__file__), 'Qt', 'plugins')
except ImportError:
    pyqt5_plugin_dir = ""

a = Analysis(
    ['V8_1.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 图标和数据库
        ('reagent.ico', '.'),
    ],
    hiddenimports=[
        'sqlite3',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.sip',
        'csv',
        'datetime',
        'random'
    ],
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
    name='ReagentManagementSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='reagent.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)