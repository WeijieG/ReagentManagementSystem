# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 构建配置 —— onedir 布局。

为什么是 onedir 而不是原来的 onefile：
  1. QtWebEngine 的辅助进程 QtWebEngineProcess.exe 需要相对主程序可预测的路径，
     onefile 每次解压到 _MEIxxxx 临时目录，会报 "Could not find
     QtWebEngineProcess.exe" / icu_util.cc FATAL。
  2. Chromium 运行时约 150MB，onefile 每次启动都要解压一遍，Win7 机械盘上
     可达数十秒。

UPX 必须关闭：对 Qt5WebEngineCore.dll / QtWebEngineProcess.exe 上 UPX 是崩溃与
杀软误报的常见来源，且对已优化的 75MB Chromium DLL 也压不动。

--------------------------------------------------------------------------
备选：改用 onefile
--------------------------------------------------------------------------
旧版（37MB，PyQt5 无 WebEngine）是 onefile，在 Win7 上一直能跑。如果 onedir
在目标机上出问题，可以退回 onefile —— 改法：

  1. EXE(...) 里删掉 `exclude_binaries=True`，并把第 3 个参数从 `[]` 改回
     `a.binaries, a.zipfiles, a.datas`
  2. 删掉末尾整个 COLLECT(...)

注意 onefile 对 QtWebEngine 有已知风险（辅助进程路径问题、每次启动解压
约 150MB），所以默认不走这条路。详见 docs/BUILD.md 第 2 节。
"""

import os

block_cipher = None

# QtWebEngine 语言包裁剪：共 53 个 .pak 约 17MB，实际只用中英两种。
# ⚠️ en-US 必须保留 —— Chromium 的回退依赖它，缺了是硬失败。
KEEP_LOCALES = {"en-US.pak", "zh-CN.pak"}

# 明确用不到、且体积可观的模块，排掉能省不少空间。
# 注意：PyQt5.QtQml / QtQuick / QtPositioning / QtNetwork / QtWebChannel
# 都是 QtWebEngine 的依赖，**绝不能**放进这里。
EXCLUDES = [
    "tkinter",
    "unittest",
    "pydoc_data",
    "matplotlib",
    "numpy",
    "scipy",
    "PIL",
    "pytest",
    "PyQt5.QtBluetooth",
    "PyQt5.QtNfc",
    "PyQt5.QtDesigner",
    "PyQt5.QtHelp",
    "PyQt5.QtTest",
    "PyQt5.QtSql",
    "PyQt5.QtMultimedia",
    "PyQt5.QtMultimediaWidgets",
    "PyQt5.QtSensors",
    "PyQt5.QtSerialPort",
    "PyQt5.QtRemoteObjects",
    "PyQt5.QtTextToSpeech",
    "PyQt5.QtWebKit",
    "PyQt5.QtWebKitWidgets",
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('reagent.ico', '.'),
        ('config.ini', '.'),
        # Web UI 资产。paths.get_ui_dir() 在打包后会到 exe 同级的 ui/ 找。
        ('app/ui', 'ui'),
        ('images', 'images'),
    ],
    hiddenimports=[
        'sqlite3', 'csv', 'json', 're', 'hashlib', 'shutil', 'tempfile',
        'configparser', 'subprocess', 'datetime', 'traceback',
        # QtWebEngine 的传递依赖，必须显式声明
        'PyQt5.sip', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets',
        'PyQt5.QtNetwork', 'PyQt5.QtWebChannel',
        'PyQt5.QtWebEngineCore', 'PyQt5.QtWebEngineWidgets',
        'PyQt5.QtQml', 'PyQt5.QtQuick', 'PyQt5.QtPositioning',
        'PyQt5.QtPrintSupport',
        # 业务依赖
        'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
        'requests', 'urllib3', 'certifi', 'idna', 'charset_normalizer',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 裁剪语言包（在 PYZ 之前处理，datas 里已经含 Qt 的文件）
_kept = []
_dropped = 0
for entry in a.datas:
    dest = entry[0]
    if 'qtwebengine_locales' in dest:
        if os.path.basename(dest) not in KEEP_LOCALES:
            _dropped += 1
            continue
    # DevTools 资源只在调试时用得上，正式包里不留
    if 'qtwebengine_devtools_resources' in dest:
        _dropped += 1
        continue
    _kept.append(entry)
a.datas = _kept
print(f"[build] 已裁剪 QtWebEngine 语言包/调试资源 {_dropped} 个文件")

# ---------------------------------------------------------------------------
# 排除 MSVC 运行库（MSVCP140*.dll / VCRUNTIME140*.dll / CONCRT140*.dll）
#
# 这一条不是优化，是**修 bug**。
#
# PyInstaller 会从**打包机的 System32** 里抓这些文件，于是打包机上装了什么
# 版本就带什么版本；而开发机的 Visual Studio 总是最新的。实测踩过一次：
#
#   开发机 System32 的 MSVCP140.dll 是 14.51（Visual Studio 2026），
#   而这一版**已经放弃 Windows 7** —— 它静态导入了 Win8 才有的
#   GetSystemTimePreciseAsFileTime，于是在 Win7 上加载即失败：
#
#       ReagentManagementSystem.exe - 无法找到入口
#       无法定位程序输入点 GetSystemTimePreciseAsFileTime 于 KERNEL32.dll 上。
#
# 最阴的地方在于：这个 app-local 副本会**盖住** System32 里那份好的
# （应用目录在 DLL 搜索顺序里最靠前），所以目标机就算装了正确的 VC++
# 运行库也没用 —— 装了完整安装包反而不如只发一个 exe 的补丁包能跑。
#
# 正确做法是交给系统，由安装包里的 VC_redist.x64.exe 负责 ——
# 那份是刻意钉死在 14.44 的，最后一个还支持 Win7 的版本。
# 版本一致性也一并解决了：以前产物里是 MSVCP140 14.51 + MSVCP140_1 14.26
# 这种混搭，本身就不该出现。
#
# tools/build_installer.py 里有对应的构建期检查，会扫产物有没有静态导入
# Win8+ 的 API，防止这类问题再溜过去。
MSVC_RUNTIME_PREFIXES = ("msvcp140", "vcruntime140", "concrt140", "vccorlib140")

_kept_binaries = []
_dropped_msvc = []
for entry in a.binaries:
    base = os.path.basename(entry[0]).lower()
    if base.startswith(MSVC_RUNTIME_PREFIXES):
        _dropped_msvc.append(entry[0])
        continue
    _kept_binaries.append(entry)
a.binaries = _kept_binaries

if _dropped_msvc:
    print(f"[build] 已排除 MSVC 运行库 {len(_dropped_msvc)} 个（交给系统/VC_redist）: "
          + "、".join(os.path.basename(p) for p in _dropped_msvc))

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,          # onedir 的关键
    name='ReagentManagementSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                      # 见文件头说明，必须关闭
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

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ReagentManagementSystem',
)
