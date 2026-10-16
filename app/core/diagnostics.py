# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""启动/运行失败时的环境自检。

打包后的程序没有控制台，一旦启动就崩，用户看到的只是「闪一下」，
什么线索都没有。这里把环境信息和堆栈一起写成文件，让用户能直接发回来。

**本模块刻意只依赖标准库里最轻的几个**（os/sys/ctypes/platform/struct），
绝不 import socket / requests 之类 —— 那些正是可能坏掉的东西，诊断代码
自己踩上去就什么都报不出来了。
"""

import os
import sys


def get_app_dir():
    """exe 所在目录（打包后）或仓库根目录（开发时）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


# 程序目录里值得逐一测试加载的文件
RUNTIME_FILES = [
    "python38.dll",
    "VCRUNTIME140.dll",
    "VCRUNTIME140_1.dll",
    "MSVCP140.dll",
    "MSVCP140_1.dll",
    "ucrtbase.dll",
    "api-ms-win-crt-runtime-l1-1-0.dll",
    "api-ms-win-crt-string-l1-1-0.dll",
    "_socket.pyd",
    "_ssl.pyd",
    "_hashlib.pyd",
    "_sqlite3.pyd",
    "sqlite3.dll",
    "libffi-7.dll",
]

QT_FILES = [
    "Qt5Core.dll",
    "Qt5Gui.dll",
    "Qt5Network.dll",
    "Qt5WebEngineCore.dll",
    "Qt5WebEngineWidgets.dll",
    "libEGL.dll",
    "libGLESv2.dll",
    "opengl32sw.dll",
    "d3dcompiler_47.dll",
]

SYSTEM_DLLS = [
    ("ucrtbase.dll", "KB2999226 通用 C 运行库"),
    ("vcruntime140.dll", "VC++ 2015-2022 运行库"),
    ("msvcp140.dll", "VC++ C++ 标准库"),
    ("WS2_32.dll", "Winsock（_socket 依赖）"),
    ("IPHLPAPI.DLL", "IP 助手 API（_socket 依赖）"),
    ("NSI.dll", "网络存储接口（IPHLPAPI 依赖）"),
    ("KERNEL32.dll", "核心"),
]


def _try_load(path):
    """尝试加载 DLL，返回 (是否成功, 错误描述)。"""
    try:
        import ctypes
        ctypes.WinDLL(path)
        return True, ""
    except OSError as e:
        return False, f"WinError={getattr(e, 'winerror', '?')} {e.strerror or e}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def _probe_dir(base, names, label, lines):
    lines.append(f"--- {label} ---")
    if not os.path.isdir(base):
        lines.append(f"  （目录不存在: {base}）")
        lines.append("")
        return
    for name in names:
        full = os.path.join(base, name)
        if not os.path.isfile(full):
            lines.append(f"  [不存在] {name}")
            continue
        ok, err = _try_load(full)
        lines.append(f"  [{'可加载' if ok else '失败  '}] {name}"
                     + (f"   <- {err}" if err else ""))
    lines.append("")


def collect_diagnostics(app_dir=None):
    """收集环境信息。任何一段失败都不会影响其它段。"""
    app_dir = app_dir or get_app_dir()
    lines = []

    lines.append("--- 环境 ---")
    try:
        import platform
        import struct
        lines.append(f"Python      : {sys.version.split()[0]}")
        lines.append(f"位数        : {struct.calcsize('P') * 8} 位")
        lines.append(f"操作系统    : {platform.platform()}")
        lines.append(f"系统版本    : {platform.version()}")
    except Exception as e:
        lines.append(f"（读取系统信息失败: {e}）")
    lines.append(f"frozen      : {getattr(sys, 'frozen', False)}")
    lines.append(f"程序目录    : {app_dir}")
    lines.append(f"APPDATA     : {os.environ.get('APPDATA', '(无)')}")

    # Win7 兼容层状态 —— 这直接关系到 Qt 能否加载
    try:
        from .. import compat
        lines.append(compat.describe())
    except Exception as e:
        lines.append(f"（兼容层状态读取失败: {e}）")

    lines.append("")

    # 系统目录里的关键 DLL —— 只报告存在性与可加载性
    lines.append("--- 系统目录关键 DLL ---")
    sys32 = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32")
    lines.append(f"  （{sys32}）")
    for name, why in SYSTEM_DLLS:
        full = os.path.join(sys32, name)
        if not os.path.isfile(full):
            lines.append(f"  [缺失  ] {name:<24} {why}")
            continue
        ok, err = _try_load(full)
        lines.append(f"  [{'可加载' if ok else '失败  '}] {name:<24} {why}"
                     + (f"   <- {err}" if err else ""))
    lines.append("")

    _probe_dir(app_dir, RUNTIME_FILES, "程序目录里的运行库", lines)
    _probe_dir(app_dir, QT_FILES, "Qt 相关原生 DLL", lines)

    lines.append("--- 逐模块导入自检 ---")
    for name in ("ctypes", "sqlite3", "json", "csv", "hashlib", "zlib",
                 "_socket", "socket", "ssl", "select"):
        try:
            __import__(name)
            lines.append(f"  [成功] import {name}")
        except Exception as e:
            lines.append(f"  [失败] import {name}   <- {type(e).__name__}: {e}")
    lines.append("")

    return "\n".join(lines)


def _desktop_dir():
    """桌面路径。报告放这里用户最容易找到。"""
    try:
        return os.path.join(os.environ["USERPROFILE"], "Desktop")
    except Exception:
        return ""


def write_report(header, exc_text, filename="startup-error.txt", app_dir=None):
    """写一份包含堆栈与环境自检的报告。返回文件路径，失败返回空串。

    同时写**程序目录**和**桌面**两处：前者是常规位置，后者是因为
    用户常常在程序目录里翻不到文件，放桌面一眼就能看见。
    """
    app_dir = app_dir or get_app_dir()

    parts = [
        "=" * 64,
        " ReagentManagementSystem — 故障报告",
        "=" * 64,
        "",
        header,
        "",
        "--- 完整堆栈 ---",
        exc_text,
        "",
    ]
    try:
        parts.append(collect_diagnostics(app_dir))
    except Exception as e:
        parts.append(f"（环境自检本身出错: {type(e).__name__}: {e}）")
    parts += ["", "=" * 64, " 请把本文件发回。", "=" * 64]

    text = "\n".join(parts)

    written = ""
    for base in (app_dir, _desktop_dir(),
                 os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else "."):
        if not base or not os.path.isdir(base):
            continue
        try:
            path = os.path.join(base, filename)
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            written = written or path
        except Exception:
            continue
    return written


def summary_for_dialog(exc_text, app_dir=None):
    """提炼一小段直接放进错误弹窗的文字。

    截到 1200 字以内 —— 用户截一张图就能把所有关键信息给我，
    不必从弹窗里复制多行文本（实测那件事很难做成）。
    """
    lines = []

    # 异常的最后几行最有用，放在最前面
    tail = [ln for ln in (exc_text or "").strip().splitlines() if ln.strip()]
    lines.append("【错误】")
    lines.extend(tail[-4:] if len(tail) > 4 else tail)
    lines.append("")

    try:
        from .. import compat
        lines.append("【Win7 兼容层】")
        if compat.DLL_DIRECTORY_SHIM_ACTIVE:
            lines.append(f"  已启用（{compat.DLL_DIRECTORY_SHIM_REASON}）")
            lines.append(f"  预加载 运行库 {len(compat.RUNTIME_PRELOAD_LOADED)}/{len(compat.RUNTIME_PRELOAD)}"
                         f"，Qt {len(compat.QT_PRELOAD_LOADED)}/{len(compat.QT_PRELOAD_ORDER)}")
            for label, failed in (("运行库", compat.RUNTIME_PRELOAD_FAILED),
                                  ("Qt", compat.QT_PRELOAD_FAILED)):
                if failed:
                    lines.append(f"  {label}失败: " + "、".join(n for n, _ in failed[:6]))
        else:
            lines.append("  未启用（本机原生支持 add_dll_directory）")
    except Exception as e:
        lines.append(f"【Win7 兼容层】读取失败: {e}")
    lines.append("")

    lines.append("【系统】")
    try:
        import platform
        lines.append(f"  {platform.platform()}")
    except Exception:
        pass
    lines.append(f"  {os.environ.get('SystemRoot', '?')}  "
                 f"ucrtbase.dll {'有' if os.path.isfile(os.path.join(os.environ.get('SystemRoot',''), 'System32', 'ucrtbase.dll')) else '缺'}")

    text = "\n".join(lines)
    return text[:1200]


def show_error_box(title, message):
    """弹一个系统错误框。打包的是窗口程序，没有控制台可用。"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
    except Exception:
        pass
