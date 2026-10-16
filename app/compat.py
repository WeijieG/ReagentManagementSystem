# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""Windows 7 兼容层。**必须在导入 PyQt5 之前调用 apply_all()。**

它处理的是这个问题：`os.add_dll_directory()` 底层调 Windows API
`AddDllDirectory`，而该 API 在 Win7 上要装了 KB2533623 才有。没装时调用抛
`OSError: [WinError 127]`，而 PyQt5 的 `find_qt()` 只捕获 `AttributeError`，
异常一路冒出去，`import PyQt5` 直接失败。

⚠️ **但别指望它替代 KB2533623。** 没有那个更新时，真正致命的不是
PyQt5 少捕获了一个异常，而是 CPython 3.8 自己加载不了任何扩展模块：
`dynload_win.c` 用 `LoadLibraryExW` 带 `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR`
加载 `.pyd`，这组标志正是 KB2533623 引入的。缺了它，`import _ctypes` 就报
`DLL load failed ... 参数错误`（ERROR_INVALID_PARAMETER）。

而本模块要干活就得先 `import ctypes` —— 恰恰是加载不了的那个模块。
所以到那一步会看到：

    预加载运行库 0/18、Qt 0/15

**那不是"预加载策略需要改进"，是这台机器根本缺前提条件，得装更新。**
完整分析见 docs/BUILD.md「KB2533623 是硬性前提」。

本模块真正解决的是另一半问题：`AddDllDirectory` **可用**、
但 `QtCore.pyd` 仍然找不到同目录 `Qt5Core.dll` 的那些情况。做法是无条件用
绝对路径把运行库和 Qt 先装进进程 —— 加载器按**基名**匹配已加载模块，
模块一旦在进程里，谁再要它都直接复用，不走搜索路径。
"""

import os
import sys

# 兼容层是否被启用（供启动报告使用）
DLL_DIRECTORY_SHIM_ACTIVE = False
DLL_DIRECTORY_SHIM_REASON = ""


def _probe_native_support():
    """探测本机的 AddDllDirectory 是否可用。"""
    if not hasattr(os, "add_dll_directory"):
        return False, "Python 未提供 os.add_dll_directory"
    try:
        # 用一个必定存在且无害的目录做探测
        probe = os.path.dirname(os.path.abspath(sys.executable))
        if not os.path.isdir(probe):
            probe = os.environ.get("SystemRoot", "C:\\Windows")
        os.add_dll_directory(probe)
        return True, ""
    except OSError as e:
        winerr = getattr(e, "winerror", None)
        if winerr == 127:
            return False, (
                "系统缺少 KB2533623：kernel32.dll 里没有 AddDllDirectory "
                "（WinError 127）"
            )
        return False, f"调用 AddDllDirectory 失败: {e}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def _install_path_fallback():
    """把 add_dll_directory 换成改写 PATH 的老式实现。

    返回值被调用方丢弃（PyQt5 就没用），所以直接返回 None 即可。
    """
    def _add_dll_directory_fallback(path):
        try:
            path = os.path.abspath(path)
            current = os.environ.get("PATH", "")
            if path and path not in current.split(os.pathsep):
                os.environ["PATH"] = path + os.pathsep + current
        except Exception:
            pass
        return None

    os.add_dll_directory = _add_dll_directory_fallback


# 预加载顺序：被依赖的排在前面。两轮加载，第一轮解不开的依赖第二轮补齐。
QT_PRELOAD_ORDER = [
    "Qt5Core.dll",
    "Qt5Gui.dll",
    "Qt5Network.dll",
    "Qt5Qml.dll",
    "Qt5QmlModels.dll",
    "Qt5Quick.dll",
    "Qt5Positioning.dll",
    "Qt5WebChannel.dll",
    "Qt5WebSockets.dll",
    "Qt5PrintSupport.dll",
    "Qt5Svg.dll",
    "Qt5Widgets.dll",
    "Qt5WebEngine.dll",
    "Qt5WebEngineCore.dll",
    "Qt5WebEngineWidgets.dll",
]

# 运行库同样需要预加载。`_socket.pyd` 之类扩展模块失败，根因和 Qt 是同一个：
# 引导器按绝对路径加载了 python38.dll 所以解释器能起来，但之后 Python 需要
# 自己把 exe 所在目录注册进 DLL 搜索路径 —— 这一步在缺 KB2533623 的机器上失效，
# 于是所有扩展模块都找不到同目录下的依赖。
RUNTIME_PRELOAD = [
    "ucrtbase.dll",
    "api-ms-win-crt-runtime-l1-1-0.dll",
    "api-ms-win-crt-string-l1-1-0.dll",
    "api-ms-win-crt-stdio-l1-1-0.dll",
    "api-ms-win-crt-heap-l1-1-0.dll",
    "api-ms-win-crt-convert-l1-1-0.dll",
    "api-ms-win-crt-filesystem-l1-1-0.dll",
    "api-ms-win-crt-math-l1-1-0.dll",
    "api-ms-win-crt-time-l1-1-0.dll",
    "api-ms-win-crt-locale-l1-1-0.dll",
    "api-ms-win-crt-conio-l1-1-0.dll",
    "api-ms-win-crt-environment-l1-1-0.dll",
    "api-ms-win-crt-process-l1-1-0.dll",
    "api-ms-win-crt-utility-l1-1-0.dll",
    "VCRUNTIME140.dll",
    "VCRUNTIME140_1.dll",
    "MSVCP140.dll",
    "MSVCP140_1.dll",
]

QT_PRELOAD_LOADED = []
QT_PRELOAD_FAILED = []
RUNTIME_PRELOAD_LOADED = []
RUNTIME_PRELOAD_FAILED = []


def _search_dirs():
    """预加载时到哪里找 DLL，按优先级排列。"""
    dirs = []
    if getattr(sys, "frozen", False):
        # 打包后所有 DLL 都在 exe 旁边
        dirs.append(os.path.dirname(sys.executable))
    else:
        # 开发时 Qt 在 site-packages/PyQt5/Qt5/bin。
        # 用 find_spec 而不是 import —— 导入 PyQt5 正是要绕开的动作。
        try:
            import importlib.util
            spec = importlib.util.find_spec("PyQt5")
            if spec and spec.origin:
                pyqt_dir = os.path.dirname(spec.origin)
                dirs.append(os.path.join(pyqt_dir, "Qt5", "bin"))
                dirs.append(os.path.join(pyqt_dir, "Qt", "bin"))
        except Exception:
            pass
        dirs.append(os.path.dirname(os.path.abspath(__file__)))
    return [d for d in dirs if d and os.path.isdir(d)]


def _preload(names, loaded):
    """按名字逐个用绝对路径加载，返回确实加载失败的 [(名字, 错误)]。

    跑两轮：第一轮可能因为依赖尚未就绪而失败，第二轮补齐顺序。
    """
    # ctypes 本身也可能加载不了（_ctypes.pyd 依赖 libffi-7.dll）。
    # 那种情况下没法做预加载 —— 但绝不能让兼容层把启动流程带崩，
    # 因为主程序也许还能在别的前提下跑起来。
    try:
        import ctypes
    except Exception as e:
        return [("<ctypes 不可用>", f"{type(e).__name__}: {e}")]

    dirs = _search_dirs()
    errors = {}

    for _ in range(2):
        for name in names:
            if name in loaded:
                continue
            for base_dir in dirs:
                path = os.path.join(base_dir, name)
                if not os.path.isfile(path):
                    continue
                try:
                    ctypes.WinDLL(path)
                    loaded.append(name)
                    errors.pop(name, None)
                except OSError as e:
                    errors[name] = str(e)
                break

    return [(n, errors[n]) for n in names if n in errors]


def preload_qt_dlls():
    """用绝对路径把 Qt 与运行库预先加载进来。

    为什么需要：`os.add_dll_directory` 不可用时，QtCore.pyd 找不到
    Qt5Core.dll，报 "DLL load failed while importing QtCore"。

    改写 PATH 解决不了 —— Python 3.8 在 Windows 上会调用
    SetDefaultDllDirectories，之后 DLL 搜索就忽略 PATH 了。

    但 Windows 加载模块时是**按基名**匹配的：只要 Qt5Core.dll 已经在
    进程里，后面 QtCore.pyd 再要它就直接复用，根本不走搜索路径。
    所以这里用绝对路径先把它装进去。
    """
    # 先运行库再 Qt —— 后者依赖前者
    RUNTIME_PRELOAD_FAILED[:] = _preload(RUNTIME_PRELOAD, RUNTIME_PRELOAD_LOADED)
    QT_PRELOAD_FAILED[:] = _preload(QT_PRELOAD_ORDER, QT_PRELOAD_LOADED)


def apply_all():
    """应用所有兼容处理。必须在 `import PyQt5` 之前调用。

    整体包了保护：兼容层出问题不应该把启动流程一起带崩 ——
    没有它程序在正常机器上本来就能跑。
    """
    global DLL_DIRECTORY_SHIM_ACTIVE, DLL_DIRECTORY_SHIM_REASON

    try:
        ok, reason = _probe_native_support()
        if not ok:
            _install_path_fallback()
            DLL_DIRECTORY_SHIM_ACTIVE = True
            DLL_DIRECTORY_SHIM_REASON = reason

        # ⚠️ 预加载**无条件执行**，不要改回「探测通过就 return」。
        #
        # 原来的写法只在探测失败时才预加载。结果在一台干净机器上出了
        # "DLL load failed while importing QtCore"，而 PE 导入表全量扫下来
        # （tools/pe_imports.py，203 个 PE 文件）**依赖一个不少** ——
        # 文件都在，坏的是加载器的搜索路径。
        #
        # 而探测成功只说明 AddDllDirectory 这个 API 存在，并不能保证之后
        # QtCore.pyd 找得到 Qt5Core.dll。与其去猜哪台机器的搜索路径有问题，
        # 不如一律用绝对路径把依赖先装进进程：加载器是按**基名**匹配的，
        # 模块一旦在进程里，后面谁再要它都是直接复用，根本不走搜索路径。
        preload_qt_dlls()

        if DLL_DIRECTORY_SHIM_ACTIVE:
            print(f"[compat] os.add_dll_directory 不可用（{reason}）")
        print(f"[compat] 预加载 运行库 {len(RUNTIME_PRELOAD_LOADED)}/{len(RUNTIME_PRELOAD)}、"
              f"Qt {len(QT_PRELOAD_LOADED)}/{len(QT_PRELOAD_ORDER)}")
        for label, failed in (("运行库", RUNTIME_PRELOAD_FAILED),
                              ("Qt", QT_PRELOAD_FAILED)):
            if failed:
                print(f"[compat] {label}预加载失败: "
                      + "、".join(n for n, _ in failed[:8]))
    except Exception as e:
        print(f"[compat] 兼容层自身出错，已跳过: {type(e).__name__}: {e}")


def preload_summary():
    """一行摘要，供启动失败报告在最坏情况下也能拿到关键信息。"""
    return (f"os.add_dll_directory: "
            f"{'兼容层（' + DLL_DIRECTORY_SHIM_REASON + '）' if DLL_DIRECTORY_SHIM_ACTIVE else '原生可用'}；"
            f"预加载 运行库 {len(RUNTIME_PRELOAD_LOADED)}/{len(RUNTIME_PRELOAD)}、"
            f"Qt {len(QT_PRELOAD_LOADED)}/{len(QT_PRELOAD_ORDER)}"
            + (f"；失败: {'、'.join(n for n, _ in (RUNTIME_PRELOAD_FAILED + QT_PRELOAD_FAILED)[:8])}"
               if (RUNTIME_PRELOAD_FAILED or QT_PRELOAD_FAILED) else ""))


def describe():
    """给启动报告用的说明。

    注意：这里报告的是**探测时**的结果。探测之后 os.add_dll_directory
    可能已经被替换掉了，再去调用它验证只会得到兼容层的假象，
    所以结论以这里记录的为准。
    """
    lines = ["--- DLL 预加载与 add_dll_directory ---"]

    if DLL_DIRECTORY_SHIM_ACTIVE:
        lines.append(f"  [兼容层] os.add_dll_directory 不可用：{DLL_DIRECTORY_SHIM_REASON}")
        lines.append("           建议装 KB2533623（认准 Windows6.1，6.0 是 Vista 的）：")
        lines.append("           https://www.catalog.update.microsoft.com/Search.aspx?q=KB2533623")
    else:
        lines.append("  [原生可用] 本机支持 AddDllDirectory（已装 KB2533623 或系统为 Win8+）")

    # 预加载现在是无条件执行的，所以不管上面走哪条分支都要报 —— 判断
    # "QtCore 为什么起不来"时，这行数字是最关键的线索。
    lines.append(f"  预加载运行库 {len(RUNTIME_PRELOAD_LOADED)}/{len(RUNTIME_PRELOAD)}、"
                 f"Qt {len(QT_PRELOAD_LOADED)}/{len(QT_PRELOAD_ORDER)}")

    for label, failed in (("运行库", RUNTIME_PRELOAD_FAILED),
                          ("Qt", QT_PRELOAD_FAILED)):
        if failed:
            lines.append(f"  ！{label}预加载失败 {len(failed)} 个：")
            for name, err in failed[:8]:
                lines.append(f"      {name}  <- {err}")

    return "\n".join(lines)
