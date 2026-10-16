# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""Win7 环境诊断（控制台程序，独立打包成小 exe）。

用途：主程序在目标机上起不来时，用它把确切原因打出来 —— 是缺系统补丁、
缺运行库，还是随包的 DLL 不兼容。

设计要点：**只依赖 ctypes / os / sys**，不导入 socket / ssl / requests。
一旦引入这些，诊断程序自己就会踩到和主程序一样的坑，什么也报不出来。
按段输出，每段自成一体，前面失败不影响后面。
"""

import os
import platform
import struct
import sys

# ⚠️ ctypes 必须**可失败地**导入。
#
# 这个诊断程序最该跑的场景，恰恰是"机器缺 KB2533623"——而那种机器上
# python38.dll 加载不了任何扩展模块，`import ctypes` 自己就会抛
# ImportError: DLL load failed while importing _ctypes。
# 写成顶层硬导入的话，诊断程序会在最需要它的机器上直接崩掉。
try:
    import ctypes
    import ctypes.wintypes as wt
    CTYPES_ERROR = ""
except BaseException as _e:          # ImportError，也可能是更底层的失败
    ctypes = None
    wt = None
    CTYPES_ERROR = f"{type(_e).__name__}: {_e}"

_LOG_PATH = ""


class _Tee:
    """把输出同时写到控制台和文件。

    为什么需要：双击运行时，程序一结束控制台窗口就关了，输出全没了。
    落一份文件最稳妥 —— 哪怕窗口被关掉，也还能把文件发回来。
    """

    def __init__(self, *streams):
        self._streams = streams

    def write(self, data):
        for s in self._streams:
            try:
                s.write(data)
            except Exception:
                pass

    def flush(self):
        for s in self._streams:
            try:
                s.flush()
            except Exception:
                pass


def line(title=""):
    print("=" * 62)
    if title:
        print(title)
        print("=" * 62)


def section(title):
    print()
    line(title)


# ---------------------------------------------------------------- 基本信息
def report_system():
    section("1. 系统信息")
    try:
        print(f"  Python      : {sys.version.split()[0]}")
        print(f"  解释器位数  : {struct.calcsize('P') * 8} 位")
        print(f"  操作系统    : {platform.platform()}")
        print(f"  版本号      : {platform.version()}")
        print(f"  机器        : {platform.machine()}")
        print(f"  APPDATA     : {os.environ.get('APPDATA', '(无)')}")

        # 是否 64 位系统
        try:
            is64 = ctypes.windll.kernel32.IsWow64Process(
                ctypes.windll.kernel32.GetCurrentProcess(), ctypes.byref(wt.BOOL()))
            print("  （进程运行在 64 位系统上）" if is64 else "  （32 位系统）")
        except Exception:
            pass
    except Exception as e:
        print(f"  读取失败: {e}")


# ------------------------------------------------------- 系统 DLL 是否存在
CRITICAL_SYSTEM_DLLS = [
    ("ucrtbase.dll",                    "KB2999226 通用 C 运行库"),
    ("api-ms-win-crt-runtime-l1-1-0.dll", "UCRT 运行时 API 集"),
    ("api-ms-win-crt-string-l1-1-0.dll",  "UCRT 字符串 API 集"),
    ("api-ms-win-crt-stdio-l1-1-0.dll",   "UCRT 标准 IO API 集"),
    ("api-ms-win-crt-heap-l1-1-0.dll",    "UCRT 堆 API 集"),
    ("api-ms-win-crt-convert-l1-1-0.dll", "UCRT 转换 API 集"),
    ("api-ms-win-core-file-l1-2-0.dll",   "Win8+ 文件 API 集"),
    ("api-ms-win-core-synch-l1-2-0.dll",  "Win8+ 同步 API 集"),
    ("api-ms-win-core-processthreads-l1-1-1.dll", "Win8+ 线程 API 集"),
    ("api-ms-win-core-localization-l1-2-0.dll",   "Win8+ 本地化 API 集"),
    ("vcruntime140.dll",                "VC++ 2015-2022 运行库"),
    ("msvcp140.dll",                    "VC++ C++ 标准库"),
    ("WS2_32.dll",                      "Winsock（socket 依赖）"),
    ("IPHLPAPI.DLL",                    "IP 助手 API（socket 依赖）"),
]

WINDIR = os.environ.get("SystemRoot", r"C:\Windows")
SYS32 = os.path.join(WINDIR, "System32")
SYSWOW = os.path.join(WINDIR, "SysWOW64")


def report_system_dlls():
    section("2. 系统目录里的关键 DLL")
    print(f"  查找目录: {SYS32}")
    print()
    missing = []
    for name, why in CRITICAL_SYSTEM_DLLS:
        path = os.path.join(SYS32, name)
        exists = os.path.isfile(path)
        if not exists:
            missing.append((name, why))
        mark = "有  " if exists else "缺失"
        print(f"  [{mark}] {name:<45} {why}")

    # 这几项单独实测「能不能加载」，比只看文件在不在更有说服力 ——
    # 文件在但加载失败，说明它自己还有依赖没满足。
    print()
    print("  实测加载（绝对路径）：")
    for name in ("WS2_32.dll", "IPHLPAPI.DLL", "KERNEL32.dll"):
        path = os.path.join(SYS32, name)
        if not os.path.isfile(path):
            print(f"    [缺失  ] {name}")
            continue
        ok, err = try_load(path)
        print(f"    [{'可加载' if ok else '失败  '}] {name}" + (f"  <- {err}" if err else ""))

    print()
    if missing:
        print(f"  >>> 系统目录里没有 {len(missing)} 个 API 集文件：")
        for name, why in missing:
            print(f"        {name}  ({why})")
        print()
        print("  说明：从 Win8 起，api-ms-win-crt-* 与 api-ms-win-core-* 这类")
        print("        「API 集」由系统加载器虚拟解析，System32 里本来就没有实体文件。")
        print("        Win7 没有这个机制，它要求这些必须是磁盘上的真实文件。")
        print()
        print("  处理方式：")
        print("    · 缺 api-ms-win-crt-*       -> 装 KB2999226（或在系统上运行一次")
        print("                                   包里的 VC_redist.x64.exe）")
        print("    · 缺 api-ms-win-core-*-l1-2-0 等 Win8+ 集")
        print("                                -> Win7 上永远不会存在，属正常现象。")
        print("                                   只有在某个 DLL 加载失败并指向它时才需处理。")
        print("    · 缺 vcruntime140 / msvcp140 -> 装 VC_redist.x64.exe")
    else:
        print("  系统目录关键 DLL 齐全。")
    return missing


# ------------------------------------------- Windows 更新 / 加载器能力
def _has_export(module, proc):
    """kernel32 里是否导出了某个函数。只查导出表，不调用。

    返回 (是否有, 错误说明)。查不到不等于失败 —— 要看调用方怎么处理。
    """
    if ctypes is None:
        return None, "ctypes 不可用"

    try:
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.GetModuleHandleW.restype = ctypes.c_void_p
        k32.GetModuleHandleW.argtypes = [ctypes.c_wchar_p]
        k32.GetProcAddress.restype = ctypes.c_void_p
        k32.GetProcAddress.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    except OSError as e:
        return None, f"kernel32 不可用: {e}"

    handle = k32.GetModuleHandleW(module)
    if not handle:
        return False, f"模块未加载: {module}"

    addr = k32.GetProcAddress(handle, proc.encode("ascii"))
    return bool(addr), ""


def report_loader_capabilities():
    section("3. Windows 更新与加载器能力")

    if CTYPES_ERROR:
        # 这一条本身就是决定性证据，不必再去查导出表。
        # python38.dll 加载扩展模块时用 LoadLibraryExW 带的那组搜索标志
        # 是 KB2533623 引入的；缺了它，连 _ctypes 都加载不了。
        print("  >>> ctypes 加载失败，这就是根因。")
        print()
        print(f"      {CTYPES_ERROR}")
        print()
        print("      本机缺 KB2533623（或其替代更新 KB3063858）。")
        print("      Python 3.8 在 Win7 上没有这个更新就加载不了任何扩展模块，")
        print("      本程序（以及本诊断工具的大部分检查）都无法工作。")
        print()
        print("      注意：下面第 4/5/6 节的 DLL 加载测试**全部无效** ——")
        print("      它们都得靠 ctypes，测出来的是 ctypes 故障，不是那些 DLL 的问题。")
        print()
        print("      处理：装 KB3063858 后**重启**。仓库 Dependencies\\ 下带了这个 .msu。")
        print()
        return ["KB2533623"]

    print("  KB2533623 提供的一组 DLL 搜索 API：")
    kb2533623 = [
        ("AddDllDirectory",
         "Python 的 os.add_dll_directory 直接调它"),
        ("SetDefaultDllDirectories",
         "Python 3.8 加载扩展模块时会用"),
        ("RemoveDllDirectory", "配套的清理函数"),
    ]
    missing_kb = []
    for proc, why in kb2533623:
        has, _ = _has_export("kernel32.dll", proc)
        if has is None:
            print(f"    [未知  ] {proc}")
            continue
        if not has:
            missing_kb.append(proc)
        print(f"    [{'有    ' if has else '缺失  '}] {proc:<26} {why}")

    print()
    if missing_kb:
        print("  >>> 本机没有装 KB2533623。**这是本程序跑不起来的最主要原因。**")
        print()
        print("      没有它，CPython 3.8 加载不了任何扩展模块（dynload_win.c 用")
        print("      LoadLibraryExW 带 LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR 加载 .pyd，")
        print("      那组标志正是这个更新引入的）。表现是：")
        print()
        print("          ImportError: DLL load failed while importing _ctypes: 参数错误")
        print()
        print("      app/compat.py 救不了 —— 它本身就要用 ctypes。")
        print()
        print("      处理：装 KB2533623（已被微软下架），或它的替代更新 KB3063858，")
        print("            装完**重启**才生效。仓库里的 Dependencies\\ 下带了这个 .msu，")
        print("            安装包也会自动装。")
        print()
        print("      下载时认准 Windows6.1（6.0 是 Vista 的，装不上）：")
        print("      https://www.catalog.update.microsoft.com/Search.aspx?q=KB2533623")
    else:
        print("  KB2533623 已安装（或系统本身自带这套 API）。")

    # UCRT 单独说明一次：这是 Win7 上最常缺的一环，但本程序自带了一份。
    print()
    sys_ucrt = os.path.isfile(os.path.join(SYS32, "ucrtbase.dll"))
    app_ucrt = os.path.isfile(os.path.join(app_dir(), "ucrtbase.dll"))
    print(f"  通用 C 运行库 (UCRT)：")
    print(f"    系统目录 System32\\ucrtbase.dll : {'有' if sys_ucrt else '没有'}"
          f"{'' if sys_ucrt else '   （KB2999226，系统级缺失）'}")
    print(f"    程序目录 ucrtbase.dll          : {'有' if app_ucrt else '没有'}"
          f"{'' if app_ucrt else '   （随包副本也缺失，才是真问题）'}")
    if not sys_ucrt and app_ucrt:
        print("      → 系统级没有，但程序自带了一份，可以跑。")
        print("        要补系统级：装包里的 VC_redist.x64.exe，或 KB2999226。")

    return missing_kb


# ------------------------------------------------- 随程序的 DLL 能否加载
MAIN_EXE = "ReagentManagementSystem.exe"

# 主程序目录。因为诊断程序自己也是打包过的，不能再假设「自己所在目录就是
# 主程序目录」—— 那样检查的会是诊断程序自己的 DLL，得出错误结论。
APP_DIR = ""


def find_app_dir():
    """定位主程序目录。

    依次尝试：命令行参数 -> 自己所在目录 -> 自己的子目录 -> 上一级及其子目录。
    实在找不到就退回自己所在目录，至少还能报出系统信息。
    """
    if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
        return os.path.abspath(sys.argv[1])

    frozen = getattr(sys, "frozen", False)
    here = (os.path.dirname(sys.executable) if frozen
            else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    subnames = ("ReagentManagementSystem", "ReagentManagementSystem_x64")
    candidates = []

    # 从自己所在目录往上找两级，每级都看看自己和自己的子目录。
    # 用户把诊断包放哪儿都能找到主程序，不必要求特定位置。
    base = here
    for _ in range(3):
        candidates.append(base)
        for sub in subnames:
            candidates.append(os.path.join(base, sub))
        try:
            for entry in os.listdir(base):
                full = os.path.join(base, entry)
                if os.path.isdir(full):
                    candidates.append(full)
                    for sub in subnames:
                        candidates.append(os.path.join(full, sub))
        except OSError:
            pass
        parent = os.path.dirname(base)
        if parent == base:
            break
        base = parent

    for path in candidates:
        if os.path.isfile(os.path.join(path, MAIN_EXE)):
            return path
    return here


def app_dir():
    return APP_DIR or os.path.dirname(os.path.abspath(__file__))


def try_load(path):
    """尝试加载 DLL，返回 (是否成功, 错误描述)。

    ctypes 不可用时返回的第三条信息是**"测不了"而不是"坏了"** ——
    这两者必须区分开，否则报告里每一行都显示失败，看着像所有 DLL 都损坏，
    实际测的是 ctypes 本身。
    """
    if ctypes is None:
        return False, f"<无法测试：ctypes 不可用>"
    try:
        ctypes.WinDLL(path)
        return True, ""
    except OSError as e:
        return False, f"WinError={getattr(e, 'winerror', '?')} {e.strerror or e}"
    except BaseException as e:
        return False, f"{type(e).__name__}: {e}"


def report_bundled_dlls():
    section("4. 随程序的 DLL 能否加载")
    here = app_dir()
    print(f"  程序目录: {here}")

    # 优先检查这几个 —— 它们是本次故障的核心嫌疑
    important = [
        "ucrtbase.dll",
        "api-ms-win-crt-runtime-l1-1-0.dll",
        "api-ms-win-crt-string-l1-1-0.dll",
        "VCRUNTIME140.dll",
        "MSVCP140.dll",
        "python38.dll",
        "_socket.pyd",
        "_ssl.pyd",
    ]

    print()
    failed = []
    for name in important:
        path = os.path.join(here, name)
        if not os.path.isfile(path):
            print(f"  [不存在] {name}")
            continue
        ok, err = try_load(path)
        if ok:
            print(f"  [可加载] {name}")
        else:
            print(f"  [失败  ] {name}  <- {err}")
            failed.append((name, err))

    return failed


# --------------------------------------------------------- 逐项导入自检
def report_imports():
    section("5. 关键模块导入自检")
    print("  依次尝试导入，第一次失败不影响后面。")
    print()
    # 只列本诊断程序自带的标准库。sqlite3 / json 在主程序里有打包，
    # 但本诊断程序没有，列出来只会制造噪声。
    modules = [
        "ctypes", "csv", "hashlib", "socket", "ssl",
        "zlib", "select", "encodings.idna",
    ]
    first_failure = None
    for name in modules:
        try:
            __import__(name)
            print(f"  [成功] import {name}")
        except Exception as e:
            print(f"  [失败] import {name}  <- {type(e).__name__}: {e}")
            if first_failure is None:
                first_failure = (name, e)
    return first_failure


def report_qt():
    """单独检查 QtWebEngine 的加载能力。

    刻意**不去 import PyQt5** —— 诊断程序自己不带 Qt，import 必然失败，
    那样报出来的结果是误导。改成直接尝试加载它依赖的原生 DLL，
    这才是 Win7 上真正会出问题的环节。
    """
    section("6. QtWebEngine 原生 DLL 加载自检")
    here = app_dir()

    qt_dlls = [
        "Qt5Core.dll",
        "Qt5Gui.dll",
        "Qt5Network.dll",
        "Qt5Qml.dll",
        "Qt5Quick.dll",
        "Qt5WebChannel.dll",
        "Qt5Positioning.dll",
        "Qt5WebEngine.dll",
        "Qt5WebEngineWidgets.dll",
        "Qt5WebEngineCore.dll",
        "d3dcompiler_47.dll",
        "libEGL.dll",
        "libGLESv2.dll",
        "opengl32sw.dll",
    ]

    failed = []
    for name in qt_dlls:
        path = os.path.join(here, name)
        if not os.path.isfile(path):
            # PyInstaller 也可能把它们放在别处，不算致命
            print(f"  [未找到] {name}")
            continue
        ok, err = try_load(path)
        if ok:
            print(f"  [可加载] {name}")
        else:
            print(f"  [失败  ] {name}  <- {err}")
            failed.append((name, err))

    # 辅助进程是 onefile 布局下最容易出问题的一环，这里顺便看一眼位置
    helper = os.path.join(here, "PyQt5", "Qt5", "bin", "QtWebEngineProcess.exe")
    print()
    print(f"  辅助进程: {helper}")
    print(f"            {'存在' if os.path.isfile(helper) else '未找到'}")

    return failed


def _harden_stdout():
    """让控制台编码问题不至于把诊断程序本身搞崩。

    Windows 7 上 cmd 的代码页处理很脆弱（尤其切到 65001 之后）。
    如果某个字符在控制台代码页里表示不了，默认会抛 UnicodeEncodeError，
    那诊断程序自己就挂了 —— 比什么都不报还糟。这里改成用占位符代替。
    文件那份日志是 UTF-8，不受影响。
    """
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass


def _start_log():
    """把 stdout 分一路到文件。返回日志路径（失败则返回空串）。

    文件名刻意用纯 ASCII —— 目标机可能是英文 Windows，中文文件名
    在命令行环境下会变成乱码，用户根本找不到文件。
    """
    for base in (app_dir(), os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else "."):
        try:
            path = os.path.join(base, "win7diag-result.txt")
            handle = open(path, "w", encoding="utf-8")
            sys.stdout = _Tee(sys.__stdout__, handle)
            return path
        except Exception:
            continue
    return ""


def _pause():
    """等用户看完再关窗口。

    这里**不加 isatty 判断** —— 之前就是因为那个判断没成立，
    窗口一闪而过什么都没留下。除非显式设了环境变量关闭暂停，
    否则一律停住。
    """
    if os.environ.get("DIAG_NO_PAUSE"):
        return
    try:
        print()
        input("按回车键关闭本窗口…")
    except Exception:
        # 没有可用的 stdin（例如被当作窗口程序启动）时只能作罢，
        # 但结果文件已经写好了，不影响排查
        print()
        print("（无法等待按键，请查看上面提到的结果文件）")


def main():
    global _LOG_PATH, APP_DIR
    APP_DIR = find_app_dir()
    _LOG_PATH = _start_log()

    # 开机横幅用纯英文：目标机可能是英文 Windows，中文在控制台上会变乱码，
    # 用户就看不到这个程序到底跑没跑、结果文件在哪。详细内容仍然是中文，
    # 写在 UTF-8 的结果文件里，那部分由开发者阅读。
    print("=" * 62)
    print(" ReagentManagementSystem - Windows 7 Diagnostic")
    print("=" * 62)
    print()
    print(f"Target app dir : {APP_DIR}")
    if not os.path.isfile(os.path.join(APP_DIR, MAIN_EXE)):
        print(f"  !! {MAIN_EXE} not found there.")
        print("     Please put this diagnostic next to the app folder and rerun.")
    print()
    if _LOG_PATH:
        print("Full report saved to:")
        print(f"  {_LOG_PATH}")
        print()
        print(">>> PLEASE SEND THAT FILE BACK <<<")
    else:
        print("!! Could not write the report file. Please copy this window instead.")
    print()
    print("Running checks ...")
    print()

    report_system()
    missing_sys = report_system_dlls()
    missing_kb = report_loader_capabilities()

    if CTYPES_ERROR:
        # 缺 KB2533623 时这两个检查只会输出一屏"失败"，而失败原因全都是同一个
        # （ctypes 测不了），看着像所有 DLL 都损坏了。跳过，别淹没真正的结论。
        section("4. 随程序的 DLL 能否加载 —— 已跳过")
        print("  ctypes 不可用，这些测试无法进行。原因见第 3 节。")
        section("6. QtWebEngine 原生 DLL 加载自检 —— 已跳过")
        print("  ctypes 不可用，这些测试无法进行。原因见第 3 节。")
        failed_bundled, failed_qt = [], []
    else:
        failed_bundled = report_bundled_dlls()
        failed_qt = report_qt()

    # 逐模块导入自检照跑 —— 即使 ctypes 挂了它也有价值：
    # "json/csv/zlib 成功但 ctypes/_socket/_sqlite3 全失败"这个pattern本身就是铁证。
    first_fail = report_imports()

    section("结论")

    # 这里只把「真问题」列出来。api-ms-win-* 在 System32 里不存在属于
    # Win8+ 的正常现象，不当成故障报，否则会误导排查方向。
    real_problems = []

    if first_fail:
        name, err = first_fail
        real_problems.append(f"模块 `{name}` 导入失败：{type(err).__name__}: {err}")

    for name, err in failed_bundled:
        real_problems.append(f"随程序的 {name} 加载失败：{err}")

    for name, err in failed_qt:
        real_problems.append(f"Qt 的 {name} 加载失败：{err}")

    if real_problems:
        print("  发现以下问题：")
        for item in real_problems:
            print(f"    · {item}")
        print()
        print("  请把本窗口全部内容回传。")
    else:
        print("  环境自检未发现故障。")
        print("  若主程序仍然起不来，请把主程序的完整报错一并回传。")

    print()
    if CTYPES_ERROR:
        # 缺 KB2533623 时，上面基于 ctypes 的检查结果全是噪声。必须说清楚，
        # 否则用户会照着"Qt5Core.dll 加载失败"去查，方向完全错。
        print("  ⚠️ 本机缺 KB2533623，上面第 4/5/6 节的加载测试结果全部无效")
        print("     （它们依赖 ctypes，而 ctypes 正是加载不了的那个模块）。")
        print("     装 KB3063858 并重启后重跑本诊断，结果才有意义。")
    elif missing_kb:
        print("  ⚠️ 本机缺 KB2533623。这是本程序起不来的直接原因，务必先装：")
        print("     没有它，Python 3.8 加载不了任何扩展模块。")
        print("     装 KB3063858 并重启后重跑本诊断。")

    print()
    print("  提示：System32 里没有 api-ms-win-* 文件属正常现象，")
    print("        从 Win8 起这些 API 集由加载器虚拟解析，不落盘。")
    print("        程序目录里带了一份实体副本，Win7 靠它们工作。")

    print()
    print("=" * 62)
    print(" Diagnostic finished.")
    if _LOG_PATH:
        print(f" Report file: {_LOG_PATH}")
        print(" >>> Please send that file back. <<<")
    print("=" * 62)
    try:
        sys.stdout.flush()
    except Exception:
        pass
    _pause()
    return 0


if __name__ == "__main__":
    sys.exit(main())
