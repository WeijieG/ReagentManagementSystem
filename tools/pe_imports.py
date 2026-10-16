# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""列出 PE 文件（exe/dll/pyd）的导入表 —— 它依赖哪些 DLL。

用途：目标机上报 "DLL load failed while importing XXX" 时，光看错误信息
是查不出根因的，得知道那个模块到底要谁。这个工具把那层揭开：

    python tools/pe_imports.py Dist/ReagentManagementSystem/PyQt5/QtCore.pyd
    python tools/pe_imports.py --check-dir Dist/ReagentManagementSystem Qt5Core.dll

`--check-dir` 会把每个依赖按「产物里有 / 只能靠系统」分类，一眼看出打包漏了谁。

纯标准库实现，不依赖 pefile。**故意不用 ctypes 去 LoadLibrary 探测** ——
要查的就是"加载失败"，用加载去验证加载是循环论证。
"""

import argparse
import os
import struct
import sys

# Win7 SP1 起系统自带的 DLL，不必随包。缺了这些才是系统问题。
KNOWN_SYSTEM_DLLS = {
    "kernel32.dll", "user32.dll", "gdi32.dll", "advapi32.dll", "shell32.dll",
    "ole32.dll", "oleaut32.dll", "shlwapi.dll", "version.dll", "ws2_32.dll",
    "iphlpapi.dll", "dnsapi.dll", "winmm.dll", "imm32.dll", "netapi32.dll",
    "userenv.dll", "psapi.dll", "rpcrt4.dll", "comdlg32.dll", "comctl32.dll",
    "uxtheme.dll", "dwmapi.dll", "secur32.dll", "crypt32.dll", "bcrypt.dll",
    "ncrypt.dll", "d3d9.dll", "d3d11.dll", "dxgi.dll", "d2d1.dll",
    "dwrite.dll", "mpr.dll", "mswsock.dll", "winspool.drv", "opengl32.dll",
    "glu32.dll", "msimg32.dll", "wtsapi32.dll", "powrprof.dll", "setupapi.dll",
    "cfgmgr32.dll", "usp10.dll", "wininet.dll", "normaliz.dll", "authz.dll",
    "propsys.dll", "dbghelp.dll", "hid.dll", "imagehlp.dll", "wldap32.dll",
    "oleacc.dll", "pdh.dll", "sensorsapi.dll", "wevtapi.dll", "wer.dll",
    # 下面这些是 Qt5WebEngineCore（Chromium）要的，Win7 SP1 系统目录里都有。
    # 漏了会把它们误报成"缺失"，反而把排查方向带偏。
    "winhttp.dll", "urlmon.dll", "dxva2.dll", "dhcpcsvc.dll",
    "dhcpcsvc6.dll", "bluetoothapis.dll", "msvcrt.dll", "msvcp_win.dll",
    "ucrtbase.dll", "combase.dll", "cryptbase.dll", "dnsapi.dll",
    "ncrypt.dll", "bcryptprimitives.dll", "rpcns4.dll", "winusb.dll",
}


def _looks_like_dll_name(name):
    """判断字符串像不像一个 DLL 文件名。

    用于过滤畸形导入表读出的二进制垃圾：真实名字只有可打印 ASCII，
    且以 .dll/.drv/.exe/.ocx 结尾。
    """
    if not name or len(name) > 64:
        return False
    if not name.isascii() or any(c in name for c in "\r\n\t"):
        return False
    if not all(c.isalnum() or c in "._-+ " for c in name):
        return False
    return name.lower().endswith((".dll", ".drv", ".exe", ".ocx"))


def _read(path):
    with open(path, "rb") as f:
        return f.read()


def parse_imports(path):
    """返回 {'machine':..., 'imports': [...], 'delay': [...]}；不是 PE 返回 None。"""
    data = _read(path)
    if len(data) < 0x40 or data[:2] != b"MZ":
        return None

    e_lfanew = struct.unpack_from("<I", data, 0x3C)[0]
    if data[e_lfanew:e_lfanew + 4] != b"PE\0\0":
        return None

    coff = e_lfanew + 4
    (machine, nsec, _ts, _sym, _nsym, opt_size, _chars) = struct.unpack_from(
        "<HHIIIHH", data, coff)
    opt = coff + 20

    magic = struct.unpack_from("<H", data, opt)[0]
    if magic == 0x20B:       # PE32+
        dd_off = opt + 112
    elif magic == 0x10B:     # PE32
        dd_off = opt + 96
    else:
        return None

    sec_off = opt + opt_size
    sections = []
    for i in range(nsec):
        o = sec_off + i * 40
        vsize, vaddr, rawsize, rawptr = struct.unpack_from("<IIII", data, o + 8)
        sections.append((vaddr, vsize, rawptr, rawsize))

    def rva_to_off(rva):
        for vaddr, vsize, rawptr, rawsize in sections:
            if vaddr <= rva < vaddr + max(vsize, rawsize):
                return rawptr + (rva - vaddr)
        return None

    def cstr(off):
        if off is None:
            return None
        end = data.find(b"\0", off)
        if end < 0:
            return None
        return data[off:end].decode("latin1")

    def walk(directory_index):
        if dd_off + directory_index * 8 + 8 > len(data):
            return []
        rva, _size = struct.unpack_from("<II", data, dd_off + directory_index * 8)
        if not rva:
            return []
        off = rva_to_off(rva)
        if off is None:
            return []

        names = []
        for i in range(512):            # 上限，防畸形文件死循环
            entry = off + i * 20
            if entry + 20 > len(data):
                break
            desc = struct.unpack_from("<IIIII", data, entry)
            if desc == (0, 0, 0, 0, 0):
                break
            name = cstr(rva_to_off(desc[3]))
            if name and _looks_like_dll_name(name):
                names.append(name)
            # 不像文件名的直接丢掉。Chromium 的 Qt5WebEngineCore.dll 用了
            # 非标准的延迟导入布局，照标准解析会读出一堆二进制垃圾 ——
            # 报出来只会淹没真正有用的信息。
        return names

    return {
        "machine": "x64" if machine == 0x8664 else
                   "x86" if machine == 0x14C else f"0x{machine:04x}",
        "imports": walk(1),      # IMAGE_DIRECTORY_ENTRY_IMPORT
        "delay": walk(13),       # IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT
    }


def parse_import_symbols(path):
    """返回 [(dll, 符号), ...]。序号导入的符号写作 '#N'。

    只走**标准导入表**（目录项 1）。不看延迟导入表（13）：
    * 延迟导入的格式随编译器而变，Chromium 那类非标准布局容易解析出二进制垃圾；
    * 更重要的是，「无法定位程序输入点 XXX」这种启动期错误只可能来自静态导入 ——
      延迟导入是首次调用时才去解析的。
    """
    data = _read(path)
    if len(data) < 0x40 or data[:2] != b"MZ":
        return None

    e_lfanew = struct.unpack_from("<I", data, 0x3C)[0]
    if data[e_lfanew:e_lfanew + 4] != b"PE\0\0":
        return None

    coff = e_lfanew + 4
    (machine, nsec, _ts, _sym, _nsym, opt_size, _chars) = struct.unpack_from(
        "<HHIIIHH", data, coff)
    opt = coff + 20
    magic = struct.unpack_from("<H", data, opt)[0]
    if magic == 0x20B:
        dd_off, thunk_size, thunk_fmt, ordinal_flag = opt + 112, 8, "<Q", 1 << 63
    elif magic == 0x10B:
        dd_off, thunk_size, thunk_fmt, ordinal_flag = opt + 96, 4, "<I", 1 << 31
    else:
        return None

    sec_off = opt + opt_size
    sections = []
    for i in range(nsec):
        o = sec_off + i * 40
        vsize, vaddr, rawsize, rawptr = struct.unpack_from("<IIII", data, o + 8)
        sections.append((vaddr, vsize, rawptr, rawsize))

    def rva_to_off(rva):
        for vaddr, vsize, rawptr, rawsize in sections:
            if vaddr <= rva < vaddr + max(vsize, rawsize):
                return rawptr + (rva - vaddr)
        return None

    def cstr(off):
        if off is None:
            return None
        end = data.find(b"\0", off)
        return data[off:end].decode("latin1") if end >= 0 else None

    imp_rva = struct.unpack_from("<I", data, dd_off + 8)[0]
    if not imp_rva:
        return []
    off = rva_to_off(imp_rva)
    if off is None:
        return []

    out = []
    for i in range(512):
        entry = off + i * 20
        if entry + 20 > len(data):
            break
        desc = struct.unpack_from("<IIIII", data, entry)
        if desc == (0, 0, 0, 0, 0):
            break

        dll = cstr(rva_to_off(desc[3])) or f"<rva 0x{desc[3]:x}>"
        ilt = desc[0] or desc[4]          # OriginalFirstThunk，缺省退回 FirstThunk
        toff = rva_to_off(ilt)
        if toff is None:
            continue

        for j in range(8192):
            p = toff + j * thunk_size
            if p + thunk_size > len(data):
                break
            val = struct.unpack_from(thunk_fmt, data, p)[0]
            if val == 0:
                break
            if val & ordinal_flag:
                out.append((dll, "#%d" % (val & 0xFFFF)))
            else:
                # ⚠️ 别漏了前面那 2 个字节的 Hint。
                # 结构是 IMAGE_IMPORT_BY_NAME { WORD Hint; CHAR Name[]; }，
                # 从 val 直接读会得到 "}RegEnumValueW" 这种带前缀垃圾的名字，
                # 然后所有按名字精确匹配的查询都会静默失败 —— 这个坑踩过一次，
                # 当时据此得出了"没有静态导入该符号"的错误结论。
                off2 = rva_to_off(val)
                if off2 is not None:
                    name = cstr(off2 + 2)
                    if name:
                        out.append((dll, name))
    return out


def collect_dist_files(dist_dir):
    """产物里所有文件的**小写基名** -> 相对路径。加载器是按基名匹配的。"""
    index = {}
    for root, dirs, files in os.walk(dist_dir):
        dirs.sort()
        for name in files:
            index.setdefault(name.lower(), []).append(
                os.path.relpath(os.path.join(root, name), dist_dir).replace("\\", "/"))
    return index


def classify(names, dist_index):
    """把依赖分成 产物里有 / 靠系统 两类。"""
    in_dist, system, unknown = [], [], []
    for n in names:
        key = n.lower()
        if key in dist_index:
            in_dist.append((n, dist_index[key]))
        elif key in KNOWN_SYSTEM_DLLS or key.startswith("api-ms-win-"):
            system.append(n)
        else:
            unknown.append(n)
    return in_dist, system, unknown


def find_symbol(dist_dir, symbol):
    """在目录里搜「谁导入了这个函数」。

    典型用途：Win7 上弹「无法定位程序输入点 XXX 于 KERNEL32.dll」——
    说明某个 DLL 静态导入了 Win8+ 才有的 API。要改就得先知道是谁。
    """
    symbol_lower = symbol.lower()
    hits = []
    scanned = 0

    for root, dirs, files in os.walk(dist_dir):
        dirs.sort()
        for name in sorted(files):
            if os.path.splitext(name)[1].lower() not in (".dll", ".pyd", ".exe"):
                continue
            path = os.path.join(root, name)
            symbols = parse_import_symbols(path)
            if symbols is None:
                continue
            scanned += 1
            for dll, sym in symbols:
                if sym.lower() == symbol_lower:
                    hits.append((os.path.relpath(path, dist_dir).replace("\\", "/"),
                                 dll))

    print(f"\n在 {scanned} 个 PE 文件里搜索 {symbol}\n")
    if not hits:
        print("  没有静态导入这个符号的文件。")
        print("  （如果在 Win7 上仍报「无法定位程序输入点」，可能是延迟导入，")
        print("    或者来自系统目录里被替换过的 DLL。）")
        return 0

    print(f"  找到 {len(hits)} 处：\n")
    for path, dll in hits:
        print(f"    {path}")
        print(f"        <- 从 {dll} 导入")
    print("\n  这就是要处理的文件：它在 Windows 7 上会让整个进程起不来。")
    return 0


def main():
    parser = argparse.ArgumentParser(description="列出 PE 文件的导入依赖")
    parser.add_argument("target", nargs="?", help="exe / dll / pyd 路径")
    parser.add_argument("--find-symbol", metavar="NAME",
                        help="配合 --check-dir，搜哪个文件导入了这个函数")
    parser.add_argument("--check-dir", metavar="DIR",
                        help="产物目录，用来判断依赖是随包还是靠系统")
    parser.add_argument("--legacy", action="store_true",
                        help="只列 Win7 上真正需要留意的（跳过 api-ms-win-* 噪声）")
    args = parser.parse_args()

    if args.find_symbol:
        if not args.check_dir:
            print("--find-symbol 必须配合 --check-dir 使用")
            return 1
        return find_symbol(args.check_dir, args.find_symbol)

    target = args.target
    if not target:
        parser.error("需要给一个 PE 文件路径，或用 --find-symbol 配 --check-dir")
    if not os.path.isfile(target):
        print(f"找不到文件: {target}")
        return 1

    info = parse_imports(target)
    if not info:
        print(f"不是有效的 PE 文件: {target}")
        return 1

    print(f"\n{target}  ({info['machine']})\n")

    for label, names in (("导入", info["imports"]), ("延迟导入", info["delay"])):
        if not names:
            continue
        print(f"--- {label}（{len(names)}）---")
        if args.check_dir:
            dist_index = collect_dist_files(args.check_dir)
            in_dist, system, unknown = classify(names, dist_index)
            for n, where in in_dist:
                print(f"  [产物] {n:<46} {where[0]}")
            for n in system:
                if not args.legacy:
                    print(f"  [系统] {n}")
            for n in unknown:
                print(f"  [未知] {n}   <-- 既不在产物里，也不是已知的 Win7 系统 DLL")
        else:
            for n in names:
                print(f"  {n}")
        print()

    return 0


if __name__ == "__main__":
    sys.exit(main())
