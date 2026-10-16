# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""一键构建：PyInstaller 打包 → 补合规文件 → 重算完整性哈希 → 编译安装包。

    python tools\\build_installer.py                 # 全流程
    python tools\\build_installer.py --skip-build    # 只重打安装包（改过 setup.iss 后）
    python tools\\build_installer.py --no-installer  # 只出 Dist，不编译安装包
    python tools\\build_installer.py --smoke         # 构建后跑一次界面截图自检
    python tools\\build_installer.py --version 2.2.0 # 改版本号并构建

这个脚本存在的意义是把几个**顺序不能错、错了又不容易发现**的步骤串起来：

  1. PyInstaller 会把仓库根目录的 config.ini 原样拷进产物。那份模板里哈希是空的。
  2. 重算哈希必须写进**产物里**那份 config.ini，不是仓库那份。写错地方的话，
     程序和安装包都打得出来，但完整性校验永远失效 —— 静默失败，最难查。
  3. 安装包必须在哈希写完之后编译，否则打进包里的是没有哈希的 config.ini。

版本号的唯一来源是 config.ini 的 [Application] version。
"""

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)

DIST_DIR = os.path.join(REPO, "Dist", "ReagentManagementSystem")
OUTPUT_DIR = os.path.join(REPO, "Output")
SPEC_FILE = "build.spec"
ISS_FILE = "setup.iss"
VERSION_ISS = os.path.join(REPO, "version.iss")
TEMPLATE_CONFIG = os.path.join(REPO, "config.ini")

# 产物自检。
#
# 精确路径的只有这几项 —— 它们在哪个布局下都在同样的位置。
REQUIRED_PATHS = [
    "ReagentManagementSystem.exe",
    "config.ini",
    "ui/index.html",
]

# 其余的按**文件名**在产物里搜。
#
# 不写死路径是有原因的：PyInstaller 的 PyQt5 hook 会把 40 个 Qt5*.dll 摊在
# 产物根目录，却把 QtWebEngineProcess.exe 单独放进 PyQt5\Qt5\bin\，
# 这种布局在不同 PyInstaller 版本间还会变。写死路径的结果是自检误报
# "产物不完整"，然后你去 debug 一个根本不存在的打包问题。
#
# en-US.pak 必须留着：Chromium 找不到语言包回退项时是硬失败，不是警告。
REQUIRED_BY_NAME = [
    "python38.dll",
    "Qt5Core.dll",
    "Qt5WebEngineCore.dll",
    "QtWebEngineProcess.exe",
    "en-US.pak",
]

# GPL v3 要求随二进制分发许可证与源码线索
GPL_FILES = ["LICENSE", "SOURCE_CODE"]

# Windows 7 上 Python 3.8 的硬性前提。缺了它，安装包装到干净 Win7 上，
# 程序会以「DLL load failed while importing _ctypes: 参数错误」告终 ——
# 报错方向完全是错的。详见 setup.iss 开头的说明。
KB_PACKAGE = os.path.join("Dependencies", "Windows6.1-KB3063858-x64.msu")

# Windows 7 上**不存在**、但新工具链会静默静态导入的 API。
#
# 这类问题在开发机上永远复现不了 —— 开发机是 Win10，这些 API 全都在；
# 只有拿到 Win7 上才会以「无法定位程序输入点 XXX 于 KERNEL32.dll 上」炸出来，
# 而那个报错完全看不出跟"打包机装了个新版编译器"有关。
# 实测踩过一次（MSVCP140 14.51 导入 GetSystemTimePreciseAsFileTime），
# 所以在构建期拦一道。
WIN8_ONLY_APIS = (
    "GetSystemTimePreciseAsFileTime",   # Win8
    "SetThreadDescription",             # Win10 1607
    "GetThreadDescription",             # Win10 1607
    "GetDpiForWindow",                  # Win10 1607
    "PathCchCanonicalizeEx",            # Win8
    "PathCchCombineEx",                 # Win8
    "CreateFile2",                      # Win8
    "GetTempPath2W",                    # Win10 1903
    "GetSystemTimeAdjustmentPrecise",   # Win10
)


# ---------------------------------------------------------------- 输出

def step(text):
    print("\n" + "=" * 68)
    print(f"  {text}")
    print("=" * 68)


def ok(text):
    print(f"  [OK]   {text}")


def warn(text):
    print(f"  [警告] {text}")


def fail(text):
    print(f"  [失败] {text}")


class BuildError(Exception):
    """可预期的构建失败，不需要堆栈。"""


# ---------------------------------------------------------------- 配置

def read_version():
    import configparser
    config = configparser.ConfigParser()
    config.read(TEMPLATE_CONFIG, encoding="utf-8")
    version = config.get("Application", "Version", fallback="").strip()
    if not version:
        raise BuildError(f"{TEMPLATE_CONFIG} 里没有 [Application] version")
    if not re.match(r"^\d+(\.\d+)*$", version):
        raise BuildError(f"版本号格式不对: {version!r}（应当形如 2.1.0）")
    return version


def write_version_to_config(version):
    """把版本号写回 config.ini。只替换那一行，保留其余内容与注释。"""
    with open(TEMPLATE_CONFIG, "r", encoding="utf-8") as f:
        text = f.read()

    new_text, count = re.subn(
        r"(?m)^(\s*version\s*=\s*)\S+\s*$", r"\g<1>" + version, text,
    )
    if count != 1:
        raise BuildError(f"在 config.ini 里找到 {count} 处 version= ，预期 1 处")

    with open(TEMPLATE_CONFIG, "w", encoding="utf-8") as f:
        f.write(new_text)
    ok(f"config.ini 版本号已改为 {version}")


def check_version_consistency(version):
    """提醒 server.py 的 CURRENT_VERSION 有没有跟着改。

    这两处不同步的后果很隐蔽：安装包发布出去了，但更新服务端还认为
    「已经是最新」，客户端永远收不到推送。
    """
    server_py = os.path.join(REPO, "server.py")
    if not os.path.isfile(server_py):
        return

    with open(server_py, "r", encoding="utf-8") as f:
        text = f.read()

    match = re.search(r"(?m)^CURRENT_VERSION\s*=\s*['\"]([^'\"]+)['\"]", text)
    if not match:
        return

    server_version = match.group(1)
    if server_version != version:
        warn(f"server.py 的 CURRENT_VERSION 还是 {server_version}，"
             f"与本次构建的 {version} 不一致。")
        warn("  不同步的话更新推送不会生效。改 server.py 开头那一行即可。")
    else:
        ok(f"server.py 的 CURRENT_VERSION 与版本号一致（{version}）")


def check_production_config():
    """确认即将被打进安装包的 config.ini 是"生产形态"。

    最容易犯的错是带着测试隔离配置(DataDirName=...Test)发布 —— 用户装上
    之后看到的是一个空库，会以为数据丢了。
    """
    import configparser
    config = configparser.ConfigParser()
    config.read(TEMPLATE_CONFIG, encoding="utf-8")

    problems = []

    data_dir = config.get("Application", "DataDirName", fallback="").strip()
    if data_dir and data_dir != "ReagentManagementSystem":
        problems.append(
            f"[Application] DataDirName = {data_dir}\n"
            "         这是测试用的数据目录隔离配置，正式包不能带。\n"
            "         删掉这一行（或改成 ReagentManagementSystem）再构建。"
        )

    if not config.get("Update", "Server", fallback="").strip():
        warn("[Update] server 为空，程序里的「检查更新」会直接跳过。")

    if problems:
        for p in problems:
            fail(p)
        raise BuildError("config.ini 不是生产配置，已中止")

    ok("config.ini 是生产配置（无数据目录隔离）")


def check_win7_prerequisite():
    """确认 Win7 必需更新的安装包在不在。

    不在也能构建（setup.iss 里是编译期条件），但装到干净 Win7 上会起不来，
    而报错指向"缺 DLL"而不是"缺系统更新"。所以这里必须吭声。
    """
    path = os.path.join(REPO, KB_PACKAGE)
    if os.path.isfile(path):
        ok(f"Win7 必需更新已就位（{os.path.basename(KB_PACKAGE)}，"
           f"{os.path.getsize(path) / 1024 / 1024:.1f} MB）")
        return

    warn(f"没有 {KB_PACKAGE}")
    print(
        "\n        这是 Windows 7 上 Python 3.8 的硬性前提。缺了它，安装包\n"
        "        装到干净 Win7 上会以「DLL load failed while importing _ctypes」\n"
        "        告终 —— 一个指向完全错误方向的报错。\n"
        "\n"
        "        安装脚本会退化成「装的时候提示用户自己去补」，仍然可用。\n"
        "        把它放回 Dependencies\\ 即可自动带上：\n"
        "        https://download.microsoft.com/download/0/8/E/"
        "08E0386B-F6AF-4651-8D1B-C0A95D2731F0/Windows6.1-KB3063858-x64.msu\n"
    )


# ---------------------------------------------------------------- 环境

def check_environment():
    step("1/7  检查构建环境")

    if sys.version_info[:2] != (3, 8):
        raise BuildError(
            f"当前 Python 是 {sys.version.split()[0]}，必须是 3.8.x。\n"
            "        3.8 是最后一个支持 Windows 7 的版本；用它打包出来的程序\n"
            "        才能在 Win7 上跑。更高的版本打的包在 Win7 上会起不来。"
        )
    ok(f"Python {sys.version.split()[0]}")

    if sys.maxsize <= 2 ** 32:
        raise BuildError("当前是 32 位 Python，本项目只出 64 位版本。")
    ok("64 位解释器")

    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        raise BuildError(
            "没装 PyInstaller。执行：\n"
            "        python -m pip install -r requirements-build.txt"
        )
    ok("PyInstaller 可用")


def find_iscc():
    """定位 Inno Setup 的命令行编译器 ISCC.exe。

    找不到时返回 None —— 只出 Dist 的模式下不需要它。
    """
    candidates = []

    env_path = os.environ.get("ISCC")
    if env_path:
        candidates.append(env_path)

    found = shutil.which("ISCC")
    if found:
        candidates.append(found)

    for base in filter(None, [
        os.environ.get("ProgramFiles(x86)"),
        os.environ.get("ProgramFiles"),
        os.environ.get("LocalAppData"),
    ]):
        for sub in ("Programs\\Inno Setup 6", "Programs\\Inno Setup 5",
                    "Inno Setup 6", "Inno Setup 5"):
            candidates.append(os.path.join(base, sub, "ISCC.exe"))

    # 注册表里的安装位置（winget 装的通常也能在这里找到）
    try:
        import winreg
        for hive, key in (
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1"),
        ):
            try:
                with winreg.OpenKey(hive, key) as handle:
                    location = winreg.QueryValueEx(handle, "InstallLocation")[0]
                candidates.append(os.path.join(location, "ISCC.exe"))
            except OSError:
                continue
    except ImportError:
        pass

    for path in candidates:
        if path and os.path.isfile(path):
            return os.path.abspath(path)
    return None


# 中文语言包。Inno Setup 自带 28 个语言包，**不含简体中文** ——
# 中文是非官方翻译，要单独下载。MessagesFile 指向不存在的文件时 ISCC 是
# 直接编译失败（不是给个警告），所以必须能探测、能退化。
CHINESE_ISL_NAME = "ChineseSimplified.isl"
VENDORED_ISL = os.path.join(REPO, "packaging", CHINESE_ISL_NAME)
CHINESE_ISL_URL = (
    "https://github.com/jrsoftware/issrc/raw/main/Files/Languages/Unofficial/"
    + CHINESE_ISL_NAME
)


def resolve_chinese_messages(iscc_path):
    """返回 Inno 的 MessagesFile 值；没有可用的中文包时返回空串。

    优先用仓库里自带的那份（这样换台机器也能编出中文安装界面），
    其次用 ISCC 安装目录里的，最后退回英文。
    """
    if os.path.isfile(VENDORED_ISL):
        return r"packaging\%s" % CHINESE_ISL_NAME

    if iscc_path:
        machine_copy = os.path.join(
            os.path.dirname(iscc_path), "Languages", CHINESE_ISL_NAME)
        if os.path.isfile(machine_copy):
            return r"compiler:Languages\%s" % CHINESE_ISL_NAME

    return ""


def report_missing_chinese():
    warn("没找到简体中文语言包，本次安装界面会是英文。")
    print(
        "\n        要中文界面，把 " + CHINESE_ISL_NAME + " 放到：\n"
        f"            {os.path.relpath(VENDORED_ISL, REPO)}\n"
        "        下载地址（GitHub 国内可能连不上，用你的方式拿）：\n"
        f"            {CHINESE_ISL_URL}\n"
        "        或者放到 ISCC 安装目录的 Languages\\ 下也行。\n"
        "        放好后重跑本脚本，setup.iss 会自动用上。\n"
    )


# ---------------------------------------------------------------- 构建

def clean():
    step("2/7  清理上一次的产物")
    for path in (DIST_DIR, os.path.join(REPO, "build", "pyi")):
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
            ok(f"已删除 {os.path.relpath(path, REPO)}")
    ok("清理完成")


def run_pyinstaller():
    step("3/7  PyInstaller 打包")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--distpath", "Dist",
        "--workpath", "build/pyi",
        SPEC_FILE,
    ]
    print("  " + " ".join(cmd) + "\n")
    code = subprocess.call(cmd, cwd=REPO)
    if code != 0:
        raise BuildError(f"PyInstaller 退出码 {code}")
    ok("打包完成")


def stage_gpl_files():
    """把 LICENSE / SOURCE_CODE 放进产物目录。

    这样 Dist\\ReagentManagementSystem\\ 就是一个自洽的、可以直接拷走用的
    目录，安装包也只需要一条 Source 规则。
    """
    step("4/7  补 GPL 合规文件与产物自检")

    for name in GPL_FILES:
        src = os.path.join(REPO, name)
        if not os.path.isfile(src):
            warn(f"仓库根目录没有 {name}，已跳过")
            continue
        shutil.copy2(src, os.path.join(DIST_DIR, name))
        ok(f"{name} 已放入产物")

    missing = [p for p in REQUIRED_PATHS
               if not os.path.exists(os.path.join(DIST_DIR, p.replace("/", os.sep)))]

    present_names = set()
    size = 0
    file_count = 0
    for root, _, files in os.walk(DIST_DIR):
        for name in files:
            present_names.add(name.lower())
            size += os.path.getsize(os.path.join(root, name))
            file_count += 1

    missing += [n for n in REQUIRED_BY_NAME if n.lower() not in present_names]

    if missing:
        for m in missing:
            fail(f"产物里缺少 {m}")
        raise BuildError("打包产物不完整，已中止（别去编译安装包）")

    ok(f"产物关键文件齐全（{file_count} 个文件，抽查 "
       f"{len(REQUIRED_PATHS) + len(REQUIRED_BY_NAME)} 项）")
    ok(f"产物大小 {size / 1024 / 1024:.1f} MB")

    verify_win7_api_surface()


def verify_win7_api_surface():
    """扫产物里有没有静态导入 Win8+ 才有的 API。

    为什么值得单独拦一道：这类问题的**报错和根因完全脱节**。
    「无法定位程序输入点 GetSystemTimePreciseAsFileTime 于 KERNEL32.dll 上」
    看起来像系统缺文件，实际是打包机上的编译器太新、把 Win8 的 API 写进了
    导入表。而且 app-local 的副本会盖住系统里那份好的，所以目标机装了正确的
    运行库也救不回来。

    在开发机上**永远复现不了** —— 开发机是 Win10，那些 API 都在。
    """
    sys.path.insert(0, os.path.join(REPO, "tools"))
    import pe_imports

    findings = []
    scanned = 0
    for root, dirs, files in os.walk(DIST_DIR):
        dirs.sort()
        for name in sorted(files):
            if os.path.splitext(name)[1].lower() not in (".dll", ".pyd", ".exe"):
                continue
            path = os.path.join(root, name)
            symbols = pe_imports.parse_import_symbols(path)
            if symbols is None:
                continue
            scanned += 1
            for dll, sym in symbols:
                if sym in WIN8_ONLY_APIS:
                    findings.append(
                        (os.path.relpath(path, DIST_DIR).replace("\\", "/"), dll, sym))

    if findings:
        for path, dll, sym in findings:
            fail(f"{path} 静态导入了 Win8+ 才有的 {sym}（来自 {dll}）")
        raise BuildError(
            "产物里有 Windows 7 上不存在的 API —— 装到 Win7 上会报\n"
            "        「无法定位程序输入点 XXX 于 KERNEL32.dll 上」，程序根本起不来。\n"
            "        最常见的原因是 PyInstaller 从打包机的 System32 抓了过新的\n"
            "        MSVC 运行库。检查 build.spec 里的 MSVC_RUNTIME_PREFIXES 是否\n"
            "        还在生效，以及打包机上装的是哪一版 Visual Studio。"
        )

    ok(f"API 面检查通过：{scanned} 个 PE 文件均未静态导入 Win8+ 的 API")


def regenerate_hashes():
    step("5/7  重算完整性哈希（写入产物里的 config.ini）")
    cmd = [sys.executable, os.path.join(REPO, "tools", "gen_hashes.py"),
           "--dist", DIST_DIR]
    code = subprocess.call(cmd, cwd=REPO)
    if code != 0:
        raise BuildError("gen_hashes.py 执行失败")

    # 复核一遍：哈希真的写进产物了才算数。这一步是防"写错文件"静默失败的。
    import configparser
    config = configparser.ConfigParser()
    config.read(os.path.join(DIST_DIR, "config.ini"), encoding="utf-8")
    exe_hash = config.get("Security", "MainFileHashe_x64", fallback="").strip('"')
    ui_hash = config.get("Security", "UiBundleHashe", fallback="").strip('"')
    if len(exe_hash) != 64 or len(ui_hash) != 64:
        raise BuildError(
            "产物里的 config.ini 哈希为空或不完整，完整性校验不会生效。\n"
            "        检查 gen_hashes.py 的 --config 目标是不是产物里那份。"
        )
    ok(f"exe 哈希 {exe_hash[:16]}… / ui 清单 {ui_hash[:16]}…")


def smoke_test():
    step("5.5/7  界面自检（打包产物真跑一遍）")
    shot = os.path.join(REPO, "build", "installer-smoke.png")
    os.makedirs(os.path.dirname(shot), exist_ok=True)
    if os.path.exists(shot):
        os.remove(shot)

    env = dict(os.environ)
    env["RMS_SMOKE_SHOT"] = shot
    env["RMS_SMOKE_DELAY_MS"] = "15000"

    exe = os.path.join(DIST_DIR, "ReagentManagementSystem.exe")
    try:
        code = subprocess.call([exe], cwd=DIST_DIR, env=env, timeout=180)
    except subprocess.TimeoutExpired:
        raise BuildError("自检超时（180 秒），程序可能卡在启动阶段")

    if code != 0 or not os.path.exists(shot):
        raise BuildError(
            f"自检失败（退出码 {code}，截图{'有' if os.path.exists(shot) else '无'}）。\n"
            "        产物目录和程序目录下会有 startup-error.txt。"
        )
    ok(f"自检通过，截图 {os.path.relpath(shot, REPO)}")


def write_version_iss(version):
    """生成 version.iss，setup.iss 会 #include 它。

    只放版本号 —— 语言包的探测放在 setup.iss 里自己做（ISPP 的 FileExists
    就够用），这样本文件不会因为"这台机器有没有中文包"而变来变去。
    """
    parts = (version.split(".") + ["0", "0", "0"])[:4]
    quad = ".".join(parts)

    lines = [
        "; " + "=" * 71,
        ";  本文件由 tools/build_installer.py 自动生成，请勿手工编辑。",
        ";",
        ";  版本号的唯一来源是 config.ini 的 [Application] version ——",
        ";  改版本请改那里，然后重跑构建脚本。",
        ";",
        ";  之所以单独生成一个文件而不是用 ISCC 的 /D 传参：ISCC 的 /D 会把",
        ";  值当表达式解析，\"2.1.0\" 不是合法表达式，容易踩坑。",
        "; " + "=" * 71,
        "",
        f'#define MyAppVersion "{version}"',
        f'#define MyAppVersionQuad "{quad}"',
        "",
        "; 只放版本号。语言包用哪个由 setup.iss 自己探测 —— 放进这里的话",
        "; 同一个仓库在两台机器上会生成出内容不同的文件。",
    ]

    with open(VERSION_ISS, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    ok(f"version.iss 已生成（版本 {version} / {quad}）")


def compile_installer(iscc_path, version):
    step("7/7  编译安装包")
    print(f"  {iscc_path}\n")
    code = subprocess.call([iscc_path, ISS_FILE], cwd=REPO)
    if code != 0:
        raise BuildError(f"ISCC 退出码 {code}")

    # 找刚生成的安装包（文件名里带版本号，不去猜）
    candidates = [
        os.path.join(OUTPUT_DIR, name)
        for name in os.listdir(OUTPUT_DIR)
        if name.endswith(".exe") and version in name
    ]
    if not candidates:
        raise BuildError(f"编译成功但在 {OUTPUT_DIR} 里找不到含 {version} 的安装包")

    setup = max(candidates, key=os.path.getmtime)
    return setup


# ---------------------------------------------------------------- 主流程

def print_iscc_help():
    fail("没有找到 Inno Setup 的命令行编译器 ISCC.exe。")
    print(
        "\n  装一个即可（任选）：\n"
        "      winget install -e --id JRSoftware.InnoSetup\n"
        "      https://jrsoftware.org/isdl.php\n"
        "\n  装完重跑本脚本；也可以先用 --no-installer 只出 Dist 目录。\n"
        "  装好之后确认 Languages\\ChineseSimplified.isl 在不在 ——\n"
        "  缺了会让安装界面变成英文（脚本会自动退化，不会编译失败）。"
    )


def main():
    parser = argparse.ArgumentParser(
        description="一键构建打包并编译安装包",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", help="改版本号（写回 config.ini）后再构建")
    parser.add_argument("--skip-build", action="store_true",
                        help="跳过 PyInstaller，只用现有 Dist 重打安装包")
    parser.add_argument("--no-installer", action="store_true",
                        help="只生成 Dist，不编译安装包")
    parser.add_argument("--smoke", action="store_true",
                        help="构建后启动一次产物并截图，确认界面真的渲染出来")
    args = parser.parse_args()

    os.chdir(REPO)

    try:
        if args.version:
            write_version_to_config(args.version)
        version = read_version()

        print(f"\n  药剂仓库管理系统 ——构建 {version}")
        print(f"  仓库: {REPO}")

        check_production_config()
        check_win7_prerequisite()
        check_version_consistency(version)

        iscc_path = None
        chinese_messages = ""
        if not args.no_installer:
            iscc_path = find_iscc()
            if not iscc_path:
                print_iscc_help()
                return 1
            ok(f"找到 ISCC: {iscc_path}")

            chinese_messages = resolve_chinese_messages(iscc_path)
            if chinese_messages:
                ok(f"安装界面语言包: {chinese_messages}")
            else:
                report_missing_chinese()

        check_environment()

        if not args.skip_build:
            clean()
            run_pyinstaller()
        else:
            step("3/7  跳过 PyInstaller（--skip-build）")

        stage_gpl_files()
        regenerate_hashes()

        if args.smoke:
            smoke_test()

        if args.no_installer:
            step("7/7  跳过安装包（--no-installer）")
            ok(f"产物目录: {os.path.relpath(DIST_DIR, REPO)}")
            return 0

        write_version_iss(version)
        setup = compile_installer(iscc_path, version)

        step("完成")
        size_mb = os.path.getsize(setup) / 1024 / 1024
        print(f"  安装包   : {setup}")
        print(f"  大小     : {size_mb:.1f} MB")
        digest = hashlib.sha256(open(setup, "rb").read()).hexdigest()
        print(f"  SHA-256  : {digest}")
        print(
            "\n  下一步：\n"
            "    1. 先在 Win7 上装一遍、跑通\n"
            "    2. 发布更新时，把安装包拷进 server.py 同级目录的 updates\\，\n"
            "       并把 server.py 的 CURRENT_VERSION 改成 " + version + "\n"
            "    3. updates\\changelog.txt 里写更新说明（没有该文件时用通用文案）"
        )
        return 0

    except BuildError as e:
        print()
        fail(str(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
