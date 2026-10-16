# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""构建配置的回归测试。

这里只盯**咬过人的几个点**，不对 setup.iss 做全面解析 —— 那种测试很脆，
改个注释就红，最后会被无视。

真正防的是这三件事：

  1. 版本号漂移。config.ini 改了、version.iss 没跟着生成，
     打出来的安装包版本号是旧的。
  2. setup.iss 又去单独 Source 仓库根目录的 config.ini。那会把产物里
     已经写好完整性哈希的那份覆盖掉，校验永久失效 —— 而且完全不报错。
  3. 少了 CloseApplications=force。升级时 QtWebEngine 的辅助进程占着
     Qt5WebEngineCore.dll，会留下新旧混装的目录。

运行：python -m unittest discover -s tests -v
"""

import configparser
import os
import re
import sys
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)

CONFIG_INI = os.path.join(REPO, "config.ini")
VERSION_ISS = os.path.join(REPO, "version.iss")
SETUP_ISS = os.path.join(REPO, "setup.iss")


def read_text(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


class TestVersionConsistency(unittest.TestCase):
    def setUp(self):
        self.config = configparser.ConfigParser()
        self.config.read(CONFIG_INI, encoding="utf-8")

    def test_version_is_well_formed(self):
        version = self.config.get("Application", "Version", fallback="")
        self.assertRegex(version, r"^\d+(\.\d+)*$",
                         "版本号必须形如 2.1.0 —— server.py 和安装包都按它比较")

    def test_version_iss_matches_config(self):
        """version.iss 由构建脚本生成；不一致说明有人手工改了版本号没重新构建。"""
        self.assertTrue(os.path.isfile(VERSION_ISS),
                        "version.iss 不存在，出安装包前请跑 tools/build_installer.py")
        match = re.search(r'(?m)^#define\s+MyAppVersion\s+"([^"]+)"', read_text(VERSION_ISS))
        self.assertIsNotNone(match, "version.iss 里找不到 MyAppVersion")
        self.assertEqual(
            match.group(1), self.config.get("Application", "Version"),
            "version.iss 与 config.ini 的版本号不一致，重跑 tools/build_installer.py",
        )

    def test_version_iss_quad_is_numeric(self):
        """VersionInfoVersion 必须是四段纯数字，否则 ISCC 直接编译失败。"""
        match = re.search(r'(?m)^#define\s+MyAppVersionQuad\s+"([^"]+)"', read_text(VERSION_ISS))
        self.assertIsNotNone(match, "version.iss 里找不到 MyAppVersionQuad")
        parts = match.group(1).split(".")
        self.assertEqual(len(parts), 4, f"必须是四段：{match.group(1)}")
        for p in parts:
            self.assertTrue(p.isdigit(), f"必须是纯数字：{match.group(1)}")

    def test_server_version_synced(self):
        """server.py 的 CURRENT_VERSION 不同步，更新推送会静默失效。"""
        server_py = os.path.join(REPO, "server.py")
        if not os.path.isfile(server_py):
            self.skipTest("没有 server.py")
        match = re.search(r"(?m)^CURRENT_VERSION\s*=\s*['\"]([^'\"]+)['\"]",
                          read_text(server_py))
        self.assertIsNotNone(match, "server.py 里找不到 CURRENT_VERSION")
        self.assertEqual(
            match.group(1), self.config.get("Application", "Version"),
            "server.py 的 CURRENT_VERSION 与 config.ini 版本号不一致，"
            "客户端将永远收不到更新推送",
        )


class TestSetupScript(unittest.TestCase):
    def setUp(self):
        self.text = read_text(SETUP_ISS)

    def test_does_not_source_repo_config_ini(self):
        """产物里的 config.ini 带哈希，仓库里那份是空模板。单独 Source 仓库的会覆盖掉。"""
        offenders = re.findall(r'(?m)^\s*Source:\s*"[.\\/]*config\.ini"', self.text)
        self.assertEqual(
            offenders, [],
            "setup.iss 不该单独 Source 仓库根目录的 config.ini —— "
            "那会把产物里写入完整性哈希的那份覆盖掉，校验静默失效",
        )

    def test_sources_whole_dist_directory(self):
        self.assertIn(r'Source: "{#MyDistDir}\*"', self.text,
                      "必须整目录分发（onedir）")

    def test_close_applications_forced(self):
        self.assertRegex(
            self.text, r"(?mi)^\s*CloseApplications\s*=\s*force\s*$",
            "升级时必须强制关闭正在运行的程序，否则会留下新旧混版的目录",
        )

    def test_does_not_set_appid(self):
        """设了 AppId 就变成另一个应用，无法原地升级旧版。"""
        self.assertIsNone(
            re.search(r"(?mi)^\s*AppId\s*=", self.text),
            "不要设置 AppId —— 不写时 Inno 用 AppName 作为 AppId，"
            "与旧版一致才能原地升级",
        )

    def test_targets_win7_sp1(self):
        self.assertRegex(self.text, r"(?mi)^\s*MinVersion\s*=\s*6\.1sp1\s*$",
                         "MinVersion 应为 6.1sp1（Win7 SP1）")

    def test_does_not_delete_user_database_on_uninstall(self):
        """[Code] 里的卸载备份只应复制，不应出现删除现役数据库的调用。"""
        self.assertIn("BackupUserData", self.text, "卸载备份逻辑不见了")
        self.assertNotIn("DeleteFile(DbFile)", self.text,
                         "卸载时不要删除现役数据库")

    def test_deletes_stale_msvc_runtime_on_upgrade(self):
        """升级时必须清掉旧版本留下的 MSVC 运行库副本。

        新版本刻意不再分发这些文件，但 Inno 不会删除"已存在而新版本不再提供"
        的文件。不显式清掉的话，覆盖安装后那个坏的副本还在程序目录里，
        照样盖住 System32 里好的那份 —— 升级等于没修。
        """
        self.assertIn("[InstallDelete]", self.text,
                      "缺 [InstallDelete]：覆盖安装不会清掉旧的 MSVC 运行库副本")
        for name in ("MSVCP140.dll", "VCRUNTIME140.dll"):
            self.assertRegex(
                self.text,
                r'(?mi)^\s*Type:\s*files;\s*Name:\s*"\{app\}\\%s"' % re.escape(name),
                f"[InstallDelete] 里没有 {name}",
            )

    def test_vcredist_check_covers_all_runtime_files(self):
        """程序不再自带运行库，四个文件缺一个都起不来，检查必须覆盖全部。"""
        for name in ("vcruntime140.dll", "vcruntime140_1.dll",
                     "msvcp140.dll", "msvcp140_1.dll"):
            self.assertIn(name, self.text,
                          f"安装脚本没有检查 {name}")

    def test_includes_version_iss(self):
        self.assertIn('#include "version.iss"', self.text)


class TestBuildScript(unittest.TestCase):
    """构建脚本里的几个纯函数。"""

    def setUp(self):
        sys.path.insert(0, os.path.join(REPO, "tools"))
        import build_installer
        self.mod = build_installer

    def test_read_version(self):
        self.assertRegex(self.mod.read_version(), r"^\d+\.\d+")

    def test_dist_checks_are_not_empty(self):
        self.assertTrue(self.mod.REQUIRED_PATHS)
        self.assertTrue(self.mod.REQUIRED_BY_NAME)

    def test_dist_checks_use_basenames_for_qt_files(self):
        """Qt 文件必须按文件名搜。写死路径会因为 PyInstaller 布局变化而误报。"""
        for name in self.mod.REQUIRED_BY_NAME:
            self.assertNotIn("/", name)
            self.assertNotIn("\\", name)

    def test_excludes_msvc_runtime_from_bundle(self):
        """MSVC 运行库必须排除。

        PyInstaller 从打包机的 System32 抓这些文件，打包机装什么版本就带什么
        版本。开发机上是 VS2026 的 14.51，而它静态导入了 Win8 才有的
        GetSystemTimePreciseAsFileTime，Win7 上加载即失败。
        """
        text = read_text(os.path.join(REPO, "build.spec"))
        self.assertIn("MSVC_RUNTIME_PREFIXES", text,
                      "build.spec 里不再排除 MSVC 运行库了？")
        for prefix in ("msvcp140", "vcruntime140"):
            self.assertIn(prefix, text.lower())

    def test_exact_paths_point_at_files_not_qt_dlls(self):
        for path in self.mod.REQUIRED_PATHS:
            self.assertFalse(path.lower().endswith(".dll"),
                             f"{path} 应该按文件名搜，不要写死路径")


class TestBuiltDist(unittest.TestCase):
    """直接检查已构建的产物。没构建过就跳过 —— 不强迫先跑一遍 PyInstaller。"""

    DIST = os.path.join(REPO, "Dist", "ReagentManagementSystem")

    def setUp(self):
        if not os.path.isdir(self.DIST):
            self.skipTest("还没构建过产物（跑 tools/build_installer.py）")
        sys.path.insert(0, os.path.join(REPO, "tools"))

    def test_no_win8_only_api_imports(self):
        """产物里不能有文件静态导入 Win8+ 才有的 API。

        这类问题在开发机上永远复现不了（开发机是 Win10，那些 API 都在），
        到 Win7 上才炸，而且报的是「无法定位程序输入点 XXX」——
        看不出跟"打包机装了新版编译器"有关。判定逻辑与构建脚本用的是同一套，
        但这个测试能在没构建的时候也把源码层面的改动兜住。
        """
        import pe_imports
        import build_installer

        findings = []
        for root, dirs, files in os.walk(self.DIST):
            for name in files:
                if os.path.splitext(name)[1].lower() not in (".dll", ".pyd", ".exe"):
                    continue
                path = os.path.join(root, name)
                symbols = pe_imports.parse_import_symbols(path)
                if symbols is None:
                    continue
                for dll, sym in symbols:
                    if sym in build_installer.WIN8_ONLY_APIS:
                        findings.append(
                            f"{os.path.relpath(path, self.DIST)} 从 {dll} 导入 {sym}")

        self.assertEqual(
            findings, [],
            "产物里有文件静态导入了 Windows 7 不存在的 API，"
            "装到 Win7 会报「无法定位程序输入点」：\n  " + "\n  ".join(findings),
        )

    def test_msvc_runtime_not_bundled(self):
        """MSVC 运行库不能随包 —— app-local 副本会盖住系统里那份好的。"""
        for name in ("MSVCP140.dll", "MSVCP140_1.dll",
                     "VCRUNTIME140.dll", "VCRUNTIME140_1.dll"):
            self.assertFalse(
                os.path.exists(os.path.join(self.DIST, name)),
                f"产物里又出现 {name} 了。它多半是从打包机的 System32 抓来的，"
                "版本可能比目标机的新，甚至不支持 Win7",
            )


if __name__ == "__main__":
    unittest.main()
