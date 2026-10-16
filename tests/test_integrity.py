# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""完整性校验测试。

这块最要命的一条是「哈希留空时放行」。改这个行为前请先想清楚：

每次重新打包 exe 的哈希都会变，而补丁/升级只替换 exe、不替换 config.ini，
于是旧哈希必然对不上。如果这时判定为失败，用户看到的是「应用不完整」——
一个指向完全错误方向的报错，而且程序直接打不开。这个坑前后踩过两次。

运行：python -m unittest discover -s tests -v
"""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import paths  # noqa: E402
from app.core import integrity  # noqa: E402


class IntegrityTestBase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="rms-integrity-")
        self.config_path = os.path.join(self.tmp, "config.ini")

        # 这些测试全部在「已打包」的前提下跑 —— 开发环境本来就不校验。
        self._frozen = paths.IS_FROZEN
        paths.IS_FROZEN = True

    def tearDown(self):
        paths.IS_FROZEN = self._frozen
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write_config(self, x86="", x64="", ui=""):
        with open(self.config_path, "w", encoding="utf-8") as f:
            f.write(
                "[Application]\nversion = 2.0.0\n\n"
                "[Security]\n"
                f'mainfilehashe_x86 = {x86}\n'
                f'mainfilehashe_x64 = {x64}\n'
                f'uibundlehashe = {ui}\n'
            )

    def check(self, current_hash="deadbeef", ui_hash=None):
        """把当前程序哈希固定住，免得测试依赖真实 exe 是否存在。"""
        original = integrity.calculate_file_hashes
        integrity.calculate_file_hashes = lambda *a, **k: current_hash
        if ui_hash is not None:
            original_ui = integrity.calculate_ui_bundle_hash
            integrity.calculate_ui_bundle_hash = lambda *a, **k: ui_hash
            self.addCleanup(setattr, integrity, "calculate_ui_bundle_hash", original_ui)
        self.addCleanup(setattr, integrity, "calculate_file_hashes", original)
        return integrity.check_file_integrity(self.config_path)


class TestEmptyHashBypass(IntegrityTestBase):
    """补丁只换 exe 不换 config.ini，所以「空哈希」必须是放行而不是拦路。"""

    def test_both_empty_passes(self):
        self.write_config(x86='""', x64='""')
        ok, msg = self.check()
        self.assertTrue(ok, f"哈希留空应当放行，实际被拦: {msg}")
        self.assertEqual(msg, "")

    def test_both_empty_passes_even_when_exe_hash_is_junk(self):
        """exe 都读不到（返回 '-1'）也不该拦住 —— 空哈希的语义是「未配置」。"""
        self.write_config(x86='""', x64='""')
        ok, _ = self.check(current_hash="-1")
        self.assertTrue(ok)

    def test_missing_keys_passes(self):
        """config.ini 里根本没有这几行，等同于未配置。"""
        with open(self.config_path, "w", encoding="utf-8") as f:
            f.write("[Application]\nversion = 2.0.0\n")
        ok, _ = self.check()
        self.assertTrue(ok)

    def test_only_x86_configured_still_checks(self):
        """只清空了一个不算未配置 —— 另一个还配着，就得照常校验。"""
        self.write_config(x86='""', x64='"abcd"')
        ok, msg = self.check(current_hash="deadbeef")
        self.assertFalse(ok, "还有一个哈希配着，不该跳过校验")
        self.assertIn("完整性", msg)


class TestHashComparison(IntegrityTestBase):
    def test_matching_hash_passes(self):
        self.write_config(x86='"deadbeef"')
        ok, _ = self.check(current_hash="deadbeef")
        self.assertTrue(ok)

    def test_mismatch_is_rejected(self):
        self.write_config(x64='"aaaa"')
        ok, msg = self.check(current_hash="bbbb")
        self.assertFalse(ok)
        self.assertIn("完整性", msg)

    def test_either_hash_matching_is_enough(self):
        """旧实现就是这样：x86/x64 匹配任意一个即通过。"""
        self.write_config(x86='"aaaa"', x64='"deadbeef"')
        ok, _ = self.check(current_hash="deadbeef")
        self.assertTrue(ok)

    def test_mismatch_message_points_at_gen_hashes(self):
        """报错要给出下一步动作，否则用户只会去重装。"""
        self.write_config(x86='"aaaa"')
        _, msg = self.check(current_hash="bbbb")
        self.assertIn("gen_hashes", msg)


class TestUiBundle(IntegrityTestBase):
    def test_ui_hash_skipped_when_unset(self):
        self.write_config(x86='"deadbeef"', ui='""')
        ok, _ = self.check(current_hash="deadbeef", ui_hash="whatever")
        self.assertTrue(ok)

    def test_ui_hash_enforced_when_set(self):
        self.write_config(x86='"deadbeef"', ui='"cccc"')
        ok, msg = self.check(current_hash="deadbeef", ui_hash="dddd")
        self.assertFalse(ok)
        self.assertIn("界面资源", msg)

    def test_ui_hash_passing(self):
        self.write_config(x86='"deadbeef"', ui='"cccc"')
        ok, _ = self.check(current_hash="deadbeef", ui_hash="cccc")
        self.assertTrue(ok)


class TestDevelopmentMode(IntegrityTestBase):
    def test_dev_mode_never_blocks(self):
        """开发环境连哈希都不该比 —— 改代码后哈希必然变。"""
        self.write_config(x86='"aaaa"')
        paths.IS_FROZEN = False
        ok, _ = self.check(current_hash="bbbb")
        self.assertTrue(ok)


class TestMissingConfig(IntegrityTestBase):
    def test_missing_config_is_rejected(self):
        """唯一一个「配置缺失即失败」的判定，沿用旧实现。"""
        ok, msg = integrity.check_file_integrity(
            os.path.join(self.tmp, "not-there.ini")
        )
        self.assertFalse(ok)
        self.assertIn("配置文件不存在", msg)


class TestRealHashCalculation(unittest.TestCase):
    """真算一遍，确认返回值契约没被改坏。"""

    def test_missing_file_returns_minus_one_string(self):
        self.assertEqual(
            integrity.calculate_file_hashes(os.path.join(tempfile.gettempdir(), "nope.exe")),
            "-1",
        )

    def test_shipped_config_has_empty_hashes(self):
        """仓库里的 config.ini 必须是空哈希，否则开发构建一打补丁就弹完整性异常。"""
        config = os.path.join(paths._DEV_ROOT, "config.ini")
        if not os.path.exists(config):
            self.skipTest("根目录无 config.ini")
        import configparser
        parser = configparser.ConfigParser()
        parser.read(config, encoding="utf-8")
        for option in ("MainFileHashe_x86", "MainFileHashe_x64", "UiBundleHashe"):
            if parser.has_option("Security", option):
                self.assertEqual(
                    integrity._read_expected(parser, "Security", option), "",
                    f"config.ini 的 {option} 应为空，否则启动会被完整性校验拦住",
                )


if __name__ == "__main__":
    unittest.main()
