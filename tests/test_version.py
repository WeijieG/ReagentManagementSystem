# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""版本比较测试。

重点覆盖旧实现在字典序比较下会判错的用例 —— 那些正是会导致更新推送失效的输入。

运行：python -m unittest discover -s tests -v
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.core import version  # noqa: E402


class TestSemanticComparison(unittest.TestCase):
    def assert_cmp(self, left, right, expected):
        self.assertEqual(
            version.compare_versions(left, right), expected,
            f"compare_versions({left!r}, {right!r}) 应为 {expected}",
        )

    def test_two_digit_segment_beats_lexicographic(self):
        """旧实现下 '2.0.10' < '2.0.9'，更新永远推不出去。"""
        self.assert_cmp("2.0.10", "2.0.9", 1)
        self.assert_cmp("2.0.9", "2.0.10", -1)

    def test_different_segment_counts(self):
        self.assert_cmp("2.1", "2.0.10", 1)
        self.assert_cmp("2.1.0", "2.1", 0)
        self.assert_cmp("3", "2.9.9", 1)

    def test_equality(self):
        self.assert_cmp("1.0.0", "1.0.0", 0)
        self.assert_cmp("2.0", "2.0.0", 0)

    def test_major_and_minor(self):
        self.assert_cmp("2.10.0", "2.9.0", 1)
        self.assert_cmp("10.0.0", "9.9.9", 1)

    def test_prerelease_is_older_than_release(self):
        self.assert_cmp("2.0.0-beta", "2.0.0", -1)
        self.assert_cmp("2.0.0", "2.0.0-beta", 1)

    def test_prerelease_ordering(self):
        self.assert_cmp("2.0.0-alpha", "2.0.0-beta", -1)
        self.assert_cmp("2.0.0-beta", "2.0.0-beta", 0)

    def test_empty_and_missing(self):
        self.assert_cmp("", "0", 0)
        self.assert_cmp("", "1.0.0", -1)

    def test_whitespace_tolerated(self):
        self.assert_cmp(" 2.0.10 ", "2.0.9", 1)


class TestHelpers(unittest.TestCase):
    def test_is_newer(self):
        self.assertTrue(version.is_newer("2.0.10", "2.0.9"))
        self.assertFalse(version.is_newer("2.0.9", "2.0.10"))
        self.assertFalse(version.is_newer("2.0.0", "2.0.0"))

    def test_is_same_or_older(self):
        """用于「忽略此版本」判定。

        调用形态是 is_same_or_older(可用版本, 已忽略版本)：
        忽略 2.0.10 之后，2.0.10 本身不再提示，**2.0.11 仍必须提示**。
        """
        self.assertTrue(version.is_same_or_older("2.0.10", "2.0.10"))
        self.assertFalse(version.is_same_or_older("2.0.11", "2.0.10"))
        self.assertTrue(version.is_same_or_older("2.0.9", "2.0.10"))
        # 回归：字典序下 2.0.10 > 2.0.9，这里必须按数值判
        self.assertTrue(version.is_same_or_older("2.0.9", "2.0.10"))


class TestRegressionAgainstOldBehaviour(unittest.TestCase):
    """记录旧实现的错误判定，避免将来有人"改回去"。"""

    def test_old_lexicographic_would_have_failed(self):
        old_result = "2.0.10" > "2.0.9"          # 旧实现就是这么比的
        new_result = version.is_newer("2.0.10", "2.0.9")
        self.assertFalse(old_result, "字典序确实会判错")
        self.assertTrue(new_result, "新实现必须判对")


if __name__ == "__main__":
    unittest.main(verbosity=2)
