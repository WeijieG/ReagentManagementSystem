# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""GS1 解析器行为对照测试。

把旧实现（ReagentMGS.GS1Parser）与新实现（app.core.gs1）放在同一语料上跑，
断言两者输出字典逐键相等 —— 这是重构期间判断"行为没变"的依据。

运行：python -m unittest discover -s tests -v
"""

import os
import sys
import unittest

_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _REPO)
# 旧实现已归档到 tests/legacy_reference/ —— 它只作为本测试的对照基准存在，
# 不再是程序的一部分。见该目录下的 README。
sys.path.insert(0, os.path.join(_REPO, "tests", "legacy_reference"))

from app.core import gs1 as new_gs1  # noqa: E402

# 旧实现在导入时会初始化 PyQt 相关模块；解析器本身与 Qt 无关。
import ReagentMGS  # noqa: E402


# 必须与旧实现输出完全一致的语料
PARITY_CORPUS = [
    # 标准括号形式
    "(01)06936415919665(17)260218(10)A2409",
    "(01)06936415919665(11)240819(17)260819(10)97105-002512-00",
    # 全角括号
    "（01）06936415919665（17）260218",
    # 只有部分字段
    "(10)BATCH-001",
    "(17)260218",
    # 未知 AI
    "(71)ABC123",
    "(99)INTERNAL-CODE",
    # 定长数据不足 / 超长
    "(01)12345",
    "(01)06936415919665000000",
    # 小数点 AI
    "(3103)001234",
    "(3100)000500",
    # 空与无括号输入（旧实现均返回空字典，新实现也必须如此）
    "",
    "   ",
    "no-ai-here",
    "abc-123-xyz",
]

# 有意为之的行为差异：旧实现对这些输入返回空字典，新实现解析出字段。
# 逐条固定住，任何意外扩大都要让测试失败。
DOCUMENTED_DIVERGENCES = [
    # 裸 GTIN（扫码枪输出原始 GS1-128 但单个定长 AI 后无 FNC1）
    ("0106936415919665", {}, {"GTIN": "06936415919665"}),
    # 裸变长字段
    ("10LOT-2024", {}, {"Batch/Lot Number": "LOT-2024"}),
]


class TestGs1Parity(unittest.TestCase):
    def test_new_matches_old_on_corpus(self):
        """语料上新旧输出必须逐键相等。"""
        mismatches = []
        for raw in PARITY_CORPUS:
            old = ReagentMGS.GS1Parser.parse_barcode(raw)
            new = new_gs1.parse_barcode(raw)
            if old != new:
                mismatches.append((raw, old, new))

        self.assertEqual(
            mismatches, [],
            "新旧解析器输出不一致:\n" + "\n".join(
                f"  输入 {r!r}\n    旧: {o}\n    新: {n}" for r, o, n in mismatches
            ),
        )

    def test_documented_divergences_are_exactly_as_intended(self):
        """新旧行为差异必须是已知且有限的，不能悄悄扩散。"""
        for raw, expected_old, expected_new in DOCUMENTED_DIVERGENCES:
            self.assertEqual(
                ReagentMGS.GS1Parser.parse_barcode(raw), expected_old,
                f"旧实现行为变了: {raw!r}",
            )
            self.assertEqual(
                new_gs1.parse_barcode(raw), expected_new,
                f"新实现输出不符预期: {raw!r}",
            )

    def test_parenthesised_path_unchanged_for_known_example(self):
        """固定样本的逐字段断言，防止对照测试本身被改坏。"""
        parsed = new_gs1.parse_barcode("(01)06936415919665(17)260218(10)A2409")
        self.assertEqual(parsed["GTIN"], "06936415919665")
        self.assertEqual(parsed["Expiration Date"], "260218")
        self.assertEqual(parsed["Batch/Lot Number"], "A2409")


class TestGs1RawFnc1(unittest.TestCase):
    """原始 GS1-128（FNC1 分隔）支持 —— 本次新增的增量能力。"""

    def test_raw_fnc1_is_parsed(self):
        """旧实现对此输入返回空字典；新实现应正确解析。"""
        raw = (
            "0106936415919665\x1d"
            "11240819\x1d"
            "17260218\x1d"
            "102024050123\x1d"
            "97105-002512-00\x1d"
            "98AAX202405012303531"
        )

        self.assertEqual(ReagentMGS.GS1Parser.parse_barcode(raw), {},
                         "旧实现本就不支持原始格式，这是本次修复的出发点")

        parsed = new_gs1.parse_barcode(raw)
        self.assertEqual(parsed["GTIN"], "06936415919665")
        self.assertEqual(parsed["Production Date"], "240819")
        self.assertEqual(parsed["Expiration Date"], "260218")
        self.assertEqual(parsed["Batch/Lot Number"], "2024050123")

    def test_raw_without_fnc1_for_trailing_fixed_field(self):
        """变长字段在末尾、后面还有定长字段时，FNC1 缺失不应误吞后续 AI。"""
        # 10 为变长且不在末尾：必须靠 FNC1 定界
        parsed = new_gs1.parse_barcode("0106936415919665\x1d102024050123\x1d17260218")
        self.assertEqual(parsed["Batch/Lot Number"], "2024050123")
        self.assertEqual(parsed["Expiration Date"], "260218")

    def test_raw_truncated_batch_stops_at_end(self):
        """变长字段在末尾且无 FNC1 时，读到字符串结尾。"""
        parsed = new_gs1.parse_barcode("0106936415919665\x1d10LOT-99")
        self.assertEqual(parsed["Batch/Lot Number"], "LOT-99")

    def test_parenthesised_still_wins_when_both_possible(self):
        """括号形式优先，原始路径不干扰既有输入。"""
        parsed = new_gs1.parse_barcode("(01)06936415919665")
        self.assertEqual(parsed, {"GTIN": "06936415919665"})


if __name__ == "__main__":
    unittest.main(verbosity=2)
