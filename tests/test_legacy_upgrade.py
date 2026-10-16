# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""老库升级路径测试。

这里用**真正的旧代码**（ReagentMGS.ReagentDatabase）写出数据库，再用新代码打开，
验证升级能正确接管 —— 而不是靠手写 SQL 模拟老库，那样测不到旧代码的真实行为。

重点验证：旧实现下外键从未开启，某些操作会静默留下孤立子行；新代码开启外键后
必须能识别并修复它们，否则那些记录之后就写不动了。

运行：python -m unittest discover -s tests -v
"""

import os
import sys
import tempfile
import unittest

_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _REPO)
# 旧实现已归档到 tests/legacy_reference/，只作为对照基准使用。
sys.path.insert(0, os.path.join(_REPO, "tests", "legacy_reference"))

from app.core.db import ReagentDatabase  # noqa: E402

# 旧实现。导入会拉起 PyQt5，但解析器与数据库类本身与 Qt 无关。
import ReagentMGS  # noqa: E402


def build_db_with_legacy_code(path):
    """用旧代码建库并写入数据，返回库路径。

    旧类的库路径取自模块级全局，这里临时改写它。
    结束后恢复，避免影响同进程里的其它测试。
    """
    original = ReagentMGS.DB_FILE
    ReagentMGS.DB_FILE = path
    try:
        legacy = ReagentMGS.ReagentDatabase()

        # 正常数据
        legacy.add_new_reagent("生理盐水", "A2409", "2024-01-01", "2026-01-01",
                               0, "06936415919665", None)
        rid = legacy.get_reagent_details("生理盐水", "A2409")[0]
        legacy.add_inbound(rid, "生理盐水", "A2409", "2024-05-01 09:00:00:000",
                           100, "张三", "06936415919665")
        legacy.add_outbound(rid, "生理盐水", "A2409", "2024-06-01 09:00:00:000",
                            30, "李四")

        # 触发旧实现的数据缺陷：gtin 已被「生理盐水」占用，
        # 旧代码的 `INSERT OR IGNORE INTO reagent_names (name, gtin)` 会被整条忽略，
        # 名称行建不出来，但 reagents 行照样插进去 —— 留下一条孤立子行。
        legacy.add_new_reagent("手工新名称", "B777", "2024-01-01", "2026-01-01",
                               0, "06936415919665", None)

        return legacy
    finally:
        ReagentMGS.DB_FILE = original


class TestLegacyUpgrade(unittest.TestCase):
    def setUp(self):
        self.path = os.path.join(tempfile.mkdtemp(prefix="rms_upg_"), "legacy.db")
        # 旧库确实含有故意制造的孤立行 —— 先确认前提成立
        self.legacy = build_db_with_legacy_code(self.path)

    def test_legacy_code_did_leave_an_orphan(self):
        """先确认前提：旧代码真的会留下孤立行，否则后面的测试是空的。"""
        self.legacy.cursor.execute('''
            SELECT COUNT(*) FROM reagents
            WHERE name NOT IN (SELECT name FROM reagent_names)
        ''')
        self.assertGreater(
            self.legacy.cursor.fetchone()[0], 0,
            "旧代码本应因 gtin 冲突而漏建父行；前提不成立说明测试写法有问题",
        )

    def test_new_code_opens_and_reads_legacy_db(self):
        db = ReagentDatabase(self.path)
        names = db.get_reagent_names()
        self.assertIn("生理盐水", names)

        detail = db.get_reagent_details("生理盐水", "A2409")
        self.assertIsNotNone(detail)
        self.assertEqual(detail[3], 70)          # 100 入 - 30 出

        recs, total = db.list_records("全部记录")
        self.assertEqual(total, 2)

    def test_new_code_repairs_the_orphan(self):
        db = ReagentDatabase(self.path)
        report = db.foreign_key_report()

        self.assertIn("手工新名称", report["repaired"],
                      "旧代码留下的孤立名称应被自动补进父表")
        self.assertEqual(report["remaining"], [])
        self.assertIn("手工新名称", db.get_reagent_names())

    def test_orphaned_reagent_becomes_writable_after_repair(self):
        """修复的实际意义：这条旧记录之后能正常出入库。"""
        db = ReagentDatabase(self.path)
        db.foreign_key_report()

        rid = db.get_reagent_details("手工新名称", "B777")[0]
        db.add_inbound(rid, "手工新名称", "B777", "2024-07-01 09:00:00:000",
                       25, "王五", None)
        self.assertEqual(db.get_reagent_details("手工新名称", "B777")[3], 25)

    def test_orphan_impact_is_rename_not_stock_updates(self):
        """明确孤儿行的真实影响面。

        SQLite 只在**被修改的外键列**上做校验，所以改 quantity 这类不碰 name 的
        更新其实不会报错 —— 早先我以为整行都写不动，是错的。真正会坏的是改名：

          * 改 quantity -> 不触发检查，照常成功
          * 改 name     -> 触发检查，父表里没有对应名称，失败

        这也解释了为什么修复仍值得做：它让这些记录重新可改名，
        并且让 `foreign_key_check` 恢复干净，后续诊断才有意义。
        """
        db = ReagentDatabase(self.path)
        rid = db.get_reagent_details("手工新名称", "B777")[0]

        # 不碰 name 的更新照常成功
        db.cursor.execute(
            'UPDATE reagents SET quantity = quantity + 1 WHERE id = ?', (rid,))
        db.conn.commit()
        self.assertEqual(db.get_reagent_details("手工新名称", "B777")[3], 1)

        # 改名则失败
        self.assertFalse(
            db.update_reagent_name_and_gtin("手工新名称", "改过的名字", None, None),
            "未修复的孤立行改名应当失败",
        )
        self.assertIsNotNone(db.get_reagent_details("手工新名称", "B777"),
                             "失败后应回滚，名称保持不变")

    def test_rename_works_after_repair(self):
        """同一操作在修复之后应当成功 —— 这就是修复的价值。"""
        db = ReagentDatabase(self.path)
        db.foreign_key_report()

        self.assertTrue(
            db.update_reagent_name_and_gtin("手工新名称", "改过的名字", None, None))
        self.assertIsNotNone(db.get_reagent_details("改过的名字", "B777"))
        self.assertIn("改过的名字", db.get_reagent_names())

    def test_schema_is_unchanged_so_legacy_db_is_compatible(self):
        """新库与旧库的表结构必须一致，否则老库升上来会缺列。"""
        db = ReagentDatabase(self.path)

        def columns(conn, table):
            conn.execute(f"PRAGMA table_info({table})")
            return [row[1] for row in conn.fetchall()]

        for table in ("reagent_names", "reagents", "inbound", "outbound"):
            self.assertEqual(
                columns(db.cursor, table),
                columns(self.legacy.cursor, table),
                f"表 {table} 的列与旧库不一致",
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
