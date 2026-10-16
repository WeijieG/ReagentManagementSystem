# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""处理器层测试。

入库/出库是整套代码里唯一会改数据的部分，也是风险最高的地方。
这里用桩件替代 Qt 依赖，直接在 Python 层验证提交语义与校验规则，
包括旧实现里的几个边界行为。

运行：python -m unittest discover -s tests -v
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.bridge import BridgeError, ErrorCode  # noqa: E402
from app.core.db import ReagentDatabase  # noqa: E402
from app.handlers.actions import build_action_handlers, now_timestamp  # noqa: E402
from app.handlers.queries import build_query_handlers  # noqa: E402


class StubApi:
    """收集 emit_event 的调用，用于断言「改动后有没有通知前端」。"""

    def __init__(self):
        self.events = []

    def emit_event(self, event, data=None):
        self.events.append((event, data or {}))

    def toast(self, kind, text):
        self.emit_event("toast", {"kind": kind, "text": text})


class StubUi:
    def __init__(self):
        self.saved = []

    def confirm(self, title, text):
        return True

    def warn(self, title, text):
        pass

    def info(self, title, text):
        pass

    def choose_save_file(self, title, default_name, file_filter):
        return self.saved.pop(0) if self.saved else ""

    def choose_open_file(self, title, file_filter):
        return ""

    def quit_app(self):
        pass


class StubCtx:
    pass


def make_ctx():
    path = os.path.join(tempfile.mkdtemp(prefix="rms_h_"), "h.db")
    ctx = StubCtx()
    ctx.db = ReagentDatabase(path)
    ctx.api = StubApi()
    ctx.ui = StubUi()
    ctx.version = "2.0.0"
    ctx.config_file = None
    return ctx


class TestInbound(unittest.TestCase):
    def setUp(self):
        self.ctx = make_ctx()
        self.h = build_action_handlers(self.ctx)

    def test_creates_reagent_then_adds_stock(self):
        res = self.h["inbound.submit"]({
            "name": "生理盐水", "batch": "A2409",
            "quantity": 100, "operator": "张三",
            "gtin": "06936415919665",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        self.assertEqual(res["quantity"], 100)

        detail = self.ctx.db.get_reagent_details("生理盐水", "A2409")
        self.assertIsNotNone(detail, "应自动创建试剂记录")
        self.assertEqual(detail[3], 100)

        # 名称也应进 reagent_names
        self.assertIn("生理盐水", self.ctx.db.get_reagent_names())

    def test_second_inbound_accumulates(self):
        base = {"name": "生理盐水", "batch": "A2409", "operator": "张三",
                "productionDate": "2024-01-01", "expiryDate": "2026-01-01"}
        self.h["inbound.submit"](dict(base, quantity=100))
        self.h["inbound.submit"](dict(base, quantity=50))
        self.assertEqual(
            self.ctx.db.get_reagent_details("生理盐水", "A2409")[3], 150)

    def test_remaining_quantity_recorded_on_row(self):
        """inbound 行上的 remaining_quantity 应是入库后的库存快照。"""
        self.h["inbound.submit"]({
            "name": "质控液", "batch": "B1", "quantity": 40, "operator": "",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        rows = self.ctx.db.get_inbound_records()
        self.assertEqual(rows[-1][6], 40)

    def test_rejects_missing_fields(self):
        for payload in (
            {"name": "", "batch": "B", "quantity": 1},
            {"name": "N", "batch": "", "quantity": 1},
            {"name": "N", "batch": "B", "quantity": ""},
        ):
            with self.assertRaises(BridgeError) as cm:
                self.h["inbound.submit"](payload)
            self.assertEqual(cm.exception.code, ErrorCode.VALIDATION)

    def test_rejects_non_positive_and_non_numeric_quantity(self):
        for qty in (0, -5, "abc", "1.5"):
            with self.assertRaises(BridgeError) as cm:
                self.h["inbound.submit"]({
                    "name": "N", "batch": "B", "quantity": qty,
                    "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
                })
            self.assertEqual(cm.exception.code, ErrorCode.VALIDATION)

    def test_notifies_frontend(self):
        self.h["inbound.submit"]({
            "name": "N", "batch": "B", "quantity": 1, "operator": "",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        events = [e for e, _ in self.ctx.api.events]
        self.assertIn("db.changed", events)


class TestOutbound(unittest.TestCase):
    def setUp(self):
        self.ctx = make_ctx()
        self.h = build_action_handlers(self.ctx)
        self.h["inbound.submit"]({
            "name": "生理盐水", "batch": "A2409", "quantity": 100, "operator": "张三",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })

    def test_reduces_stock(self):
        res = self.h["outbound.submit"]({
            "name": "生理盐水", "batch": "A2409", "quantity": 30, "operator": "李四",
        })
        self.assertEqual(res["quantity"], 30)
        self.assertEqual(
            self.ctx.db.get_reagent_details("生理盐水", "A2409")[3], 70)

    def test_exact_stock_is_allowed(self):
        """边界：出库数量恰好等于库存，应当成功（旧实现是 current < qty 才拒绝）。"""
        self.h["outbound.submit"]({
            "name": "生理盐水", "batch": "A2409", "quantity": 100, "operator": "",
        })
        self.assertEqual(
            self.ctx.db.get_reagent_details("生理盐水", "A2409")[3], 0)

    def test_overdraw_is_refused(self):
        """边界：多出 1 个就要拒绝，且库存不得变动。"""
        with self.assertRaises(BridgeError) as cm:
            self.h["outbound.submit"]({
                "name": "生理盐水", "batch": "A2409", "quantity": 101, "operator": "",
            })
        self.assertEqual(cm.exception.code, ErrorCode.STOCK_INSUFFICIENT)
        self.assertEqual(
            self.ctx.db.get_reagent_details("生理盐水", "A2409")[3], 100,
            "被拒绝时库存必须原样不动")

    def test_unknown_reagent_is_not_found(self):
        with self.assertRaises(BridgeError) as cm:
            self.h["outbound.submit"]({
                "name": "不存在", "batch": "X", "quantity": 1, "operator": "",
            })
        self.assertEqual(cm.exception.code, ErrorCode.NOT_FOUND)


class TestNamesCrud(unittest.TestCase):
    def setUp(self):
        self.ctx = make_ctx()
        self.h = build_action_handlers(self.ctx)

    def test_add_then_duplicate(self):
        self.h["names.add"]({"name": "生理盐水", "gtin": "06936415919665"})
        with self.assertRaises(BridgeError) as cm:
            self.h["names.add"]({"name": "生理盐水"})
        self.assertEqual(cm.exception.code, ErrorCode.DUPLICATE)

    def test_delete_guarded_when_stock_exists(self):
        self.h["inbound.submit"]({
            "name": "生理盐水", "batch": "A2409", "quantity": 10, "operator": "",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        with self.assertRaises(BridgeError) as cm:
            self.h["names.delete"]({"name": "生理盐水"})
        self.assertEqual(cm.exception.code, ErrorCode.GUARDED)

    def test_delete_succeeds_when_unused(self):
        self.h["names.add"]({"name": "临时名称"})
        self.h["names.delete"]({"name": "临时名称"})
        self.assertNotIn("临时名称", self.ctx.db.get_reagent_names())

    def test_rename_cascades_to_records(self):
        """改名要同步到 reagents / inbound / outbound —— 这是最容易被打坏的地方。"""
        self.h["inbound.submit"]({
            "name": "旧名", "batch": "B1", "quantity": 10, "operator": "",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        self.h["names.update"]({
            "oldName": "旧名", "oldGtin": None, "newName": "新名", "newGtin": None,
        })

        self.assertIn("新名", self.ctx.db.get_reagent_names())
        self.assertNotIn("旧名", self.ctx.db.get_reagent_names())
        self.assertIsNotNone(self.ctx.db.get_reagent_details("新名", "B1"))
        self.assertEqual(self.ctx.db.get_inbound_records()[-1][1], "新名")

    def test_rename_rejects_conflict(self):
        self.h["names.add"]({"name": "A"})
        self.h["names.add"]({"name": "B"})
        with self.assertRaises(BridgeError) as cm:
            self.h["names.update"]({
                "oldName": "A", "oldGtin": None, "newName": "B", "newGtin": None,
            })
        self.assertEqual(cm.exception.code, ErrorCode.DUPLICATE)


class TestGtinUniqueness(unittest.TestCase):
    def setUp(self):
        self.ctx = make_ctx()
        self.h = build_action_handlers(self.ctx)

    def test_two_names_cannot_share_gtin(self):
        self.h["names.add"]({"name": "A", "gtin": "111"})
        with self.assertRaises(BridgeError):
            self.h["names.add"]({"name": "B", "gtin": "111"})

    def test_inbound_with_gtin_bound_to_another_name(self):
        """扫描到的 GTIN 已被别的名称占用时，仍应能入库，只是不绑定该 GTIN。

        回归用例：原来的写法是 `INSERT OR IGNORE INTO reagent_names (name, gtin)`，
        gtin 冲突会让**整条插入**被忽略，名称行建不出来，紧接着插 reagents 时
        外键找不到父行而失败 —— 表现为「添加新试剂失败」，用户无法继续。
        """
        self.h["names.add"]({"name": "生理盐水", "gtin": "06936415919665"})

        # 用户扫了属于「生理盐水」的码，却手输了另一个名称
        res = self.h["inbound.submit"]({
            "name": "手工新名称", "batch": "B1", "quantity": 5, "operator": "",
            "gtin": "06936415919665",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        self.assertEqual(res["quantity"], 5)

        # 名称与库存都应建好
        self.assertIn("手工新名称", self.ctx.db.get_reagent_names())
        self.assertEqual(
            self.ctx.db.get_reagent_details("手工新名称", "B1")[3], 5)

        # 但 GTIN 仍归原名称所有，不能被抢走
        self.assertEqual(
            self.ctx.db.get_name_by_gtin("06936415919665"), "生理盐水")

    def test_inbound_binds_gtin_to_existing_unbound_name(self):
        """名称已存在但未绑定 GTIN 时，扫码入库应顺手把它绑上。"""
        self.h["names.add"]({"name": "未绑定试剂"})
        self.assertEqual(
            self.ctx.db.get_gtin_bindings().get("未绑定试剂"), "")

        self.h["inbound.submit"]({
            "name": "未绑定试剂", "batch": "B9", "quantity": 3, "operator": "",
            "gtin": "12345678901234",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        self.assertEqual(
            self.ctx.db.get_gtin_bindings().get("未绑定试剂"), "12345678901234")

    def test_no_orphan_reagents_created(self):
        """任何入库路径都不应产生「reagents 里的名称在 reagent_names 里不存在」。"""
        self.h["names.add"]({"name": "A", "gtin": "999"})
        for name, gtin in (("B", "999"), ("C", None), ("A", "999")):
            self.h["inbound.submit"]({
                "name": name, "batch": "B1", "quantity": 1, "operator": "",
                "gtin": gtin,
                "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
            })
        self.assertEqual(self.ctx.db.check_foreign_keys(), [],
                         "不应存在外键违例（孤立子行）")


class TestQueries(unittest.TestCase):
    def setUp(self):
        self.ctx = make_ctx()
        self.actions = build_action_handlers(self.ctx)
        self.queries = build_query_handlers(self.ctx)

    def test_gs1_date_expansion_is_always_20xx(self):
        """沿用旧实现：两位年份恒解析为 20xx（"19" 分支不可达）。"""
        res = self.queries["gs1.parse"]({"raw": "(01)06936415919665(17)260218"})
        self.assertEqual(res["expiryDate"], "2026-02-18")

    def test_gs1_expansion_handles_88(self):
        """88 应得到 2088 而不是 1988 —— 如实反映现有行为。"""
        res = self.queries["gs1.parse"]({"raw": "(17)880101"})
        self.assertEqual(res["expiryDate"], "2088-01-01")

    def test_gs1_reports_unrecognised(self):
        res = self.queries["gs1.parse"]({"raw": "not-a-barcode"})
        self.assertFalse(res["recognised"])

    def test_inventory_page_rejects_absurd_page_size(self):
        with self.assertRaises(BridgeError) as cm:
            self.queries["inventory.page"]({"pageSize": 99999})
        self.assertEqual(cm.exception.code, ErrorCode.VALIDATION)

    def test_inventory_sort_rejects_unknown_column(self):
        """未知列名应回退到 id，而不是拼进 SQL。"""
        res = self.queries["inventory.page"]({
            "pageSize": 10, "sort": {"col": "id; DROP TABLE reagents", "dir": "asc"},
        })
        self.assertEqual(res["total"], 0)
        # 表还在，说明没有被注入
        self.ctx.db.cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='reagents'")
        self.assertIsNotNone(self.ctx.db.cursor.fetchone())

    def test_expiry_state_thresholds(self):
        """阈值必须与旧 UI 逐档一致。

        旧实现是 `if days_remaining < 0: 已过期 elif days_remaining <= 30: 橙 ...`，
        所以**当天到期（剩 0 天）落在橙色档，不算已过期**。这是历史行为，
        本次重构如实保留，测试也据此固定下来。
        """
        cases = [(-1, "expired"), (0, "warn30"), (1, "warn30"), (30, "warn30"),
                 (31, "warn90"), (90, "warn90"), (91, "ok"), (300, "ok")]
        for days, expected in cases:
            self.actions["inbound.submit"]({
                "name": f"N{days}", "batch": "B", "quantity": 1, "operator": "",
                "productionDate": "2000-01-01",
                # 用一个固定基准日反推有效期，避免测试随时间漂移
                "expiryDate": _date_offset(days),
            })

        res = self.queries["inventory.page"]({"pageSize": 100})
        by_name = {r["name"]: r["state"] for r in res["rows"]}
        for days, expected in cases:
            self.assertEqual(by_name.get(f"N{days}"), expected,
                             f"剩余 {days} 天的状态应为 {expected}")


def _date_offset(days):
    import datetime
    return (datetime.date.today() + datetime.timedelta(days=days)).isoformat()


class TestLegacyDatabaseRepair(unittest.TestCase):
    """老库升级路径。

    历史库是在外键关闭的状态下写出来的，可能带着孤立子行进来。
    开启 `PRAGMA foreign_keys` 不会追溯校验既有数据，但一旦对孤立行做写操作
    就会报 FOREIGN KEY constraint failed。所以开库时要自检并修复。
    """

    def _make_legacy_db(self):
        """造一个含孤立 reagent_names 引用的老库（外键关闭状态下写入）。"""
        path = os.path.join(tempfile.mkdtemp(prefix="rms_legacy_"), "legacy.db")
        db = ReagentDatabase(path)
        db.cursor.execute("PRAGMA foreign_keys = OFF")
        db.cursor.execute('''
            INSERT INTO reagents (name, batch, production_date, expiry_date, quantity)
            VALUES ('孤儿名称', 'B1', '2024-01-01', '2026-01-01', 5)
        ''')
        db.conn.commit()
        db.cursor.execute("PRAGMA foreign_keys = ON")
        return db

    def test_violation_is_detected(self):
        db = self._make_legacy_db()
        violations = db.check_foreign_keys()
        self.assertEqual(len(violations), 1, f"应检出 1 条违例，实际 {violations}")
        self.assertEqual(violations[0][0], "reagents")

    def test_report_repairs_missing_parent(self):
        db = self._make_legacy_db()
        report = db.foreign_key_report()
        self.assertEqual(report["repaired"], ["孤儿名称"])
        self.assertEqual(report["remaining"], [], "修复后不应残留违例")
        self.assertIn("孤儿名称", db.get_reagent_names())

    def test_repaired_row_becomes_writable(self):
        """修复的意义在于：这条记录之后能正常出入库。"""
        db = self._make_legacy_db()
        db.foreign_key_report()

        ctx = StubCtx()
        ctx.db = db
        ctx.api = StubApi()
        ctx.ui = StubUi()
        ctx.version = "2.0.0"
        handlers = build_action_handlers(ctx)

        handlers["inbound.submit"]({
            "name": "孤儿名称", "batch": "B1", "quantity": 10, "operator": "",
            "productionDate": "2024-01-01", "expiryDate": "2026-01-01",
        })
        self.assertEqual(db.get_reagent_details("孤儿名称", "B1")[3], 15)

    def test_clean_database_reports_nothing(self):
        path = os.path.join(tempfile.mkdtemp(prefix="rms_clean_"), "clean.db")
        db = ReagentDatabase(path)
        self.assertEqual(db.foreign_key_report(), {"repaired": [], "remaining": []})

    def test_orphan_inbound_row_is_not_auto_repaired(self):
        """指向不存在试剂的行属于数据损坏，不能自作主张删除，只应报告出来。"""
        path = os.path.join(tempfile.mkdtemp(prefix="rms_orph_"), "orphan.db")
        db = ReagentDatabase(path)
        db.cursor.execute("PRAGMA foreign_keys = OFF")
        db.cursor.execute('''
            INSERT INTO inbound (reagent_id, reagent_name, reagent_batch, date,
                                 quantity, operator, gtin, remaining_quantity)
            VALUES (99999, '幽灵', 'B1', '2024-01-01 00:00:00:000', 1, '', NULL, 1)
        ''')
        db.conn.commit()
        db.cursor.execute("PRAGMA foreign_keys = ON")

        report = db.foreign_key_report()
        self.assertEqual(report["repaired"], [])
        self.assertTrue(report["remaining"],
                        "修不了的违例必须报出来，而不是静默放过")


class TestTimestampFormat(unittest.TestCase):
    def test_matches_legacy_format(self):
        """旧实现用 QDateTime.toString("yyyy-MM-dd HH:mm:ss:zzz")。"""
        ts = now_timestamp()
        self.assertEqual(len(ts), 23, ts)      # 2025-01-01 10:00:00:123
        self.assertEqual(ts[10], " ")
        self.assertEqual(ts[19], ":")
        self.assertTrue(ts[20:].isdigit())


if __name__ == "__main__":
    unittest.main(verbosity=2)
