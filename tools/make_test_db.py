# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""生成一份用于界面与功能测试的数据库。

用法：
    python tools/make_test_db.py [输出路径]

默认输出 testdata/ReagentTestData.db，重复运行会覆盖。
数据刻意覆盖：已过期 / 30 天内到期 / 90 天内到期 / 正常效期、
零库存、GTIN 绑定与未绑定、以及足够多的出入库记录（可翻页）。

生成完可以直接在「设置 → 导入数据库」里选用；原有数据库会自动备份。
"""

import datetime
import os
import random
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from app.core.db import ReagentDatabase  # noqa: E402

DEFAULT_OUT = os.path.join(ROOT, "testdata", "ReagentTestData.db")

NAMES = [
    ("无水乙醇", "6901234567890"),
    ("甲醇", "6901234567891"),
    ("盐酸", None),
    ("氢氧化钠", "6901234567892"),
    ("硫酸铜", None),
    ("氯化钠", "6901234567893"),
    ("冰醋酸", None),
    ("丙酮", "6901234567894"),
    ("碘化钾", None),
    ("硫酸", None),
    ("高锰酸钾", "6901234567895"),
    ("乙二胺四乙酸二钠", None),
]

# 只建名称、没有库存的名称，用来测试「试剂名称」页
EXTRA_NAMES = [
    ("硼酸", "6901234567896"),
    ("硝酸银", None),
]

OPERATORS = ["张伟", "李娜", "王强", "刘洋", "陈静"]

# 效期相对今天的天数：负数为已过期，覆盖 30 天、90 天与正常三档
EXPIRY_OFFSETS = [-400, -120, -35, -5, 4, 11, 22, 27, 48, 83, 160, 330, 720]


def ts(day, rng):
    """记录时间戳，格式与程序一致：yyyy-MM-dd HH:mm:ss:zzz"""
    return "%s %02d:%02d:%02d:000" % (
        day.isoformat(), rng.randint(8, 19), rng.randint(0, 59), rng.randint(0, 59))


def build(out_path):
    if os.path.exists(out_path):
        os.remove(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    rng = random.Random(20251017)
    db = ReagentDatabase(out_path)
    today = datetime.date.today()

    for name, gtin in EXTRA_NAMES:
        db.add_reagent_name(name, gtin)

    rows = 0
    for name_index, (name, gtin) in enumerate(NAMES):
        batch_count = 4 if name_index < 6 else 5
        for batch_index in range(batch_count):
            offset = EXPIRY_OFFSETS[(name_index * 4 + batch_index) % len(EXPIRY_OFFSETS)]
            # 轻微打散，避免大量同一天到期
            offset += rng.randint(-2, 2)
            expiry = today + datetime.timedelta(days=offset)
            shelf = (365, 540, 730)[(name_index + batch_index) % 3]
            production = expiry - datetime.timedelta(days=shelf)
            # 生产日期必须在过去，否则入库记录会落到未来日期
            if production > today - datetime.timedelta(days=30):
                production = today - datetime.timedelta(days=shelf)
            batch = "L%02d%02d%02d-%02d" % (
                production.year % 100, production.month, production.day, batch_index + 1)

            ok = db.add_new_reagent(
                name, batch, production.isoformat(), expiry.isoformat(), 0, gtin)
            if not ok:
                continue

            db.cursor.execute(
                "SELECT id FROM reagents WHERE name = ? AND batch = ?", (name, batch))
            reagent_id = db.cursor.fetchone()[0]

            # 第一次入库：日期落在生产之后、最近 300 天以内
            first_qty = rng.choice([20, 30, 50, 60, 80, 100, 120, 150, 200])
            earliest = max(production + datetime.timedelta(days=5),
                           today - datetime.timedelta(days=300))
            latest = max(earliest, today - datetime.timedelta(days=6))
            span = max((latest - earliest).days, 0)
            in_day = earliest + datetime.timedelta(days=rng.randint(0, span))
            db.add_inbound(reagent_id, name, batch, ts(in_day, rng),
                           first_qty, rng.choice(OPERATORS), gtin)

            total_in = first_qty
            last_day = in_day

            # 部分批号再补一次入库
            if rng.random() < 0.35:
                second_day = last_day + datetime.timedelta(
                    days=rng.randint(5, max(6, min(60, (today - last_day).days))))
                if second_day < today:
                    qty = rng.randint(10, 60)
                    db.add_inbound(reagent_id, name, batch, ts(second_day, rng),
                                   qty, rng.choice(OPERATORS), gtin)
                    total_in += qty
                    last_day = second_day

            # 出库：先算好数量，保证剩余不为负
            out_day = last_day + datetime.timedelta(
                days=rng.randint(3, max(4, min(45, (today - last_day).days + 1))))
            if out_day >= today:
                out_day = today
            if rng.random() < 0.65:
                qty = rng.randint(1, max(1, total_in // 2))
                db.add_outbound(reagent_id, name, batch, ts(out_day, rng),
                                qty, rng.choice(OPERATORS))

            # 最后一行做成零库存，用于测试「无库存」筛选与展示
            if name_index == len(NAMES) - 1 and batch_index == batch_count - 1:
                db.cursor.execute(
                    "SELECT quantity FROM reagents WHERE id = ?", (reagent_id,))
                remain = db.cursor.fetchone()[0]
                if remain > 0:
                    db.add_outbound(reagent_id, name, batch, ts(out_day, rng),
                                    remain, "张伟")

            rows += 1

    return db, rows


def main():
    out_path = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_OUT
    db, rows = build(out_path)

    db.cursor.execute("SELECT COUNT(*) FROM reagent_names")
    name_count = db.cursor.fetchone()[0]
    db.cursor.execute("SELECT COUNT(*) FROM inbound")
    inbound_count = db.cursor.fetchone()[0]
    db.cursor.execute("SELECT COUNT(*) FROM outbound")
    outbound_count = db.cursor.fetchone()[0]
    summary = db.dashboard_summary()
    fk_issues = db.check_foreign_keys()
    db.conn.close()

    print("测试数据库: %s" % out_path)
    print("试剂名称 %d 个 / 库存记录 %d 条 / 入库记录 %d 条 / 出库记录 %d 条"
          % (name_count, rows, inbound_count, outbound_count))
    print("在库品规 %d，库存总量 %d，已过期 %d，30 天内到期 %d"
          % (summary["skuCount"], summary["totalQuantity"],
             summary["expiredCount"], summary["expiringSoonCount"]))
    if fk_issues:
        print("警告：存在外键违例 %r" % (fk_issues,))
        return 1
    print("外键校验通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
