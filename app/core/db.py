# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""SQLite 数据访问层。

实现体照搬自 ReagentMGS.py 的 ReagentDatabase（原文件 129-460 行），行为逐条对齐，
未做"顺手优化"。唯一的差异是把打开失败时的 QMessageBox 换成异常，使本模块不依赖 PyQt5。
"""

import datetime
import os
import sqlite3

from .. import paths


class DatabaseError(Exception):
    """数据库无法打开。由调用方决定如何呈现给用户。"""


class ReagentDatabase:
    def __init__(self, db_file=None):
        self.db_file = db_file or paths.DB_FILE
        print(f"Database path: {self.db_file}")

        db_dir = os.path.dirname(self.db_file)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)

        try:
            self.conn = sqlite3.connect(self.db_file)
            self.cursor = self.conn.cursor()
            # 建表声明里一直写着 ON DELETE CASCADE，但从没开启过外键，
            # 那些级联其实是死条款。这里显式打开。
            # 注意：必须在任何事务之外设置，事务内设置会被忽略。
            self.cursor.execute("PRAGMA foreign_keys = ON")
            self.create_tables()
        except sqlite3.Error as e:
            raise DatabaseError(
                f"无法打开数据库文件:\n{self.db_file}\n错误: {str(e)}"
            ) from e

    def create_tables(self):
        # 创建试剂名称表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS reagent_names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                gtin TEXT UNIQUE
            )
        ''')

        # 创建试剂信息表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS reagents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                batch TEXT NOT NULL,
                production_date TEXT NOT NULL,
                expiry_date TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                gtin TEXT,
                qrcode TEXT,
                FOREIGN KEY (name) REFERENCES reagent_names(name) ON DELETE CASCADE,
                UNIQUE (name, batch)
            )
        ''')

        # 创建入库记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS inbound (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                reagent_id INTEGER NOT NULL,
                reagent_name TEXT NOT NULL,
                reagent_batch TEXT NOT NULL,
                date TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                operator TEXT,
                gtin TEXT,
                remaining_quantity INTEGER NOT NULL,
                FOREIGN KEY (reagent_id) REFERENCES reagents (id)
            )
        ''')

        # 创建出库记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS outbound (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                reagent_id INTEGER NOT NULL,
                reagent_name TEXT NOT NULL,
                reagent_batch TEXT NOT NULL,
                date TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                operator TEXT,
                gtin TEXT,
                remaining_quantity INTEGER NOT NULL,
                FOREIGN KEY (reagent_id) REFERENCES reagents (id)
            )
        ''')
        self.conn.commit()

    def check_foreign_keys(self):
        """返回当前库里的外键违例。每项为 (表名, 行id, 父表名, 约束序号)。

        注意：开启 `PRAGMA foreign_keys` 不会追溯校验既有数据，
        所以老库可能带着违例进来。这个方法用来把它们查出来。
        """
        self.cursor.execute("PRAGMA foreign_key_check")
        return self.cursor.fetchall()

    def repair_orphan_reagent_names(self):
        """把 reagents 里存在、但 reagent_names 里缺失的名称补进父表。

        这是老库最可能出现的一类不一致 —— 旧代码在 gtin 冲突时会漏建父行。
        补父行是纯增量操作，不修改也不删除任何既有数据。

        返回补入的名称列表。
        """
        self.cursor.execute('''
            SELECT DISTINCT name FROM reagents
            WHERE name NOT IN (SELECT name FROM reagent_names)
        ''')
        missing = [row[0] for row in self.cursor.fetchall()]
        if missing:
            self.cursor.executemany(
                'INSERT OR IGNORE INTO reagent_names (name) VALUES (?)',
                [(n,) for n in missing],
            )
            self.conn.commit()
        return missing

    def foreign_key_report(self):
        """开库时的一致性自检。

        返回 {'repaired': [...], 'remaining': [...]}。
        remaining 非空时应当提示用户，但不必阻止启动 —— 只读查询仍然正常，
        只有碰到具体那些行时才可能报错。
        """
        if not self.check_foreign_keys():
            return {"repaired": [], "remaining": []}

        repaired = self.repair_orphan_reagent_names()
        return {"repaired": repaired, "remaining": self.check_foreign_keys()}

    def get_reagents(self):
        self.cursor.execute(
            'SELECT id, name, batch, production_date, expiry_date, quantity FROM reagents'
        )
        return self.cursor.fetchall()

    def get_reagent_names(self):
        self.cursor.execute('SELECT name FROM reagent_names ORDER BY name')
        return [item[0] for item in self.cursor.fetchall()]

    def get_all_reagent_names(self):
        return self.get_reagent_names()

    def add_reagent_name(self, name, gtin=None):
        try:
            self.cursor.execute('''
                INSERT INTO reagent_names (name, gtin) VALUES (?, ?)
            ''', (name, gtin))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    def get_name_by_gtin(self, gtin):
        self.cursor.execute('''
            SELECT name
            FROM reagent_names
            WHERE gtin = ?
        ''', (gtin,))
        result = self.cursor.fetchone()
        return result[0] if result else None

    def delete_reagent_name(self, name):
        # 检查是否有库存记录
        self.cursor.execute('SELECT COUNT(*) FROM reagents WHERE name = ?', (name,))
        count = self.cursor.fetchone()[0]

        if count > 0:
            return False, "该试剂名称有库存记录，无法删除"

        # 检查是否有出入库记录
        self.cursor.execute('SELECT COUNT(*) FROM inbound WHERE reagent_name = ?', (name,))
        count += self.cursor.fetchone()[0]

        self.cursor.execute('SELECT COUNT(*) FROM outbound WHERE reagent_name = ?', (name,))
        count += self.cursor.fetchone()[0]

        if count > 0:
            return False, "该试剂名称有出入库记录，无法删除"

        self.cursor.execute('DELETE FROM reagent_names WHERE name = ?', (name,))
        self.conn.commit()
        return True, "删除成功"

    def get_batches_by_name(self, name):
        self.cursor.execute('SELECT batch FROM reagents WHERE name = ? ORDER BY batch', (name,))
        return [item[0] for item in self.cursor.fetchall()]

    def get_reagent_details(self, name, batch):
        self.cursor.execute('''
            SELECT id, production_date, expiry_date, quantity
            FROM reagents
            WHERE name = ? AND batch = ?
        ''', (name, batch))
        return self.cursor.fetchone()

    def get_name_by_batch(self, batch):
        self.cursor.execute('''
            SELECT name
            FROM reagents
            WHERE batch = ?
        ''', (batch,))
        result = self.cursor.fetchone()
        return result[0] if result else None

    def get_inbound_records(self, name_filter="", batch_filter="", start_date=None, end_date=None):
        query = '''
            SELECT id, reagent_name, reagent_batch, date, quantity, operator, remaining_quantity
            FROM inbound
            WHERE 1=1
        '''
        params = []

        if name_filter:
            query += " AND reagent_name LIKE ?"
            params.append(f"%{name_filter}%")

        if batch_filter:
            query += " AND reagent_batch LIKE ?"
            params.append(f"%{batch_filter}%")

        if start_date:
            query += " AND date >= ?"
            params.append(start_date)

        if end_date:
            query += " AND date <= ?"
            params.append(end_date)

        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def get_outbound_records(self, name_filter="", batch_filter="", start_date=None, end_date=None):
        query = '''
            SELECT id, reagent_name, reagent_batch, date, quantity, operator, remaining_quantity
            FROM outbound
            WHERE 1=1
        '''
        params = []

        if name_filter:
            query += " AND reagent_name LIKE ?"
            params.append(f"%{name_filter}%")

        if batch_filter:
            query += " AND reagent_batch LIKE ?"
            params.append(f"%{batch_filter}%")

        if start_date:
            query += " AND date >= ?"
            params.append(start_date)

        if end_date:
            query += " AND date <= ?"
            params.append(end_date)

        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def get_matching_reagent_names(self, prefix):
        self.cursor.execute('''
            SELECT name
            FROM reagent_names
            WHERE name LIKE ?
            ORDER BY name
        ''', (f"{prefix}%",))
        return [item[0] for item in self.cursor.fetchall()]

    def get_matching_batches(self, prefix):
        self.cursor.execute('''
            SELECT DISTINCT batch
            FROM reagents
            WHERE batch LIKE ?
            ORDER BY batch
        ''', (f"{prefix}%",))
        return [item[0] for item in self.cursor.fetchall()]

    def add_inbound(self, reagent_id, name, batch, date, quantity, operator, gtin=None):
        # 更新库存前先查询当前库存
        self.cursor.execute('SELECT quantity FROM reagents WHERE id = ?', (reagent_id,))
        current_quantity = self.cursor.fetchone()[0]

        remaining_quantity = current_quantity + quantity

        self.cursor.execute('''
            UPDATE reagents
            SET quantity = quantity + ?
            WHERE id = ?
        ''', (quantity, reagent_id))

        # 添加试剂名称到名称表（如果不存在）
        self.cursor.execute('''
            INSERT OR IGNORE INTO reagent_names (name) VALUES (?)
        ''', (name,))

        self.cursor.execute('''
            INSERT INTO inbound (reagent_id, reagent_name, reagent_batch, date, quantity, operator, gtin, remaining_quantity)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (reagent_id, name, batch, date, quantity, operator, gtin, remaining_quantity))
        self.conn.commit()

    def add_outbound(self, reagent_id, name, batch, date, quantity, operator):
        # 检查库存是否足够
        self.cursor.execute('SELECT quantity FROM reagents WHERE id = ?', (reagent_id,))
        current_quantity = self.cursor.fetchone()[0]

        if current_quantity < quantity:
            return False, "库存不足"

        remaining_quantity = current_quantity - quantity

        self.cursor.execute('''
            UPDATE reagents
            SET quantity = quantity - ?
            WHERE id = ?
        ''', (quantity, reagent_id))

        self.cursor.execute('''
            INSERT INTO outbound (reagent_id, reagent_name, reagent_batch, date, quantity, operator, gtin, remaining_quantity)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (reagent_id, name, batch, date, quantity, operator, None, remaining_quantity))
        self.conn.commit()
        return True, "出库成功"

    def add_new_reagent(self, name, batch, production_date, expiry_date, quantity, gtin=None, qrcode=None):
        """新增一条试剂（名称+批号）。

        ⚠️ 父表插入必须**只插名称**。
        原先写的是 `INSERT OR IGNORE INTO reagent_names (name, gtin)`：当 gtin
        已被别的名称占用时，整条 INSERT 因 gtin 的 UNIQUE 约束被忽略 ——
        名称行没建出来，紧接着插 reagents 时外键找不到父行而失败，用户看到的是
        「添加新试剂失败，可能已存在相同名称和批号的试剂」这种牛头不对马嘴的提示。
        （外键未开启时更糟：会静默留下一条孤立子行。）

        拆成两步后语义也更合理：名称一定要建出来，GTIN 只在没被占用时绑定。
        """
        try:
            self.cursor.execute(
                'INSERT OR IGNORE INTO reagent_names (name) VALUES (?)', (name,))

            # 仅在 gtin 空闲时绑定，且不覆盖已有的绑定。
            # 用 `UPDATE OR IGNORE`：若该 gtin 已被别的名称占用，SQLite 会
            # 跳过这一行而不是抛错 —— 正是这里想要的行为（不去抢别人的 gtin）。
            if gtin:
                self.cursor.execute('''
                    UPDATE OR IGNORE reagent_names SET gtin = ?
                    WHERE name = ? AND gtin IS NULL
                ''', (gtin, name))

            self.cursor.execute('''
                INSERT INTO reagents (name, batch, production_date, expiry_date, quantity, gtin, qrcode)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (name, batch, production_date, expiry_date, quantity, gtin, qrcode))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            self.conn.rollback()
            return False

    def update_reagent_name_and_gtin(self, old_name, new_name, old_gtin, new_gtin):
        """更新试剂名称和GTIN，并级联更新所有相关表。

        外键说明：`reagents.name` 引用 `reagent_names.name`，但声明里只有
        ON DELETE CASCADE、**没有 ON UPDATE CASCADE**。本方法先改父表再改子表，
        开启外键后第一步就会让所有子行暂时成为孤儿。

        因此这里用 `PRAGMA defer_foreign_keys = ON` 把校验推迟到 COMMIT ——
        届时父子表已经一致，检查自然通过。这个 PRAGMA 只在事务内有效，
        所以必须放在 BEGIN 之后。
        """
        try:
            self.conn.execute("BEGIN TRANSACTION")
            self.conn.execute("PRAGMA defer_foreign_keys = ON")

            self.cursor.execute('''
                UPDATE reagent_names
                SET name = ?, gtin = ?
                WHERE name = ? OR gtin = ?
            ''', (new_name, new_gtin, old_name, old_gtin))

            self.cursor.execute('''
                UPDATE reagents
                SET name = ?, gtin = ?
                WHERE name = ? OR gtin = ?
            ''', (new_name, new_gtin, old_name, old_gtin))

            self.cursor.execute('''
                UPDATE inbound
                SET reagent_name = ?, gtin = ?
                WHERE reagent_name = ? OR gtin = ?
            ''', (new_name, new_gtin, old_name, old_gtin))

            self.cursor.execute('''
                UPDATE outbound
                SET reagent_name = ?, gtin = ?
                WHERE reagent_name = ? OR gtin = ?
            ''', (new_name, new_gtin, old_name, old_gtin))

            self.conn.commit()
            return True
        except sqlite3.Error as e:
            self.conn.rollback()
            print(f"更新失败: {str(e)}")
            return False

    # ------------------------------------------------------------------
    # 以下为 Web UI 新增的查询辅助方法，均为对既有表的只读聚合，
    # 不改变任何写入行为。
    # ------------------------------------------------------------------

    def get_gtin_bindings(self):
        """返回 {试剂名称: GTIN}，未绑定的值为空字符串。"""
        self.cursor.execute('SELECT name, gtin FROM reagent_names')
        return {row[0]: (row[1] or "") for row in self.cursor.fetchall()}

    @staticmethod
    def days_until_expiry(expiry_date, today=None):
        """距有效期还有几天。

        与旧实现语义一致（ReagentMGS.py:3173）：无法解析的日期返回 -1，
        调用方据此判为「已过期」。
        """
        try:
            expiry = datetime.date.fromisoformat(expiry_date)
        except (ValueError, TypeError):
            return -1
        base = today or datetime.date.today()
        return (expiry - base).days

    def count_inventory(self, name_filter="", batch_filter="", expiry_filter="all"):
        """库存总数（用于分页）。

        沿用旧实现的口径：只统计 quantity > 0 的记录。
        """
        sql, params = self._inventory_where(name_filter, batch_filter, expiry_filter)
        self.cursor.execute(f"SELECT COUNT(*) FROM reagents WHERE 1=1{sql}", params)
        return self.cursor.fetchone()[0]

    # 允许排序的列白名单。前端传来的列名在这里查表，绝不直接拼进 SQL。
    INVENTORY_SORT_COLUMNS = {
        "id": "id",
        "name": "name",
        "batch": "batch",
        "productionDate": "production_date",
        "expiryDate": "expiry_date",
        "quantity": "quantity",
    }

    def list_inventory(self, name_filter="", batch_filter="", expiry_filter="all",
                       page=1, page_size=30, sort_col="id", sort_dir="asc"):
        """分页查询库存。返回 (rows, total)。

        rows 每项为 (id, name, batch, production_date, expiry_date, quantity, gtin)。
        """
        total = self.count_inventory(name_filter, batch_filter, expiry_filter)

        sql, params = self._inventory_where(name_filter, batch_filter, expiry_filter)

        column = self.INVENTORY_SORT_COLUMNS.get(sort_col, "id")
        direction = "DESC" if str(sort_dir).lower() == "desc" else "ASC"
        # 次级排序固定为 id，保证分页时行序稳定（否则翻页可能重复/漏行）
        order_by = f"{column} {direction}, id ASC"

        offset = max(0, (page - 1) * page_size)
        self.cursor.execute(
            f'''
                SELECT id, name, batch, production_date, expiry_date, quantity, gtin
                FROM reagents
                WHERE 1=1{sql}
                ORDER BY {order_by}
                LIMIT ? OFFSET ?
            ''',
            params + [page_size, offset],
        )
        return self.cursor.fetchall(), total

    def _inventory_where(self, name_filter, batch_filter, expiry_filter):
        """构造库存查询的 WHERE 片段。

        效期筛选在 SQL 层按 expiry_date 与本机当天比较，保持与旧 UI 的
        "剩余天数 = 有效期至 - 今天" 口径一致。
        """
        sql = " AND quantity > 0"
        params = []

        if name_filter:
            sql += " AND name LIKE ?"
            params.append(f"%{name_filter}%")

        if batch_filter:
            sql += " AND batch LIKE ?"
            params.append(f"%{batch_filter}%")

        # 使用 date('now','localtime') 保证按本地日期比较，与 QDate.currentDate() 一致
        if expiry_filter == "expired":
            sql += " AND date(expiry_date) < date('now','localtime')"
        elif expiry_filter == "soon":
            sql += (" AND date(expiry_date) >= date('now','localtime')"
                    " AND julianday(expiry_date) - julianday(date('now','localtime')) <= 30")
        elif expiry_filter == "near":
            sql += (" AND julianday(expiry_date) - julianday(date('now','localtime')) > 30"
                    " AND julianday(expiry_date) - julianday(date('now','localtime')) <= 90")
        elif expiry_filter == "valid":
            sql += " AND julianday(expiry_date) - julianday(date('now','localtime')) > 90"

        return sql, params

    def dashboard_summary(self, attention_days=30):
        """概览页数据。全部由既有表现算，不新增采集。"""
        today = "date('now','localtime')"

        self.cursor.execute(f'''
            SELECT COUNT(*), COALESCE(SUM(quantity), 0)
            FROM reagents
            WHERE quantity > 0
        ''')
        sku_count, total_quantity = self.cursor.fetchone()

        self.cursor.execute(f'''
            SELECT COUNT(*)
            FROM reagents
            WHERE quantity > 0 AND date(expiry_date) < {today}
        ''')
        expired = self.cursor.fetchone()[0]

        self.cursor.execute(f'''
            SELECT COUNT(*)
            FROM reagents
            WHERE quantity > 0
              AND date(expiry_date) >= {today}
              AND julianday(expiry_date) - julianday({today}) <= ?
        ''', (attention_days,))
        expiring_soon = self.cursor.fetchone()[0]

        self.cursor.execute(f'''
            SELECT name, batch, expiry_date, quantity,
                   CAST(julianday(expiry_date) - julianday({today}) AS INTEGER) AS days_left
            FROM reagents
            WHERE quantity > 0
              AND julianday(expiry_date) - julianday({today}) <= ?
            ORDER BY days_left ASC
        ''', (attention_days,))
        attention_rows = self.cursor.fetchall()

        self.cursor.execute(f'''
            SELECT '入库' AS kind, reagent_name, reagent_batch, date, quantity, operator
            FROM inbound
            UNION ALL
            SELECT '出库' AS kind, reagent_name, reagent_batch, date, quantity, operator
            FROM outbound
            ORDER BY date DESC
            LIMIT 8
        ''')
        recent = self.cursor.fetchall()

        return {
            "skuCount": sku_count,
            "totalQuantity": total_quantity,
            "expiredCount": expired,
            "expiringSoonCount": expiring_soon,
            "attention": [
                {
                    "name": r[0],
                    "batch": r[1],
                    "expiryDate": r[2],
                    "quantity": r[3],
                    "daysLeft": r[4],
                }
                for r in attention_rows
            ],
            "recent": [
                {
                    "kind": r[0],
                    "name": r[1],
                    "batch": r[2],
                    "date": r[3],
                    "quantity": r[4],
                    "operator": r[5] or "",
                }
                for r in recent
            ],
        }

    def count_records(self, record_type, name_filter="", batch_filter="",
                      start_date=None, end_date=None):
        """出入库记录总数，用于分页。"""
        tables = []
        if record_type in ("全部记录", "all", "入库记录", "inbound"):
            tables.append("inbound")
        if record_type in ("全部记录", "all", "出库记录", "outbound"):
            tables.append("outbound")

        total = 0
        for table in tables:
            sql = f"SELECT COUNT(*) FROM {table} WHERE 1=1"
            params = []
            if name_filter:
                sql += " AND reagent_name LIKE ?"
                params.append(f"%{name_filter}%")
            if batch_filter:
                sql += " AND reagent_batch LIKE ?"
                params.append(f"%{batch_filter}%")
            if start_date:
                sql += " AND date >= ?"
                params.append(start_date)
            if end_date:
                sql += " AND date <= ?"
                params.append(end_date)
            self.cursor.execute(sql, params)
            total += self.cursor.fetchone()[0]
        return total

    def list_records(self, record_type, name_filter="", batch_filter="",
                     start_date=None, end_date=None, page=1, page_size=30):
        """分页查询出入库记录。

        返回与旧 CombinedRecordPage 相同的行结构：
        (类型, id, 名称, 批号, 日期, 数量, 操作员, 剩余库存)。
        """
        clauses = []
        params = []
        if name_filter:
            clauses.append("reagent_name LIKE ?")
            params.append(f"%{name_filter}%")
        if batch_filter:
            clauses.append("reagent_batch LIKE ?")
            params.append(f"%{batch_filter}%")
        if start_date:
            clauses.append("date >= ?")
            params.append(start_date)
        if end_date:
            clauses.append("date <= ?")
            params.append(end_date)

        where = (" AND " + " AND ".join(clauses)) if clauses else ""

        selects = []
        select_params = []
        if record_type in ("全部记录", "all", "入库记录", "inbound"):
            selects.append(f"""
                SELECT '入库', id, reagent_name, reagent_batch, date, quantity, operator, remaining_quantity
                FROM inbound WHERE 1=1{where}
            """)
            select_params.extend(params)
        if record_type in ("全部记录", "all", "出库记录", "outbound"):
            selects.append(f"""
                SELECT '出库', id, reagent_name, reagent_batch, date, quantity, operator, remaining_quantity
                FROM outbound WHERE 1=1{where}
            """)
            select_params.extend(params)

        if not selects:
            return [], 0

        union = " UNION ALL ".join(selects)
        self.cursor.execute(f"SELECT COUNT(*) FROM ({union})", select_params)
        total = self.cursor.fetchone()[0]

        offset = max(0, (page - 1) * page_size)
        # 与旧 CombinedRecordPage 一致：仅按日期倒序
        self.cursor.execute(
            f"SELECT * FROM ({union}) ORDER BY 5 DESC LIMIT ? OFFSET ?",
            select_params + [page_size, offset],
        )
        return self.cursor.fetchall(), total
