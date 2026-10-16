# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""数据操作：备份、导入、导出 Excel。

流程照搬旧实现（ReagentMGS.py:2009-2103 的 backup/import，2105-2396 的导出），
差异只有两点，都是为了适应 Web UI：
  * 文件选择回到 Python 侧的 QFileDialog —— 页面在 file:// 源下拿不到原生路径
  * 导出成功不再弹窗，改为返回值由前端 toast

已知限制：这几个 handler 都在主线程同步执行。实验室量级的数据（几百到几千条）
耗时远小于一秒，所以没做后台任务。若将来数据量变大导致界面卡顿，应按桥的设计
改成 submit() 异步任务 + job.progress 推送。
"""

import os
import shutil
from datetime import datetime

from .. import paths
from ..bridge import BridgeError, ErrorCode
from ..core import excel, updater


def _timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def build_data_handlers(ctx):
    db = ctx.db
    ui = ctx.ui

    def data_backup(payload):
        default_name = f"ReagentWarehouseData_{_timestamp()}.db"
        target = ui.choose_save_file("备份数据库", default_name, "数据库文件 (*.db)")
        if not target:
            return {"cancelled": True}

        try:
            shutil.copyfile(paths.DB_FILE, target)
        except Exception as e:
            raise BridgeError(ErrorCode.IO, f"备份失败: {e}") from e

        return {"path": target}

    def data_import(payload):
        source = ui.choose_open_file("选择要导入的数据库", "数据库文件 (*.db)")
        if not source:
            return {"cancelled": True}

        if not os.path.exists(source):
            raise BridgeError(ErrorCode.IO, "所选文件不存在。")

        # 先做一次 SQLite 头校验。旧实现直接把文件覆盖上去，
        # 若选错了文件（比如选到一张图片），用户要等重启后才发现打不开。
        with open(source, "rb") as f:
            header = f.read(16)
        if not header.startswith(b"SQLite format 3"):
            raise BridgeError(ErrorCode.VALIDATION, "所选文件不是有效的 SQLite 数据库。")

        if not ui.confirm(
            "确认导入",
            "此操作将用所选文件替换当前数据库。\n"
            "原有数据库会先自动备份到 Backups 目录。\n\n确定继续吗？",
        ):
            return {"cancelled": True}

        # 备份现库，与旧实现一致放在 DB_DIR\Backups
        os.makedirs(paths.BACKUP_DIR, exist_ok=True)
        backup_path = os.path.join(
            paths.BACKUP_DIR, f"ReagentWarehouseData_backup_{_timestamp()}.db",
        )
        try:
            shutil.copyfile(paths.DB_FILE, backup_path)
            shutil.copyfile(source, paths.DB_FILE)
        except Exception as e:
            raise BridgeError(ErrorCode.IO, f"导入失败: {e}") from e

        return {"backupPath": backup_path, "needRestart": True}

    def data_restart(payload):
        """重启程序使导入生效。

        旧实现用 os.execl 就地重启；那里没关掉 sqlite 连接。这里先关闭连接
        再交给 updater.restart_application。
        """
        try:
            db.conn.close()
        except Exception:
            pass
        ui.quit_app()
        updater.restart_application()
        return {}

    def records_export(payload):
        start = (payload.get("start") or "").strip()
        end = (payload.get("end") or "").strip()
        if not start or not end:
            raise BridgeError(ErrorCode.VALIDATION, "导出需要指定日期范围。")

        default_name = excel.default_filename(start, end)
        target = ui.choose_save_file("导出出入库记录", default_name, "Excel 文件 (*.xlsx)")
        if not target:
            # 前端据此区分「用户取消」与「失败」，取消不该弹错误提示
            return {"cancelled": True}

        rows = excel.collect_export_rows(db, start, end)
        try:
            excel.write_workbook(rows, target)
        except PermissionError as e:
            raise BridgeError(
                ErrorCode.IO,
                "无法保存文件，请检查：\n"
                "1. 目标位置是否有写入权限\n"
                "2. 文件是否被其他程序（如 Excel）打开\n"
                "3. 是否选择了正确的保存位置",
            ) from e
        except Exception as e:
            raise BridgeError(ErrorCode.IO, f"导出数据时出错: {e}") from e

        return {"path": target, "rowCount": max(0, len(rows) - 1)}

    return {
        "data.backup": data_backup,
        "data.import": data_import,
        "data.restart": data_restart,
        "records.export": records_export,
    }
