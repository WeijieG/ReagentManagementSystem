# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""写入类 handler。

入库/出库的提交顺序与校验全部照搬旧实现（ReagentMGS.py:1682-1731 / 1906-1948），
包括「先查是否已有该批号，没有就先建一条数量为 0 的试剂记录」这一步。

每个写操作成功后都会推送 db.changed，前端据此刷新当前页 —— 这替代了旧实现里
对话框关闭后由父窗口回头调 refresh_widget() 的耦合写法。
"""

import datetime

from ..bridge import BridgeError, ErrorCode
from ..core import integrity, version as version_mod


def now_timestamp():
    """与旧实现一致的记录时间戳格式：yyyy-MM-dd HH:mm:ss:zzz

    旧代码用 QDateTime.toString("yyyy-MM-dd HH:mm:ss:zzz")，毫秒为 3 位。
    """
    now = datetime.datetime.now()
    return now.strftime("%Y-%m-%d %H:%M:%S:") + f"{now.microsecond // 1000:03d}"


def _require(payload, key, label):
    value = (payload.get(key) or "")
    if not isinstance(value, str) or not value.strip():
        raise BridgeError(ErrorCode.VALIDATION, f"请填写{label}")
    return value.strip()


def _positive_int(payload, key, label):
    raw = payload.get(key)
    try:
        value = int(raw)
    except (TypeError, ValueError):
        raise BridgeError(ErrorCode.VALIDATION, f"请输入有效的{label}（正整数）")
    if value <= 0:
        raise BridgeError(ErrorCode.VALIDATION, f"请输入有效的{label}（正整数）")
    return value


def build_action_handlers(ctx):
    db = ctx.db
    api = ctx.api

    def notify_changed(*tables):
        api.emit_event("db.changed", {"tables": list(tables)})

    # ------------------------------------------------------------------
    # 应用级
    # ------------------------------------------------------------------
    def style_set(payload):
        # 具体风格名单由前端的 STYLES 注册表维护，这里只校验标识格式
        style = (payload.get("style") or "light").strip()
        if not integrity.is_valid_ui_style(style):
            raise BridgeError(ErrorCode.VALIDATION, f"不支持的界面风格: {style}")
        integrity.set_ui_style(style)
        return {"style": style}

    def disclaimer_accept(payload):
        accepted = bool(payload.get("accepted", True))
        integrity.set_disclaimer_accepted(accepted)
        return {"accepted": accepted}

    def app_quit(payload):
        ctx.ui.quit_app()
        return {}

    # ------------------------------------------------------------------
    # 试剂名称
    # ------------------------------------------------------------------
    def names_add(payload):
        name = _require(payload, "name", "试剂名称")
        gtin = (payload.get("gtin") or "").strip() or None

        if not db.add_reagent_name(name, gtin):
            raise BridgeError(ErrorCode.DUPLICATE, f"试剂名称 '{name}' 或该 GTIN 已被占用")

        notify_changed("reagent_names")
        return {"name": name, "gtin": gtin or ""}

    def names_update(payload):
        old_name = _require(payload, "oldName", "当前名称")
        old_gtin = (payload.get("oldGtin") or "").strip() or None
        new_name = (payload.get("newName") or "").strip() or old_name
        new_gtin = (payload.get("newGtin") or "").strip() or old_gtin

        if not new_name and not new_gtin:
            raise BridgeError(ErrorCode.VALIDATION, "未提供任何修改")

        # 冲突检查与旧实现一致（ReagentMGS.py:1456-1467）
        if new_name != old_name:
            db.cursor.execute('SELECT COUNT(*) FROM reagent_names WHERE name = ?', (new_name,))
            if db.cursor.fetchone()[0] > 0:
                raise BridgeError(ErrorCode.DUPLICATE, f"试剂名称 '{new_name}' 已存在")

        if new_gtin and new_gtin != old_gtin:
            db.cursor.execute('SELECT COUNT(*) FROM reagent_names WHERE gtin = ?', (new_gtin,))
            if db.cursor.fetchone()[0] > 0:
                raise BridgeError(ErrorCode.DUPLICATE, f"GTIN '{new_gtin}' 已被使用")

        if not db.update_reagent_name_and_gtin(old_name, new_name, old_gtin, new_gtin):
            raise BridgeError(ErrorCode.DB, "更新失败，请检查数据库连接")

        notify_changed("reagent_names", "reagents", "inbound", "outbound")
        return {"name": new_name, "gtin": new_gtin or ""}

    def names_delete(payload):
        name = _require(payload, "name", "试剂名称")
        success, message = db.delete_reagent_name(name)
        if not success:
            # 旧实现把「有库存/有记录」当业务规则拒绝，不是异常
            raise BridgeError(ErrorCode.GUARDED, message)

        notify_changed("reagent_names")
        return {"name": name}

    # ------------------------------------------------------------------
    # 入库
    # ------------------------------------------------------------------
    def inbound_submit(payload):
        name = _require(payload, "name", "试剂名称")
        batch = _require(payload, "batch", "批号")
        quantity = _positive_int(payload, "quantity", "数量")
        operator = (payload.get("operator") or "").strip()
        gtin = (payload.get("gtin") or "").strip() or None
        qrcode = (payload.get("qrcode") or "").strip() or None
        production_date = _require(payload, "productionDate", "生产日期")
        expiry_date = _require(payload, "expiryDate", "有效期")

        # 与旧实现一致：已存在该 (名称, 批号) 就复用，否则先建记录
        detail = db.get_reagent_details(name, batch)
        if detail:
            reagent_id = detail[0]
        else:
            if not db.add_new_reagent(name, batch, production_date,
                                      expiry_date, 0, gtin, qrcode):
                raise BridgeError(
                    ErrorCode.DUPLICATE,
                    "添加新试剂失败，可能已存在相同名称和批号的试剂",
                )
            detail = db.get_reagent_details(name, batch)
            if not detail:
                raise BridgeError(ErrorCode.DB, "新建试剂后无法读取记录")
            reagent_id = detail[0]

        db.add_inbound(reagent_id, name, batch, now_timestamp(),
                       quantity, operator, gtin)

        notify_changed("reagents", "inbound", "reagent_names")
        return {"name": name, "batch": batch, "quantity": quantity}

    # ------------------------------------------------------------------
    # 出库
    # ------------------------------------------------------------------
    def outbound_submit(payload):
        name = _require(payload, "name", "试剂名称")
        batch = _require(payload, "batch", "批号")
        quantity = _positive_int(payload, "quantity", "数量")
        operator = (payload.get("operator") or "").strip()

        detail = db.get_reagent_details(name, batch)
        if not detail:
            raise BridgeError(ErrorCode.NOT_FOUND, "找不到指定的试剂")

        success, message = db.add_outbound(
            detail[0], name, batch, now_timestamp(), quantity, operator,
        )
        if not success:
            # 「库存不足」是业务拒绝，不是系统故障
            raise BridgeError(ErrorCode.STOCK_INSUFFICIENT, message)

        notify_changed("reagents", "outbound")
        return {"name": name, "batch": batch, "quantity": quantity}

    return {
        "style.set": style_set,
        "disclaimer.accept": disclaimer_accept,
        "app.quit": app_quit,
        "names.add": names_add,
        "names.update": names_update,
        "names.delete": names_delete,
        "inbound.submit": inbound_submit,
        "outbound.submit": outbound_submit,
    }
