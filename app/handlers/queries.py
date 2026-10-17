# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""只读类 handler。

统一约定：入参是 payload 字典，返回可 JSON 化的数据；需要报错时抛 BridgeError。
"""

from .. import paths
from ..bridge import BridgeError, ErrorCode
from ..core import gs1, integrity, text

# 每页条数上限。防止前端塞个巨大值让一次过桥传全表。
MAX_PAGE_SIZE = 500

# 效期阈值，与旧实现一致（ReagentMGS.py:3204-3207）
WARN30_DAYS = 30
WARN90_DAYS = 90


def _int(payload, key, default, minimum=None, maximum=None):
    """从 payload 取整数并做范围校验。"""
    raw = payload.get(key, default)
    try:
        value = int(raw)
    except (TypeError, ValueError):
        raise BridgeError(ErrorCode.VALIDATION, f"参数 {key} 必须是整数，收到 {raw!r}")
    if minimum is not None and value < minimum:
        raise BridgeError(ErrorCode.VALIDATION, f"参数 {key} 不能小于 {minimum}")
    if maximum is not None and value > maximum:
        raise BridgeError(ErrorCode.VALIDATION, f"参数 {key} 不能大于 {maximum}")
    return value


def _expiry_state(days_left):
    """把剩余天数映射成状态枚举，阈值与旧 UI 完全一致。"""
    if days_left < 0:
        return "expired"
    if days_left <= WARN30_DAYS:
        return "warn30"
    if days_left <= WARN90_DAYS:
        return "warn90"
    return "ok"


# 网络组件自检只做一次，结果缓存在这里
_NETWORK_CHECK = {"done": False, "ok": True, "error": ""}


def check_network_components():
    """检查网络组件是否可用。

    `requests` 已经改成按需导入，所以即使这里的 socket 坏了，主程序照样能用，
    只是"检查更新"会用不了。此时写一份诊断报告，用户可以直接发回来。
    """
    if _NETWORK_CHECK["done"]:
        return _NETWORK_CHECK["ok"], _NETWORK_CHECK["error"]

    _NETWORK_CHECK["done"] = True
    try:
        import socket  # noqa: F401
        import ssl      # noqa: F401
        _NETWORK_CHECK["ok"] = True
    except Exception as e:
        _NETWORK_CHECK["ok"] = False
        _NETWORK_CHECK["error"] = f"{type(e).__name__}: {e}"
        try:
            from ..core import diagnostics
            diagnostics.write_report(
                "网络组件（socket / ssl）不可用。\n"
                "影响范围：仅「检查更新」功能，库存与出入库不受影响。",
                diagnostics.collect_diagnostics(),
                filename="network-error.txt",
            )
        except Exception:
            pass
    return _NETWORK_CHECK["ok"], _NETWORK_CHECK["error"]


def build_query_handlers(ctx):
    db = ctx.db

    # ------------------------------------------------------------------
    # 启动引导
    # ------------------------------------------------------------------
    def app_bootstrap(payload):
        report = getattr(ctx, "db_report", None) or {}
        remaining = report.get("remaining") or []
        net_ok, net_err = check_network_components()
        return {
            "version": ctx.version,
            "dbPath": paths.DB_FILE,
            "disclaimerAccepted": integrity.is_disclaimer_accepted(),
            "theme": integrity.get_theme(),
            "style": integrity.get_ui_style(),
            "pageSizes": [20, 50, 100, 200],
            "maxPageSize": MAX_PAGE_SIZE,
            # 老库可能带着修不了的外键违例进来，只读可用但写入那些行会报错，
            # 必须让用户知道而不是等他自己撞上
            "dbRepaired": report.get("repaired") or [],
            "dbIssues": [
                {"table": row[0], "rowId": row[1], "parent": row[2]}
                for row in remaining
            ],
            # 网络组件不可用时，更新功能会失效，要在界面上说清楚
            "networkOk": net_ok,
            "networkError": net_err,
        }

    # ------------------------------------------------------------------
    # 概览
    # ------------------------------------------------------------------
    def dashboard_summary(payload):
        data = db.dashboard_summary(attention_days=WARN30_DAYS)
        for item in data["attention"]:
            item["state"] = _expiry_state(item["daysLeft"])
        return data

    # ------------------------------------------------------------------
    # 库存
    # ------------------------------------------------------------------
    def inventory_page(payload):
        filters = payload.get("filters") or {}
        sort = payload.get("sort") or {}
        page = _int(payload, "page", 1, minimum=1)
        page_size = _int(payload, "pageSize", 50, minimum=1, maximum=MAX_PAGE_SIZE)

        rows, total = db.list_inventory(
            name_filter=(filters.get("name") or "").strip(),
            batch_filter=(filters.get("batch") or "").strip(),
            expiry_filter=filters.get("expiry") or "all",
            page=page,
            page_size=page_size,
            sort_col=sort.get("col") or "id",
            sort_dir=sort.get("dir") or "asc",
        )

        items = []
        for row in rows:
            days_left = db.days_until_expiry(row[4])
            items.append({
                "id": row[0],
                "name": row[1],
                "batch": row[2],
                "productionDate": row[3],
                "expiryDate": row[4],
                "quantity": row[5],
                "gtin": row[6] or "",
                "daysLeft": days_left,
                "state": _expiry_state(days_left),
            })

        pages = max(1, (total + page_size - 1) // page_size)
        return {"rows": items, "total": total, "page": page,
                "pageSize": page_size, "pages": pages}

    # ------------------------------------------------------------------
    # 出入库记录
    # ------------------------------------------------------------------
    def records_page(payload):
        filters = payload.get("filters") or {}
        page = _int(payload, "page", 1, minimum=1)
        page_size = _int(payload, "pageSize", 50, minimum=1, maximum=MAX_PAGE_SIZE)

        record_type = filters.get("type") or "all"
        start = (filters.get("start") or "").strip() or None
        end = (filters.get("end") or "").strip() or None
        # 结束日期补足到当天末尾，才能覆盖带时间戳的记录（旧实现同样处理）
        if end:
            end = end + " 23:59:59:999"

        rows, total = db.list_records(
            record_type,
            name_filter=(filters.get("name") or "").strip(),
            batch_filter=(filters.get("batch") or "").strip(),
            start_date=start,
            end_date=end,
            page=page,
            page_size=page_size,
        )

        items = [{
            "kind": row[0],
            "id": row[1],
            "name": row[2],
            "batch": row[3],
            "date": row[4],
            "quantity": row[5],
            "operator": row[6] or "",
            "remaining": row[7],
        } for row in rows]

        pages = max(1, (total + page_size - 1) // page_size)
        return {"rows": items, "total": total, "page": page,
                "pageSize": page_size, "pages": pages}

    # ------------------------------------------------------------------
    # 试剂名称
    # ------------------------------------------------------------------
    def names_list(payload):
        names = db.get_reagent_names()
        bindings = db.get_gtin_bindings()
        return {
            "rows": [
                {"name": n, "gtin": bindings.get(n, "")} for n in names
            ]
        }

    def names_suggest(payload):
        prefix = (payload.get("prefix") or "").strip()
        limit = _int(payload, "limit", 30, minimum=1, maximum=200)
        names = db.get_reagent_names()
        if prefix:
            lowered = prefix.lower()
            names = [n for n in names if lowered in n.lower()]
        return {"items": names[:limit]}

    def batches_list(payload):
        name = (payload.get("name") or "").strip()
        if not name:
            return {"items": []}
        return {"items": db.get_batches_by_name(name)}

    def reagent_details(payload):
        name = (payload.get("name") or "").strip()
        batch = (payload.get("batch") or "").strip()
        if not name or not batch:
            return {"found": False}
        detail = db.get_reagent_details(name, batch)
        if not detail:
            return {"found": False}
        return {
            "found": True,
            "id": detail[0],
            "productionDate": detail[1],
            "expiryDate": detail[2],
            "quantity": detail[3],
        }

    # ------------------------------------------------------------------
    # GS1 解析
    # ------------------------------------------------------------------
    def gs1_parse(payload):
        """解析条码。

        解析器始终跑在 Python 侧，绝不在 JS 里重写一遍 —— 它那些边界行为
        （括号形式限定、定长截断告警键名、310x 小数插入）是数据正确性的依赖。
        """
        raw = payload.get("raw") or ""
        parsed = gs1.parse_barcode(raw)

        result = {
            "raw": raw,
            "fields": parsed,
            "gtin": parsed.get("GTIN", ""),
            "batch": parsed.get("Batch/Lot Number", ""),
            "productionDate": _expand_gs1_date(parsed.get("Production Date")),
            "expiryDate": _expand_gs1_date(parsed.get("Expiration Date")),
            "recognised": bool(parsed),
        }

        # 命中 GTIN 时顺带把绑定的试剂名带回去，省一次往返
        if result["gtin"]:
            result["name"] = db.get_name_by_gtin(result["gtin"]) or ""
        else:
            result["name"] = ""

        return result

    def _expand_gs1_date(yymmdd):
        """把 GS1 的 6 位日期（YYMMDD）补成 ISO 的 yyyy-MM-dd。

        沿用旧实现的世纪规则（ReagentMGS.py:1634）：
            century = "20" if int(year_part) < 100 else "19"
        int(两位数字) 恒小于 100，所以 "19" 分支不可达 —— 实际永远是 20xx。
        这里保留该行为，不擅自改成滚动窗口。
        """
        if not yymmdd or len(yymmdd) != 6 or not yymmdd.isdigit():
            return ""
        return f"20{yymmdd[0:2]}-{yymmdd[2:4]}-{yymmdd[4:6]}"

    # ------------------------------------------------------------------
    # 说明文本
    # ------------------------------------------------------------------
    def app_disclaimer(payload):
        return {"text": text.DISCLAIMER_TEXT}

    def about_text(payload):
        return {"text": text.ABOUT_TEXT}

    return {
        "app.bootstrap": app_bootstrap,
        "app.disclaimer": app_disclaimer,
        "about.text": about_text,
        "dashboard.summary": dashboard_summary,
        "inventory.page": inventory_page,
        "records.page": records_page,
        "names.list": names_list,
        "names.suggest": names_suggest,
        "batches.list": batches_list,
        "reagent.details": reagent_details,
        "gs1.parse": gs1_parse,
    }
