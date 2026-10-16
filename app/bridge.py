# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""QWebChannel 桥接层。

设计原则：**桥要薄**。不做"每个业务方法一个 @pyqtSlot"，而是只用两个通道：

  请求（JS → Python）：invoke(action, payload_json) -> envelope_json
  推送（Python → JS）：notify(json) 信号，靠 event 字段区分事件类型

所有跨语言数据统一走 JSON 字符串，避免 Qt 的 QVariant 类型塌陷。
大结果集一律服务端分页，不过桥传全表。

注意：本模块运行在 Qt 主线程。任何耗时操作（下载、导出）都必须交给
worker 线程，否则整个窗口与通道都会冻结。
"""

import json
import traceback

from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot


# 统一错误码。JS 侧只在这几个分支上做处理。
class ErrorCode:
    VALIDATION = "VALIDATION"
    NOT_FOUND = "NOT_FOUND"
    DUPLICATE = "DUPLICATE"
    STOCK_INSUFFICIENT = "STOCK_INSUFFICIENT"
    GUARDED = "GUARDED"
    IO = "IO"
    DB = "DB"
    NETWORK = "NETWORK"
    HASH_MISMATCH = "HASH_MISMATCH"
    CANCELLED = "CANCELLED"
    INTERNAL = "INTERNAL"


class BridgeError(Exception):
    """处理器可以抛这个把错误码与消息直接送达前端。"""

    def __init__(self, code, message, detail=""):
        super().__init__(message)
        self.code = code
        self.message = message
        self.detail = detail


def ok(data=None):
    return {"ok": True, "data": data if data is not None else {}}


def fail(code, message, detail=""):
    return {"ok": False, "error": {"code": code, "message": message, "detail": detail}}


def _jsonify(value):
    """把常见不可序列化的类型转成 JSON 友好的形式。"""
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, dict):
        return {str(k): _jsonify(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonify(v) for v in value]
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace")
    if hasattr(value, "isoformat"):
        return value.isoformat()
    if isinstance(value, (set, frozenset)):
        return [_jsonify(v) for v in value]
    return str(value)


def _map_exception(exc):
    """把内核抛出的异常映射成错误信封。"""
    if isinstance(exc, BridgeError):
        return fail(exc.code, exc.message, exc.detail)

    try:
        import sqlite3
        if isinstance(exc, sqlite3.Error):
            return fail(ErrorCode.DB, f"数据库操作失败: {exc}")
    except ImportError:
        pass

    if isinstance(exc, (OSError, IOError)):
        return fail(ErrorCode.IO, f"文件操作失败: {exc}")

    if isinstance(exc, (ValueError, TypeError, KeyError)):
        return fail(ErrorCode.VALIDATION, f"参数有误: {exc}")

    return fail(ErrorCode.INTERNAL, "内部错误", detail=traceback.format_exc())


class Api(QObject):
    """暴露给 JS 的单个对象。

    JS 侧通过 channel.objects.api 访问，使用方式：

        api.invoke('inventory.page', JSON.stringify({page: 1}), function (reply) {
            var res = JSON.parse(reply);
            if (res.ok) { ... } else { toast(res.error.message); }
        });
        api.notify.connect(function (raw) { var evt = JSON.parse(raw); ... });
    """

    # 单推送信号。事件类型放在 JSON 的 event 字段里。
    notify = pyqtSignal(str)

    def __init__(self, handlers, parent=None):
        super().__init__(parent)
        self._handlers = handlers

    # ------------------------------------------------------------------
    # JS → Python
    # ------------------------------------------------------------------
    @pyqtSlot(str, str, result=str)
    def invoke(self, action, payload_json):
        """同步分发一次请求。返回 JSON 信封字符串。"""
        try:
            payload = json.loads(payload_json) if payload_json else {}
            if not isinstance(payload, dict):
                payload = {"value": payload}
        except ValueError as e:
            return json.dumps(fail(ErrorCode.VALIDATION, f"请求参数不是合法 JSON: {e}"),
                              ensure_ascii=False)

        handler = self._handlers.get(action)
        if handler is None:
            return json.dumps(fail(ErrorCode.NOT_FOUND, f"未知操作: {action}"),
                              ensure_ascii=False)

        try:
            result = handler(payload)
            return json.dumps(ok(_jsonify(result)), ensure_ascii=False)
        except Exception as e:  # noqa: BLE001 —— 边界处必须兜住一切
            envelope = _map_exception(e)
            if envelope["error"]["code"] == ErrorCode.INTERNAL:
                traceback.print_exc()
            return json.dumps(envelope, ensure_ascii=False)

    # ------------------------------------------------------------------
    # Python → JS
    # ------------------------------------------------------------------
    def emit_event(self, event, data=None):
        """向页面推送一条事件。仅传 JSON 字符串。"""
        self.notify.emit(json.dumps(
            {"event": event, "data": _jsonify(data or {})}, ensure_ascii=False,
        ))

    def toast(self, kind, text):
        self.emit_event("toast", {"kind": kind, "text": text})


def build_api(handlers, parent=None):
    return Api(handlers, parent)
