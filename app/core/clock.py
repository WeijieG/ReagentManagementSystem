# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""全局时间基准。

网络校时（handlers/timecheck.py）成功后，写入类操作的时间戳使用
「本机时间 + 偏差」，使出入库记录不会因为系统时钟跑偏而记错日期；
断网或校时失败时回退本机时间。

两条额外的保护：
  * 校时结果超过 MAX_AGE_SECONDS 不再使用（时钟组件每 30 分钟会重新校时）；
  * 检测到本机时钟在两次校时之间被手动改动（偏差超过 CLOCK_JUMP_MS），
    改为从上次网络时间按单调时钟推进，避免旧偏差或刚改的本机时间
    把记录带偏。
"""

import datetime
import threading
import time

_lock = threading.Lock()

_state = {
    "network_ok": False,
    "offset_ms": 0,
    "synced_network_ms": 0,
    "synced_system_ms": 0,
    "synced_monotonic": 0.0,
    "source": "",
    "error": "",
}

# 校时结果的有效期。时钟组件每 30 分钟重新校时，这里留足余量，
# 防止窗口最小化时定时器被节流导致短暂没有基准。
MAX_AGE_SECONDS = 2 * 60 * 60

# 本机时钟在两次校时之间被改动超过这个值，就认为旧偏差已不可信。
CLOCK_JUMP_MS = 5000


def set_network_time(network_epoch_ms, system_epoch_ms, source=""):
    """记录一次成功的网络校时结果。"""
    with _lock:
        _state.update({
            "network_ok": True,
            "offset_ms": int(network_epoch_ms) - int(system_epoch_ms),
            "synced_network_ms": int(network_epoch_ms),
            "synced_system_ms": int(system_epoch_ms),
            "synced_monotonic": time.monotonic(),
            "source": source or "",
            "error": "",
        })


def set_network_failure(error=""):
    """网络时间不可用：从现在起回退本机时间。"""
    with _lock:
        _state.update({
            "network_ok": False,
            "offset_ms": 0,
            "error": error or "",
        })


def offset_ms():
    """当前可用的时间偏差（毫秒）；没有有效网络时间时返回 0。"""
    now_ms = int(time.time() * 1000)
    mono = time.monotonic()

    with _lock:
        if not _state["network_ok"]:
            return 0

        elapsed = mono - _state["synced_monotonic"]
        if elapsed > MAX_AGE_SECONDS:
            return 0

        # 若本机时钟在这段时间内被手动改过，旧的偏差只会帮倒忙
        expected_ms = _state["synced_system_ms"] + elapsed * 1000
        if abs(now_ms - expected_ms) > CLOCK_JUMP_MS:
            return 0

        return _state["offset_ms"]


def current_datetime():
    """返回校准后的本地 naive datetime，口径与 datetime.now() 一致。

    正常情况用「本机时间 + 偏差」（对休眠最稳）；一旦发现本机时钟
    在两次校时之间被改动，就改为从上次网络时间按单调时钟推进，
    保证用户改完系统时间后立即入库也不会被带偏。
    """
    now_ts = time.time()
    now = datetime.datetime.fromtimestamp(now_ts)
    now_ms = int(now_ts * 1000)
    mono = time.monotonic()

    with _lock:
        if not _state["network_ok"]:
            return now

        elapsed = mono - _state["synced_monotonic"]
        if elapsed > MAX_AGE_SECONDS:
            return now

        expected_system_ms = _state["synced_system_ms"] + elapsed * 1000
        if abs(now_ms - expected_system_ms) <= CLOCK_JUMP_MS:
            return now + datetime.timedelta(milliseconds=_state["offset_ms"])

        # 本机时钟被手动改动：旧偏差不可信，但网络基准仍然有效
        network_ms = _state["synced_network_ms"] + elapsed * 1000

    return datetime.datetime.fromtimestamp(network_ms / 1000.0)


def state():
    """当前校时状态，便于诊断与测试。"""
    with _lock:
        return dict(_state)
