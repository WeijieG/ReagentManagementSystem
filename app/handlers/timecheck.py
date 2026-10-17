# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""网络时间校准。

程序本身可以离线运行。顶栏时钟与出入库时间戳会优先使用公开站点的
时间；断网或请求失败时回退本机时间，并提示用户自行核对系统时钟。

取 HTTP 响应里的 Date 头，而不是某个具体的 JSON 时间 API：
  * 不需要密钥，也不绑定单一服务的长期可用性；
  * 主流站点的 Date 头由服务器生成，用来判断"系统时间是否明显跑偏"
    已经足够（秒级、不含时区歧义 —— Date 头固定是 GMT）。

**必须放在后台线程**：QWebChannel 的 invoke 跑在 Qt 主线程，同步等
网络超时会把整个窗口和通道一起冻住。
"""

import datetime
import email.utils

from PyQt5.QtCore import QThread, pyqtSignal

from ..core import clock

# 公开、无需密钥的站点；按国内网络可达性排序。
# 只读 Date 头，不解析页面内容。
TIME_SOURCES = (
    "https://www.baidu.com/",
    "https://www.taobao.com/",
    "https://www.cloudflare.com/",
)

# 单个来源的超时。断网时三个来源依次失败，最坏约 6 秒，全程在后台线程。
REQUEST_TIMEOUT = 2.0

# 本机时间与网络时间相差超过这个值就提示用户校正系统时钟
WARN_THRESHOLD_SECONDS = 120


def fetch_network_time():
    """尝试公开站点，返回 (UTC datetime, 来源 URL, 错误描述)。

    全部失败时 datetime 为 None。失败原因合并成一句，供界面提示排查。
    """
    try:
        import requests
    except Exception as exc:  # 目标机器可能缺网络组件
        return None, "", f"网络组件不可用：{type(exc).__name__}: {exc}"

    headers = {
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "User-Agent": "ReagentManagementSystem-TimeCheck",
    }

    errors = []
    for url in TIME_SOURCES:
        try:
            resp = requests.head(
                url, timeout=REQUEST_TIMEOUT, allow_redirects=True, headers=headers,
            )
            raw = resp.headers.get("Date", "")
            if not raw:
                raise ValueError("响应中没有 Date 头")
            parsed = email.utils.parsedate_to_datetime(raw)
            if parsed is None:
                raise ValueError(f"无法解析 Date 头: {raw!r}")
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=datetime.timezone.utc)
            return parsed.astimezone(datetime.timezone.utc), url, ""
        except Exception as exc:  # noqa: BLE001 —— 换下一个来源继续试
            host = url.split("//")[-1].rstrip("/")
            errors.append(f"{host}: {type(exc).__name__}: {exc}")

    return None, "", "；".join(errors)


class TimeCheckWorker(QThread):
    """后台访问公开站点并算出本机时间偏差。"""

    done = pyqtSignal(object)

    def run(self):
        network_dt, source, error = fetch_network_time()
        system_now = datetime.datetime.now().astimezone()

        if network_dt is not None:
            offset_seconds = int((network_dt - system_now).total_seconds())
        else:
            offset_seconds = 0

        self.done.emit({
            "networkOk": network_dt is not None,
            "systemTime": system_now.isoformat(timespec="seconds"),
            "systemEpochMs": int(system_now.timestamp() * 1000),
            "networkTime": (
                network_dt.astimezone().isoformat(timespec="seconds")
                if network_dt is not None else ""
            ),
            "networkEpochMs": int(network_dt.timestamp() * 1000) if network_dt else 0,
            "offsetSeconds": offset_seconds,
            "warnThresholdSeconds": WARN_THRESHOLD_SECONDS,
            "source": source,
            "error": error,
        })


def build_time_handlers(ctx):
    api = ctx.api
    state = {"worker": None}

    def time_check(payload):
        worker = state["worker"]
        if worker is not None and worker.isRunning():
            return {"started": False}

        worker = TimeCheckWorker()
        state["worker"] = worker

        def on_done(result):
            # 同步给写入层：出入库时间戳也从这里取校准后的时间
            if (result.get("networkOk") and result.get("networkEpochMs")
                    and result.get("systemEpochMs")):
                clock.set_network_time(
                    result["networkEpochMs"], result["systemEpochMs"],
                    result.get("source", ""),
                )
            else:
                clock.set_network_failure(result.get("error", ""))
            api.emit_event("time.synced", result)

        def on_finished():
            # 等线程真正结束再放开引用，避免 QThread 在运行中被回收
            state["worker"] = None

        worker.done.connect(on_done)
        worker.finished.connect(on_finished)
        worker.start()
        return {"started": True}

    return {
        "time.check": time_check,
    }
