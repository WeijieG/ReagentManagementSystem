# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""检查更新与下载安装。

检查是同步的（timeout=5，可接受）；下载必须放到后台线程，否则会冻结整个界面
—— QWebChannel 的调用跑在主线程上，主线程一卡，通道也跟着卡。

进度与结果通过 notify 信号推送给前端。
"""

import os
import tempfile

from PyQt5.QtCore import QThread, pyqtSignal

from .. import paths
from ..bridge import BridgeError, ErrorCode
from ..core import integrity, updater


class DownloadWorker(QThread):
    """后台下载更新包。"""

    progress = pyqtSignal(int)
    done = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, url, dest, md5=None, pinned="", parent=None):
        super().__init__(parent)
        self._url = url
        self._dest = dest
        self._md5 = md5
        self._pinned = pinned
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        try:
            updater.download_update(
                self._url,
                self._dest,
                md5=self._md5,
                progress_cb=self.progress.emit,
                pinned_fingerprint=self._pinned,
                cancel_check=lambda: self._cancelled,
            )
            self.done.emit(self._dest)
        except updater.UpdateError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit(f"下载失败: {e}")


def build_update_handlers(ctx):
    ui = ctx.ui
    api = ctx.api
    state = {"worker": None}

    def _on_progress(percent):
        api.emit_event("update.progress", {"percent": percent})

    def update_check(payload):
        ignored = integrity.get_ignored_version()
        info = updater.check_for_update(ignored_version=ignored)
        if not info:
            return {"available": False}
        return {"available": True, "info": info}

    def update_ignore(payload):
        version = (payload.get("version") or "").strip()
        if version:
            integrity.set_ignored_version(version)
        return {"version": version}

    def update_start(payload):
        # 重新查一次，避免前端传来被篡改的下载地址
        ignored = integrity.get_ignored_version()
        info = updater.check_for_update(ignored_version=ignored)
        if not info:
            raise BridgeError(ErrorCode.NOT_FOUND, "当前没有可用的更新。")

        url = info.get("url")
        if not url:
            raise BridgeError(ErrorCode.VALIDATION, "更新信息里缺少下载地址。")

        if state["worker"] and state["worker"].isRunning():
            raise BridgeError(ErrorCode.VALIDATION, "已有下载任务在进行中。")

        dest = os.path.join(tempfile.gettempdir(), os.path.basename(url) or "update.exe")
        pinned = updater.get_pinned_fingerprint()

        worker = DownloadWorker(url, dest, md5=info.get("md5"), pinned=pinned)
        state["worker"] = worker

        def on_done(path):
            api.emit_event("update.progress", {"percent": 100})
            api.emit_event("update.installing", {})
            try:
                updater.install_update(path)
            except updater.UpdateError as e:
                api.emit_event("update.error", {"message": str(e)})
                return
            # 立即硬退出。用 QApplication.quit() 的话，事件循环退出后
            # QtWebEngineProcess.exe 仍会占着 Qt5WebEngineCore.dll，
            # 静默安装会因文件被锁而半途失败，留下混版安装。
            os._exit(0)

        def on_failed(message):
            api.emit_event("update.error", {"message": message})

        worker.progress.connect(_on_progress)
        worker.done.connect(on_done)
        worker.failed.connect(on_failed)
        worker.start()

        return {"started": True}

    return {
        "update.check": update_check,
        "update.ignore": update_ignore,
        "update.start": update_start,
    }
