# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""应用入口。

加载顺序是有讲究的，不要重排：
    1. 设置 QTWEBENGINE_* / QT_OPENGL 环境变量 —— 必须在导入 PyQt5 之前，
       事后设置会被忽略
    2. 设置 AA_ShareOpenGLContexts —— 必须在 QApplication 之前
    3. 导入 QtWebEngineWidgets —— 必须在 QApplication 之前
    4. 创建 QApplication

渲染模式失败的降级见 _resolve_render_mode()。
"""

import os
import sys

# 让 `python app/main.py` 与 `python -m app.main` 都能找到包
if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import paths  # noqa: E402
from app.core import integrity, updater  # noqa: E402

# ⚠️ 必须在导入 PyQt5 之前。Windows 7 缺 KB2533623 时 os.add_dll_directory
# 会抛 OSError(WinError 127)，而 PyQt5 只捕获 AttributeError，导致
# `import PyQt5` 整体失败、程序起不来。详见 app/compat.py。
from app import compat  # noqa: E402
compat.apply_all()


def _resolve_render_mode():
    """决定用硬件还是软件渲染，并据此设置环境变量。

    老 Win7 显卡（Intel HD 3000/4000 一代）在 InitializeGLOneOffPlatform 上
    可能死锁，表现为**纯白界面且不抛任何异常**。降级结果记在
    userData.ini [UI] RenderMode，这样只需失败一次就能永久避开。
    """
    mode = integrity.get_render_mode()  # 'hardware' | 'software'

    if mode == "software":
        os.environ.setdefault(
            "QTWEBENGINE_CHROMIUM_FLAGS",
            "--disable-gpu --disable-gpu-compositing "
            "--disable-direct-composition --disable-accelerated-2d-canvas",
        )
        # 让 Qt 走自带的 opengl32sw.dll（Mesa llvmpipe）
        os.environ.setdefault("QT_OPENGL", "software")
        os.environ.setdefault("QT_QUICK_BACKEND", "software")
    else:
        # 保持 Chromium 沙箱开启。只有从 UNC 路径运行时 Chromium 才拒绝沙箱，
        # 那种情况由下面单独处理 —— 不要为了方便默认关掉它。
        if paths.APPLICATION_PATH.startswith("\\\\"):
            os.environ["QTWEBENGINE_DISABLE_SANDBOX"] = "1"
            print("检测到从网络路径运行，已关闭 Chromium 沙箱。")

    return mode


_RENDER_MODE = _resolve_render_mode()

from PyQt5 import QtCore  # noqa: E402
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_ShareOpenGLContexts)

from PyQt5.QtCore import QUrl  # noqa: E402
from PyQt5.QtGui import QIcon  # noqa: E402
from PyQt5.QtWebEngineWidgets import (  # noqa: E402
    QWebEngineView, QWebEnginePage, QWebEngineSettings,
)
from PyQt5.QtWebChannel import QWebChannel  # noqa: E402
from PyQt5.QtWidgets import (  # noqa: E402
    QApplication, QMainWindow, QFileDialog, QMessageBox,
)

from app.bridge import build_api  # noqa: E402
from app.handlers import Context, UiHost, build_registry  # noqa: E402

APP_TITLE = "试剂库存管理系统"


class DebuggablePage(QWebEnginePage):
    """把页面里的 JS 报错打到标准输出。

    没有这一层，前端异常就只表现为白屏，排查代价极高。
    仅在 RMS_DEBUG=1 时启用，正常运行时输出会被丢弃。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._verbose = os.environ.get("RMS_DEBUG") == "1"

    def javaScriptConsoleMessage(self, level, message, line, source):
        if self._verbose:
            print(f"[JS] {source}:{line} {message}")


class MainWindow(QMainWindow, UiHost):
    """承载 Web UI 的主窗口，同时实现 UiHost 提供的原生交互。"""

    def __init__(self, db, db_report=None):
        super().__init__()
        self.db = db
        self.db_report = db_report or {"repaired": [], "remaining": []}

        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 760)
        self.setMinimumSize(940, 620)
        # 打包后没有这个会显示 Python 的默认图标
        if os.path.exists(paths.ICON_PATH):
            self.setWindowIcon(QIcon(paths.ICON_PATH))

        self.view = QWebEngineView(self)
        self.view.setPage(DebuggablePage(self.view))
        self.setCentralWidget(self.view)

        self._configure_web_settings()

        # 先建 api 壳，再用它自己组装 handler —— handler 需要靠 api 推送事件。
        # api / channel 挂在 self 上是刻意的：被 GC 掉会表现为 JS 侧
        # 「unknown object」，且极难排查。
        version = updater.get_current_version()
        self.api = build_api({}, self)
        self.api._handlers = build_registry(Context(
            db=self.db, api=self.api, ui=self, version=version,
            db_report=self.db_report,
        ))

        # 通道必须在 load() 之前挂上，否则页面拿不到 qt.webChannelTransport
        self.channel = QWebChannel(self.view.page())
        self.channel.registerObject("api", self.api)
        self.view.page().setWebChannel(self.channel)

        self._load_ui()

    def _configure_web_settings(self):
        s = self.view.settings()
        # 页面与静态资源都在 file:// 下，必须允许访问本地文件；
        # 但绝不允许访问远程 URL —— 本应用完全离线。
        s.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        s.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, False)
        s.setAttribute(QWebEngineSettings.JavascriptCanOpenWindows, False)
        s.setAttribute(QWebEngineSettings.PluginsEnabled, False)
        s.setAttribute(QWebEngineSettings.FullScreenSupportEnabled, False)
        s.setAttribute(QWebEngineSettings.HyperlinkAuditingEnabled, False)
        s.setAttribute(QWebEngineSettings.ScrollAnimatorEnabled, False)
        s.setAttribute(QWebEngineSettings.ErrorPageEnabled, False)
        s.setAttribute(QWebEngineSettings.WebGLEnabled, False)
        # 远程调试不走 QWebEngineSettings：设置 QTWEBENGINE_REMOTE_DEBUGGING=9222
        # 后用 Chrome/Edge 打开 http://127.0.0.1:9222 即可。只在排查时开。

    def _load_ui(self):
        """从真实文件加载界面。

        不能用 setHtml：那样 Qt 不会注入 qt.webChannelTransport，页面会报
        "Cannot read property 'send' of undefined"，整个界面是死的。
        """
        index = os.path.join(paths.get_ui_dir(), "index.html")
        if not os.path.exists(index):
            QMessageBox.critical(self, "界面文件缺失",
                                 f"找不到界面文件：\n{index}")
            return
        self.view.load(QUrl.fromLocalFile(index))

    # ------------------------------------------------------------------
    # UiHost 实现 —— Web 层做不了的原生交互
    # ------------------------------------------------------------------
    def choose_save_file(self, title, default_name, file_filter):
        path, _ = QFileDialog.getSaveFileName(self, title, default_name, file_filter)
        return path or ""

    def choose_open_file(self, title, file_filter):
        path, _ = QFileDialog.getOpenFileName(self, title, "", file_filter)
        return path or ""

    def confirm(self, title, text):
        reply = QMessageBox.question(
            self, title, text, QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        return reply == QMessageBox.Yes

    def warn(self, title, text):
        QMessageBox.warning(self, title, text)

    def info(self, title, text):
        QMessageBox.information(self, title, text)

    def quit_app(self):
        QApplication.quit()


def main():
    # 完整性自校验。开发环境（未打包）下内部会直接放行。
    ok, message = integrity.check_file_integrity()
    if not ok:
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "应用不完整！", message)
        return 1

    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    if os.path.exists(paths.ICON_PATH):
        app.setWindowIcon(QIcon(paths.ICON_PATH))

    try:
        from app.core.db import ReagentDatabase
        db = ReagentDatabase()
    except Exception as e:
        QMessageBox.critical(None, "数据库错误", str(e))
        return 1

    # 开库一致性自检。历史库是在外键关闭的状态下写出来的，可能带着孤立子行，
    # 那些行一旦被写就会报 FOREIGN KEY constraint failed。
    fk_report = db.foreign_key_report()
    if fk_report["repaired"]:
        print(f"已修复 {len(fk_report['repaired'])} 个孤立试剂名称："
              f"{'、'.join(fk_report['repaired'][:5])}")
    if fk_report["remaining"]:
        print(f"警告: 仍有 {len(fk_report['remaining'])} 条外键违例，"
              f"涉及的表：{sorted({v[0] for v in fk_report['remaining']})}")

    # 上次更新残留的安装包，启动时清掉
    updater.cleanup_old_installer()

    window = MainWindow(db, db_report=fk_report)
    window.show()

    # 打包自检：设 RMS_SMOKE_SHOT=<图片路径> 后，启动完成即截图并退出。
    # 打包后的 exe 没有控制台，这是唯一能在目标机上确认「界面真的渲染出来了」
    # 而不只是「进程没崩」的手段。
    shot_path = os.environ.get("RMS_SMOKE_SHOT")
    if shot_path:
        from PyQt5.QtCore import QTimer

        delay = int(os.environ.get("RMS_SMOKE_DELAY_MS", "12000"))

        def _capture_and_quit():
            pixmap = window.grab()
            ok = pixmap.save(shot_path)
            print(f"自检截图: {shot_path} ({'成功' if ok else '失败'}) "
                  f"{pixmap.width()}x{pixmap.height()}")
            app.quit()

        QTimer.singleShot(delay, _capture_and_quit)

    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
