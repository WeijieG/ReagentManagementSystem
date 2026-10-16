# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""QtWebEngine + QWebChannel 环境探针。

用途：在目标机器上确认「内嵌 Chromium 能渲染」且「Python 与 JS 能双向通信」，
即方案阶段 0 的前置闸门。渲染结果截图到指定路径。

运行：
    python tools/qtwebengine_probe.py                 # 默认 probe_render.png
    python tools/qtwebengine_probe.py --out hw.png    # 指定输出路径

渲染模式由环境变量控制，需在启动前设置：
    QTWEBENGINE_CHROMIUM_FLAGS / QT_OPENGL
详见 docs/BUILD.md「渲染模式」一节。

这个探针与另外两个工具的分工：
    win7_diag.py    —— 查 DLL 能不能加载（程序起不来时用它）
    ui_smoke_test.py —— 驱动真实主窗口遍历各页面并截图
    本文件          —— 单独验证「Chromium 能渲染」+「Python/JS 能双向通信」，
                       把 QtWebEngine 从整个程序里摘出来测，便于定位渲染问题
"""

import json
import os
import sys

# --- 安装正确性自检 -------------------------------------------------------
# 必须使用 PyQtWebEngine（纯绑定）+ PyQtWebEngine-Qt5（Qt 运行时）这一组合，
# 它们会把 Qt5WebEngine*.dll 装进 PyQt5/Qt5/bin —— 与 PyQt5-Qt5 同一棵树。
# 单体轮子 PyQtWebEngine==5.15.2 会装到 PyQt5/Qt/bin，导致
# "DLL load failed while importing QtWebEngineWidgets"。
def _check_qt_layout():
    try:
        import PyQt5
    except ImportError:
        return
    pyqt_dir = os.path.dirname(PyQt5.__file__)
    same_tree = os.path.join(pyqt_dir, "Qt5", "bin", "QtWebEngineProcess.exe")
    stray_tree = os.path.join(pyqt_dir, "Qt", "bin", "Qt5WebEngineCore.dll")
    if os.path.exists(stray_tree):
        raise SystemExit(
            "依赖装错了：检测到 PyQt5/Qt/bin 下存在 WebEngine，说明装了单体轮子。\n"
            "请执行：pip uninstall PyQtWebEngine\n"
            "        pip install PyQtWebEngine==5.15.4 PyQtWebEngine-Qt5==5.15.2"
        )
    if not os.path.exists(same_tree):
        raise SystemExit(
            "未找到 QtWebEngineProcess.exe，PyQtWebEngine-Qt5 可能未安装。"
        )


_check_qt_layout()

# QtWebEngineWidgets 必须在创建 QApplication 之前导入
from PyQt5.QtCore import QObject, QUrl, pyqtSlot, QTimer  # noqa: E402
from PyQt5.QtWidgets import QApplication  # noqa: E402
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEnginePage  # noqa: E402
from PyQt5.QtWebChannel import QWebChannel  # noqa: E402

PROBE_HTML = """<!doctype html>
<html><head><meta charset="utf-8"><title>probe</title>
<style>
  html,body{margin:0;height:100%;font-family:"Microsoft YaHei",Segoe UI,sans-serif}
  body{display:flex;align-items:center;justify-content:center;
       background:linear-gradient(135deg,#1e293b,#2563eb);color:#fff}
  .card{background:rgba(255,255,255,.12);padding:28px 40px;border-radius:14px;
        text-align:center;min-width:320px}
  h1{font-size:20px;margin:0 0 14px}
  .row{font-size:13px;opacity:.9;margin:6px 0}
  button{margin-top:14px;padding:9px 20px;border:0;border-radius:8px;
         background:#fff;color:#1e293b;font-size:14px;font-weight:600;cursor:pointer}
  #result{margin-top:12px;font-weight:600;color:#a7f3d0;min-height:20px}
</style></head>
<body><div class="card">
  <h1>内嵌 Chromium 探针</h1>
  <div class="row">渲染引擎：<span id="ua"></span></div>
  <div class="row">栅格能力：<span id="css">检测中…</span></div>
  <button id="ping">调用 Python</button>
  <div id="result"></div>
</div>
<script src="qwebchannel.js"></script>
<script>
  document.getElementById('ua').textContent = navigator.userAgent.match(/Chrome\\/[\\d.]+/)[0];
  // 用 grid + gap 探测 CSS 能力（Chromium 83 支持 grid gap；须真正入 DOM 才有计算值）
  var probe = document.createElement('div');
  probe.style.cssText = 'display:grid;grid-template-columns:1fr 1fr;gap:4px';
  document.body.appendChild(probe);
  document.getElementById('css').textContent =
      getComputedStyle(probe).gap ? 'grid+gap 可用' : '受限';

  // qt.webChannelTransport 由 Qt 注入，时机晚于本脚本。必须轮询等待，
  // 否则会在构造 QWebChannel 时读到 undefined。
  var attempts = 0;
  function initChannel() {
    if (typeof qt === 'undefined' || !qt.webChannelTransport) {
      if (++attempts > 200) {
        document.getElementById('result').textContent = '通道未就绪（超时）';
        return;
      }
      return setTimeout(initChannel, 25);
    }
    new QWebChannel(qt.webChannelTransport, function (channel) {
      var api = channel.objects.probe;
      api.jsReady(navigator.userAgent, function (reply) {
        document.getElementById('result').textContent = 'Python 已握手: ' + reply;
      });
      document.getElementById('ping').onclick = function () {
        api.ping('按钮点击于 ' + new Date().toLocaleTimeString(), function (r) {
          document.getElementById('result').textContent = r;
        });
      };
    });
  }
  initChannel();
</script></body></html>
"""


class DebugPage(QWebEnginePage):
    """把页面里的 JS 报错与 console 输出打到标准输出。

    没有这一层，前端一旦抛异常就只表现为「白屏/无反应」，排查代价极高。
    """

    def javaScriptConsoleMessage(self, level, message, line, source):
        print(f"[JS] {source}:{line} {message}")

    def javaScriptAlert(self, frame, msg):
        print(f"[JS alert] {msg}")


class Probe(QObject):
    """暴露给 JS 的对象。所有方法通过 QWebChannel 调用。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.handshake_payload = None

    @pyqtSlot(str, result=str)
    def jsReady(self, user_agent):
        self.handshake_payload = user_agent
        return "OK"

    @pyqtSlot(str, result=str)
    def ping(self, message):
        return f"Python 收到: {message}"


def _parse_out_path(argv):
    """取 --out 指定的输出路径，默认放在脚本同目录。"""
    if "--out" in argv:
        idx = argv.index("--out")
        if idx + 1 < len(argv):
            return os.path.abspath(argv[idx + 1])
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "probe_render.png")


def _describe_render_mode():
    """把当前渲染模式相关的环境变量打印出来，便于对照结果。"""
    flags = os.environ.get("QTWEBENGINE_CHROMIUM_FLAGS", "")
    opengl = os.environ.get("QT_OPENGL", "")
    parts = []
    parts.append("禁用GPU" if "--disable-gpu" in flags else "启用GPU")
    parts.append(f"QT_OPENGL={opengl or '(默认)'}")
    return "，".join(parts)


def main():
    out_png = _parse_out_path(sys.argv)
    print(f"渲染模式: {_describe_render_mode()}")
    print(f"输出路径: {out_png}")

    app = QApplication(sys.argv)

    view = QWebEngineView()
    view.resize(560, 360)
    view.setPage(DebugPage(view))

    channel = QWebChannel()
    probe = Probe()
    channel.registerObject("probe", probe)
    view.page().setWebChannel(channel)

    # 关键：必须用 load() 从真实文件加载，不能用 setHtml。
    # setHtml 下 Qt 不会注入 qt.webChannelTransport，页面会报
    # "Cannot read property 'send' of undefined"。生产环境同样走 load()。
    page_dir = _prepare_page_dir()
    index_path = os.path.join(page_dir, "index.html")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(PROBE_HTML)

    view.load(QUrl.fromLocalFile(index_path))

    view.show()

    def capture():
        diag = (
            "JSON.stringify({"
            " qt: typeof qt,"
            " transport: (typeof qt !== 'undefined' && qt) ? typeof qt.webChannelTransport : 'n/a',"
            " send: (typeof qt !== 'undefined' && qt && qt.webChannelTransport)"
            "       ? typeof qt.webChannelTransport.send : 'n/a',"
            " QWebChannel: typeof window.QWebChannel"
            "})"
        )

        def on_diag(value):
            print(f"页面状态: {value}")
            pixmap = view.grab()
            ok = pixmap.save(out_png)
            print(f"截图: {out_png} ({'成功' if ok else '失败'})")
            print(f"尺寸: {pixmap.width()}x{pixmap.height()}")
            print(f"JS 握手: {'完成' if probe.handshake_payload else '未完成'}")
            if probe.handshake_payload:
                print(f"  UA: {probe.handshake_payload}")
            app.quit()

        view.page().runJavaScript(diag, on_diag)

    QTimer.singleShot(4000, capture)
    return app.exec_()


def _prepare_page_dir():
    """准备一个临时页面目录：抽出 qwebchannel.js（模拟随包内置）。"""
    import tempfile
    from PyQt5.QtCore import QFile, QIODevice

    page_dir = tempfile.mkdtemp(prefix="qwc_")
    target = os.path.join(page_dir, "qwebchannel.js")

    f = QFile(":/qtwebchannel/qwebchannel.js")
    if not f.exists() or not f.open(QIODevice.ReadOnly):
        print("警告: 未能从 Qt 资源取得 qwebchannel.js")
        return page_dir

    with open(target, "wb") as out:
        out.write(bytes(f.readAll()))
    f.close()
    return page_dir


if __name__ == "__main__":
    sys.exit(main())
