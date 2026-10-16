# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""从 Qt 资源里导出 qwebchannel.js 到 app/ui/vendor/。

为什么必须落盘一份：`file://` 源的页面加载不了 `qrc://` 资源，
引用 `qrc:///qtwebchannel/qwebchannel.js` 会静默失败 —— 页面照常渲染，
但 QWebChannel 是 undefined，整个界面变成死的，且界面上没有任何报错。

PyQt5 不带这个 .js 文件，只能从 Qt 资源里抽。
运行：python tools/extract_qwebchannel.py
"""

import os
import sys

from PyQt5.QtCore import QCoreApplication, QFile, QIODevice

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGET = os.path.join(REPO, "app", "ui", "vendor", "qwebchannel.js")

RESOURCE_PATH = ":/qtwebchannel/qwebchannel.js"


def main():
    # QFile 访问 Qt 资源需要 QCoreApplication 存在；QtWebChannel 模块的
    # 导入会把对应资源注册进来。
    app = QCoreApplication(sys.argv)  # noqa: F841
    import PyQt5.QtWebChannel  # noqa: F401

    f = QFile(RESOURCE_PATH)
    if not f.exists() or not f.open(QIODevice.ReadOnly):
        print(f"无法打开 Qt 资源 {RESOURCE_PATH}")
        print("请确认已安装 PyQtWebEngine 与 PyQtWebEngine-Qt5。")
        return 1

    data = bytes(f.readAll())
    f.close()

    if not data:
        print("资源为空。")
        return 1

    os.makedirs(os.path.dirname(TARGET), exist_ok=True)
    with open(TARGET, "wb") as out:
        out.write(data)

    print(f"已导出 {len(data)} 字节 -> {TARGET}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
