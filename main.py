# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""打包入口 —— 带启动失败自诊断。

PyInstaller 对「包内部的脚本」处理起来容易出岔子（__package__ 为 None 时
相对导入会失败），所以单独放一个顶层入口。

除此之外这里还有一件事：**把启动阶段的任何失败落成一个报告文件**。
打包后的程序没有控制台，一旦启动就崩，用户看到的只是「闪一下」，
什么线索都没有。这里的处理会在程序目录里写出 startup-error.txt。
"""

import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _compat_state():
    """尽量拿到兼容层的状态。

    诊断模块导入失败时这是唯一能解释「QtCore 为什么起不来」的线索 ——
    compat 在 PyQt5 之前就跑过了，所以即使 PyQt5 挂了它也在 sys.modules 里。
    """
    try:
        from app import compat
        return compat.preload_summary()
    except BaseException as e:
        return f"<取不到兼容层状态: {type(e).__name__}: {e}>"


def _fallback_report(exc_text):
    """连诊断模块都导入不了时的兜底。

    ⚠️ 必须挨个试**多个**可写位置。装在 Program Files 时程序目录对普通用户
    是只读的，以前只写程序目录、失败就返回空串 —— 结果是崩了却什么都没留下，
    只剩一个用户看不懂的堆栈。
    """
    try:
        app_dir = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
                   else os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        app_dir = ""

    candidates = [app_dir,
                  os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
                  os.path.join(os.environ.get("APPDATA", ""), "..", "Desktop")]

    text = ("=" * 64 + "\n ReagentManagementSystem — 启动失败\n" + "=" * 64
            + "\n\n--- 兼容层状态 ---\n" + _compat_state()
            + "\n\n--- 完整堆栈 ---\n" + exc_text + "\n")

    for base in candidates:
        if not base or not os.path.isdir(base):
            continue
        try:
            path = os.path.join(base, "startup-error.txt")
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            return path
        except Exception:
            continue
    return ""


def _hint_for(exc_text):
    """把几个「报错信息完全指错方向」的已知情况翻译成人话。

    最典型的是干净 Win7 SP1 缺 KB2533623：真实原因是系统缺一个 2011 年的
    更新，但报出来是 `DLL load failed while importing _ctypes`，
    看着像打包缺文件，于是会去反复重打包、重装运行库。见 docs/BUILD.md。
    """
    if "importing _ctypes" in exc_text or "importing QtCore" in exc_text \
            or "importing _socket" in exc_text:
        return ("\n【提示】这多半不是打包问题，而是系统缺 Windows 7 更新 "
                "KB2533623（或其替代更新 KB3063858）。\n"
                "Python 3.8 没有这个更新就加载不了任何扩展模块，症状正是 "
                "「DLL load failed while importing ...」。\n"
                "装上并重启后再试。\n")
    return ""


def _handle_failure(header, exc_text):
    """写报告 + 弹提示，返回报告路径。"""
    hint = _hint_for(exc_text)

    try:
        from app.core.diagnostics import (
            write_report, show_error_box, summary_for_dialog,
        )
    except BaseException:
        path = _fallback_report(exc_text)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                None, f"程序启动失败。\n\n报告: {path or '(写入失败)'}",
                "ReagentManagementSystem", 0x10)
        except Exception:
            pass
        return path

    path = write_report(header + hint, exc_text)

    # 把关键信息直接放进弹窗 —— 实测从弹窗里复制多行文本很难做成，
    # 而截一张图很容易。所以让弹窗自身就带够信息。
    try:
        summary = summary_for_dialog(exc_text)
    except Exception:
        summary = ""

    show_error_box(
        "ReagentManagementSystem — 启动失败",
        "程序启动失败。请把本窗口截图发回。\n"
        "（完整报告也写在桌面和程序目录下，文件名 startup-error.txt）\n"
        + hint + "\n" + summary,
    )
    return path


try:
    from app.main import main
except BaseException:
    _path = _handle_failure(
        "失败发生在**导入阶段**（还没进到主程序）", traceback.format_exc())
    sys.stderr.write(f"\n启动失败，报告已写入: {_path}\n")
    raise


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        _path = _handle_failure("失败发生在**运行阶段**", traceback.format_exc())
        raise
