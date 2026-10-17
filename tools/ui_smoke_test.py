# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""Web UI 冒烟测试。

用一份带数据的临时库启动真实主窗口，依次访问每个路由页并检查：
  * 页面是否渲染（有没有掉进致命错误面板）
  * 已知的关键元素是否出现（表格、卡片、徽章等）
  * 桥是否握手、主题是否正确

最后把指定页面截图。不需要人工点击，可直接在 Win7 虚拟机上跑。

运行：
    python tools/ui_smoke_test.py                  # 明亮风格，截图库存页
    python tools/ui_smoke_test.py --style night
    python tools/ui_smoke_test.py --style vivid
    python tools/ui_smoke_test.py --shot dashboard

--theme 仅在 --style default（兼容旧用法）时用于选明亮/暗夜；其余皮肤主题固定。
为兼容旧用法，--style 仍接受 light/dark/lively/minimal/ios/simple/default。
"""

import datetime
import json
import os
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)

OUT_PNG = os.path.join(REPO, "tools", "ui_smoke.png")

# 每个路由要检查的关键选择器。空字符串表示只要求页面渲染出来即可。
ROUTE_CHECKS = [
    ("dashboard", ".stat-card", 4, "概览"),
    ("inventory", ".table tbody tr", 1, "库存管理"),
    ("records",   ".table tbody tr", 1, "出入库记录"),
    ("names",     ".table tbody tr", 1, "试剂名称"),
    ("settings",  ".setting-row", 8, "设置"),
]


def build_fixture_db():
    """造一份覆盖三种效期状态的数据，让各页面都有内容可看。"""
    from app.core.db import ReagentDatabase

    path = os.path.join(tempfile.mkdtemp(prefix="rms_smoke_"), "smoke.db")
    db = ReagentDatabase(path)

    today = datetime.date.today()

    def iso(days):
        return (today + datetime.timedelta(days=days)).isoformat()

    fixtures = [
        ("生理盐水",   "A2409", iso(-400), iso(-10), 120, 30, "06936415919665"),
        ("质控液",     "B2411", iso(-200), iso(20),   45,  0, "06936415919666"),
        ("葡萄糖试剂", "C2501", iso(-120), iso(75),   60,  5, ""),
        ("校准品",     "D2502", iso(-60),  iso(300),  80,  0, ""),
    ]

    for name, batch, prod, exp, qty_in, qty_out, gtin in fixtures:
        db.add_new_reagent(name, batch, prod, exp, 0, gtin or None, None)
        rid = db.get_reagent_details(name, batch)[0]
        db.add_inbound(rid, name, batch, f"{prod} 09:00:00:000", qty_in, "测试员", gtin or None)
        if qty_out:
            db.add_outbound(rid, name, batch, f"{exp} 09:00:00:000", qty_out, "测试员")

    return db


def collect_state_script(selector):
    """生成一段读取页面状态的 JS。"""
    return (
        "(function () {"
        "  var fatal = document.getElementById('fatal-panel');"
        "  var sel = " + json.dumps(selector) + ";"
        "  return JSON.stringify({"
        "    fatalVisible: !!(fatal && fatal.style.display !== 'none'),"
        "    fatalText: fatal ? document.getElementById('fatal-message').textContent : '',"
        "    matches: sel ? document.querySelectorAll(sel).length : 0,"
        "    emptyState: document.querySelectorAll('.empty').length,"
        "    navActive: (document.querySelector('.nav-item.is-active') || {}).textContent || '',"
        "    theme: document.documentElement.getAttribute('data-theme'),"
        "    style: document.documentElement.getAttribute('data-style')"
        "  });"
        "}())"
    )


def main():
    theme = "light"
    style = "default"
    shot_route = "inventory"
    for flag, setter in (("--theme", "theme"), ("--style", "style"), ("--shot", "shot_route")):
        if flag in sys.argv:
            idx = sys.argv.index(flag)
            if idx + 1 < len(sys.argv):
                if setter == "theme":
                    theme = sys.argv[idx + 1]
                elif setter == "style":
                    style = sys.argv[idx + 1]
                else:
                    shot_route = sys.argv[idx + 1]

    from app.core import integrity

    # --style 传 data-style 标识，每个皮肤的主题是固定的。
    # 旧标识仍可用，先统一映射到新标识；已移除的皮肤回落到明亮。
    legacy = {
        "default": None, "light": "bright", "dark": "night",
        "lively": "vivid",
        "minimal": "bright", "ios": "bright", "ios-bright": "bright",
        "ios-dark": "bright", "ios-night": "bright",
        "simple": "bright", "simple-bright": "bright", "simple-night": "bright",
    }
    if style in legacy:
        style = legacy[style] or ("night" if theme == "dark" else "bright")

    expected_style = style
    expected_theme = {
        "bright": "light", "night": "dark",
        "tech": "dark", "vivid": "light", "soft": "light",
        "depth": "light", "glass": "dark", "paper": "light",
    }.get(style, theme)
    preference = style

    integrity.set_ui_style(preference)
    # 免责声明是会挡住整个界面的闸门，测试里先放行，否则截到的都是它
    integrity.set_disclaimer_accepted(True)

    from PyQt5.QtCore import QTimer
    from PyQt5.QtWidgets import QApplication

    from app.main import MainWindow

    app = QApplication(sys.argv)
    db = build_fixture_db()
    window = MainWindow(db)
    window.resize(1180, 760)
    window.show()

    problems = []
    findings = []

    SETTLE_MS = 900       # 切换路由后等渲染
    RETRY_MS = 1200       # 重试前多等一会
    MAX_ATTEMPTS = 3      # 偶发竞态不该让测试报假阳性

    def probe(route, selector, minimum, on_state, attempt=1):
        """切到路由并读一次页面状态；元素数不足时自动重试。"""
        def after_query(raw):
            try:
                state = json.loads(raw or "{}")
            except ValueError:
                state = {}

            # 只有"元素没到齐"才重试 —— 那多半是渲染还没完成。
            # 致命错误是确定性失败，重试无意义。
            not_ready = (minimum and state.get("matches", 0) < minimum
                         and not state.get("fatalVisible"))
            if not_ready and attempt < MAX_ATTEMPTS:
                print(f"  ({route} 第 {attempt} 次未就绪，重试…)")
                return QTimer.singleShot(
                    RETRY_MS, lambda: probe(route, selector, minimum, on_state, attempt + 1))

            on_state(state)

        delay = SETTLE_MS if attempt == 1 else RETRY_MS
        window.view.page().runJavaScript(
            f"location.hash = '#/{route}';",
            lambda _: QTimer.singleShot(delay, lambda: window.view.page().runJavaScript(
                collect_state_script(selector), after_query)),
        )

    def check_route(index):
        if index >= len(ROUTE_CHECKS):
            print()
            return check_drawers()

        route, selector, minimum, label = ROUTE_CHECKS[index]

        def on_state(state):
            findings.append((label, state))

            if state.get("fatalVisible"):
                problems.append(f"[{label}] 致命错误面板被触发: {state.get('fatalText')}")
            if state.get("theme") != expected_theme:
                problems.append(
                    f"[{label}] 主题应为 {expected_theme}，实际 {state.get('theme')}")
            if state.get("style") != expected_style:
                problems.append(
                    f"[{label}] 风格应为 {expected_style}，实际 {state.get('style')}")
            if minimum and state.get("matches", 0) < minimum:
                problems.append(
                    f"[{label}] 选择器 {selector!r} 期望至少 {minimum} 个元素，"
                    f"实际 {state.get('matches')}"
                )

            if route == shot_route:
                shot = window.view.grab()
                saved = shot.save(OUT_PNG)
                print(f"截图（{label}）: {OUT_PNG} ({'成功' if saved else '失败'}) "
                      f"{shot.width()}x{shot.height()}")

            QTimer.singleShot(700, lambda: check_route(index + 1))

        probe(route, selector, minimum, on_state)

    def run_js(script, callback):
        window.view.page().runJavaScript(script, callback)

    def click_topbar(label, callback):
        """点击顶栏上文案为 label 的按钮。"""
        run_js(
            "(function(){"
            "  var btns = document.querySelectorAll('.topbar-actions .btn');"
            "  for (var i = 0; i < btns.length; i++) {"
            "    if (btns[i].textContent.trim() === " + json.dumps(label) + ") {"
            "      btns[i].click(); return 'clicked'; }"
            "  }"
            "  return 'not-found';"
            "}())",
            callback,
        )

    def measure_drawer(callback):
        """统计当前抽屉里的表单字段数，并顺带看有没有掉进致命错误面板。"""
        run_js(
            "(function(){"
            "  var d = document.querySelector('.drawer');"
            "  var fatal = document.getElementById('fatal-panel');"
            "  return JSON.stringify({"
            "    fields: d ? d.querySelectorAll('.field').length : 0,"
            "    fatalVisible: !!(fatal && fatal.style.display !== 'none'),"
            "    fatalText: fatal ? document.getElementById('fatal-message').textContent : ''"
            "  });"
            "}())",
            callback,
        )

    def close_drawer(callback):
        """点遮罩关闭抽屉（抽屉的关闭方式之一）。"""
        run_js(
            "(function(){"
            "  var s = document.querySelector('.drawer-scrim');"
            "  if (s) { s.dispatchEvent(new MouseEvent('mousedown', {bubbles: true})); }"
            "  return 'ok';"
            "}())",
            callback,
        )

    def check_drawers():
        """入库/出库是主工作流，必须确认抽屉真的能打开并渲染出来。"""
        results = {}

        def step_open_inbound():
            def opened(res):
                if res != "clicked":
                    results["in"] = {"fields": 0, "missing": True}
                    return step_open_outbound()
                QTimer.singleShot(1200, step_measure_inbound)
            click_topbar("入库", opened)

        def step_measure_inbound():
            def measured(raw):
                results["in"] = _loads(raw)
                window.view.grab().save(OUT_PNG.replace(".png", "_inbound.png"))
                QTimer.singleShot(300, step_close_inbound)
            measure_drawer(measured)

        def step_close_inbound():
            close_drawer(lambda _: QTimer.singleShot(700, step_open_outbound))

        def step_open_outbound():
            def opened(res):
                if res != "clicked":
                    results["out"] = {"fields": 0, "missing": True}
                    return step_report()
                QTimer.singleShot(1200, step_measure_outbound)
            click_topbar("出库", opened)

        def step_measure_outbound():
            def measured(raw):
                results["out"] = _loads(raw)
                window.view.grab().save(OUT_PNG.replace(".png", "_outbound.png"))
                QTimer.singleShot(300, step_report)
            measure_drawer(measured)

        def step_report():
            for key, label, minimum in (("in", "入库抽屉", 5), ("out", "出库抽屉", 4)):
                info = results.get(key) or {}
                fields = info.get("fields", 0)
                print(f"  {label}: 表单字段 {fields} 个"
                      + ("（按钮未找到）" if info.get("missing") else ""))
                if fields < minimum:
                    problems.append(f"[{label}] 表单字段不足: {fields} < {minimum}")
                if info.get("fatalVisible"):
                    problems.append(f"[{label}] 致命错误面板被触发: {info.get('fatalText')}")
            finish()

        step_open_inbound()

    def _loads(raw):
        try:
            return json.loads(raw or "{}")
        except ValueError:
            return {}

    def finish():
        print()
        for label, state in findings:
            print(f"  {label}: 元素 {state.get('matches')} 个, "
                  f"空状态 {state.get('emptyState')} 个, "
                  f"主题 {state.get('theme')}, 风格 {state.get('style')}")

        if problems:
            print("\n冒烟测试失败:")
            for p in problems:
                print("  - " + p)
        else:
            print(f"\n冒烟测试通过（主题 {expected_theme}，风格 {expected_style}，"
                  f"共 {len(ROUTE_CHECKS)} 个页面）")
        app.quit()

    # 等首屏加载与 bootstrap 完成
    QTimer.singleShot(4000, lambda: check_route(0))
    app.exec_()
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
