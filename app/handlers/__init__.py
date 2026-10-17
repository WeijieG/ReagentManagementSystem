# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""invoke() 的动作实现。

每个 handler 接收 payload 字典、返回可 JSON 化的数据；出错时抛 BridgeError
由 bridge 统一转成错误信封。

handler 不直接弹 Qt 对话框，需要原生交互时通过 ctx.ui 调用（见 UiHost）。
"""

from .queries import build_query_handlers
from .actions import build_action_handlers
from .dataops import build_data_handlers
from .update import build_update_handlers
from .timecheck import build_time_handlers


class UiHost:
    """主窗口需要实现的少量原生交互接口。

    Web UI 无法自行弹出原生文件选择框（file:// 源下也不该尝试），
    因此这几件事必须回到 Python 侧做。
    """

    def choose_save_file(self, title, default_name, file_filter):
        raise NotImplementedError

    def choose_open_file(self, title, file_filter):
        raise NotImplementedError

    def confirm(self, title, text):
        raise NotImplementedError

    def warn(self, title, text):
        raise NotImplementedError

    def info(self, title, text):
        raise NotImplementedError

    def quit_app(self):
        raise NotImplementedError


class Context:
    """handler 能接触到的一切。"""

    def __init__(self, db, api, ui, version, config_file=None, db_report=None):
        self.db = db
        self.api = api
        self.ui = ui
        self.version = version
        self.config_file = config_file
        # 开库一致性自检结果，由 app.bootstrap 带给前端用于提示
        self.db_report = db_report or {"repaired": [], "remaining": []}


def build_registry(ctx):
    """组装 action -> handler 映射。新增动作只需在这里加一行。"""
    registry = {}
    registry.update(build_query_handlers(ctx))
    registry.update(build_action_handlers(ctx))
    registry.update(build_data_handlers(ctx))
    registry.update(build_update_handlers(ctx))
    registry.update(build_time_handlers(ctx))
    return registry
