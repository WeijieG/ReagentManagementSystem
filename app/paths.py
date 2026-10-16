# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""路径解析。

打包后（PyInstaller onedir）应用目录是 exe 所在目录；开发环境下是仓库根目录。
数据库与用户配置始终放在用户目录，不随安装位置变化。
"""

import configparser
import os
import sys

# app/paths.py -> app/ -> 仓库根目录
_DEV_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DEV_APP_DIR = os.path.dirname(os.path.abspath(__file__))

IS_FROZEN = getattr(sys, "frozen", False)

DEFAULT_DATA_DIR_NAME = "ReagentManagementSystem"


def get_application_path():
    """exe 所在目录（打包后）或仓库根目录（开发时）。"""
    if IS_FROZEN:
        return os.path.dirname(sys.executable)
    return _DEV_ROOT


APPLICATION_PATH = get_application_path()
CONFIG_FILE = os.path.join(APPLICATION_PATH, "config.ini")
ICON_PATH = os.path.join(APPLICATION_PATH, "reagent.ico")


def get_ui_dir():
    """Web UI 静态资产目录。

    打包后由 build.spec 释放到 exe 同级的 ui/ 下；开发时位于 app/ui/。
    """
    if IS_FROZEN:
        return os.path.join(APPLICATION_PATH, "ui")
    return os.path.join(_DEV_APP_DIR, "ui")


def get_data_dir_name():
    """用户数据目录名。

    默认 "ReagentManagementSystem"。允许 config.ini 里用
    `[Application] DataDirName` 覆盖 —— 出测试版时改成别的名字，
    就能和正式数据完全隔离，不至于测试版一跑就动到真实库存。
    """
    try:
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE, encoding="utf-8")
        name = config.get("Application", "DataDirName", fallback="").strip()
        if name:
            return name
    except Exception:
        pass
    return DEFAULT_DATA_DIR_NAME


def get_db_dir():
    """数据库所在目录，不存在则创建。"""
    if IS_FROZEN:
        appdata_path = os.getenv("APPDATA")
        db_dir = (
            os.path.join(appdata_path, get_data_dir_name())
            if appdata_path
            else APPLICATION_PATH
        )
    else:
        db_dir = APPLICATION_PATH

    os.makedirs(db_dir, exist_ok=True)
    return db_dir


DB_DIR = get_db_dir()
DB_FILE = os.path.join(DB_DIR, "ReagentWarehouseData.db")
USER_DATA_FILE = os.path.join(DB_DIR, "userData.ini")
BACKUP_DIR = os.path.join(DB_DIR, "Backups")
SAVED_PATH_FILE = os.path.join(DB_DIR, "saved_path.txt")


def get_image_path(file_name):
    """获取图片的绝对路径，并检查文件是否存在。"""
    image_path = os.path.join(APPLICATION_PATH, "images", file_name)

    if not os.path.exists(image_path):
        print(f"警告: 图片文件不存在 - {image_path}")
        return ""

    # 规范化为正斜杠，便于 Qt 样式表引用
    return os.path.normpath(image_path).replace("\\", "/")
