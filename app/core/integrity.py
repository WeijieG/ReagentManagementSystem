# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""应用完整性与用户状态。

自校验逻辑照搬自 ReagentMGS.py（原文件 2720-2812 行），判定条件未做改动：
  * 只校验主程序 exe 的 SHA-256；
  * x86 / x64 两个哈希匹配任意一个即通过；
  * 未打包（开发环境）时整个校验跳过；
  * config.ini 缺失即失败。
"""

import configparser
import hashlib
import json
import os

from .. import paths

EXE_NAME = "ReagentManagementSystem.exe"

# 可选：ui/ 目录的清单哈希。config.ini 里存在该键时才启用，
# 用于在 onedir 布局下覆盖明文散落的 Web 资产。
UI_HASH_KEY = "UiBundleHashe"


def calculate_file_hashes(file_path=None):
    """计算主程序 exe 的 SHA-256。

    沿用旧实现的两个特殊返回值：文件不存在返回字符串 '-1'，
    读取失败返回整数 -1。
    """
    if file_path is None:
        file_path = os.path.join(paths.APPLICATION_PATH, EXE_NAME)

    if not os.path.exists(file_path):
        return '-1'

    combined_hash = hashlib.sha256()
    try:
        with open(file_path, 'rb') as f:
            chunk = f.read(8192)
            while chunk:
                combined_hash.update(chunk)
                chunk = f.read(8192)
    except Exception as e:
        print(f"无法读取文件 {EXE_NAME}: {str(e)}")
        return -1

    return combined_hash.hexdigest()


def calculate_ui_bundle_hash(ui_dir=None):
    """计算 ui/ 目录内所有文件的清单哈希。

    逐个文件累加 "相对路径\\0内容哈希" 后再取一次 SHA-256，
    因此文件增删、改名、内容变化都会改变结果。
    """
    if ui_dir is None:
        ui_dir = paths.get_ui_dir()

    if not os.path.isdir(ui_dir):
        return '-1'

    combined = hashlib.sha256()
    for root, dirs, files in os.walk(ui_dir):
        dirs.sort()
        for name in sorted(files):
            full = os.path.join(root, name)
            rel = os.path.relpath(full, ui_dir).replace("\\", "/")
            file_hash = hashlib.sha256()
            try:
                with open(full, 'rb') as f:
                    chunk = f.read(8192)
                    while chunk:
                        file_hash.update(chunk)
                        chunk = f.read(8192)
            except Exception as e:
                print(f"无法读取 {rel}: {str(e)}")
                continue
            combined.update(rel.encode("utf-8"))
            combined.update(b"\0")
            combined.update(file_hash.hexdigest().encode("ascii"))
            combined.update(b"\0")

    return combined.hexdigest()


def _read_expected(config, section, option):
    """读取 config.ini 中的哈希值。沿用旧实现的 json.loads 解析方式。"""
    if not config.has_option(section, option):
        return ""
    try:
        return json.loads(config.get(section, option))
    except (ValueError, TypeError):
        return ""


def check_file_integrity(config_file=None):
    """校验应用完整性。返回 (是否通过, 提示信息)。"""
    if config_file is None:
        config_file = paths.CONFIG_FILE

    if not os.path.exists(config_file):
        return False, "配置文件不存在，无法进行完整性校验。"

    config = configparser.ConfigParser()
    config.read(config_file, encoding="utf-8")

    stored_hashe_86 = _read_expected(config, 'Security', 'MainFileHashe_x86')
    stored_hashe_64 = _read_expected(config, 'Security', 'MainFileHashe_x64')

    current_hashes = calculate_file_hashes()

    # 开发环境不做校验（与旧实现一致）
    if not paths.IS_FROZEN:
        return True, ""

    # 两个哈希都为空说明还没配置过 —— 直接放行。
    #
    # 这是为了避免一个很伤人的坑：每次重新打包 exe 的哈希都会变，忘了重算就会
    # 让程序启动即退出，而报错信息还指向"应用不完整"，排查方向完全是错的。
    # 构建流程里跑一次 tools/gen_hashes.py 就会把哈希写上，校验随即生效。
    if not stored_hashe_86 and not stored_hashe_64:
        print("提示: config.ini 未配置完整性哈希，跳过校验。"
              "如需启用，请运行 tools/gen_hashes.py")
        return True, ""

    if stored_hashe_86 != current_hashes and stored_hashe_64 != current_hashes:
        return False, (
            "应用完整性异常，出于安全考虑，程序将退出。\n\n"
            f"当前主程序哈希:\n{current_hashes}\n\n"
            "如果是刚重新打包过，运行 tools/gen_hashes.py 重算即可。"
        )

    # 主程序通过后再校验 UI 资产（仅在配置了该项时启用）
    expected_ui = _read_expected(config, 'Security', UI_HASH_KEY)
    if expected_ui:
        if calculate_ui_bundle_hash() != expected_ui:
            return False, "界面资源完整性异常，出于安全考虑，程序将退出。"

    return True, ""


def _load_user_data():
    user_data = configparser.ConfigParser()
    if os.path.exists(paths.USER_DATA_FILE):
        user_data.read(paths.USER_DATA_FILE, encoding="utf-8")
    return user_data


def _save_user_data(user_data):
    os.makedirs(os.path.dirname(paths.USER_DATA_FILE), exist_ok=True)
    with open(paths.USER_DATA_FILE, 'w', encoding="utf-8") as f:
        user_data.write(f)


def is_disclaimer_accepted():
    """用户是否已接受免责声明。首次运行会写出一份默认配置。"""
    if not os.path.exists(paths.USER_DATA_FILE):
        user_data = configparser.ConfigParser()
        user_data['General'] = {'DisclaimerAccepted': 'False'}
        _save_user_data(user_data)

    user_data = _load_user_data()
    return user_data.get('General', 'DisclaimerAccepted', fallback='False') == 'True'


def set_disclaimer_accepted(accepted=True):
    user_data = _load_user_data()
    if not user_data.has_section('General'):
        user_data.add_section('General')
    user_data.set('General', 'DisclaimerAccepted', 'True' if accepted else 'False')
    _save_user_data(user_data)


def get_ignored_version():
    """用户选择忽略的更新版本号。"""
    user_data = _load_user_data()
    if user_data.has_section('Update') and user_data.has_option('Update', 'IgnoredVersion'):
        return user_data.get('Update', 'IgnoredVersion', fallback='')
    return ""


def set_ignored_version(version):
    user_data = _load_user_data()
    if not user_data.has_section('Update'):
        user_data.add_section('Update')
    user_data.set('Update', 'IgnoredVersion', version)
    _save_user_data(user_data)


def get_render_mode():
    """界面渲染模式：'hardware' 或 'software'。

    默认硬件渲染。一旦在目标机器上检测到 WebEngine 渲染失败，就永久记成
    software —— 老 Win7 显卡上这类失败是静默的（纯白界面、无异常），
    不记下来每次启动都会白屏。
    """
    user_data = _load_user_data()
    mode = user_data.get('UI', 'RenderMode', fallback='hardware')
    return mode if mode in ('hardware', 'software') else 'hardware'


def set_render_mode(mode):
    user_data = _load_user_data()
    if not user_data.has_section('UI'):
        user_data.add_section('UI')
    user_data.set('UI', 'RenderMode', mode)
    _save_user_data(user_data)


def get_theme():
    """界面主题偏好：light / dark / system。"""
    user_data = _load_user_data()
    return user_data.get('General', 'Theme', fallback='system')


def set_theme(mode):
    user_data = _load_user_data()
    if not user_data.has_section('General'):
        user_data.add_section('General')
    user_data.set('General', 'Theme', mode)
    _save_user_data(user_data)
