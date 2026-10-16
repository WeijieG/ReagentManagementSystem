# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""构建后重算完整性哈希，写回**产物目录里**的 config.ini。

应用启动时会比对主程序的 SHA-256 与 config.ini 里记录的值
（app/core/integrity.py）。每次重新打包 exe 的哈希都会变。

留空 = 跳过校验，所以漏跑这一步不会让程序打不开 —— 但校验也就形同虚设。
正式发布前必须跑，tools/build_installer.py 已经把它串进流程了。

⚠️ 写入目标是**产物里**那份 config.ini，不是仓库根目录那份。
仓库那份是构建模板，哈希恒为空，不会被分发；写错地方等于白算。

这里同时算一份 ui/ 目录的清单哈希 —— onedir 之后界面是明文散文件，
只校验 exe 等于没校验界面。

用法：
    python tools/gen_hashes.py                     # 自动找 Dist 下的产物
    python tools/gen_hashes.py --dist <目录>
    python tools/gen_hashes.py --config <文件>      # 覆盖写入目标
"""

import argparse
import hashlib
import json
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)

from app.core import integrity  # noqa: E402

DEFAULT_DIST = os.path.join(REPO, "Dist")


def hash_file(path, chunk_size=8192):
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        chunk = f.read(chunk_size)
        while chunk:
            digest.update(chunk)
            chunk = f.read(chunk_size)
    return digest.hexdigest()


def find_dist_dir(explicit=None):
    """定位打包产物目录。"""
    if explicit:
        return os.path.abspath(explicit)

    candidates = [
        os.path.join(DEFAULT_DIST, "ReagentManagementSystem"),
        os.path.join(DEFAULT_DIST, "ReagentManagementSystem_x64"),
        os.path.join(REPO, "dist", "ReagentManagementSystem"),
    ]
    for path in candidates:
        if os.path.isfile(os.path.join(path, integrity.EXE_NAME)):
            return path
    return None


def update_config(config_path, exe_hash, ui_hash):
    """把哈希写进 config.ini。

    config.ini 里存的是 **JSON 字符串**（带引号），因为读取端用 json.loads 解析
    （ReagentMGS.py:2741）。这里保持该格式不变，免得新旧不兼容。
    """
    import configparser

    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")

    if not config.has_section("Security"):
        config.add_section("Security")

    # x86 / x64 两把都写同一个值：判定逻辑是「任意一个匹配即通过」
    quoted = json.dumps(exe_hash)
    config.set("Security", "MainFileHashe_x64", quoted)
    config.set("Security", "MainFileHashe_x86", quoted)

    if ui_hash:
        config.set("Security", integrity.UI_HASH_KEY, json.dumps(ui_hash))

    with open(config_path, "w", encoding="utf-8") as f:
        config.write(f)


def main():
    parser = argparse.ArgumentParser(description="重算完整性哈希并写回 config.ini")
    parser.add_argument("--dist", help="打包产物目录（含 ReagentManagementSystem.exe）")
    parser.add_argument("--config", default=None,
                        help="要写入的 config.ini，默认写产物目录里的那份")
    parser.add_argument("--ui-in-dist", action="store_true", default=True,
                        help="ui/ 取打包产物里的那份（默认）")
    args = parser.parse_args()

    dist = find_dist_dir(args.dist)
    if not dist:
        print("找不到打包产物。请先运行 PyInstaller，或用 --dist 指定目录。")
        print(f"已尝试：{DEFAULT_DIST}\\ReagentManagementSystem\\{integrity.EXE_NAME}")
        return 1

    exe_path = os.path.join(dist, integrity.EXE_NAME)
    if not os.path.isfile(exe_path):
        print(f"找不到主程序：{exe_path}")
        return 1

    exe_hash = hash_file(exe_path)
    print(f"产物目录 : {dist}")
    print(f"主程序   : {integrity.EXE_NAME}")
    print(f"SHA-256  : {exe_hash}")

    ui_dir = os.path.join(dist, "ui")
    ui_hash = ""
    if os.path.isdir(ui_dir):
        ui_hash = integrity.calculate_ui_bundle_hash(ui_dir)
        print(f"ui/ 清单 : {ui_hash}")
    else:
        print("警告: 产物里没有 ui/ 目录，界面资源未纳入校验")

    # 默认写**产物目录里**的 config.ini，而不是仓库根目录那份。
    #
    # 仓库那份只是构建时的模板（哈希恒为空）；真正被打进安装包的是产物里的
    # 副本。以前默认写仓库的，结果哈希全写在了不会被分发的那一份上，
    # 装出来的包校验永远为空 —— 这个功能等于没生效。
    config_path = args.config or os.path.join(dist, "config.ini")
    if not os.path.isfile(config_path):
        print(f"找不到要写入的 config.ini：{config_path}")
        return 1

    update_config(config_path, exe_hash, ui_hash)
    print(f"\n已写回 {config_path}")
    print("  提示：重算后 config.ini 本身也会变，若安装包已经打好需要重新打包。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
