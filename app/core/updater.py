# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""自动更新。

协议与流程照搬自 ReagentMGS.py（原文件 463-733、3290-3304 行）：
  * GET {server}/update.json?version={current}
  * 响应字段 update / version / url / md5 / changelog / size
  * 下载后校验 MD5，再以 /SILENT /RESTARTAPPLICATIONS 静默运行安装包
  * 下载路径写入 saved_path.txt，下次启动时清理

本模块不导入 PyQt5；进度通过回调上报，由调用方决定如何呈现。

两处与旧实现不同（本次修复）：
  1. 版本比较改为语义化（见 version.py）。
  2. 支持按证书指纹固定 TLS 对端。未配置指纹时退回旧行为（不校验证书）。
"""

import configparser
import hashlib
import os
import subprocess
import sys

from . import version as version_mod
from .. import paths

# `requests` 刻意**不在模块级导入**。
#
# 它只服务于"检查更新"这一个功能，却会拉起 urllib3 -> socket -> _socket.pyd。
# 一旦目标机上 _socket 加载不了（老 Windows 上确实会遇到），整个程序会在
# 启动阶段就崩掉 —— 一个用不到的功能把主功能拦在门外。
#
# 改成按需导入后：更新用不了只会让"检查更新"失败并给出提示，
# 库存、出入库、导出这些核心功能照常可用。


def _requests():
    """按需导入 requests。缺失时抛 UpdateError，由调用方兜住。"""
    try:
        import requests
    except ImportError as e:
        raise UpdateError(f"缺少网络组件，无法检查更新: {e}") from e
    return requests

# 证书指纹配置项（config.ini [Security]）。为空表示不做固定。
CERT_FINGERPRINT_KEY = "UpdateCertFingerprint"


class UpdateError(Exception):
    """更新流程中的可预期失败。"""


def _load_config(config_file=None):
    config = configparser.ConfigParser()
    config.read(config_file or paths.CONFIG_FILE, encoding="utf-8")
    return config


def get_server_url(config_file=None):
    config = _load_config(config_file)
    return config.get('Update', 'Server', fallback='')


def get_current_version(config_file=None):
    config = _load_config(config_file)
    return config.get('Application', 'Version', fallback='1.0.0')


def get_pinned_fingerprint(config_file=None):
    """读取固定的证书 SHA-256 指纹（小写十六进制，冒号可选）。"""
    config = _load_config(config_file)
    raw = config.get('Security', CERT_FINGERPRINT_KEY, fallback='')
    if not raw:
        return ""
    return normalize_fingerprint(raw)


def normalize_fingerprint(value):
    """把 "AB:CD:..." 或带引号的写法统一成 64 位小写十六进制。"""
    text = (value or "").strip().strip('"').strip("'")
    text = text.replace(":", "").replace(" ", "").lower()
    return text


def peer_fingerprint(response):
    """从已建立的连接里取出对端证书的 SHA-256 指纹。

    取不到时返回空字符串（调用方应据此拒绝，而不是放行）。
    """
    raw = getattr(response, "raw", None)
    connection = getattr(raw, "_connection", None)
    if connection is None:
        connection = getattr(raw, "connection", None)

    sock = getattr(connection, "sock", None)
    if sock is None:
        return ""

    try:
        der = sock.getpeercert(binary_form=True)
    except (AttributeError, ValueError, OSError):
        return ""

    if not der:
        return ""

    return hashlib.sha256(der).hexdigest()


def _verify_pinned(response, pinned_fingerprint):
    """校验对端证书指纹。未配置指纹时不做校验（与旧行为一致）。"""
    if not pinned_fingerprint:
        return

    actual = peer_fingerprint(response)
    if not actual:
        raise UpdateError("无法获取服务器证书，已中止更新。")
    if actual != pinned_fingerprint:
        raise UpdateError("服务器证书指纹不匹配，已中止更新。")


def _get(url, pinned_fingerprint="", timeout=5, stream=False):
    """统一的 GET：按需校验证书指纹后再返回响应。"""
    requests = _requests()
    response = requests.get(url, timeout=timeout, verify=False, stream=stream)
    response.raise_for_status()
    _verify_pinned(response, pinned_fingerprint)
    return response


def check_for_update(ignored_version="", config_file=None, timeout=5):
    """检查更新。

    返回更新信息字典；无更新或检查失败时返回 None。
    失败不抛异常 —— 与旧实现一致，检查更新失败不应影响程序启动。
    """
    try:
        update_url = get_server_url(config_file)
        if not update_url:
            return None

        current_version = get_current_version(config_file)
        pinned = get_pinned_fingerprint(config_file)

        response = _get(
            f"{update_url}/update.json?version={current_version}",
            pinned_fingerprint=pinned,
            timeout=timeout,
        )
        update_info = response.json()
        update_version = update_info['version']

        # 用户已忽略的版本不提示（对齐旧实现 update_version <= ignored_version）
        if ignored_version and version_mod.is_same_or_older(update_version, ignored_version):
            return None

        if version_mod.is_newer(update_version, current_version):
            return update_info
        return None
    except Exception as e:
        print(f"检查更新失败: {str(e)}")
        return None


def calculate_md5(file_path, chunk_size=4096):
    md5 = hashlib.md5()
    with open(file_path, 'rb') as f:
        chunk = f.read(chunk_size)
        while chunk:
            md5.update(chunk)
            chunk = f.read(chunk_size)
    return md5.hexdigest()


def download_update(url, dest_path, md5=None, progress_cb=None,
                    pinned_fingerprint="", timeout=30, cancel_check=None):
    """下载更新包到 dest_path。

    progress_cb(percent) 上报进度；cancel_check() 返回 True 时中止。
    失败抛 UpdateError。
    """
    try:
        response = _get(url, pinned_fingerprint=pinned_fingerprint,
                        timeout=timeout, stream=True)
    except UpdateError:
        raise
    except Exception as e:
        raise UpdateError(f"下载失败: {str(e)}") from e

    total_length = response.headers.get('content-length')
    total = int(total_length) if total_length and total_length.isdigit() else 0

    written = 0
    try:
        with open(dest_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=1024):
                if cancel_check and cancel_check():
                    raise UpdateError("下载已取消。")
                if not chunk:
                    continue
                f.write(chunk)
                written += len(chunk)
                if progress_cb and total:
                    progress_cb(min(99, int(written * 100 / total)))
    except UpdateError:
        raise
    except Exception as e:
        raise UpdateError(f"下载失败: {str(e)}") from e
    finally:
        response.close()

    if md5:
        if calculate_md5(dest_path) != md5:
            try:
                os.remove(dest_path)
            except OSError:
                pass
            raise UpdateError("文件校验失败: MD5 不匹配")

    if progress_cb:
        progress_cb(100)

    return dest_path


def write_saved_path(file_path):
    """记录安装包位置，供下次启动清理。"""
    try:
        with open(paths.SAVED_PATH_FILE, 'w', encoding="utf-8") as f:
            f.write(file_path)
    except OSError as e:
        print(f"无法记录安装包路径: {str(e)}")


def install_update(file_path):
    """静默运行安装包并退出当前程序。"""
    write_saved_path(file_path)
    try:
        subprocess.Popen([file_path, '/SILENT', '/RESTARTAPPLICATIONS'])
    except Exception as e:
        raise UpdateError(f"无法启动安装程序: {str(e)}") from e


def cleanup_old_installer():
    """删除上次更新留下的安装包。"""
    if not os.path.exists(paths.SAVED_PATH_FILE):
        return None

    try:
        with open(paths.SAVED_PATH_FILE, 'r', encoding="utf-8") as f:
            file_path = f.read()
    except OSError:
        return None

    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
            print(f"已删除安装包: {file_path}")
        except OSError as e:
            print(f"删除安装包失败: {str(e)}")
    return None


def restart_application():
    """就地重启当前程序。"""
    os.execl(sys.executable, sys.executable, *sys.argv)
