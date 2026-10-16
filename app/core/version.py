# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""版本号比较。

不依赖任何第三方库，客户端（updater.py）与服务端（server.py）共用同一份实现，
避免两端判断不一致导致更新循环。

替换掉原先的字典序字符串比较 —— 那种比较下 "2.0.10" < "2.0.9"，
会让 2.0.10 的更新推送不出去。
"""

import re

_VERSION_RE = re.compile(r'^(\d+)(.*)$')


def _parse(version):
    """把版本号切成可比较的元组。

    每段取前导数字作为数值；缺段补 0，因此 "2.1" == "2.1.0"。
    数字段之后的残余（如 "2.0.0-beta" 的 "-beta"）作为预发布标记，
    没有预发布标记的版本更新。
    """
    text = (version or "").strip()
    parts = text.split('.')
    numbers = []
    prerelease = ""

    for index, part in enumerate(parts):
        match = _VERSION_RE.match(part.strip())
        if match:
            numbers.append(int(match.group(1)))
            rest = match.group(2)
            if rest and not prerelease:
                prerelease = rest
        else:
            # 整段都不是数字：剩下的都当预发布标记
            if part.strip() and not prerelease:
                prerelease = part.strip()
            numbers.append(0)

    return numbers, prerelease


def compare_versions(left, right):
    """比较两个版本号。

    返回 -1 / 0 / 1，语义同 (left 与 right 的关系)。
    """
    left_numbers, left_pre = _parse(left)
    right_numbers, right_pre = _parse(right)

    length = max(len(left_numbers), len(right_numbers))
    left_numbers += [0] * (length - len(left_numbers))
    right_numbers += [0] * (length - len(right_numbers))

    if left_numbers != right_numbers:
        return -1 if left_numbers < right_numbers else 1

    # 数字段相同：无预发布标记的为正式版，更新
    if left_pre == right_pre:
        return 0
    if left_pre and not right_pre:
        return -1
    if right_pre and not left_pre:
        return 1
    return -1 if left_pre < right_pre else 1


def is_newer(candidate, current):
    """candidate 是否比 current 新。"""
    return compare_versions(candidate, current) > 0


def is_same_or_older(candidate, baseline):
    """candidate 是否不高于 baseline。

    用于「忽略此版本」判定：忽略 2.0.10 后，2.0.10 及更旧的不再提示，
    但 2.0.11 仍必须提示。
    """
    return compare_versions(candidate, baseline) <= 0
