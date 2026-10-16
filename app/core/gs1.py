# Copyright (C) [2025] [阿炜]
# 本程序是自由软件：你可以根据 GNU GPL v3 协议重新分发或修改它。
"""GS1 条码解析。

`_parse_parenthesised` 的实现体照搬自 ReagentMGS.py（原文件 1028-1137 行），
输出字典逐键对齐，未做任何行为改动 —— 这是医院实际数据正确性的依赖。

`parse_barcode` 在其上增加了一层：当括号形式解析不出任何字段时，回退到
原始 GS1-128（FNC1 分隔）解析。原始路径是纯增量，不影响既有输出。
"""

import re

# AI 表。照搬自原实现，顺序与内容均未改动。
AI_LIST = [
    # 固定长度AI
    {"ai": "00", "name": "SSCC", "length": 18, "type": "numeric"},
    {"ai": "01", "name": "GTIN", "length": 14, "type": "numeric"},
    {"ai": "02", "name": "Content GTIN", "length": 14, "type": "numeric"},
    {"ai": "11", "name": "Production Date", "length": 6, "type": "numeric"},
    {"ai": "12", "name": "Due Date", "length": 6, "type": "numeric"},
    {"ai": "13", "name": "Packaging Date", "length": 6, "type": "numeric"},
    {"ai": "15", "name": "Best Before Date", "length": 6, "type": "numeric"},
    {"ai": "17", "name": "Expiration Date", "length": 6, "type": "numeric"},
    {"ai": "20", "name": "Product Variant", "length": 2, "type": "numeric"},
    {"ai": "30", "name": "Variable Count", "length": 8, "type": "numeric"},

    # 可变长度AI
    {"ai": "10", "name": "Batch/Lot Number", "length": None, "type": "alphanumeric"},
    {"ai": "21", "name": "Serial Number", "length": None, "type": "alphanumeric"},
    {"ai": "22", "name": "Consumer Product Variant", "length": None, "type": "alphanumeric"},
    {"ai": "240", "name": "Additional Item Identification", "length": None, "type": "alphanumeric"},
    {"ai": "241", "name": "Customer Part Number", "length": None, "type": "alphanumeric"},
    {"ai": "242", "name": "Made-to-Order Variation Number", "length": None, "type": "numeric"},
    {"ai": "250", "name": "Secondary Serial Number", "length": None, "type": "alphanumeric"},
    {"ai": "251", "name": "Reference to Source Entity", "length": None, "type": "alphanumeric"},
    {"ai": "253", "name": "Global Document Type Identifier", "length": 17, "type": "numeric"},
    {"ai": "254", "name": "GLN Extension Component", "length": None, "type": "alphanumeric"},
    {"ai": "255", "name": "Global Coupon Number", "length": 25, "type": "numeric"},

    # 带小数点的AI
    {"ai": "3100", "name": "Net Weight (kg)", "length": 6, "type": "numeric"},
    {"ai": "3101", "name": "Net Weight (kg)", "length": 6, "type": "numeric"},
    {"ai": "3102", "name": "Net Weight (kg)", "length": 6, "type": "numeric"},
    {"ai": "3103", "name": "Net Weight (kg)", "length": 6, "type": "numeric"},
    {"ai": "3104", "name": "Net Weight (kg)", "length": 6, "type": "numeric"},
    {"ai": "3105", "name": "Net Weight (kg)", "length": 6, "type": "numeric"},
    {"ai": "3920", "name": "Price", "length": None, "type": "numeric"},
    {"ai": "3921", "name": "Price", "length": None, "type": "numeric"},
    {"ai": "3922", "name": "Price", "length": None, "type": "numeric"},

    # 内部使用标识
    {"ai": "90", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "91", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "92", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "93", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "94", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "95", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "96", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "97", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "98", "name": "Internal", "length": None, "type": "alphanumeric"},
    {"ai": "99", "name": "Internal", "length": None, "type": "alphanumeric"},
]

AI_MAP = {ai_info["ai"]: ai_info for ai_info in AI_LIST}

# GS1-128 的字段分隔符（FNC1）。扫码枪常以 GS(29) / \x1d 形式输出。
FNC1 = "\x1d"

# 原始格式里，解析 AI 时按长度从长到短尝试，避免把 3100 误判成 31/310
_AI_BY_LENGTH_DESC = sorted(AI_MAP.keys(), key=len, reverse=True)

# 定长 AI 的编号长度集合：2 位（01/10/17...）、3 位（240...）、4 位（3100...）
_AI_NUMBER_LENGTHS = sorted({len(a) for a in AI_MAP}, reverse=True)


def _apply_ai(ai_str, data, parsed_data):
    """把单个 (AI, 数据) 对写入结果字典。键名规则与旧实现完全一致。"""
    ai_info = AI_MAP.get(ai_str)

    if not ai_info:
        # 未知AI，使用原始格式存储
        parsed_data[f"Unknown_AI({ai_str})"] = data
        return

    if ai_info["length"] is not None:
        if len(data) < ai_info["length"]:
            processed_data = data
            key = f"ERROR_Short_{ai_info['name']}"
        elif len(data) > ai_info["length"]:
            processed_data = data[:ai_info["length"]]
            key = f"WARNING_Truncated_{ai_info['name']}"
        else:
            processed_data = data
            key = ai_info["name"]

        # 处理带小数点的AI (310X系列)
        if ai_str.startswith("310") and len(ai_str) == 4 and processed_data.isdigit():
            decimal_pos = int(ai_str[3])
            if decimal_pos > 0 and len(processed_data) > decimal_pos:
                processed_data = processed_data[:-decimal_pos] + '.' + processed_data[-decimal_pos:]
    else:
        processed_data = data
        key = ai_info["name"]

    parsed_data[key] = processed_data


def _parse_parenthesised(gs1_string):
    """括号形式：(01)0693...(17)260218。旧实现，行为逐条保留。"""
    gs1_string = gs1_string.replace("（", "(")
    gs1_string = gs1_string.replace("）", ")")

    pattern = r'\((\d+)\)([^\(]*)'
    matches = re.findall(pattern, gs1_string)

    parsed_data = {}
    for ai_str, data in matches:
        _apply_ai(ai_str, data, parsed_data)
    return parsed_data


def _match_ai_at(text, pos):
    """在 text[pos:] 处尝试匹配一个 AI，返回 (ai_str, ai_info) 或 None。"""
    for number_length in _AI_NUMBER_LENGTHS:
        candidate = text[pos:pos + number_length]
        if len(candidate) < number_length:
            continue
        if not candidate.isdigit():
            continue
        ai_info = AI_MAP.get(candidate)
        if ai_info:
            return candidate, ai_info
    return None


def _parse_raw_gs1(gs1_string):
    """原始 GS1-128：0106936415919665\\x1d11240819\\x1d17...

    定长字段按表内长度推进；变长字段读到 FNC1 或字符串结尾为止。
    与括号形式共用 _apply_ai，因此键名、截断告警、310x 小数处理完全一致。
    """
    # 扫码枪有时把 FNC1 打成 GS(29) 之外的可见形式，统一成 \x1d
    text = gs1_string.strip()
    if not text:
        return {}

    parsed_data = {}
    pos = 0
    length = len(text)

    while pos < length:
        # 跳过相邻分隔符
        if text[pos] == FNC1:
            pos += 1
            continue

        matched = _match_ai_at(text, pos)
        if not matched:
            # 一个字段都没解析出来 —— 这不是 GS1 数据，返回空字典，
            # 与旧实现对非 GS1 输入的输出保持一致。
            if not parsed_data:
                return {}
            # 已经解析出字段，说明是尾部无法识别的残留：
            # 保留下来，便于排查扫码枪配置。
            remainder = text[pos:].strip(FNC1)
            if remainder:
                parsed_data["Unknown_AI(?)"] = remainder
            break

        ai_str, ai_info = matched
        pos += len(ai_str)

        if ai_info["length"] is not None:
            data = text[pos:pos + ai_info["length"]]
            pos += ai_info["length"]
        else:
            end = text.find(FNC1, pos)
            if end == -1:
                end = length
            data = text[pos:end]
            pos = end + 1 if end < length else length

        _apply_ai(ai_str, data, parsed_data)

    return parsed_data


def parse_barcode(gs1_string):
    """解析 GS1 格式编码字符串。

    优先按括号形式解析（与旧实现完全一致）；解析不出任何字段时，回退到
    原始 GS1-128 格式。因此对既有输入的输出保持不变。
    """
    if not gs1_string:
        return {}

    parsed = _parse_parenthesised(gs1_string)
    if parsed:
        return parsed

    return _parse_raw_gs1(gs1_string)


class GS1Parser:
    """兼容旧调用的包装，保持 `GS1Parser.parse_barcode(raw)` 可用。"""

    parse_barcode = staticmethod(parse_barcode)
