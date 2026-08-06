"""
错误信息翻译
===========
将 Python 异常翻译为实验室人员可理解的中文提示。

用法::

    from common.error_utils import translate

    try:
        run_pipeline(...)
    except Exception as e:
        print(translate(e))
"""

import re

# ── 翻译规则 ──────────────────────────────────────────────
# 每条规则: (匹配模式, 翻译函数)
# 模式按优先级排列 — 越靠前越优先匹配

_RULES = [
    # ── 源文件缺失 ──
    (
        re.compile(r'没有包含\s*"([^"]+)"\s*的\s*xlsx\s*文件'),
        lambda m, _e: (
            f'缺少源数据文件：\n'
            f'  需要在 input_data/ 中放入包含 "{m.group(1)}" 的 .xlsx 文件\n'
            f'  请检查文件名是否包含该关键词，或文件是否放入了正确的 input_data/ 目录'
        ),
    ),

    # ── 多文件冲突 ──
    (
        re.compile(r'中有多个包含\s*"([^"]+)"\s*的文件[，,]?\s*请只保留本批次文件[：:]\s*(.+)'),
        lambda m, _e: (
            f'input_data/ 中匹配到多个包含 "{m.group(1)}" 的文件：\n'
            f'  {m.group(2)}\n'
            f'  请只保留本批次需要的一个文件，将其他文件移出 input_data/'
        ),
    ),

    # ── A 表 sheet 结构无效 ──
    (
        re.compile(r'A表中没有找到有效的sheet名'),
        lambda _m, _e: (
            f'A 表（上机数据）的 sheet 结构无法识别。\n'
            f'  期望：sheet 名为数字（1→A面, 2→B面）或单个大写字母（A, B, C...）\n'
            f'  请确认是标准的华大 T7+ 上机表，sheet 未被重命名'
        ),
    ),

    # ── B 表模板缺少数据 sheet ──
    (
        re.compile(r'B表中没有可用的数据sheet作为复制模板'),
        lambda _m, _e: (
            f'Pooling 模板中没有可用的数据 sheet。\n'
            f'  模板应至少包含一个单字母数据 sheet（如 A），供脚本复制创建其他面\n'
            f'  请检查模板文件是否完整，或联系管理员恢复模板'
        ),
    ),

    # ── 文件被占用 ──
    (
        re.compile(r'Permission denied|PermissionError'),
        lambda _m, e: (
            f'文件被占用，无法写入。\n'
            f'  请关闭所有打开的 Excel 窗口（包括模板和输出文件）后重试\n'
            f'  原始错误: {e}'
        ),
    ),

    # ── openpyxl 加载/保存失败 ──
    (
        re.compile(r'(?:openpyxl.*Error|InvalidFileException|BadZipFile)'),
        lambda _m, e: (
            f'Excel 文件读取失败，可能原因：\n'
            f'  1. 文件损坏或格式不正确（不是真正的 .xlsx）\n'
            f'  2. 文件正在被其他程序占用\n'
            f'  3. 文件是旧版 .xls 格式（需转换为 .xlsx）\n'
            f'  原始错误: {e}'
        ),
    ),

    # ── KeyError / openpyxl sheet 不存在 ──
    (
        re.compile(r"Worksheet\s+'?([^']+)'?\s*does not exist"),
        lambda m, _e: (
            f'找不到 sheet: "{m.group(1)}"\n'
            f'  请确认模板文件未被修改或删除 sheet\n'
            f'  也可能是 A 表 sheet 映射出了问题 — 请检查 A 表 sheet 名称'
        ),
    ),

    # ── ModuleNotFoundError / ImportError ──
    (
        re.compile(r"No module named '?(openpyxl|\\w+)'?"),
        lambda m, _e: (
            f'缺少 Python 依赖: {m.group(1)}\n'
            f'  请运行: pip install {m.group(1)}'
        ),
    ),
]


def translate(error):
    """将异常翻译为友好中文提示。

    遍历所有规则，返回第一条匹配的翻译。
    无匹配则返回原始错误信息。

    Args:
        error: Exception 对象

    Returns:
        str: 翻译后的中文提示（或原始错误信息）
    """
    msg = str(error)
    error_type = type(error).__name__

    for pattern, formatter in _RULES:
        match = pattern.search(msg)
        if match:
            return formatter(match, error)

    # ── 没有匹配规则 → 返回带上下文的原始信息 ──
    return (
        f'执行时出现未预期的错误 ({error_type})：\n'
        f'  {msg}\n'
        f'  请将以上信息反馈给维护人员。'
    )
