"""
配置文件: 自动发现源表、管理路径
=================================
目录结构:
  input_data/               : 放源数据文件 (A/C/D表)
  output_result/            : 输出目录 (从模板复制, 所有修改在此进行)
  PE100_pooling模板表.xlsx  : 原始模板 (只读, 不会被修改)

A表 sheet 名映射: 数字→英文字母 (3→C, 4→D...)
"""
import os
import sys
from pathlib import Path

# 确保可以 import common 包（项目根目录）
_sys_path = str(Path(__file__).resolve().parent.parent.parent)
if _sys_path not in sys.path:
    sys.path.insert(0, _sys_path)

from common.file_utils import find_file, prepare_output as _prepare_output, reset_output as _reset_output
from common.sheet_utils import (
    is_data_sheet,
    sheet_map_a as _sheet_map_a,
    get_target_sheets as _get_target_sheets,
    sync_b_table_sheets as _sync_b_table_sheets,
)

ROOT = Path(__file__).resolve().parent.parent          # PE_100/


def _runtime_path(variable, fallback):
    configured = os.environ.get(variable)
    return Path(configured).resolve() if configured else fallback


INPUT_DATA = _runtime_path('HPLS_INPUT_DIR', ROOT / 'input_data')
OUTPUT_DIR = _runtime_path('HPLS_OUTPUT_DIR', ROOT / 'output_result')
TEMPLATE = ROOT / 'PE100_pooling模板表.xlsx'            # 原始模板 (只读)
from datetime import datetime
_OUTPUT_NAME = f'{datetime.now().strftime("%Y%m%d")}文库pooling表T7+PE100.xlsx'
DST = str(OUTPUT_DIR / _OUTPUT_NAME)


def prepare_output():
    """将原始模板复制到 output_result/ 中 (如已存在则跳过, 保留已有工作)"""
    return _prepare_output(OUTPUT_DIR, TEMPLATE, _OUTPUT_NAME)


def reset_output():
    """强制重新复制模板 (覆盖 output_result 中的副本)"""
    return _reset_output(OUTPUT_DIR, TEMPLATE, _OUTPUT_NAME)


def get_src_a():
    """A表: 上机表"""
    return find_file(INPUT_DATA, '上机')


def get_src_c():
    """C表: 外包文库质检总表"""
    return find_file(INPUT_DATA, '质检')


def get_src_d():
    """D表: 自建库出库报告"""
    return find_file(INPUT_DATA, '自建库')


def sheet_map_a():
    """扫描A表sheet名, 返回 [(源sheet名, 目标sheet名)]

    映射规则 (按优先级):
    1. 直接字母匹配: A表sheet名是单个大写字母 → 直接使用
    2. 数字映射: 1→A, 2→B, 3→C ...
    """
    return _sheet_map_a(get_src_a(), priority=('letter', 'number'))


def get_target_sheets():
    """根据A表动态获取目标B表数据sheet列表, 如 ['C', 'D']"""
    return _get_target_sheets(get_src_a(), priority=('letter', 'number'))


def sync_b_table_sheets(wb):
    """根据A表动态调整B表sheet:
    - 缺失的数据sheet: 从现有数据sheet重命名或复制创建
    - 多余的数据sheet: 重命名或删除 (无对应A表数据)
    非数据sheet (T7+制备等) 保留不动
    """
    _sync_b_table_sheets(wb, get_target_sheets(), ensure_columns=None)
