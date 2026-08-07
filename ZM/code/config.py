"""
ZM项目配置文件: 自动发现源表、管理路径
=====================================
目录结构:
  input_data/               : 放源数据文件
  output_result/            : 输出目录
  ZM_pooling表模板.xlsx     : Pooling输出模板(只读)
  ZM_PCR定量表模板.xlsx     : PCR定量表模板(只读)
"""
import os
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _runtime_path(variable, fallback):
    configured = os.environ.get(variable)
    return Path(configured).resolve() if configured else fallback


INPUT_DATA = _runtime_path('HPLS_INPUT_DIR', ROOT / 'input_data')
OUTPUT_DIR = _runtime_path('HPLS_OUTPUT_DIR', ROOT / 'output_result')

TEMPLATE_POOL = ROOT / 'ZM_pooling表模板.xlsx'
TEMPLATE_PCR = ROOT / 'ZM_PCR定量表模板.xlsx'

from datetime import datetime as _dt
_date_str = _dt.now().strftime('%Y%m%d')
DST_POOL = str(OUTPUT_DIR / f'{_date_str}文库pooling表AE0.xlsx')
DST_PCR = str(OUTPUT_DIR / f'{_date_str}PCR定量表真迈.xlsx')


def prepare_output():
    """复制模板到 output_result/"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for tmpl, dst in [(TEMPLATE_POOL, DST_POOL), (TEMPLATE_PCR, DST_PCR)]:
        if not Path(dst).exists():
            shutil.copy2(tmpl, dst)
            print(f'[OK] 已复制 -> {dst}')
        else:
            print(f'[OK] 已存在,跳过 -> {dst}')
    return DST_POOL, DST_PCR


def reset_output():
    """强制重新复制"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for tmpl, dst in [(TEMPLATE_POOL, DST_POOL), (TEMPLATE_PCR, DST_PCR)]:
        shutil.copy2(tmpl, dst)
        print(f'[OK] 已重置 -> {dst}')
    return DST_POOL, DST_PCR


def find_file(keyword):
    """在 input_data/ 中精确选择唯一的来源文件。

    同一关键词匹配多个文件时不自动猜测，避免误用错误批次。
    """
    matches = sorted(
        (p for p in INPUT_DATA.glob(f'*{keyword}*.xlsx') if not p.name.startswith('~$')),
        key=lambda path: path.name)
    if not matches:
        raise FileNotFoundError(f'input_data/ 中没有包含 "{keyword}" 的 xlsx 文件')
    if len(matches) > 1:
        names = '、'.join(path.name for path in matches)
        raise RuntimeError(f'input_data/ 中有多个包含 "{keyword}" 的文件，请只保留本批次文件：{names}')
    return str(matches[0])


def get_src_a():
    """A表: 上机表"""
    return find_file('江西上机')


# A表shee名 → Pooling表sheet名映射
SHEET_RENAME = {'A': 'S', 'B': 'P'}


def get_target_sheets():
    """根据A表动态获取目标数据sheet列表, 返回映射后的名称如['S','P']"""
    import openpyxl
    wb = openpyxl.load_workbook(get_src_a(), data_only=True)
    sheets = sorted(sn for sn in wb.sheetnames
                    if len(sn) == 1 and sn.isalpha() and sn.isascii() and sn.isupper())
    wb.close()
    return [SHEET_RENAME.get(s, s) for s in sheets]


def get_src_c():
    """C表: 外包文库质检总表"""
    return find_file('质检')


def get_src_d():
    """D表: 自建库出库报告"""
    return find_file('自建库')


def get_src_purify():
    """纯化总表"""
    return find_file('纯化')


def get_src_qpcr():
    """qPCR定量结果统计"""
    return find_file('qPCR')
