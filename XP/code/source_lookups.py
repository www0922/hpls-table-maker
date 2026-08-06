"""来源表查找工具。"""

from __future__ import annotations

import openpyxl

from config import get_src_qpcr


def clean_key(value):
    return str(value or "").strip()


def build_qpcr_lookup():
    """构建 ``文库编号 -> qPCR浓度(K列)`` 查找表。

    历史统计表中的文库编号可能出现在项目编号、样本编号、样本名称或
    辅助编号列，因此同时索引 C/E/F/P 列。遇到同一编号多次出现时，
    使用表中最后一个非空 qPCR 结果。
    """
    workbook = openpyxl.load_workbook(get_src_qpcr(), read_only=True, data_only=True)
    lookup = {}
    duplicate_keys = set()
    for worksheet in workbook.worksheets:
        for values in worksheet.iter_rows(
            min_row=2, max_col=16, values_only=True
        ):
            qpcr_value = values[10]
            if qpcr_value is None:
                continue
            for index in (2, 4, 5, 15):
                key = clean_key(values[index])
                if not key:
                    continue
                if key in lookup and lookup[key] != qpcr_value:
                    duplicate_keys.add(key)
                lookup[key] = qpcr_value
    workbook.close()
    if duplicate_keys:
        print(f"  qPCR统计: {len(duplicate_keys)}个编号存在多个结果，已使用最后一个非空值")
    print(f"  qPCR统计: {len(lookup)}个可查编号")
    return lookup
