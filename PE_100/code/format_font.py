"""
统一格式: 字体、行高、数字格式、列宽
"""
import sys
from pathlib import Path
_sys_path = str(Path(__file__).resolve().parent.parent.parent)
if _sys_path not in sys.path:
    sys.path.insert(0, _sys_path)

import openpyxl
from config import DST, get_target_sheets, is_data_sheet
from common.format_utils import apply_standard_format


def main(pool_wb=None):
    wb = pool_wb if pool_wb is not None else openpyxl.load_workbook(DST)

    apply_standard_format(
        wb,
        sheet_order=get_target_sheets() + ['T7+制备'],
        data_sheet_predicate=is_data_sheet,
        header_height_map={'T7+制备': 40},
        data_header_height=90,
        number_formats={'D': '0.00', 'E': '0.000', 'F': '0.000', 'H': '0.00', 'I': '0.000'},
    )

    if pool_wb is None:
        wb.save(DST)
    print(f'\n[DONE] 完成 -> {DST}')


if __name__ == '__main__':
    main()
