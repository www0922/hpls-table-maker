"""PE150 Pooling 表全流程一键执行（优化版）。"""
import sys
from pathlib import Path
import openpyxl

_sys_path = str(Path(__file__).resolve().parent.parent.parent)
if _sys_path not in sys.path:
    sys.path.insert(0, _sys_path)

from config import DST

from clear_pooling import main as clear_main
from step1_migrate import main as step1
from step2_lookup import main as step2
from step3_group_sort import main as step3
from step4_formula import main as step4
from step5_lane import main as step5
from step6_t7_prepare import main as step6
from step7_offline_stats import main as step7
from format_font import main as fmt_main
from validate_output import main as validate


def main():
    import openpyxl
    
    # 步骤0: 清空模板
    clear_main(pool_wb=None)
    
    # 步骤1~7 共享工作簿
    pool_wb = openpyxl.load_workbook(DST)
    step1(pool_wb=pool_wb)
    step2(pool_wb=pool_wb)
    step3(pool_wb=pool_wb)
    step4(pool_wb=pool_wb)
    step5(pool_wb=pool_wb)
    step6(pool_wb=pool_wb)
    step7(pool_wb=pool_wb)
    pool_wb.save(DST)
    pool_wb.close()
    print("[OK] 步骤1~7 完成")
    
    # 格式化 + 校验
    fmt_main(pool_wb=None)
    validate(pool_wb=None)

    print("\n[DONE] PE150 Pooling 全流程")


if __name__ == '__main__':
    main()
