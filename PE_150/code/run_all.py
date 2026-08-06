"""PE150 Pooling 表全流程一键执行。"""
import sys
from pathlib import Path
_sys_path = str(Path(__file__).resolve().parent.parent.parent)
if _sys_path not in sys.path:
    sys.path.insert(0, _sys_path)

from common.step_runner import run_pipeline, StepError

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
    try:
        run_pipeline(
            [
                ("清空模板",       clear_main),
                ("数据迁移",       step1),
                ("查找填充",       step2),
                ("分组排序",       step3),
                ("公式计算",       step4),
                ("Lane编号",       step5),
                ("T7+制备",        step6),
                ("下机数据统计",   step7),
                ("统一格式",       fmt_main),
                ("输出校验",       validate),
            ],
            title="PE150 Pooling 全流程",
        )
    except StepError:
        sys.exit(1)


if __name__ == '__main__':
    main()
