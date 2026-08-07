"""按操作手册要求依次执行完整 XP 流程。"""

from clear_pooling import main as clear_outputs
from format_font import main as format_outputs
from step1_migrate import main as step1
from step2_group_sort import main as step2
from step3_lookup import main as step3
from step4_formula import main as step4
from step5_summary import main as step5
from step6_lane_name import main as step6
from step7_mark import main as step7
from step8_dilution import main as step8
from step_stats_template import main as step_stats
from step9_qpcr import main as step9
from validate_output import main as validate_outputs
def main():
    steps = [
        ("清空输出副本", clear_outputs),
        ("迁移上机数据", step1),
        ("按 lane 分组", step2),
        ("填充来源数据", step3),
        ("计算取样体积", step4),
        ("计算汇总和稀释", step5),
        ("生成组名", step6),
        ("填写状态", step7),
        ("生成文库稀释表", step8),
        ("生成下机数据统计模版", step_stats),
        ("生成 qPCR 定量表", step9),
        ("统一格式", format_outputs),
        ("校验输出", validate_outputs),
    ]
    for index, (label, function) in enumerate(steps, start=1):
        print(f"\n{'=' * 60}\n[{index}/{len(steps)}] {label}\n{'=' * 60}")
        function()
    print("\n[DONE] XP 全流程执行完成")


if __name__ == "__main__":
    main()
