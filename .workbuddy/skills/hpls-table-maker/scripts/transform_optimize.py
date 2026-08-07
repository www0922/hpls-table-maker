#!/usr/bin/env python3
"""自动改造 XP/ZM 项目：一次加载工作簿 + 源表 read_only 优化。"""
import re
import sys
from pathlib import Path


def transform_xp_step(filepath: Path):
    """改造单个步骤脚本，使 main() 接受可选 workbook 参数。"""
    code = filepath.read_text(encoding="utf-8")

    filename = filepath.name

    # --- 1. 函数签名改造 ---
    if filename == "step9_qpcr.py":
        code = code.replace("def main():", "def main(pool_wb=None, pcr_wb=None):")
        # 替换 pool_workbook load
        code = re.sub(
            r'pool_workbook\s*=\s*openpyxl\.load_workbook\(DST_POOL,\s*data_only=False\)',
            'pool_wb_local = pool_wb if pool_wb is not None else openpyxl.load_workbook(DST_POOL, data_only=False)',
            code,
        )
        # 替换 pcr_workbook load
        code = re.sub(
            r'pcr_workbook\s*=\s*openpyxl\.load_workbook\(DST_PCR\)',
            'pcr_wb_local = pcr_wb if pcr_wb is not None else openpyxl.load_workbook(DST_PCR)',
            code,
        )
        # 替换变量名
        code = code.replace("pool_workbook", "pool_wb_local")
        code = code.replace("pcr_workbook", "pcr_wb_local")
        # save/close 守卫
        code = code.replace(
            "pcr_wb_local.save(DST_PCR)",
            "if pcr_wb is None:\n    pcr_wb_local.save(DST_PCR)",
        )
    elif filename == "format_font.py":
        code = code.replace("def main():", "def main(pool_wb=None, pcr_wb=None):")

        # pool wb load
        pool_pattern = r"(wb(?:_pool)?)\s*=\s*openpyxl\.load_workbook\(DST_POOL\)"
        match = re.search(pool_pattern, code)
        if match:
            var = match.group(1) or "wb_pool"
            code = re.sub(
                pool_pattern,
                f'{var}_local = pool_wb if pool_wb is not None else openpyxl.load_workbook(DST_POOL)',
                code,
            )
            code = code.replace(
                f"{var}.save(DST_POOL)",
                f"if pool_wb is None:\n    {var}_local.save(DST_POOL)",
            )
            code = code.replace(var, f"{var}_local")

        # pcr wb load
        pcr_pattern = r"(wb_pcr|wb2?)\s*=\s*openpyxl\.load_workbook\(DST_PCR\)"
        match2 = re.search(pcr_pattern, code)
        if match2:
            var2 = match2.group(1)
            code = re.sub(
                pcr_pattern,
                f'{var2}_local = pcr_wb if pcr_wb is not None else openpyxl.load_workbook(DST_PCR)',
                code,
            )
            code = code.replace(
                f"{var2}.save(DST_PCR)",
                f"if pcr_wb is None:\n    {var2}_local.save(DST_PCR)",
            )
            code = code.replace(var2, f"{var2}_local")
    elif filename == "validate_output.py":
        code = code.replace("def main():", "def main(pool_wb=None):")
        code = re.sub(
            r"pool_workbook\s*=\s*openpyxl\.load_workbook\(DST_POOL\)",
            "pool_workbook = pool_wb if pool_wb is not None else openpyxl.load_workbook(DST_POOL)",
            code,
        )
        code = re.sub(
            r"pool_workbook\.close\(\)",
            "if pool_wb is None:\n        pool_workbook.close()",
            code,
        )
    elif filename == "clear_pooling.py":
        # 不改造
        pass
    else:
        # 通用步骤：只操作 DST_POOL
        code = code.replace("def main():", "def main(pool_wb=None):")

        # 匹配各种变量名的 load_workbook(DST_POOL)
        load_expr = r"(\w+)\s*=\s*openpyxl\.load_workbook\(DST_POOL\)"
        match = re.search(load_expr, code)
        if match:
            var = match.group(1)
            code = re.sub(
                load_expr,
                f'{var} = pool_wb if pool_wb is not None else openpyxl.load_workbook(DST_POOL)',
                code,
            )
            # save 守卫
            code = code.replace(
                f"{var}.save(DST_POOL)",
                f"if pool_wb is None:\n    {var}.save(DST_POOL)",
            )
            # close 守卫（如果存在）
            close_line = f"{var}.close()"
            if close_line in code:
                code = code.replace(
                    close_line,
                    f"if pool_wb is None:\n    {var}.close()",
                )

    # --- 2. 源表 read_only 优化 ---
    # 匹配加载源表（非 DST_POOL/DST_PCR）的 load_workbook，加 read_only=True
    # 但有些已经用了 data_only=True 或 read_only=True，需要跳过
    def add_read_only(m):
        full = m.group(0)
        if "read_only" in full or "data_only" in full:
            return full
        # 在最后一个 ) 前插入 , read_only=True
        return full[:-1] + ", read_only=True)"

    # 匹配 openpyxl.load_workbook( 非输出文件路径 )
    code = re.sub(
        r'openpyxl\.load_workbook\((?!DST_POOL|DST_PCR|pool_dst|pcr_dst)[^)]+\)',
        add_read_only,
        code,
    )

    # --- 3. 写入 ---
    filepath.write_text(code, encoding="utf-8")
    print(f"  ✓ {filename}")


def transform_xp_run_all(filepath: Path):
    """改造 XP run_all.py：统一加载一次。"""
    code = filepath.read_text(encoding="utf-8")

    new_main = '''def main():
    """优化版：统一加载输出工作簿一次，各步骤共享。"""
    import openpyxl

    # 步骤0：清空输出副本（独立运行，负责文件创建）
    clear_outputs()

    # ── 加载 Pooling 工作簿 ──
    pool_wb = openpyxl.load_workbook(DST_POOL)

    # 步骤1~8 + stats: 全部操作 Pooling
    step1(pool_wb=pool_wb)
    step2(pool_wb=pool_wb)
    step3(pool_wb=pool_wb)
    step4(pool_wb=pool_wb)
    step5(pool_wb=pool_wb)
    step6(pool_wb=pool_wb)
    step7(pool_wb=pool_wb)
    step8(pool_wb=pool_wb)
    pool_wb.save(DST_POOL)
    pool_wb.close()
    print("\\n[OK] 步骤1~8完成，已保存")

    # 下机数据统计模版（需要最新 Pooling 状态）
    pool2 = openpyxl.load_workbook(DST_POOL)
    step_stats(pool_wb=pool2)
    pool2.save(DST_POOL)
    pool2.close()
    print("\\n[OK] 下机数据统计模版完成，已保存")

    # 步骤9: qPCR 定量表（读 Pooling，写 PCR）
    pool3 = openpyxl.load_workbook(DST_POOL)
    pcr_wb = openpyxl.load_workbook(DST_PCR)
    step9(pool_wb=pool3, pcr_wb=pcr_wb)
    pool3.close()
    pcr_wb.save(DST_PCR)
    print("\\n[OK] 步骤9 qPCR定量表完成，已保存")

    # 格式化
    pool4 = openpyxl.load_workbook(DST_POOL)
    pcr2 = openpyxl.load_workbook(DST_PCR)
    format_outputs(pool_wb=pool4, pcr_wb=pcr2)
    pool4.save(DST_POOL)
    pcr2.save(DST_PCR)
    pool4.close()
    pcr2.close()
    print("\\n[OK] 格式化完成，已保存")

    # 校验（只读）
    pool5 = openpyxl.load_workbook(DST_POOL)
    validate_outputs(pool_wb=pool5)
    pool5.close()
    print("\\n[OK] 校验通过")

    print("\\n[DONE] XP 全流程执行完成")'''

    # 替换旧的 main() 函数
    code = re.sub(r'def main\(\).*?(?=\n\nif __name__|\nif __name__|\Z)', new_main, code, flags=re.DOTALL)

    filepath.write_text(code, encoding="utf-8")
    print("  ✓ run_all.py")


def main():
    if len(sys.argv) < 2:
        print("用法: python transform.py <XP|ZM>")
        sys.exit(1)

    project = sys.argv[1].upper()
    code_dir = Path(f"E:/HP_Project/{project}/code")
    if not code_dir.is_dir():
        print(f"目录不存在: {code_dir}")
        sys.exit(1)

    print(f"\n--- 改造 {project} ---")

    step_files = sorted(code_dir.glob("step*.py"))
    for f in step_files:
        transform_xp_step(f)

    for name in ["clear_pooling.py", "format_font.py", "validate_output.py"]:
        f = code_dir / name
        if f.is_file():
            transform_xp_step(f)

    run_all = code_dir / "run_all.py"
    if run_all.is_file():
        transform_xp_run_all(run_all)

    # 语法检查
    print("\n--- 语法检查 ---")
    import subprocess
    import py_compile
    for f in sorted(code_dir.glob("*.py")):
        try:
            py_compile.compile(str(f), doraise=True)
            print(f"  ✓ {f.name}")
        except py_compile.PyCompileError as e:
            print(f"  ✗ {f.name}: {e}")

    print(f"\n{DONE}")


if __name__ == "__main__":
    main()
