#!/usr/bin/env python3
"""统一检查并运行 PE100、PE150、XP 文库表流水线。"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[4]


@dataclass(frozen=True)
class TableConfig:
    project_dir: str
    required_keywords: tuple[str, ...]
    templates: tuple[str, ...]
    output_names: tuple[str, ...]
    requires_common: bool = False


CONFIGS = {
    "pe100": TableConfig(
        project_dir="PE_100",
        required_keywords=("上机", "质检", "自建库"),
        templates=("PE100_pooling模板表.xlsx",),
        output_names=("{date}文库pooling表T7+PE100.xlsx",),
        requires_common=True,
    ),
    "pe150": TableConfig(
        project_dir="PE_150",
        required_keywords=("上机", "质检", "自建库"),
        templates=("PE150_pooling模板表.xlsx",),
        output_names=("{date}文库pooling表T7+PE150.xlsx",),
        requires_common=True,
    ),
    "xp": TableConfig(
        project_dir="XP",
        required_keywords=("Xplus上机", "质检", "自建库", "纯化", "qPCR"),
        templates=("XP_pooling表模板.xlsx", "XP_PCR定量表模板.xlsx"),
        output_names=("{date}文库pooling表.xlsx", "{date}PCR定量表.xlsx"),
    ),
    "zhenmai": TableConfig(
        project_dir="XP",
        required_keywords=("Xplus上机", "质检", "自建库", "纯化", "qPCR"),
        templates=("XP_pooling表模板.xlsx", "XP_PCR定量表模板.xlsx"),
        output_names=("{date}文库pooling表.xlsx", "{date}PCR定量表.xlsx"),
    ),
}

# 别名映射（处理逻辑完全相同的型号）
ALIASES = {"zhenmai": "xp"}

COMMON_MODULES = (
    "__init__.py",
    "file_utils.py",
    "sheet_utils.py",
    "format_utils.py",
    "step_runner.py",
    "preflight.py",
    "validate_utils.py",
    "error_utils.py",
    "d_split.py",
)


def discover_sources(input_dir: Path, keywords: tuple[str, ...]) -> tuple[dict[str, Path], list[str]]:
    matches: dict[str, Path] = {}
    errors: list[str] = []
    if not input_dir.is_dir():
        return matches, [f"输入目录不存在: {input_dir}"]

    candidates = [
        path for path in input_dir.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    for keyword in keywords:
        found = sorted(
            (path for path in candidates if keyword.lower() in path.name.lower()),
            key=lambda path: path.name.lower(),
        )
        if not found:
            errors.append(f'缺少文件名包含“{keyword}”的 xlsx 文件')
        elif len(found) > 1:
            errors.append(
                f'关键词“{keyword}”匹配到多个文件: ' + "、".join(path.name for path in found)
            )
        else:
            matches[keyword] = found[0]
    return matches, errors


def preflight(table_type: str, input_dir: Path, output_dir: Path) -> tuple[TableConfig, list[Path]]:
    config = CONFIGS[table_type]
    project_dir = PROJECT_ROOT / config.project_dir
    errors: list[str] = []

    if not project_dir.is_dir():
        errors.append(f"项目目录不存在: {project_dir}")

    entry = project_dir / "code" / "run_all.py"
    if not entry.is_file():
        errors.append(f"入口脚本不存在: {entry}")

    for template in config.templates:
        template_path = project_dir / template
        if not template_path.is_file():
            errors.append(f"模板不存在: {template_path}")

    sources, source_errors = discover_sources(input_dir, config.required_keywords)
    errors.extend(source_errors)

    if config.requires_common:
        common_dir = PROJECT_ROOT / "common"
        missing = [name for name in COMMON_MODULES if not (common_dir / name).is_file()]
        if missing:
            errors.append(f"共享模块不完整: {common_dir} 缺少 " + "、".join(missing))

    if output_dir.exists() and not output_dir.is_dir():
        errors.append(f"输出路径不是目录: {output_dir}")

    if errors:
        print("[CHECK FAILED]")
        for error in errors:
            print(f"- {error}")
        raise SystemExit(2)

    date_str = datetime.now().strftime("%Y%m%d")
    outputs = [output_dir / name.format(date=date_str) for name in config.output_names]

    print("[CHECK OK]")
    print(f"- 表型: {table_type}")
    print(f"- 项目: {project_dir}")
    print(f"- 输入: {input_dir}")
    for keyword, source in sources.items():
        print(f"  - {keyword}: {source.name}")
    print(f"- 输出: {output_dir}")
    for output in outputs:
        print(f"  - {output.name}")
    return config, outputs


def run(table_type: str, input_dir: Path, output_dir: Path) -> list[Path]:
    config, outputs = preflight(table_type, input_dir, output_dir)
    existing = [output for output in outputs if output.exists()]
    if existing:
        print("[OUTPUT EXISTS]")
        for output in existing:
            print(f"- {output}")
        print("现有项目流程可能覆盖同名文件。请更换 --output-dir，或确认后追加 --overwrite。")
        raise SystemExit(3)

    output_dir.mkdir(parents=True, exist_ok=True)
    project_dir = PROJECT_ROOT / config.project_dir
    entry = project_dir / "code" / "run_all.py"
    env = os.environ.copy()
    env["HPLS_INPUT_DIR"] = str(input_dir)
    env["HPLS_OUTPUT_DIR"] = str(output_dir)

    print(f"[RUN] {entry}")
    completed = subprocess.run(
        [sys.executable, str(entry)],
        cwd=str(entry.parent),
        env=env,
        check=False,
    )
    if completed.returncode != 0:
        print(f"[FAILED] 流程返回码: {completed.returncode}")
        raise SystemExit(completed.returncode)

    missing_outputs = [output for output in outputs if not output.is_file()]
    if missing_outputs:
        print("[FAILED] 流程结束但未找到预期输出:")
        for output in missing_outputs:
            print(f"- {output}")
        raise SystemExit(4)

    print("[DONE]")
    for output in outputs:
        print(output)
    return outputs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成 HPLS PE100、PE150 或 XP 文库表")
    parser.add_argument("--type", required=True, choices=sorted(CONFIGS), dest="table_type")
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--check", action="store_true", help="仅执行前置检查")
    parser.add_argument("--overwrite", action="store_true", help="允许现有项目流程覆盖当天同名输出")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = CONFIGS[args.table_type]
    project_dir = PROJECT_ROOT / config.project_dir
    input_dir = (args.input_dir or project_dir / "input_data").resolve()
    output_dir = (args.output_dir or project_dir / "output_result").resolve()

    if args.check:
        preflight(args.table_type, input_dir, output_dir)
        return

    if args.overwrite:
        config_checked, outputs = preflight(args.table_type, input_dir, output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        entry = PROJECT_ROOT / config_checked.project_dir / "code" / "run_all.py"
        env = os.environ.copy()
        env["HPLS_INPUT_DIR"] = str(input_dir)
        env["HPLS_OUTPUT_DIR"] = str(output_dir)
        completed = subprocess.run([sys.executable, str(entry)], cwd=str(entry.parent), env=env)
        if completed.returncode != 0:
            raise SystemExit(completed.returncode)
        missing = [output for output in outputs if not output.is_file()]
        if missing:
            for output in missing:
                print(f"[MISSING OUTPUT] {output}")
            raise SystemExit(4)
        print("[DONE]")
        for output in outputs:
            print(output)
        return

    run(args.table_type, input_dir, output_dir)


if __name__ == "__main__":
    main()
