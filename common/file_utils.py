"""
文件发现与路径管理
=================
合并 PE100 / PE150 / XP 三版本的最佳特性：

- find_file: 关键词搜索 + ~$ 排除 (PE150) + 多匹配报错 (XP)
- prepare_output / reset_output: PE100/PE150 完全相同的模板复制逻辑
"""

import shutil
from pathlib import Path


def find_file(directory, keyword, *, exclude_temp=True, unique=True):
    """在 directory 中搜索 ``*{keyword}*.xlsx`` 文件，返回绝对路径。

    Args:
        directory: 搜索目录 (Path 或 str)
        keyword: 文件名关键词
        exclude_temp: 排除 ``~$`` 开头的 Excel 临时文件（默认 True）
        unique: 多文件匹配时抛出 RuntimeError（默认 True）

    Returns:
        str: 匹配文件的绝对路径

    Raises:
        FileNotFoundError: 无匹配文件
        RuntimeError: unique=True 且匹配到多个文件
    """
    dir_path = Path(directory) if not isinstance(directory, Path) else directory
    matches = sorted(dir_path.glob(f'*{keyword}*.xlsx'), key=lambda p: p.name)

    if exclude_temp:
        matches = [m for m in matches if not m.name.startswith('~$')]

    if not matches:
        raise FileNotFoundError(
            f'{directory}/ 中没有包含 "{keyword}" 的 xlsx 文件'
        )

    if unique and len(matches) > 1:
        names = '、'.join(p.name for p in matches)
        raise RuntimeError(
            f'{directory}/ 中有多个包含 "{keyword}" 的文件，'
            f'请只保留本批次文件：{names}'
        )

    return str(matches[0])


def ensure_gitignore(output_dir):
    """确保输出目录中存在 .gitignore（防止 rm -rf 清掉后误提交结果表）。

    Args:
        output_dir: 输出目录 (Path 或 str)
    """
    gi = Path(output_dir) / '.gitignore'
    if not gi.exists():
        gi.write_text('*\n!.gitignore\n', encoding='utf-8')


def prepare_output(output_dir, template_path, output_name):
    """复制模板到输出目录（已存在则跳过，保留已有工作）。

    Args:
        output_dir: 输出目录 (Path)
        template_path: 模板文件路径 (Path)
        output_name: 输出文件名 (str)

    Returns:
        str: 输出文件的绝对路径
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    ensure_gitignore(output_dir)
    dst_path = output_dir / output_name
    if not dst_path.exists():
        shutil.copy2(template_path, dst_path)
        print(f'[OK] 已复制模板 -> {dst_path}')
    else:
        print(f'[OK] 工作副本已存在, 跳过复制 -> {dst_path}')
    return str(dst_path)


def reset_output(output_dir, template_path, output_name):
    """强制重新复制模板到输出目录（覆盖已有副本）。

    Args:
        output_dir: 输出目录 (Path)
        template_path: 模板文件路径 (Path)
        output_name: 输出文件名 (str)

    Returns:
        str: 输出文件的绝对路径
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    ensure_gitignore(output_dir)
    dst_path = output_dir / output_name
    shutil.copy2(template_path, dst_path)
    print(f'[OK] 已重置模板 -> {dst_path}')
    return str(dst_path)
