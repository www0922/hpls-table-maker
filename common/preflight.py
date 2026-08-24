"""
前置检查脚本
===========
在修改任何数据之前，一次性检查所有前置条件。

用法::

    python common/preflight.py              # 自动检测项目
    python common/preflight.py PE150        # 指定项目
    python common/preflight.py --list       # 列出可用项目

也可导入::

    from common.preflight import run_preflight, detect_project
    ok, checks = run_preflight('PE150')
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# ── 项目定义 ──────────────────────────────────────────────

PROJECTS = {
    'PE100': {
        'root':        PROJECT_ROOT / 'PE_100',
        'input_dir':   PROJECT_ROOT / 'PE_100' / 'input_data',
        'templates':   [('模板', PROJECT_ROOT / 'PE_100' / 'PE100_pooling模板表.xlsx')],
        'output_dir':  PROJECT_ROOT / 'PE_100' / 'output_result',
        'sources': [
            ('A表（上机数据）', ['上机']),
            ('C表（外包质检）', ['质检']),
            ('D表（自建库）',   ['自建库']),
        ],
    },
    'PE150': {
        'root':        PROJECT_ROOT / 'PE_150',
        'input_dir':   PROJECT_ROOT / 'PE_150' / 'input_data',
        'templates':   [('模板', PROJECT_ROOT / 'PE_150' / 'PE150_pooling模板表.xlsx')],
        'output_dir':  PROJECT_ROOT / 'PE_150' / 'output_result',
        'sources': [
            ('A表（上机数据）', ['上机']),
            ('C表（外包质检）', ['质检']),
            ('D表（自建库）',   ['自建库']),
        ],
    },
    'XP': {
        'root':        PROJECT_ROOT / 'XP',
        'input_dir':   PROJECT_ROOT / 'XP' / 'input_data',
        'templates': [
            ('Pooling模板', PROJECT_ROOT / 'XP' / 'XP_pooling表模板.xlsx'),
            ('PCR定量模板',  PROJECT_ROOT / 'XP' / 'XP_PCR定量表模板.xlsx'),
        ],
        'output_dir': PROJECT_ROOT / 'XP' / 'output_result',
        'sources': [
            ('A表（上机数据）', ['Xplus上机']),
            ('C表（外包质检）', ['质检']),
            ('D表（自建库）',   ['自建库']),
            ('纯化总表',        ['纯化']),
            ('qPCR定量统计',   ['qPCR']),
        ],
    },
}


# ── 检查函数 ──────────────────────────────────────────────

def _find_files(directory, keywords):
    """在目录中搜索包含任一关键词的 xlsx 文件，排除临时文件。"""
    if not directory.exists():
        return []
    results = []
    for kw in keywords:
        results.extend(
            p for p in directory.glob(f'*{kw}*.xlsx')
            if not p.name.startswith('~$')
        )
    # 去重（同一文件可能匹配多个关键词），保持排序
    seen = set()
    unique = []
    for p in sorted(results, key=lambda x: x.name):
        if p.name not in seen:
            seen.add(p.name)
            unique.append(p)
    return unique


def _check(label, passed, message, detail=None):
    return {'label': label, 'passed': passed, 'message': message, 'detail': detail}


def _check_source(label, keywords, input_dir):
    """检查单个源数据文件。"""
    matches = _find_files(input_dir, keywords)
    if not matches:
        kw_str = ' / '.join(keywords)
        return _check(
            label, False,
            f'未找到文件（关键词: {kw_str}）',
            f'请在 {input_dir}/ 中放入包含上述关键词的 .xlsx 文件'
        )
    if len(matches) > 1:
        names = '、'.join(p.name for p in matches)
        return _check(
            label, False,
            f'匹配到多个文件: {names}',
            f'请只保留一个本批次文件，移除多余的'
        )
    return _check(label, True, matches[0].name)


def _check_template(label, path):
    """检查模板文件是否存在。"""
    if path.exists():
        return _check(label, True, path.name)
    return _check(label, False, f'模板不存在: {path.name}', f'请确认 {path} 文件未被删除或改名')


def _check_a_sheets(project_name, cfg):
    """检查 A 表的 sheet 结构是否可解析。"""
    import openpyxl

    input_dir = cfg['input_dir']
    if project_name == 'XP':
        matches = _find_files(input_dir, ['Xplus上机'])
    else:
        matches = _find_files(input_dir, ['上机'])

    if not matches or len(matches) > 1:
        return _check('A表sheet结构', None, '跳过（A表未就绪）')

    try:
        wb = openpyxl.load_workbook(str(matches[0]), data_only=True)
        sheets = list(wb.sheetnames)
        wb.close()

        # 尝试映射 sheet
        mapped = []
        for sn in sheets:
            try:
                num = int(sn)
                letter = chr(ord('A') + num - 1)
                mapped.append(f'{sn}→{letter}')
            except ValueError:
                if len(sn) == 1 and sn.isalpha() and sn.isascii() and sn.isupper():
                    mapped.append(sn)
        if not mapped:
            return _check(
                'A表sheet结构', False,
                f'无法识别有效的 sheet 名: {sheets}',
                '请确认是标准的华大 T7+ 上机表（sheet 为数字或单字母）'
            )
        return _check(
            'A表sheet结构', True,
            f'{len(mapped)} 个 sheet ({", ".join(mapped)})'
        )
    except Exception as e:
        return _check('A表sheet结构', False, f'读取失败: {e}')


def _check_dependency():
    """检查 Python 依赖。"""
    try:
        import openpyxl
        # 尝试获取版本号
        version = getattr(openpyxl, '__version__', '已安装')
        return _check('Python依赖', True, f'openpyxl {version}')
    except ImportError:
        return _check(
            'Python依赖', False,
            'openpyxl 未安装',
            '请运行: pip install openpyxl'
        )


def _check_output_dir(output_dir):
    """检查输出目录是否可写。"""
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        test_file = output_dir / '._preflight_write_test'
        test_file.touch()
        test_file.unlink()
        return _check('输出目录', True, str(output_dir))
    except (OSError, PermissionError) as e:
        return _check('输出目录', False, f'无法写入: {e}')


# ── 公开 API ──────────────────────────────────────────────

def detect_project():
    """自动检测当前应该运行哪个项目。

    通过扫描各项目 input_data 目录中有无 .xlsx 文件来判断。

    Returns:
        str | None | list[str]: 唯一项目名 / None（无数据）/ 列表（多个候选）
    """
    found = []
    for name, cfg in PROJECTS.items():
        input_dir = cfg['input_dir']
        if not input_dir.exists():
            continue
        xlsx_files = [
            f for f in input_dir.glob('*.xlsx')
            if not f.name.startswith('~$')
        ]
        if xlsx_files:
            found.append(name)
    if len(found) == 1:
        return found[0]
    if len(found) == 0:
        return None
    return found


def run_preflight(project):
    """运行指定项目的前置检查。

    Args:
        project: 'PE100' | 'PE150' | 'XP'

    Returns:
        (all_passed: bool, checks: list[dict])
        每个 check: {'label', 'passed', 'message', 'detail'}
        passed 可为 True / False / None（跳过）
    """
    if project not in PROJECTS:
        return False, [_check(
            '项目识别', False,
            f'未知项目 "{project}"，可选: {", ".join(PROJECTS)}'
        )]

    cfg = PROJECTS[project]
    checks = []

    # 1. 输入目录存在
    if not cfg['input_dir'].exists():
        checks.append(_check(
            '输入目录', False,
            f'{cfg["input_dir"]} 不存在',
            '请确认项目目录结构完整'
        ))
        return False, checks
    checks.append(_check('输入目录', True, str(cfg['input_dir'])))

    # 2. 源数据文件
    for label, keywords in cfg['sources']:
        checks.append(_check_source(label, keywords, cfg['input_dir']))

    # 3. 模板文件
    for label, path in cfg['templates']:
        checks.append(_check_template(label, path))

    # 4. A表 sheet 结构（依赖 A表已就绪）
    checks.append(_check_a_sheets(project, cfg))

    # 5. Python 依赖
    checks.append(_check_dependency())

    # 6. 输出目录
    checks.append(_check_output_dir(cfg['output_dir']))

    all_ok = all(c['passed'] is not False for c in checks)
    return all_ok, checks


# ── CLI ───────────────────────────────────────────────────

def _print_results(project, checks, all_ok):
    """美化打印检查结果。"""
    print(f'\n{"─" * 50}')
    print(f'  前置检查: {project}')
    print(f'{"─" * 50}')

    for c in checks:
        passed = c['passed']
        if passed is True:
            icon = '[OK]'
        elif passed is False:
            icon = '[FAIL]'
        else:
            icon = '[SKIP]'

        print(f'  {icon} {c["label"]:<20s}  {c["message"]}')
        if c.get('detail'):
            print(f'       -> {c["detail"]}')

    print(f'{"─" * 50}')
    failed = sum(1 for c in checks if c['passed'] is False)
    skipped = sum(1 for c in checks if c['passed'] is None)
    if all_ok and not skipped:
        print(f'  结果: 全部通过 [OK]  可以开始执行 pipeline\n')
    elif all_ok:
        print(f'  结果: {failed} 失败, {skipped} 跳过 — 部分检查未完成\n')
    else:
        print(f'  结果: {failed} 个问题 [FAIL]  请修复后重试\n')


def main():
    import argparse

    parser = argparse.ArgumentParser(description='HPLS 前置检查')
    parser.add_argument(
        'project', nargs='?', default=None,
        help=f'项目名: {", ".join(PROJECTS)}（省略则自动检测）'
    )
    parser.add_argument(
        '--list', action='store_true',
        help='列出所有可用项目'
    )
    args = parser.parse_args()

    if args.list:
        print('可用项目:')
        for name, cfg in PROJECTS.items():
            print(f'  {name}: {cfg["root"]}')
        return

    project = args.project
    if project is None:
        print('  自动检测项目...')
        detected = detect_project()
        if detected is None:
            print('  [FAIL] 未在任何 input_data/ 目录中找到数据文件')
            print(f'  请将源数据放入对应项目的 input_data/ 目录')
            print(f'  可用项目: {", ".join(PROJECTS)}')
            sys.exit(1)
        if isinstance(detected, list):
            print(f'  [FAIL] 多个项目有数据: {", ".join(detected)}')
            print(f'  请指定项目: python common/preflight.py <项目名>')
            sys.exit(1)
        project = detected

    all_ok, checks = run_preflight(project)
    _print_results(project, checks, all_ok)
    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()
