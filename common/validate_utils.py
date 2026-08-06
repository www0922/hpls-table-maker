"""
输出校验工具
===========
PE100 / PE150 共用的校验逻辑。
"""

from pathlib import Path


def is_summary_row(ws, row):
    """判断是否为汇总行：B 列为纯数字（组内条数），且 D 列有值。"""
    b_val = str(ws.cell(row=row, column=2).value or '').strip()
    d_val = ws.cell(row=row, column=4).value
    return b_val.isdigit() and d_val is not None


def is_blank_row(ws, row, max_col=20):
    """判断该行在业务列范围内是否全为空。"""
    return all(
        ws.cell(row=row, column=col).value is None
        for col in range(1, max_col + 1)
    )


def read_groups(ws, max_col=20):
    """从数据 sheet 读取分组结构。

    约定：
    - 连续的数据行 + 1 个汇总行 = 1 组
    - 汇总行：B 列为纯数字（组大小），D 列有值
    - 空白行分隔不同的 lane

    Returns:
        list[dict]: [
            {
                'data_rows': [row_idx, ...],  # 数据行号列表
                'summary_row': int | None,     # 汇总行号
                'lane_name': str,              # A 列 lane 名称（如 A1）
            }, ...
        ]
    """
    groups = []
    pending = []
    max_row = ws.max_row

    for row in range(2, max_row + 1):
        if is_blank_row(ws, row, max_col):
            # 空白行：清空 pending（单样本兜底）
            for data_row in pending:
                lane = str(ws.cell(row=data_row, column=1).value or '').strip()
                groups.append({
                    'data_rows': [data_row],
                    'summary_row': None,
                    'lane_name': lane,
                })
            pending = []
            continue

        if is_summary_row(ws, row):
            if pending:
                lane = str(ws.cell(row=pending[0], column=1).value or '').strip()
                groups.append({
                    'data_rows': list(pending),
                    'summary_row': row,
                    'lane_name': lane,
                })
            pending = []
            continue

        # 有 B 列值 = 数据行
        if ws.cell(row=row, column=2).value is not None:
            pending.append(row)

    # 尾部未收尾的数据行 → 单样本
    for data_row in pending:
        lane = str(ws.cell(row=data_row, column=1).value or '').strip()
        groups.append({
            'data_rows': [data_row],
            'summary_row': None,
            'lane_name': lane,
        })

    return groups


def validate_data_sheet(ws, sheet_letter):
    """校验单个数据 sheet 的结构。

    检查项：
    - 有数据（组数 > 0）
    - 每组有汇总行（单样本除外）
    - Lane 编号连续（如 A1, A2, A3...）
    - 组间无空行
    - 汇总行 B 列数字与组内行数一致
    - 组内 B 列（文库编号）无重复

    Returns:
        (errors: list[str], warnings: list[str], group_count: int)
    """
    errors = []
    warnings = []
    groups = read_groups(ws)

    if not groups:
        errors.append(f'Sheet [{sheet_letter}]: 未识别到任何分组')
        return errors, warnings, 0

    # ── Lane 编号连续性 ──
    expected_seq = 1
    for g in groups:
        name = g['lane_name']
        if name.startswith(sheet_letter):
            try:
                seq = int(name[len(sheet_letter):])
                if seq != expected_seq:
                    errors.append(
                        f'Lane {name}: 编号不连续，期望 {sheet_letter}{expected_seq}'
                    )
                expected_seq = max(expected_seq, seq) + 1
            except ValueError:
                warnings.append(f'Lane {name}: 无法解析编号，跳过连续性检查')

    # ── 逐组校验 ──
    prev_summary_row = None
    for g in groups:
        rows = g['data_rows']
        summary = g['summary_row']
        lane = g['lane_name']
        first_data_row = rows[0]

        # 检查组内 B 列无重复
        b_values = [
            str(ws.cell(row=r, column=2).value or '').strip()
            for r in rows
        ]
        b_values = [v for v in b_values if v]  # 排除空值
        if len(b_values) != len(set(b_values)):
            duplicates = [v for v in b_values if b_values.count(v) > 1]
            errors.append(
                f'{lane} 行{first_data_row}: B列(文库编号)重复 {list(set(duplicates))}'
            )

        # 多样本组必须有汇总行
        if len(rows) > 1 and summary is None:
            errors.append(
                f'{lane} 行{first_data_row}: 多样本组({len(rows)}行)缺少汇总行'
            )

        # 汇总行 B 列数字应与组内行数一致
        # 注：PE100/PE150 所有组（含单样本组）均有汇总行，这是正常设计
        if summary is not None:
            try:
                count_in_summary = int(str(ws.cell(row=summary, column=2).value or '0'))
                if count_in_summary != len(rows):
                    errors.append(
                        f'{lane} 汇总行{summary}: B列={count_in_summary}，'
                        f'但组内有{len(rows)}行数据'
                    )
            except ValueError:
                pass

        # 检查组间无空行（多组连续时，上一组汇总行下一行应直接是本组首数据行）
        if prev_summary_row is not None:
            gap = first_data_row - prev_summary_row
            if gap > 1:
                errors.append(
                    f'{lane} 行{first_data_row}: 与前一组汇总行({prev_summary_row})'
                    f'之间有{gap - 1}行空隙'
                )

        prev_summary_row = summary if summary else rows[-1]

    return errors, warnings, len(groups)


def check_sheet_order(wb, expected_order):
    """校验 sheet 排列顺序。

    Args:
        wb: openpyxl Workbook
        expected_order: 期望的顺序列表，如 ['A', 'B', 'T7+制备']

    Returns:
        list[str]: 错误信息
    """
    errors = []
    actual_names = list(wb.sheetnames)

    # 检查期望的 sheet 是否存在
    for sn in expected_order:
        if sn not in actual_names:
            errors.append(f'缺少 sheet: [{sn}]')

    # 检查期望 sheet 的排列顺序
    expected_in_actual = [s for s in expected_order if s in actual_names]
    indices = {sn: actual_names.index(sn) for sn in expected_in_actual}
    for i in range(len(expected_in_actual) - 1):
        if indices[expected_in_actual[i]] > indices[expected_in_actual[i + 1]]:
            errors.append(
                f'Sheet 顺序错误: [{expected_in_actual[i]}] 应在 '
                f'[{expected_in_actual[i + 1]}] 之前'
            )

    return errors


def check_t7_sheet(ws):
    """校验 T7+制备 sheet 是否有数据。

    Returns:
        list[str]: 错误信息
    """
    errors = []
    if ws.max_row < 2:
        errors.append('T7+制备: sheet 为空')
        return errors

    # 至少有 laneID（A列）和文库编号（B列）
    has_data = False
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=2).value is not None:
            has_data = True
            break
    if not has_data:
        errors.append('T7+制备: 无数据（B列全空）')
    return errors
