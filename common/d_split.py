"""
D 列容量拆分工具
===============
通用 Best Fit Decreasing 装箱算法，保证每组 D 合计 ≤ 上限。

步骤三（PE150 step3_group_sort）和步骤四（PE150 step4_formula）共用此函数。
"""


def split_rows_by_d_capacity(rows, get_d, limit=1000):
    """按 D 列值进行容量拆分，保证每组 D 合计 ≤ limit。

    算法: Best Fit Decreasing
    1. 按 D 值从大到小排序
    2. 逐行放入"加入后最接近上限"的可容纳子组
    3. 所有现有子组都装不下时，新建子组
    4. 单条 D > limit 的记录单独成组

    Args:
        rows: 行数据列表（可以是 dict 或其他结构，由 get_d 提取 D 值）
        get_d: callable, 从单行提取 D 值（float 或可转为 float）
        limit: D 合计上限，默认 1000

    Returns:
        (groups, oversized):
          groups: [[row, ...], ...]  每个子组都是合规的（≤ limit）
          oversized: [row, ...]      单条 D > limit 的记录，各自需单独成组
    """
    if not rows:
        return [], []

    # 分离超限记录
    normal = []
    oversized = []
    for r in rows:
        d_val = _safe_d(get_d(r))
        if d_val > limit:
            oversized.append(r)
        else:
            normal.append(r)

    # Best Fit Decreasing
    sorted_rows = sorted(normal, key=lambda r: _safe_d(get_d(r)), reverse=True)

    groups = []       # [[row, ...], ...]
    group_sums = []   # [float, ...]

    for r in sorted_rows:
        d_val = _safe_d(get_d(r))

        # 在所有能容纳的子组中，选择 D 合计最大的（加入后最接近上限）
        candidates = [
            i for i, total in enumerate(group_sums)
            if total + d_val <= limit
        ]

        if candidates:
            best = max(candidates, key=lambda i: group_sums[i])
            groups[best].append(r)
            group_sums[best] += d_val
        else:
            groups.append([r])
            group_sums.append(d_val)

    return groups, oversized


def _safe_d(v):
    """安全提取 D 值，非法值返回 0（不参与容量判断但装入任意组）。"""
    if v is None:
        return 0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0
