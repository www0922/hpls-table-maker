"""
流水线编排器
===========
通用步骤执行器，提供：
- 顺序执行 + 进度打印
- 失败中止 / 继续执行
- 统一的输出格式
"""


class StepError(Exception):
    """单步执行失败异常，携带步骤名和原始异常。"""

    def __init__(self, step_name, original_error):
        self.step_name = step_name
        self.original_error = original_error
        super().__init__(f'步骤 "{step_name}" 执行失败: {original_error}')


def run_pipeline(steps, *, stop_on_error=True, title=None):
    """顺序执行步骤列表。

    Args:
        steps: ``[(步骤名, callable), ...]`` 列表。
               每个 callable 无参数，无返回值（或返回值被忽略）。
        stop_on_error: True → 任一步失败立即中止并抛出 StepError
                       False → 记录错误，继续执行后续步骤
        title: 项目标题，如 ``'PE150 Pooling 全流程'``。
               在开始和结束时打印。

    Raises:
        StepError: stop_on_error=True 且有步骤失败时

    输出格式::

        ============================================================
        [1/8] 清空模板
        ============================================================
        ... 步骤内部输出 ...

        [FAIL] 步骤 "数据迁移" 执行失败: ...
    """
    total = len(steps)

    if title:
        print(f'\n{"=" * 60}')
        print(f'  {title}')
        print(f'{"=" * 60}')

    errors = []

    for index, (label, func) in enumerate(steps, start=1):
        print(f'\n{"=" * 60}')
        print(f'[{index}/{total}] {label}')
        print(f'{"=" * 60}')

        try:
            func()
        except Exception as e:
            from .error_utils import translate
            friendly = translate(e)
            print(f'\n[FAIL] 步骤 "{label}" 执行失败:\n{friendly}')
            if stop_on_error:
                raise StepError(label, e) from e
            errors.append((label, e))

    if errors:
        failed_names = ', '.join(name for name, _ in errors)
        print(f'\n[WARN] 以下步骤失败: {failed_names}')
    else:
        suffix = f' — {title}' if title else ''
        print(f'\n[DONE] 全流程执行完成{suffix}')
