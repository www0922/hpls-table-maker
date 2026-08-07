# HPLS 表型输入输出契约

## PE100

- 项目目录：`{PROJECT_ROOT}/PE_100`
- 入口：`code/run_all.py`
- 输入目录默认值：`input_data/`
- 必需关键词：`上机`、`质检`、`自建库`
- 模板：`PE100_pooling模板表.xlsx`
- 输出：`YYYYMMDD文库pooling表T7+PE100.xlsx`
- 环境变量：`HPLS_INPUT_DIR`、`HPLS_OUTPUT_DIR`
- 额外依赖：项目根目录 `common/`

## PE150

- 项目目录：`{PROJECT_ROOT}/PE_150`
- 入口：`code/run_all.py`
- 输入目录默认值：`input_data/`
- 必需关键词：`上机`、`质检`、`自建库`
- 模板：`PE150_pooling模板表.xlsx`
- 输出：`YYYYMMDD文库pooling表T7+PE150.xlsx`
- 环境变量：`HPLS_INPUT_DIR`、`HPLS_OUTPUT_DIR`
- 额外依赖：项目根目录 `common/`

## XP

- 项目目录：`{PROJECT_ROOT}/XP`
- 入口：`code/run_all.py`
- 输入目录默认值：`input_data/`
- 必需关键词：`Xplus上机`、`质检`、`自建库`、`纯化`、`qPCR`
- 模板：`XP_pooling表模板.xlsx`、`XP_PCR定量表模板.xlsx`
- 输出：`YYYYMMDD文库pooling表.xlsx`、`YYYYMMDDPCR定量表.xlsx`
- 环境变量：`HPLS_INPUT_DIR`、`HPLS_OUTPUT_DIR`
- 额外依赖：无项目级共享包依赖

## 真迈

- 项目目录：`{PROJECT_ROOT}/ZM`
- 入口：`code/run_all.py`
- 输入目录默认值：`input_data/`
- 必需关键词：`江西上机`、`质检`、`自建库`、`纯化`、`qPCR`
- 模板：`ZM_pooling表模板.xlsx`、`ZM_PCR定量表模板.xlsx`
- 输出：`YYYYMMDD文库pooling表AE0.xlsx`、`YYYYMMDDPCR定量表真迈.xlsx`
- 环境变量：`HPLS_INPUT_DIR`、`HPLS_OUTPUT_DIR`
- 额外依赖：无项目级共享包依赖

## 执行约束

- 每个关键词必须唯一匹配一个 `.xlsx` 文件。
- 排除以 `~$` 开头的 Excel 临时文件。
- 统一使用当前 Python 解释器执行各型号的 `run_all.py`。
- 仅在返回码为 0 且预期输出均存在时判定成功。
- PE100/PE150 当前代码引用的共享模块包括：`file_utils.py`、`sheet_utils.py`、`format_utils.py`、`step_runner.py`、`validate_utils.py`、`d_split.py` 等。
