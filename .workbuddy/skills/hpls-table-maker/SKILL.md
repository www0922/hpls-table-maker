---
name: hpls-table-maker
description: 统一生成 HPLS 项目的 PE100、PE150、XP 和真迈文库表。当用户说"帮我做 PE100 表""生成 PE150 pooling 表""做 XP 表""做真迈表"，或提供相关 Excel 源表并要求生成文库 pooling/PCR 定量表时使用。识别目标表型，校验本批次输入文件，调用对应项目的 run_all.py 全流程，并返回生成的 Excel 文件。
agent_created: true
---

# HPLS 文库表统一生成

## 前置条件

本 Skill 是 HPLS 项目的包装层，依赖项目代码（`PE_100/`、`PE_150/`、`XP/`、`ZM/`、`common/`）才能运行。Skill 通过自身在项目内的位置自动探测根目录，因此：

- **自己用**：无需额外操作，Skill 在 `.workbuddy/skills/` 下即可工作。
- **同事用**：需要将整个项目（含 `.workbuddy/skills/hpls-table-maker/`）拷贝到任意路径，项目根目录自动探测，不依赖固定盘符。

## 目标

将 PE100、PE150、XP、真迈四套现有自动化代码作为一个统一能力使用。根据用户话语识别表型，调用对应入口，完成源表校验、表格生成、输出校验和结果交付。

## 表型路由

按以下优先级识别目标：

1. 用户提到 `PE100`、`PE 100`、`100 测序`：选择 `pe100`。
2. 用户提到 `PE150`、`PE 150`、`150 测序`：选择 `pe150`。
3. 用户提到 `XP`、`X Plus`、`Xplus`、`NovaSeq X Plus`：选择 `xp`。
4. 用户提到 `真迈`、`真迈表`、`ZM`：选择 `zhenmai`。
5. 无法唯一判断时，只追问一次："需要做 PE100、PE150、XP 还是真迈表？"

不要根据文件内容猜测型号后静默执行。

## 输入要求

- PE100：恰好各有一个文件名包含“上机”“质检”“自建库”的 `.xlsx` 文件。
- PE150：恰好各有一个文件名包含“上机”“质检”“自建库”的 `.xlsx` 文件。
- XP：恰好各有一个文件名包含"Xplus上机""质检""自建库""纯化""qPCR"的 `.xlsx` 文件。
- 真迈：恰好各有一个文件名包含"江西上机""质检""自建库""纯化""qPCR"的 `.xlsx` 文件。
- 忽略以 `~$` 开头的 Excel 临时文件。
- 发现缺失或同一关键词匹配多个文件时，停止执行并向用户列出具体问题；不要自行挑选批次。

若用户提供了本批次文件所在目录，将该目录作为 `--input-dir`。否则使用项目对应的 `input_data/`。不要把附件复制进项目目录，直接通过输入目录运行。

## 执行流程

### 1. 前置检查

先运行统一入口的检查模式（路径由 Skill 脚本自身定位，无需硬编码）：

```bash
python scripts/run_table.py --type <pe100|pe150|xp|zhenmai> --check
```

使用自定义源表目录时追加：

```bash
--input-dir "<源表目录>"
```

使用自定义输出目录时追加：

```bash
--output-dir "<输出目录>"
```

> 实际调用时 WorkBuddy 会自动使用项目对应的 Python 解释器，在 Skill 目录（`.workbuddy/skills/hpls-table-maker/`）下执行上述命令。

### 2. 处理前置检查结果

- 检查通过：继续执行。
- 缺少源表或存在重复匹配：把报告原样概括给用户，等待用户补齐或明确删除哪一份；不要代替用户删除文件。
- 缺少 `common/`：说明 PE100/PE150 的共享模块不完整，列出缺失模块，停止执行。XP 不依赖该目录。
- 模板缺失：说明缺少的模板文件路径，停止执行。

### 3. 运行对应流程

```bash
python scripts/run_table.py --type <pe100|pe150|xp|zhenmai>
```

按需追加 `--input-dir` 和 `--output-dir`。统一入口会设置 `HPLS_INPUT_DIR`、`HPLS_OUTPUT_DIR`，再以子进程调用：

- PE100：`PE_100/code/run_all.py`
- PE150：`PE_150/code/run_all.py`
- XP：`XP/code/run_all.py`
- 真迈：`ZM/code/run_all.py`

不要逐个调用步骤脚本，以免遗漏顺序和最终校验。

### 4. 交付结果

- 仅在子进程返回码为 0 且输出文件存在时报告成功。
- PE100 返回当天的 `文库pooling表T7+PE100.xlsx`。
- PE150 返回当天的 `文库pooling表T7+PE150.xlsx`。
- XP 同时返回当天的 `文库pooling表.xlsx` 和 `PCR定量表.xlsx`。
- 真迈同时返回当天的 `文库pooling表AE0.xlsx` 和 `PCR定量表真迈.xlsx`。
- 使用文件展示能力把所有生成文件一次性交付给用户。
- 若失败，提取最后一段错误信息，说明失败步骤和缺失项，不要声称已生成。

## 保护规则

- 不修改原始模板。
- 不删除、不重命名用户源表。
- 不在多个候选源表间自行选择。
- 不绕过现有 `validate_output.py` 校验。
- 同一天重复运行可能覆盖同名输出；执行前若目标已存在，明确告知用户并取得确认，或使用新的输出目录。

## 参考资料

需要查看输入关键词、模板和输出契约时，读取 `references/table_contracts.md`。
