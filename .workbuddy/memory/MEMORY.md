# HPLS 文库表项目记忆

## 项目概述
- 四套文库 Pooling 表自动化生成系统：PE100、PE150、XP、真迈（ZM）
- 输入：上机表、质检表、自建库表等 Excel 源文件
- 输出：文库 Pooling 表、PCR 定量表（XP/真迈）

## 使用规范
- 完成后不要调用 present_files 打开预览，用户自己去文件夹打开表格
- 每个关键词必须唯一匹配一个源文件，多个匹配时拒绝执行

## 统一 Skill
- 位置：`.workbuddy/skills/hpls-table-maker/`
- 统一入口：`scripts/run_table.py`
- 触发：用户说"做 PE100/PE150/XP/真迈表" → WorkBuddy 自动路由到对应项目入口

## 故障排查
- PE100/PE150 依赖 `E:\HP_Project\common/` 共享模块，缺失时前置检查会明确列出缺少的文件
- XP/真迈 无共享模块依赖，可直接运行

## 性能优化
- XP/真迈 已优化为统一加载模式（步骤1~8 共享 Workbook）
- PE100/PE150 保持原始逐步骤加载模式

## 业务规则（重要）
- **XP 标记逻辑**（step7_mark.py）：G 列状态标记优先级为
  1. 单文库组 + P列有值（已定量）→ 无条件标「已定量」（最高优先，压过纯化）
  2. 纯化过（U列"纯化"或 W列有B标记）→ 标「纯化」
  - 多文库组**不标「已定量」**（P列有值只发生在单文库组，多文库出现 P 值属异常）
- PE100/PE150 T7+制备表 Q 列公式已改为 `=L/I`
- **pooling 表分组机制（四项目通用，重要）**
  - 判定条件完全一致：`is_summary_row` = B 列为纯数字 且 D 列有值；`is_blank_row` = 业务列全 None
  - `read_groups` 三条分支（XP/ZM 的 pooling_utils 与 PE100/PE150 的 common/validate_utils 逻辑同构）：
    ① 空白行 → 把 pending 里的行各自当作单文库组 flush；② 汇总行 → pending 收成一个多文库组；③ B 列有值 → 累积进 pending
  - **关键差异：单文库组是否生成汇总行**
    - PE100/PE150：**单文库组也生成汇总行**（B=1），所以每组都有汇总行收尾 → **不需要空白行**
    - XP/ZM：**单文库组不生成汇总行**（`is_summary_row` 注释即"识别多文库组汇总行"）→ **必须靠空白行**界定组尾
  - 因此：XP/ZM 表结构为 数据行(N)+汇总行+空白行、单文库组为 数据行(1)+空白行；
    PE100/PE150 为 数据行(N)+汇总行，紧接下一组
  - ⚠️ **XP/ZM 的空白行不能手工删除**：删了会导致连续的单文库组被误并成一个多文库组

## Git 约定
- 远程已切换 SSH（`git@github.com:www0922/hpls-table-maker.git`），推送免密
- `input_data/` 和 `output_result/` 用目录内 `.gitignore`（内容 `*\n!.gitignore`）保留空目录并隔离数据
- 各项目 config 的 reset/prepare_output 已加 `ensure_gitignore` 保护，防止 rm -rf 误删 .gitignore

## Git 历史清理记录（2026-08-24 已完成）
- 问题：早期提交把客户源表（质检总表/自建库报告/纯化总表等 22 个 xlsx）提交进历史，.git 达 37MB，且仓库 public
- 处理：先重建本地干净仓库（filter-repo 在本机会因 .git 文件锁丢 objects，已失败两次）→ 用户网页删库重建同名空仓库 → push
- 结果：.git 37MB → 1.2MB；远程 93 文件（67 py + 11 md + 6 模板表）；
  旧提交全部 HTTP 422 不存在，客户源表已彻底不可访问
- 仓库当前仍为 **public**（用户选择不改私有）
- 本机禁忌：**不要在 E:\HP_Project 做 rebase / filter-repo 等历史重写**，会因文件锁丢对象；
  必须先备份 .git，优先用"重建仓库 + 删库重推"方案
- 无 gh CLI / GITHUB_TOKEN：仓库的删除/创建等管理操作必须由用户在网页完成，SSH key 只能 push/pull
- 第二次删库重建（2026-08-24）：历史压成单提交 `03ac6d5`，清除含客户数据的旧模板版本；
  仓库现为**单提交干净历史**，clone 约 1.1MB，无任何客户数据/内网路径/外部链接

## 模板表卫生规范（2026-08-24 建立）
- **模板表应只保留表头**，不留任何数据行：因为 `clear_pooling.py` 对所有 sheet 执行
  `delete_rows(2, max_row-1)`，模板里的数据/公式反正会被清空，留着只会夹带旧数据
- **警惕 xlsx 内嵌 externalLinks**：PE100 模板曾内嵌 4.5MB 外部链接缓存，
  含 2024 年样本编号/浓度/检测人 + 内网路径 `C:\hgc\生物技术中心\...`，且会复制进每份输出表
- 清理后：模板 704KB→10.4KB，输出表 631KB→31KB，业务输出逐单元格零差异
- 清理工具：`E:\HP_backup_20260824\clean_external_links.py`（zipfile 精确重建，保留其他条目字节）
- **新模板入库前必查**：① 有无 `xl/externalLinks/` ② 各 sheet 是否只有表头
  ③ 体积是否异常（正常应 10~45KB，超过 100KB 需排查）
