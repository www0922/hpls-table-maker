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
