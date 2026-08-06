# HPLS 文库表项目记忆

## 项目概述
- 三套文库 Pooling 表自动化生成系统：PE100、PE150、XP
- 输入：上机表、质检表、自建库表等 Excel 源文件
- 输出：文库 Pooling 表、PCR 定量表（XP）

## 统一 Skill
- 位置：`.workbuddy/skills/hpls-table-maker/`
- 打包：`.workbuddy/skills/hpls-table-maker.zip`
- 统一入口：`scripts/run_table.py`
- 触发：用户说"做 PE100/PE150/XP 表" → WorkBuddy 自动路由到对应项目入口

## 故障排查
- PE100/PE150 依赖 `E:\HP_Project\common/` 共享模块，缺失时前置检查会明确列出缺少的文件
- XP 无共享模块依赖，可直接运行
- 每个关键词必须唯一匹配一个源文件，多个匹配时拒绝执行
