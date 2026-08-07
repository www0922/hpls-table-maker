# HP Table Process

一键生成 HPLS 实验室 PE100、PE150、XP（含真迈）三种测序文库 Pooling 表，通过 WorkBuddy Skill 统一调度，说句话就能跑。

## 前置条件

- **Python 3.x** + `openpyxl`（`pip install openpyxl`）
- **WorkBuddy**（Skill 运行环境）
- 对应型号的本批次源表（上机表、质检表、自建库表等）

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/www0922/hpls-table-maker.git
cd hpls-table-maker
```

### 2. 放入源数据

将本批次的源表放入对应项目的 `input_data/` 目录：

| 型号 | 目录 | 需要文件 |
|------|------|---------|
| PE100 | `PE_100/input_data/` | 上机表、质检表、自建库表（各一个） |
| PE150 | `PE_150/input_data/` | 上机表、质检表、自建库表（各一个） |
| XP | `XP/input_data/` | Xplus上机表、质检表、自建库表、纯化表、qPCR表（各一个） |
| 真迈 | `ZM/input_data/` | 江西上机表、质检表、自建库表、纯化表、qPCR表（各一个） |

### 3. 打开 WorkBuddy，说句话就行

```
帮我做 PE100 表
帮我做 PE150 表
帮我做 XP 表
帮我做真迈表
```

Skill 会自动完成：
- 检查源文件是否齐全
- 运行完整流水线（数据迁移 → 分组排序 → 公式计算 → 格式化 → 校验）
- 在 `output_result/` 生成最终 Excel

### 4. 不想用 Skill？直接命令行跑也行

```bash
cd PE_100/code
python run_all.py
```

```bash
cd PE_150/code
python run_all.py
```

```bash
cd XP/code
python run_all.py
```

## 项目结构

```
hpls-table-maker/
├── PE_100/                  # PE100 型代码 + 模板 + 源表
│   ├── code/
│   │   ├── run_all.py       # 一键全流程
│   │   ├── step1~6_*.py     # 分步脚本
│   │   └── config.py        # 路径配置
│   ├── input_data/          # 源数据（按批次替换）
│   └── output_result/       # 输出（自动生成，不提交 Git）
├── PE_150/                  # PE150 型（结构同上）
├── XP/                      # XP 型（结构同上，输出含 PCR 定量表）
├── ZM/                      # 真迈型（结构同上，输出含 PCR 定量表）
├── common/                  # PE100/PE150 共享工具模块
├── .workbuddy/skills/
│   └── hpls-table-maker/   # WorkBuddy 统一 Skill
│       ├── SKILL.md         # Skill 入口说明
│       └── scripts/
│           └── run_table.py # 统一路由脚本
└── README.md
```

### 输出文件

| 型号 | 输出文件 |
|------|---------|
| PE100 | `output_result/YYYYMMDD文库pooling表T7+PE100.xlsx` |
| PE150 | `output_result/YYYYMMDD文库pooling表T7+PE150.xlsx` |
| XP | `output_result/YYYYMMDD文库pooling表.xlsx` + `PCR定量表.xlsx` |
| 真迈 | `output_result/YYYYMMDD文库pooling表AE0.xlsx` + `PCR定量表真迈.xlsx` |

## 同步更新

项目代码和 Skill 统一纳入 Git，同事只需 `git pull` 即可获取最新版本：

```bash
# 自己更新后推送
git add -A
git commit -m "更新内容说明"
git push

# 同事拉取最新
git pull
```

## 注意事项

- **源表必须唯一**：每个关键词（如"上机"）只匹配一个文件，多个匹配会报错。同一目录只保留当前批次。
- **模板只读**：源码不修改原始模板，所有操作在输出副本上进行。
- **日期自动**：输出文件名以当天日期为前缀。
- **同天覆盖**：同一天重复运行会覆盖当天输出，必要时换输出目录或确认后覆盖。

## 常见问题

**Q: 前置检查报了"共享模块不完整"？**
A: 检查 `common/` 目录是否完整，PE100/PE150 依赖其中所有模块。XP 不受影响。

**Q: 如何指定自定义输入/输出目录？**
A: 使用环境变量或命令行参数：
```bash
python scripts/run_table.py --type xp --input-dir "D:\本批次数据" --output-dir "D:\输出"
```

**Q: 同事克隆后没有 WorkBuddy 怎么办？**
A: 直接命令行跑 `python run_all.py`，不需要 WorkBuddy。Skill 只是提供"说句话就跑"的便利。
