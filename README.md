# HP Table Process

一键生成 PE100、PE150、XP、真迈四种测序文库 Pooling 表。支持 WorkBuddy Skill 语音调度，也支持命令行直接运行。

## 环境要求

- **WorkBuddy**（Skill 运行环境，自带 Python，无需单独安装）
- 对应型号的本批次源表

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/www0922/hpls-table-maker.git
cd hpls-table-maker
```

### 2. 放入本批次源表

将源表放入对应型号的 `input_data/` 目录，每个关键词只保留一个文件：

| 型号 | 目录 | 需要文件（文件名包含关键词） |
|------|------|---------------------------|
| PE100 | `PE_100/input_data/` | 上机、质检、自建库 |
| PE150 | `PE_150/input_data/` | 上机、质检、自建库 |
| XP | `XP/input_data/` | Xplus上机、质检、自建库、纯化、qPCR |
| 真迈 | `ZM/input_data/` | 江西上机、质检、自建库、纯化、qPCR |

### 3. 使用 WorkBuddy

打开 WorkBuddy，在项目目录下说：

```
帮我做 PE100 表
帮我做 PE150 表
帮我做 XP 表
帮我做真迈表
```

Skill 会自动校验源文件、运行完整流水线，并在 `output_result/` 生成最终 Excel。

### 4. 命令行运行（无需 WorkBuddy）

```bash
python PE_100/code/run_all.py    # PE100
python PE_150/code/run_all.py    # PE150
python XP/code/run_all.py        # XP
python ZM/code/run_all.py        # 真迈
```

首次运行如需安装 `openpyxl`，Skill 会自动处理；命令行请手动执行 `pip install openpyxl`。

## 项目结构

```
hpls-table-maker/
├── PE_100/                     # PE100 型
│   ├── code/run_all.py         # 一键全流程
│   ├── input_data/             # 源数据（每批次替换）
│   └── output_result/          # 输出（不提交 Git）
├── PE_150/                     # PE150 型（结构同上）
├── XP/                         # XP 型（输出含 PCR 定量表）
├── ZM/                         # 真迈型（输出含 PCR 定量表）
├── common/                     # PE100/PE150 共享模块
├── .workbuddy/skills/
│   └── hpls-table-maker/       # WorkBuddy 统一 Skill
└── README.md
```

## 输出文件

| 型号 | 输出 |
|------|------|
| PE100 | `PE_100/output_result/YYYYMMDD文库pooling表T7+PE100.xlsx` |
| PE150 | `PE_150/output_result/YYYYMMDD文库pooling表T7+PE150.xlsx` |
| XP | `XP/output_result/YYYYMMDD文库pooling表.xlsx` + `YYYYMMDDPCR定量表.xlsx` |
| 真迈 | `ZM/output_result/YYYYMMDD文库pooling表AE0.xlsx` + `YYYYMMDDPCR定量表真迈.xlsx` |

> 文件名中的 `YYYYMMDD` 为当天日期，如 `20260807`。

## 注意事项

- 每个关键词必须唯一匹配一个文件，多余文件会报错。同一目录只保留当前批次。
- 模板文件为只读，所有操作在输出副本上进行，不修改原始模板。
- 同一天重复运行会覆盖当天输出。

## 常见问题

**Q: 前置检查报"共享模块不完整"？**
A: 检查 `common/` 目录是否完整，PE100/PE150 依赖该目录。XP 和真迈不受影响。

**Q: 如何自定义输入/输出目录？**
A: 使用命令行参数（Skill 内部会传参，无需手动设置）：
```bash
python .workbuddy/skills/hpls-table-maker/scripts/run_table.py --type xp --input-dir "D:\本批次" --output-dir "D:\输出"
```
