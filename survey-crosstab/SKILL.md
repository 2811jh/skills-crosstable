---
name: survey-crosstab
description: >
  问卷数据交叉分析工具。自动加载问卷数据（xlsx/xls/csv），识别题目类型，
  执行交叉分析生成频数表和列百分比表，计算满意度/NPS得分，
  并导出带DataBar可视化的专业级Excel报告。支持单选题、多选题、矩阵量表题。
---

# 问卷交叉分析 (Survey Crosstab Analysis)

## 概述

本 Skill 提供问卷交叉分析 CLI 工具链。依赖在首次运行时自动安装（pandas、openpyxl、numpy）。

脚本入口: `~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py`

## 核心工作流（必须严格按此执行）

当用户要求做交叉分析时，按以下 4 步顺序执行。不要跳步。

### Step 1: 加载数据，确认分组列

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py load "<数据文件绝对路径>"
```

阅读返回的列分类：
- `single_choice_questions` — 可作为行/列变量的单选题
- `multi_choice_questions` — 多选题（以 `"Q8."` 根标识引用）
- 找到用户指定的分组维度列名（如"性别"、"年龄"、"职业"等对应的完整列名）

### Step 2: 执行分析（run-all，不传 report-text）

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py run-all "<数据文件绝对路径>" --cols "<分组列名>" --rows all --output-path "<输出文件绝对路径>"
```

**输出文件命名规则**：
- 从源文件所在文件夹名称提取调研主题（如文件夹名"【136-G79】春节版本调研"→ 提取"春节版本调研"）
- 从分组列名提取分组维度关键词（如"Q68.请问您的职业是？"→ 提取"职业"）
- 拼接为: `{调研主题}_{分组维度}交叉分析报告.xlsx`
- 示例: `春节版本调研_职业交叉分析报告.xlsx`
- 输出到源文件同目录下

**cols 参数说明**：直接传列名字符串即可，不需要 JSON 数组格式。

run-all 完成后会在终端输出差异摘要 JSON。仔细阅读它。

### Step 3: 你来撰写分析报告

这是最重要的一步。你必须基于 Step 2 输出的差异摘要 JSON，像专业的数据分析师一样撰写报告。

**撰写规则**：
1. 只分析 `max_min_diff > 0.05`（5%以上）的题目，忽略差异小的
2. 每条 finding 必须引用具体的百分比数字和分组名称
3. finding 的写法参考这个范式：
   - "在「{题目简称}」上，{高分组}选择「{选项}」的比例({高百分比})明显高于{低分组}({低百分比})，差异达{diff}个百分点。"
   - "关于「{题目简称}」，各{分组维度}间存在明显差异：{分组A}偏向「{选项X}」({百分比})，而{分组B}更倾向「{选项Y}」({百分比})。"
4. detail 字段写具体的各分组百分比数据
5. 如果有得分数据（满意度/NPS），也要写入 finding

输出格式为 JSON 数组：

```json
[
  {
    "question": "Q1.就目前的体验而言，您对这款游戏的整体满意度评价如何？",
    "finding": "在整体满意度上，学生群体选择「5分(非常满意)」的比例(42.3%)明显高于自由职业者(28.1%)，差异达14.2个百分点。",
    "detail": "学生:42.3%, 白领:35.7%, 自由职业:28.1%, 其他:31.5%"
  }
]
```

### Step 4: 导出带报告的 Excel

把你撰写的报告 JSON 写入文件，然后导出：

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py export "<数据文件绝对路径>" "<输出文件绝对路径>" --report-text "<你写的JSON报告>"
```

如果报告内容太长不适合命令行传参，可以先把 JSON 写入一个临时 .txt 文件，然后用 Python 读取调用：

```bash
python -c "
import sys; sys.path.insert(0, '<skills脚本目录>')
from crosstab_export import export_crosstab_excel
report = open('<临时报告文件>', 'r', encoding='utf-8').read()
export_crosstab_excel('<数据文件>', '<输出文件>', report)
print('Done')
"
```

注意：export 命令需要在 run-all 同一进程或者之后运行。由于 CLI 跨进程缓存不共享，推荐用上面的 Python 内联方式，或者在 Step 2 的 run-all 命令中直接加 --report-text 参数。

**更推荐的方式**：在 Step 2 时就把报告一起传：

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py run-all "<数据文件>" --cols "<分组列名>" --rows all --output-path "<输出文件>" --report-text '<你写的JSON报告>'
```

但这意味着你需要先运行一次不带 report-text 的 run-all 来获取摘要，读完摘要后写好报告，再运行一次带 report-text 的 run-all。两次运行成本不高，推荐这么做。

## 可用子命令

| 命令 | 说明 |
|---|---|
| `load` | 加载数据文件，返回列分类信息 |
| `preview` | 预览指定列的取值分布 |
| `merge` | 合并/重编码选项（如将满意度1-3归为"不满意"） |
| `crosstab` | 执行交叉分析（需配同进程使用） |
| `score` | 计算满意度加权均值或NPS得分 |
| `summary` | 获取差异摘要 |
| `export` | 导出Excel报告 |
| `run-all` | 全流程一键执行（推荐） |

## 常见请求映射

| 用户说 | 你应该做 |
|--------|----------|
| "帮我分析男女差异" | load 确认性别列名 → run-all(cols=性别列) → 读摘要写报告 → run-all(带report-text) |
| "按年龄段看分布" | load 确认年龄列名 → run-all(cols=年龄列) → 读摘要写报告 → run-all(带report-text) |
| "按职业做交叉分析" | load 确认职业列名 → run-all(cols=职业列) → 读摘要写报告 → run-all(带report-text) |
| "把满意度合并为二分类再分析" | load → merge → run-all(cols=recode列) → 读摘要写报告 → run-all(带report-text) |
| "看看Q8多选题的情况" | load → preview("Q8.") |

## 注意事项

- 文件路径必须用绝对路径
- cols 参数支持直接传字符串，不必包裹 JSON 数组
- rows 参数 `all` 会自动排除分组列和文本题
- 多选题用根标识引用（如 `"Q8."`）
- 不要对小样本（<30人分组）过度解读
- 报告内容必须由你（AI）撰写，不要用 run-all 的自动生成报告（质量不够）