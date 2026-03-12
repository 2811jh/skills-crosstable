---
name: survey-crosstab
description: >
  问卷数据交叉分析工具。自动加载问卷数据（xlsx/xls/csv），识别题目类型，
  执行交叉分析生成频数表和列百分比表，计算满意度/NPS得分，
  并导出带DataBar可视化的专业级Excel报告。支持单选题、多选题、矩阵量表题。
---

# 问卷交叉分析 (Survey Crosstab Analysis)

## 目录

- [概述](#概述)
- [适用场景](#适用场景)
- [快速开始](#快速开始)
- [参考指南](#参考指南)
- [完整工作流](#完整工作流)
- [最佳实践](#最佳实践)

## 概述

本 Skill 提供一套完整的问卷交叉分析 CLI 工具，能够自动识别题目类型、执行交叉分析、计算得分、生成差异摘要，并导出格式化的 Excel 报告。

核心脚本位于 `scripts/` 目录，依赖会在**首次运行时自动安装**（pandas、openpyxl、numpy），无需手动配置。

## 适用场景

- 用户提供问卷数据文件，需要按性别/年龄段等维度做交叉分析
- 需要对多道题目批量生成频数表和百分比表
- 需要计算满意度加权均值或 NPS 得分
- 需要找出不同人群间差异最大的题目，撰写分析报告
- 需要导出带 DataBar 可视化的专业级 Excel 报告

## 快速开始

最简用法——一键全流程：

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py run-all \
  "C:/path/to/survey.xlsx" \
  --cols '["Q17.请问您的性别是？"]' \
  --output-path "report.xlsx"
```

这会自动完成：加载数据 → 交叉分析 → 差异摘要 → 导出 Excel。

## 参考指南

详细说明在 `references/` 目录中：

| 指南 | 内容 |
|------|------|
| [CLI 命令手册](references/cli-reference.md) | 8 个子命令的完整参数说明和示例 |
| [工作流指南](references/workflow-guide.md) | 分步流程、全流程、高级用法详解 |

## 完整工作流

### Step 1: 加载数据

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py load "<数据文件绝对路径>"
```

阅读返回的 JSON 结果，确认：
- `single_choice_questions`: 单选题（可作为行/列变量）
- `multi_choice_questions`: 多选题（以 `"Q8."` 形式的根标识引用）
- `text_questions`: 文本题（自动排除）

### Step 2: 全流程执行

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py run-all "<数据文件>" \
  --cols '["分组列名"]' \
  --rows '["all"]' \
  --score-questions '["满意度/NPS题目"]' \
  --output-path "report.xlsx"
```

推荐使用 `run-all`，它在单进程内完成全部步骤，避免缓存丢失。

### Step 3: 撰写分析报告

根据 `run-all` 输出的差异摘要：
- 对 `max_min_diff > 0.05`（5%）的题目重点分析
- 引用具体百分比数字
- 用 JSON 列表格式输出：

```json
[
  {"question": "Q1.满意度", "finding": "男性满意度(4.2)高于女性(3.8)", "detail": "..."},
  {"question": "Q8.使用功能", "finding": "女性更偏好社交功能(78% vs 52%)", "detail": "..."}
]
```

### Step 4: 导出带报告的 Excel

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py run-all "<数据文件>" \
  --cols '["分组列名"]' \
  --output-path "report.xlsx" \
  --report-text '<JSON报告内容>'
```

## 可用子命令

| 命令 | 说明 | 示例 |
|---|---|---|
| `load` | 加载数据 | `load "data.xlsx"` |
| `preview` | 预览列分布 | `preview "data.xlsx" "Q17.性别"` |
| `merge` | 合并选项 | `merge "data.xlsx" "Q1" '{"不满意":[1,2,3],"满意":[4,5]}'` |
| `crosstab` | 交叉分析 | `crosstab "data.xlsx" --rows '["all"]' --cols '["Q17.性别"]'` |
| `score` | 计算得分 | `score "data.xlsx" '["Q1.满意度"]'` |
| `summary` | 差异摘要 | `summary "data.xlsx"` |
| `export` | 导出报告 | `export "data.xlsx" "report.xlsx"` |
| `run-all` | 全流程一键 | 见上方 Step 2 |

## 常见请求映射

| 用户说 | 执行方式 |
|--------|----------|
| "帮我分析男女差异" | `run-all` + cols=性别列 |
| "按年龄段看分布" | `run-all` + cols=年龄列 |
| "出一份完整报告" | `run-all` + score + 撰写报告 + export |
| "把满意度合并为二分类" | load → merge → run-all(cols=recode列) |
| "看看Q8多选题的情况" | load → preview("Q8.") → crosstab |

## 最佳实践

### ✅ 应该做

- 使用 `run-all` 全流程命令，避免跨进程缓存丢失
- 文件路径使用绝对路径
- 先 `load` 查看列分类，确认分组维度再执行分析
- 对差异 > 5% 的题目重点分析
- 引用具体百分比数字撰写报告
- 多选题用根标识引用（如 `"Q8."`）

### ❌ 不应该做

- 不要分步调用跨多个进程（缓存不共享）
- 不要忽略文本题和元数据列的自动排除
- 不要对小样本（<30人分组）过度解读差异
- 不要遗漏 NPS/满意度得分的计算
- 不要只看频数不看百分比
