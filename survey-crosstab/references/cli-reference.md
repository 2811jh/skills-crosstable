# CLI 命令手册

## 命令总览

所有命令的调用格式：

```bash
python ~/.agents/skills/survey-crosstab/scripts/crosstab_cli.py <command> [args...]
```

---

## load — 加载问卷数据

```bash
python crosstab_cli.py load "<file_path>" [--sheet <name_or_index>]
```

| 参数 | 必需 | 说明 |
|------|------|------|
| `file_path` | ✅ | 数据文件的绝对路径（.xlsx / .xls / .csv） |
| `--sheet` | ❌ | 工作表名称或编号，默认 `0`（第一个 sheet） |

**返回字段：**
- `total_rows` / `total_columns` — 数据维度
- `single_choice_questions` — 单选题列表
- `multi_choice_questions` — 多选题字典（根标识 → 子列列表）
- `matrix_scale_questions` — 矩阵量表题
- `text_questions` — 文本题（自动排除）
- `meta_columns` — 元数据列（序号、时间等，自动排除）
- `excluded_columns` — 排除列（多选题的"其他"输入框等）
- `valid_for_crosstab` — 所有可用于交叉分析的列

---

## preview — 预览列分布

```bash
python crosstab_cli.py preview "<file_path>" "<column>"
```

| 参数 | 必需 | 说明 |
|------|------|------|
| `file_path` | ✅ | 数据文件路径 |
| `column` | ✅ | 列名或多选题根标识（如 `"Q8."`） |

**返回字段：** `type`、`total`、`non_null`、`null`、`unique_values`、`distribution`

---

## merge — 合并/重编码选项

```bash
python crosstab_cli.py merge "<file_path>" "<column>" '<merge_rules_json>' [--name <new_name>]
```

| 参数 | 必需 | 说明 |
|------|------|------|
| `file_path` | ✅ | 数据文件路径 |
| `column` | ✅ | 原始列名 |
| `merge_rules` | ✅ | JSON 格式合并规则 |
| `--name` | ❌ | 新列名（默认自动生成 `recode_xxx`） |

**merge_rules 示例：**
```json
{"不满意": [1, 2, 3], "满意": [4, 5]}
{"年轻人": ["18-25", "26-30"], "中年人": ["31-40", "41-50"]}
```

---

## crosstab — 执行交叉分析

```bash
python crosstab_cli.py crosstab "<file_path>" --rows '<json>' --cols '<json>'
```

| 参数 | 必需 | 说明 |
|------|------|------|
| `file_path` | ✅ | 数据文件路径 |
| `--rows` | ✅ | 行变量 JSON 列表 |
| `--cols` | ✅ | 列变量（分组维度）JSON 列表 |

**rows 取值：**
- `'["all"]'` — 所有可分析的题目（最常用）
- `'["Q1.xxx", "Q2.xxx"]'` — 指定题目
- `'["Q8."]'` — 多选题用根标识

**cols 取值：**
- `'["Q17.请问您的性别是？"]'` — 单个分组
- `'["Q17.性别", "Q18.年龄段"]'` — 多个分组
- `'["recode_满意度"]'` — 合并后的列

---

## score — 计算满意度/NPS 得分

```bash
python crosstab_cli.py score "<file_path>" '<questions_json>'
```

| 参数 | 必需 | 说明 |
|------|------|------|
| `file_path` | ✅ | 数据文件路径 |
| `score_questions` | ✅ | 需计算得分的题目 JSON 列表 |

**自动识别规则：**
- 含"满意度"关键词 → 加权均值（选项值 × 百分比）
- 含"推荐"/"NPS"或 0-10 分制 → NPS = 推荐者%(9-10) - 贬损者%(0-6)

⚠️ 必须在 `crosstab` 之后执行，且题目必须在行变量中出现过。

---

## summary — 获取差异摘要

```bash
python crosstab_cli.py summary "<file_path>"
```

返回每道题目各选项在不同分组间的百分比及最大差异值（`max_min_diff`）。

---

## export — 导出 Excel 报告

```bash
python crosstab_cli.py export "<file_path>" "<output_path>" [--report-text '<text>']
```

| 参数 | 必需 | 说明 |
|------|------|------|
| `file_path` | ✅ | 源数据文件路径 |
| `output_path` | ✅ | 输出 Excel 路径 |
| `--report-text` | ❌ | AI 分析报告内容（JSON 列表或纯文本） |

**生成的 Excel 包含：**
- 【交叉分析】频数表 + DataBar
- 【列百分比】百分比表 + DataBar
- 【得分分析】满意度/NPS 得分（如有）
- 【分析报告】AI 撰写的报告（如提供 report_text）

---

## run-all — 全流程一键执行

```bash
python crosstab_cli.py run-all "<file_path>" --cols '<json>' \
  [--rows '<json>'] \
  [--sheet <name>] \
  [--merge-config '<json>'] \
  [--score-questions '<json>'] \
  [--output-path "<path>"] \
  [--report-text '<text>']
```

| 参数 | 必需 | 默认值 | 说明 |
|------|------|--------|------|
| `file_path` | ✅ | — | 数据文件路径 |
| `--cols` | ✅ | — | 列变量（分组维度） |
| `--rows` | ❌ | `'["all"]'` | 行变量 |
| `--sheet` | ❌ | `0` | 工作表 |
| `--merge-config` | ❌ | — | 合并配置 JSON 列表 |
| `--score-questions` | ❌ | — | 得分题目 JSON 列表 |
| `--output-path` | ❌ | — | 输出 Excel 路径 |
| `--report-text` | ❌ | — | AI 分析报告 |

**merge-config 格式：**
```json
[{"column": "Q1.xxx", "rules": {"不满意": [1,2,3], "满意": [4,5]}, "new_column_name": "recode_满意度"}]
```
