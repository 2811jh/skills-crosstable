# 工作流指南

## 场景一：标准交叉分析（最常见）

用户说："帮我把这份问卷数据按性别做交叉分析"

```bash
SKILL_DIR="$HOME/.agents/skills/survey-crosstab/scripts"

# 1. 加载数据，查看有哪些列
python "$SKILL_DIR/crosstab_cli.py" load "/path/to/survey.xlsx"

# 2. 一键全流程
python "$SKILL_DIR/crosstab_cli.py" run-all "/path/to/survey.xlsx" \
  --cols '["Q17.请问您的性别是？"]' \
  --score-questions '["Q1.满意度题目全名"]' \
  --output-path "/path/to/report.xlsx"

# 3. 根据输出的差异摘要，撰写报告，再重新导出
python "$SKILL_DIR/crosstab_cli.py" run-all "/path/to/survey.xlsx" \
  --cols '["Q17.请问您的性别是？"]' \
  --score-questions '["Q1.满意度题目全名"]' \
  --output-path "/path/to/report.xlsx" \
  --report-text '[{"question":"Q1...","finding":"男性满意度更高","detail":"..."}]'
```

## 场景二：先合并再分析

用户说："把满意度1-3归为不满意，4-5归为满意，然后按这个分组来看其他题"

```bash
SKILL_DIR="$HOME/.agents/skills/survey-crosstab/scripts"

python "$SKILL_DIR/crosstab_cli.py" run-all "/path/to/survey.xlsx" \
  --merge-config '[{"column":"Q1.就目前的体验而言，您对山头服的整体满意度评价如何？","rules":{"不满意":[1,2,3],"满意":[4,5]}}]' \
  --cols '["recode_Q1"]' \
  --output-path "report.xlsx"
```

## 场景三：多维度交叉

用户说："按性别和年龄段两个维度来分析"

```bash
SKILL_DIR="$HOME/.agents/skills/survey-crosstab/scripts"

python "$SKILL_DIR/crosstab_cli.py" run-all "/path/to/survey.xlsx" \
  --cols '["Q17.请问您的性别是？", "Q18.请问您的年龄段是？"]' \
  --output-path "report.xlsx"
```

## 场景四：只分析特定题目

用户说："只看满意度相关的3道题"

```bash
SKILL_DIR="$HOME/.agents/skills/survey-crosstab/scripts"

python "$SKILL_DIR/crosstab_cli.py" run-all "/path/to/survey.xlsx" \
  --rows '["Q1.整体满意度", "Q2.产品质量满意度", "Q3.售后服务满意度"]' \
  --cols '["Q17.请问您的性别是？"]' \
  --output-path "report.xlsx"
```

## 场景五：查看多选题

用户说："看看Q8这道多选题各选项被选了多少"

```bash
SKILL_DIR="$HOME/.agents/skills/survey-crosstab/scripts"

# 先预览
python "$SKILL_DIR/crosstab_cli.py" preview "/path/to/survey.xlsx" "Q8."

# 再交叉分析
python "$SKILL_DIR/crosstab_cli.py" run-all "/path/to/survey.xlsx" \
  --rows '["Q8."]' \
  --cols '["Q17.请问您的性别是？"]' \
  --output-path "report.xlsx"
```

## 注意事项

### 跨进程缓存问题

每次执行 CLI 命令都是一个独立的 Python 进程。引擎内部的 `_CACHE` 字典只在单进程生命周期内有效。

**解决方案：** 始终使用 `run-all` 命令，它在单进程内串联执行所有步骤。

### 文件路径

- 始终使用**绝对路径**
- 路径中包含空格或中文时，用引号包裹
- Windows 路径使用 `\` 或 `/` 均可

### 列名匹配

- 列名必须**完全匹配**数据中的实际列名（包括中文标点）
- 先用 `load` 命令查看实际列名，再复制使用
- 多选题用根标识（如 `"Q8."`），引擎会自动展开所有子列
