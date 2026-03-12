# skills-crosstable

问卷交叉分析 Agent Skill —— 让 AI 助手自动完成问卷数据的交叉分析全流程。

## 安装

```bash
npx skills add your-github-username/skills-crosstable --skill survey-crosstab
```

安装后 Skill 会被放置到 `~/.agents/skills/survey-crosstab/` 目录。

## 功能

- 📊 自动识别问卷题目类型（单选/多选/矩阵量表/文本/元数据）
- 🔄 一键交叉分析，生成频数表 + 列百分比表
- 📈 自动计算满意度得分（加权均值）和 NPS 得分
- 📋 生成差异摘要，找出分组间差异最大的题目
- 📄 导出带 DataBar 可视化的专业级 Excel 报告
- 🔧 依赖自动安装，零配置即用

## 快速开始

安装 Skill 后，在 Codex 中直接对话：

> "帮我把 C:\xxx\survey.xlsx 按性别做交叉分析"

AI 会自动调用 Skill 脚本完成全部分析流程。

## 仓库结构

```
skills-crosstable/
└── survey-crosstab/           ← Skill 目录（npx skills 安装这个）
    ├── SKILL.md               ← Skill 入口（YAML frontmatter + 指令）
    ├── scripts/
    │   ├── crosstab_cli.py    ← CLI 入口（8 个子命令）
    │   ├── crosstab_engine.py ← 核心数据处理引擎
    │   └── crosstab_export.py ← Excel 导出与格式化
    └── references/
        ├── cli-reference.md   ← 命令手册
        └── workflow-guide.md  ← 工作流指南
```

## 子命令

| 命令        | 说明                          |
| ----------- | ----------------------------- |
| `load`      | 加载问卷数据，自动分类列      |
| `preview`   | 预览指定列的取值分布          |
| `merge`     | 合并/重编码选项               |
| `crosstab`  | 执行交叉分析                  |
| `score`     | 计算满意度/NPS 得分           |
| `summary`   | 获取差异摘要                  |
| `export`    | 导出 Excel 报告               |
| `run-all`   | 全流程一键执行（推荐）        |

## License

MIT