#!/usr/bin/env python3
"""
问卷交叉分析 CLI 入口

提供 7 个子命令，映射原 MCP Server 的 7 个 Tool：
  load        加载问卷数据，自动分类列
  preview     预览指定列的取值分布
  merge       合并/重编码选项
  crosstab    执行交叉分析
  score       计算满意度/NPS 得分
  summary     获取差异摘要
  export      导出 Excel 报告

典型用法:
  python crosstab_cli.py load data.xlsx
  python crosstab_cli.py preview data.xlsx "Q17.请问您的性别是？"
  python crosstab_cli.py merge data.xlsx "Q1.满意度" '{"不满意":[1,2,3],"满意":[4,5]}'
  python crosstab_cli.py crosstab data.xlsx --rows '["all"]' --cols '["Q17.请问您的性别是？"]'
  python crosstab_cli.py score data.xlsx '["Q1.满意度..."]'
  python crosstab_cli.py summary data.xlsx
  python crosstab_cli.py export data.xlsx report.xlsx --report-text '...'
"""

import sys
import subprocess
import json
import argparse


# ---------------------------------------------------------------------------
#  自动检测并安装缺失依赖
# ---------------------------------------------------------------------------
def _ensure_dependencies():
    """检测 pandas / openpyxl / numpy 是否已安装，缺失则自动 pip install"""
    required = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "numpy": "numpy",
    }
    missing = []
    for import_name, pip_name in required.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print(f"⚙️  检测到缺失依赖: {', '.join(missing)}，正在自动安装...")
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", *missing],
                stdout=sys.stdout,
                stderr=sys.stderr,
            )
            print(f"✅ 依赖安装完成: {', '.join(missing)}")
        except subprocess.CalledProcessError as e:
            print(f"❌ 自动安装失败（exit code {e.returncode}），请手动执行:")
            print(f"   pip install {' '.join(missing)}")
            sys.exit(1)


_ensure_dependencies()

from crosstab_engine import (
    load_data,
    preview_column,
    merge_options,
    run_crosstab,
    calc_scores,
    get_crosstab_summary,
)
from crosstab_export import export_crosstab_excel


def _json_output(data: dict):
    """以格式化 JSON 输出结果到 stdout"""
    print(json.dumps(data, ensure_ascii=False, indent=2, default=str))


# ---------------------------------------------------------------------------
#  子命令处理函数
# ---------------------------------------------------------------------------

def cmd_load(args):
    """加载问卷数据"""
    result = load_data(args.file_path, args.sheet_name)
    _json_output(result)


def cmd_preview(args):
    """预览列分布"""
    result = preview_column(args.file_path, args.column)
    _json_output(result)


def cmd_merge(args):
    """合并选项"""
    try:
        rules = json.loads(args.merge_rules)
    except json.JSONDecodeError as e:
        _json_output({"error": f"merge_rules JSON 解析失败: {str(e)}"})
        sys.exit(1)
    result = merge_options(args.file_path, args.column, rules, args.new_column_name)
    _json_output(result)


def cmd_crosstab(args):
    """执行交叉分析"""
    try:
        rows = json.loads(args.rows)
    except json.JSONDecodeError as e:
        _json_output({"error": f"rows JSON 解析失败: {str(e)}"})
        sys.exit(1)
    try:
        cols = json.loads(args.cols)
    except json.JSONDecodeError as e:
        _json_output({"error": f"cols JSON 解析失败: {str(e)}"})
        sys.exit(1)
    result = run_crosstab(args.file_path, rows, cols)
    _json_output(result)


def cmd_score(args):
    """计算得分"""
    try:
        questions = json.loads(args.score_questions)
    except json.JSONDecodeError as e:
        _json_output({"error": f"score_questions JSON 解析失败: {str(e)}"})
        sys.exit(1)
    result = calc_scores(args.file_path, questions)
    _json_output(result)


def cmd_summary(args):
    """获取差异摘要"""
    result = get_crosstab_summary(args.file_path)
    _json_output(result)


def cmd_export(args):
    """导出 Excel 报告"""
    try:
        output = export_crosstab_excel(
            args.file_path,
            args.output_path,
            args.report_text or "",
        )
        _json_output({
            "status": "success",
            "output_path": output,
            "message": f"报告已成功导出至: {output}",
        })
    except Exception as e:
        _json_output({"status": "error", "error": str(e)})
        sys.exit(1)


# ---------------------------------------------------------------------------
#  全流程快捷命令
# ---------------------------------------------------------------------------

def cmd_run_all(args):
    """全流程一键执行：加载 → 交叉分析 → 得分(可选) → 摘要 → 导出"""
    print("=" * 60)
    print("📊 问卷交叉分析全流程")
    print("=" * 60)

    # 1. 加载数据
    print("\n[1/5] 加载数据...")
    load_result = load_data(args.file_path, args.sheet_name)
    print(f"  ✅ 加载成功: {load_result['total_rows']} 行, {load_result['total_columns']} 列")
    print(f"  单选题: {len(load_result['single_choice_questions'])} 个")
    print(f"  多选题: {len(load_result['multi_choice_questions'])} 个")

    # 2. 合并选项（如有）
    if args.merge_config:
        print("\n[2/5] 合并选项...")
        try:
            merge_configs = json.loads(args.merge_config)
            for mc in merge_configs:
                col = mc["column"]
                rules = mc["rules"]
                new_name = mc.get("new_column_name")
                result = merge_options(args.file_path, col, rules, new_name)
                print(f"  ✅ 合并 '{col}' → '{result['new_column']}'")
        except (json.JSONDecodeError, KeyError) as e:
            print(f"  ⚠️ 合并配置解析失败: {e}")
    else:
        print("\n[2/5] 跳过合并选项（未提供配置）")

    # 3. 交叉分析
    print("\n[3/5] 执行交叉分析...")
    try:
        rows = json.loads(args.rows)
    except json.JSONDecodeError:
        rows = ["all"]
    try:
        cols = json.loads(args.cols)
    except json.JSONDecodeError:
        print(f"  ❌ cols JSON 解析失败")
        sys.exit(1)
    ct_result = run_crosstab(args.file_path, rows, cols)
    print(f"  ✅ 交叉分析完成: {ct_result.get('row_questions_count', '?')} 个行变量, "
          f"{ct_result.get('col_conditions_count', '?')} 个列条件")

    # 4. 得分计算（如有）
    if args.score_questions:
        print("\n[4/5] 计算得分...")
        try:
            sq = json.loads(args.score_questions)
            score_result = calc_scores(args.file_path, sq)
            if score_result.get("status") == "success":
                print(f"  ✅ 得分计算完成: {list(score_result.get('score_types', {}).values())}")
            else:
                print(f"  ⚠️ {score_result.get('message', '无有效题目')}")
        except json.JSONDecodeError as e:
            print(f"  ⚠️ score_questions 解析失败: {e}")
    else:
        print("\n[4/5] 跳过得分计算")

    # 5. 摘要 + 导出
    print("\n[5/5] 生成摘要并导出...")
    summary_result = get_crosstab_summary(args.file_path)
    print(f"  ✅ 差异摘要: {summary_result.get('questions_analyzed', '?')} 道题目")

    if args.output_path:
        # 如果用户没有提供 report_text，则基于 summary 自动生成报告 JSON
        report_text = args.report_text or ""
        if not report_text and summary_result.get("status") == "success":
            report_entries = []
            q_summaries = summary_result.get("question_summaries", {})
            for q_name, q_data in q_summaries.items():
                max_opt = q_data.get("max_diff_option", "")
                max_diff = q_data.get("max_diff_value", 0)
                if max_diff < 0.05:
                    continue
                # 构造 finding 描述
                opt_data = q_data.get("options", {}).get(max_opt, {})
                pcts = opt_data.get("percentages", {})
                if pcts:
                    sorted_pcts = sorted(pcts.items(), key=lambda x: x[1], reverse=True)
                    highest = sorted_pcts[0]
                    lowest = sorted_pcts[-1]
                    h_group = highest[0].split("\n")[-1] if "\n" in highest[0] else highest[0]
                    l_group = lowest[0].split("\n")[-1] if "\n" in lowest[0] else lowest[0]
                    finding = (
                        f"选项「{max_opt}」在不同分组间差异最大（{max_diff:.1%}）。"
                        f"「{h_group}」最高（{highest[1]:.1%}），「{l_group}」最低（{lowest[1]:.1%}）。"
                    )
                else:
                    finding = f"选项「{max_opt}」在不同分组间差异达 {max_diff:.1%}。"

                report_entries.append({
                    "question": q_name,
                    "finding": finding,
                    "detail": f"差异最大选项: {max_opt}, 最大差异值: {max_diff:.4f}",
                })

            # 追加得分摘要
            score_summary = summary_result.get("score_summary", {})
            for score_name, scores in score_summary.items():
                if scores:
                    sorted_scores = sorted(scores.items(), key=lambda x: x[1], reverse=True)
                    highest = sorted_scores[0]
                    lowest = sorted_scores[-1]
                    h_group = highest[0].split("\n")[-1] if "\n" in highest[0] else highest[0]
                    l_group = lowest[0].split("\n")[-1] if "\n" in lowest[0] else lowest[0]
                    report_entries.append({
                        "question": score_name,
                        "finding": f"最高: {h_group}（{highest[1]:.2f}），最低: {l_group}（{lowest[1]:.2f}）",
                        "detail": f"各分组得分: {json.dumps({k.split(chr(10))[-1] if chr(10) in k else k: round(v, 2) for k, v in scores.items()}, ensure_ascii=False)}",
                    })

            if report_entries:
                report_text = json.dumps(report_entries, ensure_ascii=False)
                print(f"  📝 自动生成分析报告: {len(report_entries)} 条发现")

        try:
            out = export_crosstab_excel(
                args.file_path,
                args.output_path,
                report_text,
            )
            print(f"  ✅ 报告已导出: {out}")
        except Exception as e:
            print(f"  ❌ 导出失败: {e}")
    else:
        print("  ⚠️ 未指定输出路径，跳过导出")

    # 输出摘要 JSON（便于 AI 读取撰写报告）
    print("\n" + "=" * 60)
    print("📝 差异摘要（供撰写报告使用）:")
    print("=" * 60)
    _json_output(summary_result)


# ---------------------------------------------------------------------------
#  主入口
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="问卷交叉分析 CLI 工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    subparsers = parser.add_subparsers(dest="command", help="可用子命令")

    # ---- load ----
    p_load = subparsers.add_parser("load", help="加载问卷数据文件")
    p_load.add_argument("file_path", help="数据文件路径 (.xlsx/.xls/.csv)")
    p_load.add_argument("--sheet", dest="sheet_name", default=0,
                        help="工作表名称或编号 (默认 0)")
    p_load.set_defaults(func=cmd_load)

    # ---- preview ----
    p_preview = subparsers.add_parser("preview", help="预览指定列的取值分布")
    p_preview.add_argument("file_path", help="数据文件路径")
    p_preview.add_argument("column", help="列名或多选题根 (如 Q8.)")
    p_preview.set_defaults(func=cmd_preview)

    # ---- merge ----
    p_merge = subparsers.add_parser("merge", help="合并/重编码选项")
    p_merge.add_argument("file_path", help="数据文件路径")
    p_merge.add_argument("column", help="要合并的原始列名")
    p_merge.add_argument("merge_rules", help='合并规则 JSON, 如 \'{"不满意":[1,2,3],"满意":[4,5]}\'')
    p_merge.add_argument("--name", dest="new_column_name", default=None,
                         help="新列名 (可选)")
    p_merge.set_defaults(func=cmd_merge)

    # ---- crosstab ----
    p_ct = subparsers.add_parser("crosstab", help="执行交叉分析")
    p_ct.add_argument("file_path", help="数据文件路径")
    p_ct.add_argument("--rows", required=True,
                      help='行变量 JSON, 如 \'["all"]\' 或 \'["Q1.xxx","Q2.xxx"]\'')
    p_ct.add_argument("--cols", required=True,
                      help='列变量(分组) JSON, 如 \'["Q17.性别"]\'')
    p_ct.set_defaults(func=cmd_crosstab)

    # ---- score ----
    p_score = subparsers.add_parser("score", help="计算满意度/NPS 得分")
    p_score.add_argument("file_path", help="数据文件路径")
    p_score.add_argument("score_questions",
                         help='题目列表 JSON, 如 \'["Q1.满意度..."]\'')
    p_score.set_defaults(func=cmd_score)

    # ---- summary ----
    p_summary = subparsers.add_parser("summary", help="获取差异摘要")
    p_summary.add_argument("file_path", help="数据文件路径")
    p_summary.set_defaults(func=cmd_summary)

    # ---- export ----
    p_export = subparsers.add_parser("export", help="导出 Excel 报告")
    p_export.add_argument("file_path", help="源数据文件路径")
    p_export.add_argument("output_path", help="输出 Excel 文件路径")
    p_export.add_argument("--report-text", default="",
                          help="AI 分析报告内容 (JSON 列表或纯文本)")
    p_export.set_defaults(func=cmd_export)

    # ---- run-all (全流程) ----
    p_all = subparsers.add_parser("run-all", help="全流程一键执行")
    p_all.add_argument("file_path", help="数据文件路径")
    p_all.add_argument("--sheet", dest="sheet_name", default=0,
                       help="工作表名称或编号")
    p_all.add_argument("--rows", default='["all"]',
                       help='行变量 JSON (默认 "all")')
    p_all.add_argument("--cols", required=True,
                       help='列变量(分组) JSON')
    p_all.add_argument("--merge-config", default=None,
                       help='合并配置 JSON 列表, 如 \'[{"column":"Q1","rules":{"不满意":[1,2,3],"满意":[4,5]}}]\'')
    p_all.add_argument("--score-questions", default=None,
                       help='得分题目 JSON 列表')
    p_all.add_argument("--output-path", default=None,
                       help="输出 Excel 路径")
    p_all.add_argument("--report-text", default="",
                       help="AI 分析报告内容")
    p_all.set_defaults(func=cmd_run_all)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    args.func(args)


if __name__ == "__main__":
    main()
