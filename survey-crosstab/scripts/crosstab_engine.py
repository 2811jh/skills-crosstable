"""
问卷交叉分析 MCP Server - 核心数据处理引擎

包含：数据加载、列分类、变量重编码/合并、交叉分析计算、得分计算
"""

import pandas as pd
import numpy as np
import re
import warnings
from collections import defaultdict
from typing import Optional


# ---------------------------------------------------------------------------
#  全局缓存：file_path → {"df": DataFrame, "multi_dict": {...}, ...}
# ---------------------------------------------------------------------------
_CACHE: dict[str, dict] = {}


# ========================================================================= #
#                           辅助函数
# ========================================================================= #

def _extract_subcol_number(subcol: str, prefix: str) -> int:
    """从多选子列名中提取选项序号"""
    suffix = subcol.split(prefix)[1].strip()
    match = re.search(r'^(\d+)', suffix)
    return int(match.group(1)) if match else 0


def _extract_score_from_option(option) -> Optional[float]:
    """从选项文本中提取数值分数"""
    if option is None:
        return None
    s = str(option).strip()
    if not s:
        return None
    match = re.match(r'^-?\d+(?:\.\d+)?', s)
    if not match:
        match = re.search(r'-?\d+(?:\.\d+)?', s)
    return float(match.group(0)) if match else None


def _is_text_column(series: pd.Series, col_name: str) -> bool:
    """判断某列是否为文本题（非结构化答案）"""
    # 关键词精确匹配
    if "输入文本" in col_name:
        return True
    if "非必填" in col_name:
        return True

    non_null = series.dropna()
    if len(non_null) == 0:
        return True

    # 尝试判断是否为纯数值列（量表题/单选编码） → 非文本
    try:
        pd.to_numeric(non_null)
        return False  # 全是数字，不是文本题
    except (ValueError, TypeError):
        pass

    # 字符串列进一步判断
    str_vals = non_null.astype(str)
    avg_len = str_vals.str.len().mean()
    unique_rate = non_null.nunique() / len(non_null) if len(non_null) > 0 else 0

    # 高唯一率 + 平均长度较长 → 文本题
    if unique_rate > 0.6 and avg_len > 8:
        return True
    # 超长文本
    if avg_len > 20:
        return True

    return False


def _is_meta_column(col_name: str) -> bool:
    """
    判断是否为元数据列（非题目列）。
    核心原则：以 Q+数字 开头的列优先视为题目，除非是附属文本列。
    """
    col_clean = col_name.strip()

    # 1. 以 Q+数字 开头 → 大概率是题目，但先检查个人信息类
    if re.match(r'^Q\d+[\.\s]', col_clean):
        # 排除个人信息收集题
        personal_keywords = ["姓名", "手机", "电话", "微信", "邮箱", "称呼",
                             "个人信息", "联系方式", "联系到您"]
        for pk in personal_keywords:
            if pk in col_clean:
                return True
        return False

    # 2. recode_ 开头 → 合并后的列
    if col_clean.startswith('recode_'):
        return False

    # 3. 明确的元数据关键词
    meta_exact_prefixes = ["序号", "开始答题时间", "结束答题时间", "答题时长"]
    for prefix in meta_exact_prefixes:
        if col_clean.startswith(prefix):
            return True

    # 4. 包含元数据特征关键词
    meta_keywords = [
        "uid", "UID", "IP地址", "来源渠道",
        "怎么称呼", "姓名", "手机号", "电话", "微信号", "邮箱",
    ]
    for kw in meta_keywords:
        if kw in col_clean:
            return True

    # 5. 附属文本列（如 "其他:输入文本"、"想玩的模组:输入文本"）
    if "输入文本" in col_clean:
        return True

    # 6. 非 Q 开头的其余未知列 → 按元数据处理
    if not re.match(r'^Q\d+', col_clean):
        return True

    return False


def _classify_columns(df: pd.DataFrame) -> dict:
    """
    自动分类所有列：
    返回 {
        "single_choice": [...],     # 单选题（含矩阵量表题的各子列）
        "multi_choice": {root: [subcols]},  # 多选题（0/1 编码）
        "matrix_scale": {root: [subcols]},  # 矩阵量表题（1-5分等）
        "text": [...],              # 文本题
        "meta": [...],              # 元数据列
        "excluded": [...],          # 被排除的列
        "valid_for_crosstab": [...] # 可用于交叉分析的所有列（单选 + 多选根）
    }
    """
    # 1. 先识别带冒号的子列根
    multi_roots = defaultdict(list)
    for col in df.columns:
        col_clean = str(col).strip()
        match = re.match(r'^(Q\d+\.)', col_clean)
        if match:
            root = match.group(1)
            # 检查是否有冒号分隔（子列格式: "Q8.题目:选项"）
            rest = col_clean[len(root):]
            if ':' in rest or '：' in rest:
                multi_roots[root].append(col)

    # 2. 区分多选题 vs 矩阵量表题
    #    多选题: 值只有 0/1 (或 NaN)
    #    矩阵量表题: 值有 >1 的数字 (如 1-5 分)
    multi_choice_dict = {}
    matrix_scale_dict = {}
    multi_subcols = set()
    matrix_subcols = set()

    for root, subcols in multi_roots.items():
        if len(subcols) <= 1:
            continue  # 单个子列不处理
        sorted_cols = sorted(subcols, key=lambda x: _extract_subcol_number(x, root))

        # 采样前 1000 行判断取值类型，避免大文件卡顿
        is_matrix = False
        sample = df[sorted_cols].head(1000)
        for sc in sorted_cols:
            if "输入文本" in str(sc):
                continue
            try:
                unique_vals = sample[sc].dropna().unique()
                numeric_vals = set()
                for v in unique_vals:
                    try:
                        numeric_vals.add(float(v))
                    except (ValueError, TypeError):
                        pass
                # 如果存在大于 1 的值 → 量表题（不是 0/1 编码的多选）
                if numeric_vals and max(numeric_vals) > 1:
                    is_matrix = True
                    break
            except Exception:
                pass

        if is_matrix:
            matrix_scale_dict[root] = sorted_cols
            matrix_subcols.update(sorted_cols)
        else:
            multi_choice_dict[root] = sorted_cols
            multi_subcols.update(sorted_cols)

    # 3. 分类其余列
    single_choice = []
    text_cols = []
    meta_cols = []
    excluded_cols = []

    for col in df.columns:
        col_str = str(col).strip()

        # 已归入多选题子列
        if col in multi_subcols:
            if "输入文本" in col_str:
                excluded_cols.append(col)
            continue

        # 矩阵量表题子列 → 当作独立的单选/量表题
        if col in matrix_subcols:
            if "输入文本" in col_str:
                excluded_cols.append(col)
            else:
                single_choice.append(col)
            continue

        # 元数据列
        if _is_meta_column(col_str):
            meta_cols.append(col)
            continue

        # 文本题
        if _is_text_column(df[col], col_str):
            text_cols.append(col)
            continue

        # 剩余为单选题
        single_choice.append(col)

    # 4. 可用于交叉分析的列
    valid_for_crosstab = list(single_choice)
    for root in multi_choice_dict:
        valid_for_crosstab.append(root)

    return {
        "single_choice": single_choice,
        "multi_choice": multi_choice_dict,
        "matrix_scale": matrix_scale_dict,
        "text": text_cols,
        "meta": meta_cols,
        "excluded": excluded_cols,
        "valid_for_crosstab": valid_for_crosstab,
    }


# ========================================================================= #
#                        数据加载
# ========================================================================= #

def load_data(file_path: str, sheet_name=0) -> dict:
    """
    加载问卷数据并自动分类列。
    支持 .xlsx / .xls / .csv
    """
    ext = file_path.rsplit('.', 1)[-1].lower()
    if ext == 'csv':
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

    df.columns = [str(c).strip() for c in df.columns]

    classification = _classify_columns(df)

    # 缓存
    _CACHE[file_path] = {
        "df": df,
        "classification": classification,
        "merged_cols": {},  # 合并后的列信息
    }

    return {
        "file_path": file_path,
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "single_choice_questions": classification["single_choice"],
        "multi_choice_questions": {
            root: cols for root, cols in classification["multi_choice"].items()
        },
        "matrix_scale_questions": {
            root: cols for root, cols in classification.get("matrix_scale", {}).items()
        },
        "text_questions": classification["text"],
        "meta_columns": classification["meta"],
        "excluded_columns": classification["excluded"],
        "valid_for_crosstab": classification["valid_for_crosstab"],
    }


def get_cached_df(file_path: str) -> pd.DataFrame:
    """获取缓存的 DataFrame"""
    if file_path not in _CACHE:
        load_data(file_path)
    return _CACHE[file_path]["df"]


def get_cached_classification(file_path: str) -> dict:
    """获取缓存的列分类"""
    if file_path not in _CACHE:
        load_data(file_path)
    return _CACHE[file_path]["classification"]


# ========================================================================= #
#                        列预览
# ========================================================================= #

def preview_column(file_path: str, column: str) -> dict:
    """预览某一列的取值分布"""
    df = get_cached_df(file_path)
    classification = get_cached_classification(file_path)
    multi_dict = classification["multi_choice"]

    # 检查是否为多选题根
    if column in multi_dict:
        subcols = multi_dict[column]
        result = {}
        for sc in subcols:
            counts = df[sc].value_counts().to_dict()
            result[sc] = counts
        return {
            "column": column,
            "type": "multi_choice",
            "sub_columns": list(subcols),
            "distributions": result,
        }

    if column not in df.columns:
        # 检查是否为合并后的列
        if file_path in _CACHE and column in _CACHE[file_path].get("merged_cols", {}):
            series = df[column]
        else:
            return {"error": f"列 '{column}' 不存在"}
    else:
        series = df[column]

    value_counts = series.value_counts(dropna=False).to_dict()
    # 将 NaN key 转为字符串
    clean_counts = {}
    for k, v in value_counts.items():
        key = "缺失(NaN)" if pd.isna(k) else str(k)
        clean_counts[key] = int(v)

    return {
        "column": column,
        "type": "single_choice",
        "total": len(series),
        "non_null": int(series.notna().sum()),
        "null": int(series.isna().sum()),
        "unique_values": int(series.nunique()),
        "distribution": clean_counts,
    }


# ========================================================================= #
#                        合并选项 (Merge / Recode)
# ========================================================================= #

def merge_options(
    file_path: str,
    column: str,
    merge_rules: dict[str, list],
    new_column_name: Optional[str] = None,
) -> dict:
    """
    合并指定列的选项值。

    Args:
        file_path: 文件路径
        column: 原始列名
        merge_rules: 合并规则, 例如 {"不满意": [1,2,3], "满意": [4,5]}
        new_column_name: 合并后新列名，默认为 "recode_{column简称}"

    Returns:
        合并后的列信息和分布
    """
    df = get_cached_df(file_path)

    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在"}

    # 构建映射
    mapping = {}
    for label, values in merge_rules.items():
        for v in values:
            mapping[v] = label

    if new_column_name is None:
        # 简化列名
        short_name = re.sub(r'^Q\d+\.', '', column).strip()
        if len(short_name) > 20:
            short_name = short_name[:20]
        new_column_name = f"recode_{short_name}"

    df[new_column_name] = df[column].map(mapping)

    # 记录合并信息
    _CACHE[file_path]["merged_cols"][new_column_name] = {
        "source": column,
        "rules": merge_rules,
    }

    # 返回分布
    dist = df[new_column_name].value_counts(dropna=False)
    clean_dist = {}
    for k, v in dist.items():
        key = "缺失(NaN)" if pd.isna(k) else str(k)
        clean_dist[key] = int(v)

    return {
        "new_column": new_column_name,
        "source_column": column,
        "merge_rules": {k: [str(x) for x in v] for k, v in merge_rules.items()},
        "distribution": clean_dist,
        "total_mapped": int(df[new_column_name].notna().sum()),
        "total_unmapped": int(df[new_column_name].isna().sum()),
    }


# ========================================================================= #
#                        交叉分析核心
# ========================================================================= #

def run_crosstab(
    file_path: str,
    row_questions: list[str],
    col_questions: list[str],
) -> dict:
    """
    执行交叉分析。

    Args:
        file_path: 数据文件路径
        row_questions: 行变量列表。支持:
            - 具体列名 (如 "Q1.满意度...")
            - 多选题根 (如 "Q8.")
            - "all" 表示所有可交叉分析的题目
        col_questions: 列变量列表（分组维度）

    Returns:
        包含频数表和百分比表的完整结果字典
    """
    df = get_cached_df(file_path)
    classification = get_cached_classification(file_path)
    multi_dict = classification["multi_choice"]

    # --- 处理 "all" ---
    if row_questions == ["all"] or row_questions == "all":
        row_questions = list(classification["valid_for_crosstab"])
        # 排除列变量中已出现的列（含合并后的源列）
        col_sources = set()
        for cq in col_questions:
            col_sources.add(cq)
            if file_path in _CACHE:
                merged_info = _CACHE[file_path].get("merged_cols", {})
                if cq in merged_info:
                    col_sources.add(merged_info[cq]["source"])
        row_questions = [q for q in row_questions if q not in col_sources]

    # --- 识别多选题 ---
    user_multi_roots = set()
    for q in row_questions + col_questions:
        q_clean = str(q).strip()
        if re.fullmatch(r'^Q\d+\.$', q_clean):
            user_multi_roots.add(q_clean)

    # 合并自动检测的多选题
    for root in multi_dict:
        if root in user_multi_roots or any(
            str(q).strip().startswith(root) for q in row_questions + col_questions
        ):
            user_multi_roots.add(root)

    # --- 验证问题 ---
    def validate_and_classify(questions):
        valid = []
        invalid = []
        for q in questions:
            q_clean = str(q).strip()
            # 多选题根
            if q_clean in multi_dict:
                valid.append(("multi", q_clean))
            # 尝试匹配多选题根
            elif re.match(r'^Q\d+\.$', q_clean) and q_clean in multi_dict:
                valid.append(("multi", q_clean))
            # 单选题 / 合并后的列
            elif q_clean in df.columns:
                valid.append(("single", q_clean))
            else:
                invalid.append(q_clean)
        return valid, invalid

    valid_rows, invalid_rows = validate_and_classify(row_questions)
    valid_cols, invalid_cols = validate_and_classify(col_questions)

    if invalid_rows:
        warnings.warn(f"无效行问题将被跳过：{invalid_rows}")
    if invalid_cols:
        warnings.warn(f"无效列问题将被跳过：{invalid_cols}")

    # --- 列条件生成 ---
    col_conditions = []
    col_totals = {}
    seen_cols = defaultdict(int)

    for q_type, q in valid_cols:
        q_clean = str(q).strip()
        seen_cols[q_clean] += 1
        instance_id = seen_cols[q_clean]

        if q_type == "multi":
            root = q_clean
            subcols = multi_dict[root]
            # 提取题目文本
            example_subcol = subcols[0]
            rest_part = example_subcol.split(root)[1].strip()
            if ':' in rest_part:
                question_text = rest_part.split(':', 1)[0].strip()
            elif '：' in rest_part:
                question_text = rest_part.split('：', 1)[0].strip()
            else:
                question_text = rest_part
            full_question = f"{root}{question_text}"
            if instance_id > 1:
                full_question += f" #{instance_id}"

            for subcol in subcols:
                rest_subcol = subcol.split(root)[1].strip()
                if ':' in rest_subcol:
                    option_text = rest_subcol.split(':', 1)[1].strip()
                elif '：' in rest_subcol:
                    option_text = rest_subcol.split('：', 1)[1].strip()
                else:
                    option_text = rest_subcol
                label = f"{full_question}\n{option_text}"
                cond = df[subcol] == 1
                col_conditions.append((label, cond))
                col_totals[label] = int(cond.sum())

            total_label = f"{full_question}\n总计"
            total_cond = (df[subcols] == 1).any(axis=1)
            col_conditions.append((total_label, total_cond))
            col_totals[total_label] = int(total_cond.sum())

        else:  # single
            values = df[q_clean].dropna().unique()
            try:
                sorted_values = sorted(
                    values,
                    key=lambda x: int(re.match(r'^(\d+)', str(x)).group(1))
                )
            except Exception:
                sorted_values = sorted(values, key=str)

            unique_question = q_clean
            if instance_id > 1:
                unique_question += f" #{instance_id}"

            for value in sorted_values:
                label = f"{unique_question}\n{value}"
                cond = df[q_clean] == value
                col_conditions.append((label, cond))
                col_totals[label] = int(cond.sum())

            total_label = f"{unique_question}\n总计"
            total_cond = df[q_clean].notna()
            col_conditions.append((total_label, total_cond))
            col_totals[total_label] = int(total_cond.sum())

    # --- 行条件生成 ---
    row_conditions = []
    for q_type, q in valid_rows:
        if q_type == "multi":
            root = q
            subcols = multi_dict[root]
            # 从第一个子列提取题目文本，组成完整问题名
            first_rest = subcols[0].split(root)[1].strip()
            if ':' in first_rest:
                q_text = first_rest.split(':', 1)[0].strip()
            elif '：' in first_rest:
                q_text = first_rest.split('：', 1)[0].strip()
            else:
                q_text = first_rest
            full_question = f"{root}{q_text}"

            for subcol in subcols:
                rest = subcol.split(root)[1].strip()
                # 只取冒号后面的选项文本
                if ':' in rest:
                    option_text = rest.split(':', 1)[1].strip()
                elif '：' in rest:
                    option_text = rest.split('：', 1)[1].strip()
                else:
                    option_text = rest
                cond = df[subcol] == 1
                row_conditions.append(((full_question, option_text), cond))
            total_cond = (df[subcols] == 1).any(axis=1)
            row_conditions.append(((full_question, "总计"), total_cond))
        else:
            values = df[q].dropna().unique()
            try:
                sorted_values = sorted(
                    values,
                    key=lambda x: int(re.match(r'^(\d+)', str(x)).group(1))
                )
            except Exception:
                sorted_values = sorted(values, key=str)
            for value in sorted_values:
                cond = df[q] == value
                row_conditions.append(((q, str(value)), cond))
            total_cond = df[q].notna()
            row_conditions.append(((q, "总计"), total_cond))

    # --- 交叉统计计算 ---
    freq_results = []
    for (r_question, r_option), r_cond in row_conditions:
        row_data = {}
        for c_label, c_cond in col_conditions:
            count = int((r_cond & c_cond).sum())
            row_data[c_label] = count
        freq_results.append(row_data)

    # 创建多级索引
    index = pd.MultiIndex.from_tuples(
        [(rl[0], rl[1]) for rl, _ in row_conditions],
        names=["问题", "选项"]
    )
    col_labels = [cl for cl, _ in col_conditions]

    freq_df = pd.DataFrame(freq_results, index=index, columns=col_labels)

    # --- 列百分比 ---
    percent_df = freq_df.astype(float).copy()
    for question in percent_df.index.get_level_values(0).unique():
        q_mask = percent_df.index.get_level_values(0) == question
        total_idx = (question, "总计")
        if total_idx in freq_df.index:
            denom = freq_df.loc[total_idx].replace(0, np.nan)
        else:
            denom = pd.Series(col_totals).reindex(percent_df.columns).replace(0, np.nan)
        percent_df.loc[q_mask] = freq_df.loc[q_mask].div(denom, axis=1)
    percent_df = percent_df.fillna(0)

    # --- 缓存结果 ---
    result = {
        "freq_df": freq_df,
        "percent_df": percent_df,
        "col_totals": col_totals,
        "row_conditions_info": [(rl, None) for rl, _ in row_conditions],
        "valid_rows_map": {q: q_type for q_type, q in valid_rows},
        "col_labels": col_labels,
    }
    _CACHE[file_path]["crosstab_result"] = result

    # --- 返回 JSON 友好的摘要 ---
    freq_summary = {}
    for (q, opt) in freq_df.index:
        if q not in freq_summary:
            freq_summary[q] = {}
        freq_summary[q][opt] = {
            col: int(freq_df.loc[(q, opt), col])
            for col in freq_df.columns
        }

    percent_summary = {}
    for (q, opt) in percent_df.index:
        if q not in percent_summary:
            percent_summary[q] = {}
        percent_summary[q][opt] = {
            col: round(float(percent_df.loc[(q, opt), col]), 4)
            for col in percent_df.columns
        }

    return {
        "status": "success",
        "row_questions_count": len(valid_rows),
        "col_conditions_count": len(col_conditions),
        "total_cells": len(freq_df) * len(freq_df.columns),
        "invalid_rows": invalid_rows,
        "invalid_cols": invalid_cols,
        "col_totals": col_totals,
        "freq_table": freq_summary,
        "percent_table": percent_summary,
    }


# ========================================================================= #
#                      满意度 / NPS 得分计算
# ========================================================================= #

def _detect_score_type(question_name: str, df: pd.DataFrame) -> str:
    """
    自动识别题目是满意度还是 NPS。
    返回: "satisfaction" | "nps"
    """
    q_lower = question_name.lower()

    # 关键词识别
    if "nps" in q_lower or "推荐" in question_name:
        return "nps"
    if "满意度" in question_name or "满意" in question_name:
        return "satisfaction"

    # 按分值范围判断
    if question_name in df.columns:
        values = df[question_name].dropna().unique()
        numeric_vals = []
        for v in values:
            score = _extract_score_from_option(v)
            if score is not None:
                numeric_vals.append(score)
        if numeric_vals:
            min_val, max_val = min(numeric_vals), max(numeric_vals)
            if min_val >= 0 and max_val >= 9:
                return "nps"

    return "satisfaction"


def calc_scores(file_path: str, score_questions: list[str]) -> dict:
    """
    计算满意度得分或 NPS。

    自动检测题目类型:
    - 满意度: 加权均值
    - NPS: (推荐者9-10分占比 - 贬损者0-6分占比) × 100

    Args:
        file_path: 数据文件路径
        score_questions: 需计算得分的题目列表

    Returns:
        得分结果字典
    """
    if file_path not in _CACHE or "crosstab_result" not in _CACHE[file_path]:
        return {"error": "请先执行 run_crosstab 交叉分析"}

    df = get_cached_df(file_path)
    ct_result = _CACHE[file_path]["crosstab_result"]
    freq_df = ct_result["freq_df"]
    row_type_map = ct_result["valid_rows_map"]

    score_results = []
    score_index = []
    score_type_info = {}

    for q in score_questions:
        q = str(q).strip()

        # 验证
        if q not in freq_df.index.get_level_values(0).unique():
            warnings.warn(f"题目 '{q}' 不在行变量中，已跳过")
            continue
        if row_type_map.get(q) != "single":
            warnings.warn(f"得分计算仅支持单选/量表题，已跳过：{q}")
            continue

        # 自动识别类型
        score_type = _detect_score_type(q, df)
        score_type_info[q] = score_type

        q_slice = freq_df.xs(q, level=0)

        if score_type == "satisfaction":
            # --- 满意度：加权均值 ---
            value_map = {}
            for opt in q_slice.index:
                opt_str = str(opt).strip()
                if opt_str in ("总计", "合计", "Total"):
                    continue
                score_val = _extract_score_from_option(opt_str)
                if score_val is not None:
                    value_map[opt] = score_val

            if not value_map:
                warnings.warn(f"题目 '{q}' 未找到可用数值选项，已跳过")
                continue

            q_counts = q_slice.loc[list(value_map.keys())]
            weights = pd.Series(value_map)
            numerator = (q_counts.T * weights).T.sum(axis=0)
            denominator = q_counts.sum(axis=0).replace(0, np.nan)
            score = numerator / denominator

            score_results.append(score)
            score_index.append((q, "满意度得分(加权均值)"))

        else:
            # --- NPS ---
            value_map = {}
            for opt in q_slice.index:
                opt_str = str(opt).strip()
                if opt_str in ("总计", "合计", "Total"):
                    continue
                score_val = _extract_score_from_option(opt_str)
                if score_val is not None:
                    value_map[opt] = score_val

            if not value_map:
                warnings.warn(f"题目 '{q}' 未找到可用数值选项，已跳过")
                continue

            promoter_opts = [opt for opt, s in value_map.items() if s >= 9]
            detractor_opts = [opt for opt, s in value_map.items() if s <= 6]

            q_counts = q_slice.loc[list(value_map.keys())]
            total_count = q_counts.sum(axis=0).replace(0, np.nan)

            promoter_count = q_counts.loc[promoter_opts].sum(axis=0) if promoter_opts else 0
            detractor_count = q_counts.loc[detractor_opts].sum(axis=0) if detractor_opts else 0

            nps_score = ((promoter_count - detractor_count) / total_count) * 100

            score_results.append(nps_score)
            score_index.append((q, "NPS得分"))

    if not score_results:
        return {"status": "no_valid_questions", "message": "没有找到可计算得分的有效题目"}

    score_df = pd.DataFrame(
        score_results,
        index=pd.MultiIndex.from_tuples(score_index, names=["问题", "指标"]),
    )
    score_df = score_df.reindex(columns=freq_df.columns)

    # 缓存
    _CACHE[file_path]["score_df"] = score_df

    # 返回结果
    score_summary = {}
    for (q, indicator) in score_df.index:
        if q not in score_summary:
            score_summary[q] = {}
        score_summary[q][indicator] = {
            col: round(float(score_df.loc[(q, indicator), col]), 4)
            for col in score_df.columns
        }

    return {
        "status": "success",
        "score_types": score_type_info,
        "scores": score_summary,
    }


# ========================================================================= #
#                      差异摘要
# ========================================================================= #

def get_crosstab_summary(file_path: str) -> dict:
    """
    从交叉分析结果中提取关键差异摘要，帮助 AI 撰写报告。

    Returns:
        各题目在不同分组间的差异摘要
    """
    if file_path not in _CACHE or "crosstab_result" not in _CACHE[file_path]:
        return {"error": "请先执行 run_crosstab 交叉分析"}

    ct_result = _CACHE[file_path]["crosstab_result"]
    percent_df = ct_result["percent_df"]
    freq_df = ct_result["freq_df"]
    col_labels = ct_result["col_labels"]

    # 找到"总计"列和非总计列
    total_cols = [c for c in col_labels if c.endswith("\n总计")]
    non_total_cols = [c for c in col_labels if not c.endswith("\n总计")]

    summary = {}

    for question in percent_df.index.get_level_values(0).unique():
        q_data = percent_df.xs(question, level=0)

        # 跳过总计行
        option_rows = [opt for opt in q_data.index if opt != "总计"]
        if not option_rows:
            continue

        question_summary = {
            "options": {},
            "max_diff_option": None,
            "max_diff_value": 0,
        }

        for opt in option_rows:
            opt_percents = {}
            for col in non_total_cols:
                pct = float(q_data.loc[opt, col]) if opt in q_data.index else 0
                opt_percents[col] = round(pct, 4)

            # 计算差异
            pct_values = list(opt_percents.values())
            if pct_values:
                diff = max(pct_values) - min(pct_values)
            else:
                diff = 0

            question_summary["options"][str(opt)] = {
                "percentages": opt_percents,
                "max_min_diff": round(diff, 4),
            }

            if diff > question_summary["max_diff_value"]:
                question_summary["max_diff_value"] = round(diff, 4)
                question_summary["max_diff_option"] = str(opt)

        summary[question] = question_summary

    # 得分摘要
    score_summary = None
    if "score_df" in _CACHE.get(file_path, {}):
        score_df = _CACHE[file_path]["score_df"]
        score_summary = {}
        for (q, indicator) in score_df.index:
            scores_by_col = {}
            for col in non_total_cols:
                if col in score_df.columns:
                    scores_by_col[col] = round(float(score_df.loc[(q, indicator), col]), 4)
            score_summary[f"{q} - {indicator}"] = scores_by_col

    return {
        "status": "success",
        "questions_analyzed": len(summary),
        "question_summaries": summary,
        "score_summary": score_summary,
    }
