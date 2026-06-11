from __future__ import annotations

from typing import Any

import pandas as pd

from .dictionary import StandardDictionary


WEEKDAY_ORDER = ["월", "화", "수", "목", "금", "토", "일"]


def build_source_mapping_schema(df: pd.DataFrame, dictionary: StandardDictionary) -> list[dict[str, Any]]:
    schema = []
    mapped_counts: dict[str, int] = {}
    for column in df.columns:
        matched = dictionary.match(column)
        role = "ignore"
        mapped_to = None
        confidence = 0.0
        reason = "사전 매칭 없음"
        if matched:
            mapped_counts[matched] = mapped_counts.get(matched, 0) + 1
            role = "dimension" if dictionary.is_dimension(matched) else "metric"
            mapped_to = matched
            confidence = 0.98
            reason = "표준 사전 매칭"
        else:
            role, mapped_to, confidence, reason = _infer_field_mapping(df[column], column)
        schema.append(
            {
                "source_column": str(column),
                "dictionary_match": matched,
                "ai_role": role,
                "ai_mapped_to": mapped_to,
                "confidence": confidence,
                "needs_review": not matched or confidence < 0.9,
                "final_role": role,
                "final_name": mapped_to,
                "kind": role,
                "mapped_to": mapped_to,
                "reason": reason,
                "aggregation_when_duplicate": _duplicate_policy(mapped_to, dictionary),
            }
        )

    for item in schema:
        mapped_to = item.get("mapped_to")
        item["duplicate_group_size"] = mapped_counts.get(mapped_to, 0) if mapped_to else 0
    return schema


def normalize_source_dataframe(
    df: pd.DataFrame,
    dictionary: StandardDictionary,
    field_mappings: list[dict[str, Any]] | None = None,
) -> pd.DataFrame:
    if field_mappings:
        normalized = _apply_field_mappings(df, field_mappings)
    else:
        renamed = {}
        for column in df.columns:
            matched = dictionary.match(column)
            if matched:
                renamed[column] = matched
        normalized = df.rename(columns=renamed).copy()
    normalized = _consolidate_duplicate_columns(normalized, dictionary)
    normalized = _derive_time_dimensions(normalized, dictionary)
    normalized = _apply_value_mappings(normalized, dictionary)

    for metric in dictionary.metric_names:
        if metric in normalized.columns:
            normalized[metric] = _numeric_series(normalized, metric)
    for mapping in field_mappings or []:
        if mapping.get("final_role") == "metric":
            name = str(mapping.get("final_name") or mapping.get("source_column") or "")
            if name in normalized.columns:
                normalized[name] = _numeric_series(normalized, name)
    return normalized


def aggregate_table(
    df: pd.DataFrame,
    definition: dict[str, Any],
    dictionary: StandardDictionary,
) -> pd.DataFrame:
    group_by = definition.get("group_by", [])
    metrics = definition.get("metrics", [])
    if not group_by:
        raise ValueError("group_by가 비어 있습니다.")
    missing = [column for column in group_by if column not in df.columns]
    if missing:
        raise ValueError(f"원본 데이터에 그룹 기준 컬럼이 없습니다: {', '.join(missing)}")

    filtered = _apply_filters(df, definition.get("filters", []))
    grouped = filtered.groupby(group_by, dropna=False, as_index=False)
    output = grouped.size().drop(columns=["size"])

    for metric in metrics:
        spec = dictionary.metric_spec(metric)
        aggregation = spec.get("aggregation", "sum")
        if aggregation == "sum":
            output[metric] = grouped[metric].sum()[metric].values if metric in filtered.columns else 0
        elif aggregation == "weighted_avg":
            output[metric] = _weighted_metric(grouped, filtered, spec)
        else:
            output[metric] = grouped[metric].mean()[metric].values if metric in filtered.columns else 0

    output = _sort_output(output, definition, group_by)
    limit = definition.get("limit")
    if limit:
        output = output.head(int(limit))
    return output.reset_index(drop=True)


def build_total_row(df: pd.DataFrame, definition: dict[str, Any], dictionary: StandardDictionary) -> dict[str, Any]:
    metrics = definition.get("metrics", [])
    row: dict[str, Any] = {}
    for metric in metrics:
        spec = dictionary.metric_spec(metric)
        if spec.get("aggregation") == "weighted_avg":
            numerator = spec.get("numerator")
            denominator = spec.get("denominator")
            multiplier = spec.get("multiplier", 1)
            num = df[numerator].sum() if numerator in df.columns else 0
            den = df[denominator].sum() if denominator in df.columns else 0
            row[metric] = 0 if den == 0 else num / den * multiplier
        else:
            row[metric] = df[metric].sum() if metric in df.columns else 0
    return row


def _weighted_metric(grouped: Any, df: pd.DataFrame, spec: dict[str, Any]) -> list[float]:
    numerator = spec.get("numerator")
    denominator = spec.get("denominator")
    multiplier = spec.get("multiplier", 1)
    if numerator not in df.columns or denominator not in df.columns:
        return [0.0 for _ in range(grouped.ngroups)]
    sums = grouped[[numerator, denominator]].sum()
    return [
        0 if row[denominator] == 0 else row[numerator] / row[denominator] * multiplier
        for _, row in sums.iterrows()
    ]


def _consolidate_duplicate_columns(df: pd.DataFrame, dictionary: StandardDictionary) -> pd.DataFrame:
    output = pd.DataFrame(index=df.index)
    for column in dict.fromkeys(df.columns):
        values = df.loc[:, df.columns == column]
        if values.shape[1] == 1:
            output[column] = values.iloc[:, 0]
        elif dictionary.is_metric(column):
            numeric_values = values.apply(pd.to_numeric, errors="coerce").fillna(0)
            output[column] = numeric_values.sum(axis=1)
        else:
            output[column] = values.replace("", pd.NA).bfill(axis=1).iloc[:, 0]
    return output


def _numeric_series(df: pd.DataFrame, column: str) -> pd.Series:
    values = df.loc[:, df.columns == column]
    if values.shape[1] > 1:
        values = values.apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    else:
        values = values.iloc[:, 0]
    return pd.to_numeric(values, errors="coerce").fillna(0)


def _apply_field_mappings(df: pd.DataFrame, field_mappings: list[dict[str, Any]]) -> pd.DataFrame:
    selected_columns = []
    renamed_columns = []
    for mapping in field_mappings:
        role = mapping.get("final_role") or mapping.get("kind")
        if role == "ignore":
            continue
        source_column = mapping.get("source_column")
        if source_column not in df.columns:
            continue
        final_name = str(mapping.get("final_name") or source_column).strip()
        if not final_name:
            continue
        selected_columns.append(source_column)
        renamed_columns.append(final_name)
    if not selected_columns:
        return df.copy()
    output = df.loc[:, selected_columns].copy()
    output.columns = renamed_columns
    return output


def _infer_field_mapping(series: pd.Series, column: Any) -> tuple[str, str, float, str]:
    name = str(column).strip()
    non_empty = series.dropna().head(300)
    if len(non_empty) == 0:
        return "ignore", name, 0.2, "빈 컬럼"

    numeric = pd.to_numeric(non_empty, errors="coerce")
    numeric_ratio = float(numeric.notna().mean()) if len(non_empty) else 0.0
    if numeric_ratio >= 0.8:
        return "metric", name, 0.72, "숫자형 값 비율이 높음"

    if _looks_date_like_name(name):
        dates = pd.to_datetime(non_empty, errors="coerce")
        date_ratio = float(dates.notna().mean()) if len(non_empty) else 0.0
        if date_ratio >= 0.7:
            return "dimension", name, 0.72, "날짜형 값 비율이 높음"

    unique_count = non_empty.astype(str).nunique(dropna=True)
    unique_ratio = unique_count / max(len(non_empty), 1)
    if unique_count <= 50 or unique_ratio <= 0.25:
        return "dimension", name, 0.62, "반복되는 범주형 값"
    return "dimension", name, 0.45, "텍스트 컬럼"


def _looks_date_like_name(name: str) -> bool:
    lowered = name.lower()
    return any(token in lowered for token in ["날짜", "일자", "date", "day", "월", "년"])


def _duplicate_policy(mapped_to: str | None, dictionary: StandardDictionary) -> str | None:
    if not mapped_to:
        return None
    if dictionary.is_metric(mapped_to):
        return "sum_columns"
    if dictionary.is_dimension(mapped_to):
        return "first_non_empty"
    return None


def _sort_output(output: pd.DataFrame, definition: dict[str, Any], group_by: list[str]) -> pd.DataFrame:
    sort = definition.get("sort") or {}
    sort_by = sort.get("by") or (group_by[0] if group_by else None)
    ascending = sort.get("order", "desc") == "asc"
    if sort_by == "요일" and "요일" in output.columns:
        output["__weekday_order"] = pd.Categorical(output["요일"], WEEKDAY_ORDER, ordered=True)
        return output.sort_values("__weekday_order").drop(columns=["__weekday_order"])
    if sort_by and sort_by in output.columns:
        return output.sort_values(sort_by, ascending=ascending)
    return output


def _derive_time_dimensions(df: pd.DataFrame, dictionary: StandardDictionary) -> pd.DataFrame:
    if "날짜" not in df.columns:
        return df
    dates = pd.to_datetime(df["날짜"], errors="coerce")
    if "주차" in dictionary.derived_dimensions() and "주차" not in df.columns:
        start = dates - pd.to_timedelta(dates.dt.weekday, unit="D")
        end = start + pd.to_timedelta(6, unit="D")
        df["주차"] = start.dt.strftime("%Y-%m-%d") + " ~ " + end.dt.strftime("%Y-%m-%d")
    if "월" in dictionary.derived_dimensions() and "월" not in df.columns:
        df["월"] = dates.dt.strftime("%Y-%m")
    if "요일" in dictionary.derived_dimensions() and "요일" not in df.columns:
        df["요일"] = dates.dt.weekday.map(dict(enumerate(WEEKDAY_ORDER)))
    return df


def _apply_value_mappings(df: pd.DataFrame, dictionary: StandardDictionary) -> pd.DataFrame:
    for dimension, spec in dictionary.dimensions.items():
        mapping = spec.get("value_mapping", {})
        if dimension in df.columns and mapping:
            df[dimension] = df[dimension].replace(mapping)
    return df


def _apply_filters(df: pd.DataFrame, filters: list[dict[str, Any]]) -> pd.DataFrame:
    output = df
    for filter_spec in filters or []:
        field = filter_spec.get("field")
        operator = filter_spec.get("operator")
        value = filter_spec.get("value")
        if field not in output.columns:
            continue
        if operator == "==":
            output = output[output[field] == value]
        elif operator == "!=":
            output = output[output[field] != value]
        elif operator == ">":
            output = output[output[field] > value]
        elif operator == "<":
            output = output[output[field] < value]
        elif operator == ">=":
            output = output[output[field] >= value]
        elif operator == "<=":
            output = output[output[field] <= value]
        elif operator == "in":
            output = output[output[field].isin(value if isinstance(value, list) else [value])]
        elif operator == "not_in":
            output = output[~output[field].isin(value if isinstance(value, list) else [value])]
    return output
