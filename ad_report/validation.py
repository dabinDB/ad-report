from __future__ import annotations

from typing import Any

from .dictionary import StandardDictionary
from .workbook_writer import normalize_excel_column


def validate_definition(
    definition: dict[str, Any],
    schema: dict[str, Any],
    dictionary: StandardDictionary,
    field_roles: dict[str, str] | None = None,
) -> tuple[list[str], list[str]]:
    from jsonschema import Draft7Validator

    errors = []
    warnings = []
    validator = Draft7Validator(schema)
    errors.extend(error.message for error in validator.iter_errors(definition))

    group_by = definition.get("group_by", [])
    metrics = definition.get("metrics", [])
    sort_by = (definition.get("sort") or {}).get("by")
    location = definition.get("location", {})

    field_roles = field_roles or {}

    for dimension in group_by:
        if not dictionary.is_dimension(dimension) and field_roles.get(dimension) != "dimension":
            errors.append(f"group_by 값이 표준 차원 또는 확정된 dimension에 없습니다: {dimension}")
    for metric in metrics:
        if not dictionary.is_metric(metric) and field_roles.get(metric) != "metric":
            errors.append(f"metrics 값이 표준 지표 또는 확정된 metric에 없습니다: {metric}")
    if sort_by and sort_by not in group_by and sort_by not in metrics:
        warnings.append(f"sort.by가 group_by 또는 metrics에 없습니다: {sort_by}")
    if location:
        _validate_location_columns(location, errors)
    if definition.get("limit") and definition.get("total_row", {}).get("enabled"):
        warnings.append("limit이 있는 TOP 표에 합계 행이 켜져 있습니다. 의미가 모호할 수 있습니다.")
    if "요일" in group_by and sort_by != "요일":
        warnings.append("요일 표는 요일 기준 정렬을 권장합니다.")
    for metric in metrics:
        spec = dictionary.metric_spec(metric)
        if spec.get("aggregation") == "weighted_avg":
            bases = [spec.get("numerator"), spec.get("denominator")]
            missing = [base for base in bases if base and base not in metrics]
            if missing:
                warnings.append(f"{metric} 재계산에는 원본 데이터의 {', '.join(missing)}가 필요합니다.")
    return errors, warnings


def _validate_location_columns(location: dict[str, Any], errors: list[str]) -> None:
    try:
        normalize_excel_column(location.get("label_col"))
    except ValueError as exc:
        errors.append(f"location.label_col 오류: {exc}")

    for column in (location.get("columns") or {}).keys():
        try:
            normalize_excel_column(column)
        except ValueError as exc:
            errors.append(f"location.columns 오류: {exc}")
