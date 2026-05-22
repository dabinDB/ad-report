from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .aggregation import aggregate_table, build_total_row
from .dictionary import StandardDictionary


def fill_workbook(
    template_bytes: bytes,
    source_df: pd.DataFrame,
    definitions: list[dict[str, Any]],
    dictionary: StandardDictionary,
) -> bytes:
    workbook = load_workbook(BytesIO(template_bytes))
    for definition in definitions:
        location = definition.get("location", {})
        sheet = workbook[location["sheet"]]
        result = aggregate_table(source_df, definition, dictionary)
        start_row = int(location["data_start_row"])
        end_row = int(location["data_end_row"])
        label_col_idx = column_index_from_string(location["label_col"])
        metric_columns = location.get("columns", {})
        group_by = definition.get("group_by", [])

        _clear_range(sheet, start_row, end_row, [location["label_col"], *metric_columns.keys()])

        write_row = start_row
        total_spec = definition.get("total_row", {})
        if total_spec.get("enabled") and total_spec.get("position", "top") == "top":
            total = build_total_row(source_df, definition, dictionary)
            sheet.cell(write_row, label_col_idx).value = total_spec.get("label", "합계")
            _write_metric_cells(sheet, write_row, metric_columns, total)
            write_row += 1

        for _, row in result.iterrows():
            if write_row > end_row:
                break
            sheet.cell(write_row, label_col_idx).value = _label_value(row, group_by)
            _write_metric_cells(sheet, write_row, metric_columns, row.to_dict())
            write_row += 1

        if total_spec.get("enabled") and total_spec.get("position") == "bottom" and write_row <= end_row:
            total = build_total_row(source_df, definition, dictionary)
            sheet.cell(write_row, label_col_idx).value = total_spec.get("label", "합계")
            _write_metric_cells(sheet, write_row, metric_columns, total)

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _clear_range(sheet: Any, start_row: int, end_row: int, columns: list[str]) -> None:
    for row in range(start_row, end_row + 1):
        for column in columns:
            sheet.cell(row, column_index_from_string(column)).value = None


def _write_metric_cells(sheet: Any, row: int, metric_columns: dict[str, str], values: dict[str, Any]) -> None:
    for column, metric in metric_columns.items():
        value = values.get(metric)
        sheet.cell(row, column_index_from_string(column)).value = _scalar(value)


def _label_value(row: pd.Series, group_by: list[str]) -> str:
    values = [str(row.get(column, "")) for column in group_by]
    return " / ".join(value for value in values if value and value != "nan")


def _scalar(value: Any) -> Any:
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        return value.item()
    return value
