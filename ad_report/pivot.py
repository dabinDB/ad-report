from __future__ import annotations

from io import BytesIO
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

from .dictionary import StandardDictionary


NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def detect_pivot_sources(workbook_bytes: bytes) -> list[dict[str, Any]]:
    sources = []
    with ZipFile(BytesIO(workbook_bytes), "r") as archive:
        pivot_defs = [
            name
            for name in archive.namelist()
            if name.startswith("xl/pivotCache/") and name.endswith(".xml")
        ]
        for index, path in enumerate(sorted(pivot_defs), start=1):
            root = ET.fromstring(archive.read(path))
            worksheet_source = root.find(".//main:worksheetSource", NS)
            if worksheet_source is None:
                continue
            ref = worksheet_source.attrib.get("ref")
            sheet = worksheet_source.attrib.get("sheet")
            source_name = worksheet_source.attrib.get("name")
            sources.append(
                {
                    "id": f"pivot_cache_{index}",
                    "cache_definition": path,
                    "sheet": sheet,
                    "ref": ref,
                    "name": source_name,
                    "display": _source_display(sheet, ref, source_name),
                    "refresh_on_load": root.attrib.get("refreshOnLoad") == "1",
                }
            )
    return sources


def update_pivot_source_data(
    workbook_bytes: bytes,
    source_df: pd.DataFrame,
    pivot_source: dict[str, Any],
    mode: str,
    dictionary: StandardDictionary,
) -> bytes:
    workbook = load_workbook(BytesIO(workbook_bytes))
    sheet_name = pivot_source.get("sheet")
    if not sheet_name or sheet_name not in workbook.sheetnames:
        raise ValueError(f"피벗 소스 시트를 찾을 수 없습니다: {sheet_name}")

    sheet = workbook[sheet_name]
    min_col, header_row, max_col, max_row = _source_bounds(sheet, pivot_source, source_df)
    headers = [
        sheet.cell(header_row, column).value or get_column_letter(column)
        for column in range(min_col, max_col + 1)
    ]
    aligned = _align_to_pivot_headers(source_df, headers, dictionary)

    if mode == "replace":
        _clear_source_rows(sheet, header_row + 1, max(max_row, sheet.max_row), min_col, max_col)
        write_row = header_row + 1
    else:
        write_row = _last_non_empty_row(sheet, header_row, min_col, max_col) + 1

    for row_values in aligned.itertuples(index=False, name=None):
        for offset, value in enumerate(row_values):
            sheet.cell(write_row, min_col + offset).value = _scalar(value)
        write_row += 1

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def enable_pivot_refresh_on_load(workbook_bytes: bytes) -> bytes:
    input_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()
    with ZipFile(input_buffer, "r") as source_zip, ZipFile(output_buffer, "w", ZIP_DEFLATED) as target_zip:
        for item in source_zip.infolist():
            data = source_zip.read(item.filename)
            if item.filename.startswith("xl/pivotCache/") and item.filename.endswith(".xml"):
                try:
                    root = ET.fromstring(data)
                    root.set("refreshOnLoad", "1")
                    root.set("enableRefresh", "1")
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                except ET.ParseError:
                    pass
            target_zip.writestr(item, data)
    return output_buffer.getvalue()


def strip_pivot_cache_records(workbook_bytes: bytes) -> bytes:
    input_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()
    empty_records = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>'
    )
    with ZipFile(input_buffer, "r") as source_zip, ZipFile(output_buffer, "w", ZIP_DEFLATED) as target_zip:
        for item in source_zip.infolist():
            data = source_zip.read(item.filename)
            if item.filename.startswith("xl/pivotCache/pivotCacheRecords") and item.filename.endswith(".xml"):
                data = empty_records
            elif item.filename.startswith("xl/pivotCache/") and item.filename.endswith(".xml"):
                data = _set_pivot_refresh_attributes(data)
            target_zip.writestr(item, data)
    return output_buffer.getvalue()


def _set_pivot_refresh_attributes(data: bytes) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return data
    root.set("refreshOnLoad", "1")
    root.set("enableRefresh", "1")
    root.set("saveData", "0")
    if root.tag.endswith("pivotCacheDefinition"):
        root.set("recordCount", "0")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _source_display(sheet: str | None, ref: str | None, source_name: str | None) -> str:
    if sheet and ref:
        return f"{sheet}!{ref}"
    if source_name:
        return source_name
    return "피벗 소스 범위 미확인"


def _source_bounds(sheet: Any, pivot_source: dict[str, Any], source_df: pd.DataFrame) -> tuple[int, int, int, int]:
    ref = pivot_source.get("ref")
    if ref:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        return min_col, min_row, max_col, max_row
    max_col = max(1, len(source_df.columns))
    return 1, 1, max_col, max(sheet.max_row, 1)


def _align_to_pivot_headers(
    source_df: pd.DataFrame,
    headers: list[Any],
    dictionary: StandardDictionary,
) -> pd.DataFrame:
    aligned = pd.DataFrame()
    for header in headers:
        header_text = str(header).strip()
        standard_name = dictionary.match(header_text) or header_text
        if header_text in source_df.columns:
            aligned[header_text] = source_df[header_text]
        elif standard_name in source_df.columns:
            aligned[header_text] = source_df[standard_name]
        else:
            aligned[header_text] = None
    return aligned


def _clear_source_rows(sheet: Any, start_row: int, end_row: int, min_col: int, max_col: int) -> None:
    for row in range(start_row, end_row + 1):
        for column in range(min_col, max_col + 1):
            sheet.cell(row, column).value = None


def _last_non_empty_row(sheet: Any, header_row: int, min_col: int, max_col: int) -> int:
    last = header_row
    for row in range(header_row + 1, sheet.max_row + 1):
        values = [sheet.cell(row, column).value for column in range(min_col, max_col + 1)]
        if any(value is not None and str(value).strip() for value in values):
            last = row
    return last


def _scalar(value: Any) -> Any:
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        return value.item()
    return value
