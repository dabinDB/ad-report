from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, TimeoutError
import json
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import yaml

from ad_report.aggregation import aggregate_table, build_source_mapping_schema, normalize_source_dataframe
from ad_report.dictionary import StandardDictionary, load_yaml
from ad_report.pivot import (
    detect_pivot_sources,
    enable_pivot_refresh_on_load,
    strip_pivot_cache_records,
    update_pivot_source_data,
)
from ad_report.template_analyzer import analyze_template, analyze_with_gemini, list_workbook_sheets
from ad_report.validation import validate_definition
from ad_report.workbook_writer import fill_workbook, normalize_excel_column


ROOT = Path(__file__).parent
STANDARD_DIMENSIONS_PATH = ROOT / "config" / "standard_dimensions.yaml"
TABLE_SCHEMA_PATH = ROOT / "config" / "table_definition_schema.yaml"
GEMINI_TIMEOUT_SECONDS = 25


st.set_page_config(page_title="AI 광고 보고서 생성기", page_icon="bar_chart", layout="wide")


@st.cache_data
def load_configs() -> tuple[dict[str, Any], dict[str, Any]]:
    return load_yaml(STANDARD_DIMENSIONS_PATH), load_yaml(TABLE_SCHEMA_PATH)


def read_source_file(uploaded_file: Any) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith(".tsv"):
        return pd.read_csv(uploaded_file, sep="\t")
    return pd.read_excel(uploaded_file)


def apply_uploaded_media_name(df: pd.DataFrame, media_name: str, dictionary: StandardDictionary) -> pd.DataFrame:
    output = df.copy()
    if not media_name.strip():
        return output
    has_media_column = any(dictionary.match(column) == "매체" for column in output.columns)
    if not has_media_column:
        output["매체"] = media_name.strip()
    return output


def build_pivot_source_dataframe(source_raw: pd.DataFrame, source_df: pd.DataFrame) -> pd.DataFrame:
    output = source_raw.copy()
    for column in source_df.columns:
        if column not in output.columns:
            output[column] = source_df[column]
    return output


def analyze_definitions(
    template_bytes: bytes,
    dictionary: StandardDictionary,
    schema: dict[str, Any],
    use_gemini: bool,
    api_key: str,
    model: str,
    sheet_names: list[str] | None = None,
) -> list[dict[str, Any]]:
    local_definitions = analyze_template(template_bytes, dictionary, sheet_names)
    definitions = local_definitions
    if use_gemini and api_key:
        gemini_definitions = _run_gemini_with_timeout(template_bytes, dictionary, schema, api_key, model, sheet_names)
        if gemini_definitions:
            definitions = gemini_definitions
    return [normalize_definition_location(definition, dictionary) for definition in definitions]


def _run_gemini_with_timeout(
    template_bytes: bytes,
    dictionary: StandardDictionary,
    schema: dict[str, Any],
    api_key: str,
    model: str,
    sheet_names: list[str] | None = None,
) -> list[dict[str, Any]]:
    definitions: list[dict[str, Any]] = []
    batches = [[sheet_name] for sheet_name in sheet_names] if sheet_names else [None]
    for batch in batches:
        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(analyze_with_gemini, template_bytes, dictionary, schema, api_key, model, batch)
        try:
            definitions.extend(future.result(timeout=GEMINI_TIMEOUT_SECONDS))
        except TimeoutError:
            future.cancel()
        except Exception:
            pass
        finally:
            executor.shutdown(wait=False, cancel_futures=True)
    return definitions


def normalize_definition_location(definition: dict[str, Any], dictionary: StandardDictionary) -> dict[str, Any]:
    normalized = json.loads(json.dumps(definition, ensure_ascii=False))
    location = normalized.setdefault("location", {})
    columns = location.get("columns") or {}
    fixed_columns: dict[str, str] = {}
    group_by = set(normalized.get("group_by", []))

    for key, value in columns.items():
        column, field = _split_column_mapping(key, value)
        if not column or not field:
            fixed_columns[str(key)] = str(value)
            continue
        if dictionary.is_dimension(field) and (field in group_by or not location.get("label_col")):
            location["label_col"] = column
        elif dictionary.is_metric(field):
            fixed_columns[column] = field
        else:
            fixed_columns[column] = field

    if location.get("label_col") is not None:
        try:
            location["label_col"] = normalize_excel_column(location["label_col"])
        except ValueError:
            pass
    location["columns"] = fixed_columns
    normalized["total_row"] = _normalize_summary_row(normalized.get("total_row"), "합계")
    normalized["compare_row"] = _normalize_summary_row(normalized.get("compare_row"), "전월 비교")
    normalized["average_row"] = _normalize_summary_row(normalized.get("average_row"), "일평균")
    if _is_average_label(normalized["total_row"].get("label")) and not normalized["average_row"].get("enabled"):
        normalized["average_row"] = {**normalized["total_row"], "mode": "daily_average"}
        normalized["total_row"] = {"enabled": False, "label": "합계"}
    return normalized


def _normalize_summary_row(value: Any, default_label: str) -> dict[str, Any]:
    if isinstance(value, int):
        return {"enabled": True, "row": value, "label": default_label}
    if isinstance(value, dict):
        row = value.get("row")
        normalized = {
            "enabled": bool(value.get("enabled", row is not None)),
            "label": value.get("label", default_label),
        }
        if row:
            normalized["row"] = int(row)
        if value.get("position"):
            normalized["position"] = value.get("position")
        if value.get("mode"):
            normalized["mode"] = value.get("mode")
        return normalized
    return {"enabled": False, "label": default_label}


def _split_column_mapping(key: Any, value: Any) -> tuple[str | None, str | None]:
    try:
        return normalize_excel_column(key), str(value)
    except ValueError:
        try:
            return normalize_excel_column(value), str(key)
        except ValueError:
            return None, None


def _is_average_label(label: Any) -> bool:
    return "평균" in str(label or "")


def group_definitions_by_sheet(
    definitions: list[dict[str, Any]],
    sheet_names: list[str] | None = None,
) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = {sheet: [] for sheet in sheet_names or []}
    for definition in definitions:
        sheet = definition.get("location", {}).get("sheet", "시트 미지정")
        grouped.setdefault(sheet, []).append(definition)
    return grouped


def render_sheet_definition_json(
    definitions: list[dict[str, Any]],
    sheet_names: list[str] | None = None,
) -> None:
    grouped = group_definitions_by_sheet(definitions, sheet_names)
    if not grouped:
        st.info("분석 결과가 없습니다.")
        return
    tabs = st.tabs([f"{sheet} ({len(items)})" for sheet, items in grouped.items()])
    for tab, (sheet, items) in zip(tabs, grouped.items()):
        with tab:
            st.code(
                json.dumps({"sheet": sheet, "definitions": items}, ensure_ascii=False, indent=2),
                language="json",
            )


def definition_editor(definition: dict[str, Any], index: int, dictionary: StandardDictionary) -> dict[str, Any]:
    edited = normalize_definition_location(definition, dictionary)
    with st.expander(f"{index + 1}. {definition.get('name', '표 정의')}", expanded=index == 0):
        left, mid, right = st.columns([1.2, 1, 1])
        edited["name"] = left.text_input("표 이름", edited.get("name", ""), key=f"name_{index}")
        edited["group_by"] = mid.multiselect(
            "Group By",
            dictionary.dimension_names,
            default=[item for item in edited.get("group_by", []) if item in dictionary.dimension_names],
            key=f"group_{index}",
        )
        edited["metrics"] = right.multiselect(
            "Metrics",
            dictionary.metric_names,
            default=[item for item in edited.get("metrics", []) if item in dictionary.metric_names],
            key=f"metrics_{index}",
        )

        sort = edited.setdefault("sort", {})
        sort_options = edited["group_by"] + edited["metrics"]
        sort_default = sort.get("by") if sort.get("by") in sort_options else (sort_options[0] if sort_options else "")
        loc = edited.setdefault("location", {})
        c1, c2, c3 = st.columns(3)
        sort["by"] = c1.selectbox("정렬 기준", sort_options or [""], index=(sort_options or [""]).index(sort_default), key=f"sort_by_{index}")
        sort["order"] = c2.selectbox("정렬 순서", ["desc", "asc"], index=0 if sort.get("order", "desc") == "desc" else 1, key=f"sort_order_{index}")
        limit_value = c3.number_input("Limit", min_value=0, value=int(edited.get("limit") or 0), key=f"limit_{index}")
        if limit_value:
            edited["limit"] = int(limit_value)
        else:
            edited.pop("limit", None)

        total_row = edited.setdefault("total_row", {"enabled": False, "label": "합계"})
        compare_row = edited.setdefault("compare_row", {"enabled": False, "label": "전월 비교"})
        average_row = edited.setdefault("average_row", {"enabled": False, "label": "일평균"})
        r1, r2, r3, r4, r5, r6 = st.columns(6)
        total_row["enabled"] = r1.checkbox("합계 행", value=bool(total_row.get("enabled")), key=f"total_{index}")
        total_row_num = r2.number_input("합계 행 번호", min_value=0, value=int(total_row.get("row") or 0), key=f"total_row_{index}")
        average_row["enabled"] = r3.checkbox("평균 행", value=bool(average_row.get("enabled")), key=f"average_{index}")
        average_row_num = r4.number_input("평균 행 번호", min_value=0, value=int(average_row.get("row") or 0), key=f"average_row_{index}")
        compare_row["enabled"] = r5.checkbox("비교 행", value=bool(compare_row.get("enabled")), key=f"compare_{index}")
        compare_row_num = r6.number_input("비교 행 번호", min_value=0, value=int(compare_row.get("row") or 0), key=f"compare_row_{index}")
        if total_row_num:
            total_row["row"] = int(total_row_num)
        else:
            total_row.pop("row", None)
        if compare_row_num:
            compare_row["row"] = int(compare_row_num)
            compare_row.setdefault("mode", "previous_row")
        else:
            compare_row.pop("row", None)
        if average_row_num:
            average_row["row"] = int(average_row_num)
            average_row.setdefault("mode", "daily_average")
        else:
            average_row.pop("row", None)

        st.caption("템플릿 위치")
        p1, p2, p3, p4, p5 = st.columns(5)
        loc["sheet"] = p1.text_input("시트", loc.get("sheet", ""), key=f"sheet_{index}")
        loc["label_col"] = p2.text_input("라벨 컬럼", loc.get("label_col", "A"), key=f"label_{index}").upper()
        loc["header_row"] = int(p3.number_input("헤더 행", min_value=1, value=int(loc.get("header_row", 1)), key=f"header_{index}"))
        loc["data_start_row"] = int(p4.number_input("시작 행", min_value=1, value=int(loc.get("data_start_row", 2)), key=f"start_{index}"))
        loc["data_end_row"] = int(p5.number_input("끝 행", min_value=1, value=int(loc.get("data_end_row", 20)), key=f"end_{index}"))

        columns_json = st.text_area(
            "엑셀 컬럼 → 지표 매핑(JSON)",
            json.dumps(loc.get("columns", {}), ensure_ascii=False, indent=2),
            key=f"columns_{index}",
            height=120,
        )
        try:
            loc["columns"] = json.loads(columns_json)
            edited = normalize_definition_location(edited, dictionary)
        except json.JSONDecodeError:
            st.error("컬럼 매핑 JSON 형식이 올바르지 않습니다.")

        st.code(yaml.safe_dump(edited, allow_unicode=True, sort_keys=False), language="yaml")
    return edited


def main() -> None:
    standard_config, table_config = load_configs()
    dictionary = StandardDictionary(standard_config)
    schema = table_config.get("schema", {})

    st.title("AI 광고 보고서 생성 자동화")
    st.caption("엑셀 템플릿을 표 타입이 아니라 group_by, sort, limit, metrics, location 구성요소로 분석합니다.")

    with st.sidebar:
        st.header("설정")
        use_gemini = st.toggle("Gemini 보조 분석 사용", value=False)
        api_key = st.secrets.get("GEMINI_API_KEY", "")
        model = st.secrets.get("GEMINI_MODEL", "gemini-3.1-flash-lite-preview")
        if api_key:
            st.success("Gemini API 키가 Streamlit secrets에서 로드되었습니다.")
        else:
            st.warning("GEMINI_API_KEY가 없으면 휴리스틱 분석으로 동작합니다.")
        st.caption(f"Model: `{model}`")
        st.caption("기본은 빠른 로컬 분석입니다. Gemini는 제한 시간 안에 성공할 때만 결과를 대체합니다.")
        st.divider()
        st.download_button(
            "표준 차원 사전 다운로드",
            data=yaml.safe_dump(standard_config, allow_unicode=True, sort_keys=False),
            file_name="standard_dimensions.yaml",
        )
        st.download_button(
            "표 정의 스키마 다운로드",
            data=yaml.safe_dump(table_config, allow_unicode=True, sort_keys=False),
            file_name="table_definition_schema.yaml",
        )

    with st.container(border=True):
        st.subheader("1. 템플릿 분석")
        analysis_template = st.file_uploader(
            "분석할 엑셀 보고서 템플릿",
            type=["xlsx"],
            key="analysis_template_file",
        )
        analysis_sheet_names = []
        if analysis_template is not None:
            analysis_sheet_names = list_workbook_sheets(analysis_template.getvalue())
            default_sheets = analysis_sheet_names[:1] if analysis_sheet_names else []
            analysis_selected_sheets = st.multiselect(
                "분석할 시트",
                analysis_sheet_names,
                default=default_sheets,
                key="analysis_selected_sheets",
            )
        else:
            analysis_selected_sheets = []
        if analysis_template is not None:
            analysis_signature = (
                analysis_template.name,
                getattr(analysis_template, "size", None),
                tuple(analysis_selected_sheets),
            )
            if st.session_state.get("analysis_signature") != analysis_signature:
                st.session_state.pop("analysis_definitions", None)
                st.session_state.pop("analysis_result_sheets", None)
        else:
            analysis_signature = None
        analyze_clicked = st.button("템플릿 분석하기", type="primary", disabled=analysis_template is None)
        if analyze_clicked and analysis_template is None:
            st.error("분석할 엑셀 템플릿을 먼저 업로드하세요.")
        elif analyze_clicked and not analysis_selected_sheets:
            st.error("분석할 시트를 하나 이상 선택하세요.")
        elif analyze_clicked:
            with st.spinner("템플릿 구조를 분석하는 중입니다..."):
                st.session_state.analysis_definitions = analyze_definitions(
                    analysis_template.getvalue(),
                    dictionary,
                    schema,
                    use_gemini,
                    api_key,
                    model,
                    analysis_selected_sheets,
                )
                st.session_state.analysis_signature = analysis_signature
                st.session_state.analysis_result_sheets = list(analysis_selected_sheets)

        analysis_definitions = st.session_state.get("analysis_definitions")
        if analysis_definitions:
            st.success(f"표 정의 {len(analysis_definitions)}개를 찾았습니다.")
            with st.expander("시트별 템플릿 분석 결과 JSON", expanded=True):
                render_sheet_definition_json(
                    analysis_definitions,
                    st.session_state.get("analysis_result_sheets"),
                )
        else:
            st.info("템플릿만 먼저 분석해 표 위치와 구성요소를 확인할 수 있습니다.")

    with st.container(border=True):
        st.subheader("2. 보고서 생성")
        report_template_file = st.file_uploader(
            "보고서 생성에 사용할 엑셀 템플릿",
            type=["xlsx"],
            key="report_template_file",
        )
        report_sheet_names = []
        if report_template_file is not None:
            report_sheet_names = list_workbook_sheets(report_template_file.getvalue())
            report_default_sheets = report_sheet_names[:1] if report_sheet_names else []
            report_selected_sheets = st.multiselect(
                "보고서 생성에 사용할 템플릿 시트",
                report_sheet_names,
                default=report_default_sheets,
                key="report_selected_sheets",
            )
        else:
            report_selected_sheets = []
        source_file = st.file_uploader(
            "원본 매체 데이터",
            type=["xlsx", "csv", "tsv"],
            key="report_source_file",
        )
        media_name = st.text_input(
            "업로드한 원본 데이터의 매체명",
            placeholder="예: 네이버, 카카오, 구글, 메타",
            help="원본 데이터에 매체 컬럼이 없으면 이 값이 매체 기준으로 사용됩니다.",
        )

        ready_for_review = report_template_file is not None and source_file is not None and bool(report_selected_sheets)
        report_signature = None
        if ready_for_review:
            report_signature = (
                report_template_file.name,
                getattr(report_template_file, "size", None),
                tuple(report_selected_sheets),
                source_file.name,
                getattr(source_file, "size", None),
            )
            if st.session_state.get("report_signature") != report_signature:
                st.session_state.pop("report_definitions", None)
                st.session_state.pop("report_result_sheets", None)
                st.session_state.pop("generated_report_error", None)

        review_clicked = st.button("표 정의 검수 시작", type="primary", disabled=not ready_for_review)
        if review_clicked and not ready_for_review:
            st.error("보고서 생성용 템플릿, 분석할 시트, 원본 매체 데이터를 모두 선택하세요.")
        elif review_clicked:
            with st.spinner("보고서용 템플릿을 분석하는 중입니다..."):
                st.session_state.report_definitions = analyze_definitions(
                    report_template_file.getvalue(),
                    dictionary,
                    schema,
                    use_gemini,
                    api_key,
                    model,
                    report_selected_sheets,
                )
                st.session_state.report_signature = report_signature
                st.session_state.report_result_sheets = list(report_selected_sheets)

        if not ready_for_review:
            st.info("템플릿, 분석할 시트, 원본 데이터를 모두 준비하면 표 정의 검수와 보고서 생성이 가능합니다.")
            show_library(table_config)
            return

        template_bytes = report_template_file.getvalue()
        source_raw = apply_uploaded_media_name(read_source_file(source_file), media_name, dictionary)
        source_mapping_schema = build_source_mapping_schema(source_raw, dictionary)
        source_df = normalize_source_dataframe(source_raw, dictionary)
        pivot_source_df = build_pivot_source_dataframe(source_raw, source_df)
        try:
            pivot_sources = detect_pivot_sources(template_bytes)
            pivot_source_error = ""
        except Exception as exc:
            pivot_sources = []
            pivot_source_error = str(exc)

        st.subheader("원본 매핑 스키마")
        mapping_df = pd.DataFrame(source_mapping_schema)
        st.dataframe(mapping_df, use_container_width=True)
        with st.expander("원본 매핑 스키마 JSON"):
            st.code(json.dumps(source_mapping_schema, ensure_ascii=False, indent=2), language="json")

        st.subheader("정규화된 원본 데이터 미리보기")
        st.dataframe(source_df.head(20), use_container_width=True)

        if pivot_source_error:
            st.warning(f"피벗 소스 감지 중 오류가 있어 피벗 갱신 UI를 건너뜁니다: {pivot_source_error}")

        pivot_update_enabled = False
        skip_direct_fill = False
        strip_pivot_cache = False
        pivot_update_mode = "replace"
        selected_pivot_source = None
        if pivot_sources:
            st.subheader("피벗 소스 데이터 갱신")
            st.dataframe(pd.DataFrame(pivot_sources), use_container_width=True)
            pivot_update_enabled = st.checkbox("피벗 소스 데이터 갱신 사용", value=True)
            skip_direct_fill = st.checkbox("표 직접 채우기 건너뛰기", value=True)
            strip_pivot_cache = st.checkbox("피벗 캐시 제거 후 엑셀에서 새로고침", value=True)
            selected_display = st.selectbox(
                "갱신할 피벗 소스",
                [source["display"] for source in pivot_sources],
            )
            selected_pivot_source = next(
                source for source in pivot_sources if source["display"] == selected_display
            )
            pivot_update_mode_label = st.radio(
                "소스 데이터 처리 방식",
                ["교체", "행 추가"],
                horizontal=True,
            )
            pivot_update_mode = "replace" if pivot_update_mode_label == "교체" else "append"
            st.caption("피벗 정의는 유지하고, 파일을 열 때 피벗이 자동 새로고침되도록 설정합니다.")

        report_definitions = st.session_state.get("report_definitions")
        pivot_only_ready = bool(pivot_sources and pivot_update_enabled and skip_direct_fill and selected_pivot_source)
        if not report_definitions and not pivot_only_ready:
            st.warning("먼저 `표 정의 검수 시작` 버튼을 눌러 템플릿 표 정의를 추출하세요.")
            return
        edited_definitions = []
        all_errors = []
        if report_definitions:
            st.subheader("템플릿 분석 결과 스키마")
            with st.expander("시트별 템플릿 분석 결과 JSON", expanded=False):
                render_sheet_definition_json(
                    report_definitions,
                    st.session_state.get("report_result_sheets"),
                )

            st.subheader("표 정의 검수")
            for idx, definition in enumerate(report_definitions):
                edited = definition_editor(definition, idx, dictionary)
                errors, warnings = validate_definition(edited, schema, dictionary)
                for error in errors:
                    st.error(f"{edited.get('name')}: {error}")
                for warning in warnings:
                    st.warning(f"{edited.get('name')}: {warning}")
                all_errors.extend(errors)
                edited_definitions.append(edited)

            st.subheader("집계 결과 미리보기")
            preview_tabs = st.tabs([definition.get("name", f"표 {idx + 1}") for idx, definition in enumerate(edited_definitions)])
            for tab, definition in zip(preview_tabs, edited_definitions):
                with tab:
                    try:
                        result = aggregate_table(source_df, definition, dictionary)
                        st.dataframe(result, use_container_width=True)
                    except Exception as exc:
                        st.error(str(exc))
                        all_errors.append(str(exc))
        else:
            st.info("피벗 소스 데이터만 갱신하는 모드입니다. 표 직접 채우기와 표 정의 검수는 건너뜁니다.")

        generation_blocked = bool(all_errors) and not pivot_only_ready
        output_bytes = None
        if st.button("완성 보고서 생성", type="primary", disabled=generation_blocked):
            st.session_state.pop("generated_report_error", None)
            try:
                with st.spinner("템플릿 서식을 보존하며 데이터를 채우는 중입니다..."):
                    output_bytes = template_bytes
                    if pivot_sources and strip_pivot_cache:
                        output_bytes = strip_pivot_cache_records(output_bytes)
                    if pivot_update_enabled and selected_pivot_source:
                        output_bytes = update_pivot_source_data(
                            output_bytes,
                            pivot_source_df,
                            selected_pivot_source,
                            pivot_update_mode,
                            dictionary,
                        )
                    if not skip_direct_fill:
                        output_bytes = fill_workbook(output_bytes, source_df, edited_definitions, dictionary)
                    if pivot_sources and strip_pivot_cache:
                        output_bytes = strip_pivot_cache_records(output_bytes)
                    elif pivot_sources:
                        output_bytes = enable_pivot_refresh_on_load(output_bytes)
            except Exception as exc:
                st.session_state.generated_report_error = str(exc)

        if st.session_state.get("generated_report_error"):
            st.error(f"보고서 생성 중 오류가 발생했습니다: {st.session_state.generated_report_error}")

        if output_bytes:
            st.success(f"보고서 생성이 완료되었습니다. 파일 크기: {len(output_bytes) / 1024 / 1024:.2f} MB")
            st.download_button(
                "완성된 엑셀 보고서 다운로드",
                data=output_bytes,
                file_name="completed_ad_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click="ignore",
            )


def show_library(table_config: dict[str, Any]) -> None:
    st.subheader("내장 표 정의 라이브러리")
    library = table_config.get("table_library", {})
    if library:
        st.code(yaml.safe_dump(library, allow_unicode=True, sort_keys=False), language="yaml")


if __name__ == "__main__":
    main()
