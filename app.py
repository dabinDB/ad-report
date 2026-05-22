from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import yaml

from ad_report.aggregation import aggregate_table, build_source_mapping_schema, normalize_source_dataframe
from ad_report.dictionary import StandardDictionary, load_yaml
from ad_report.template_analyzer import analyze_template, analyze_with_gemini
from ad_report.validation import validate_definition
from ad_report.workbook_writer import fill_workbook


ROOT = Path(__file__).parent
STANDARD_DIMENSIONS_PATH = ROOT / "config" / "standard_dimensions.yaml"
TABLE_SCHEMA_PATH = ROOT / "config" / "table_definition_schema.yaml"


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


def analyze_definitions(
    template_bytes: bytes,
    dictionary: StandardDictionary,
    schema: dict[str, Any],
    use_gemini: bool,
    api_key: str,
    model: str,
) -> list[dict[str, Any]]:
    if use_gemini and api_key:
        return analyze_with_gemini(template_bytes, dictionary, schema, api_key, model)
    return analyze_template(template_bytes, dictionary)


def definition_editor(definition: dict[str, Any], index: int, dictionary: StandardDictionary) -> dict[str, Any]:
    edited = json.loads(json.dumps(definition, ensure_ascii=False))
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
        c1, c2, c3, c4 = st.columns(4)
        sort["by"] = c1.selectbox("정렬 기준", sort_options or [""], index=(sort_options or [""]).index(sort_default), key=f"sort_by_{index}")
        sort["order"] = c2.selectbox("정렬 순서", ["desc", "asc"], index=0 if sort.get("order", "desc") == "desc" else 1, key=f"sort_order_{index}")
        limit_value = c3.number_input("Limit", min_value=0, value=int(edited.get("limit") or 0), key=f"limit_{index}")
        edited["limit"] = int(limit_value) if limit_value else None
        edited.setdefault("total_row", {})["enabled"] = c4.checkbox("합계 행", value=bool(edited.get("total_row", {}).get("enabled")), key=f"total_{index}")

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
        use_gemini = st.toggle("Gemini로 템플릿 분석", value=True)
        api_key = st.secrets.get("GEMINI_API_KEY", "")
        model = st.secrets.get("GEMINI_MODEL", "gemini-3.1-flash-lite-preview")
        if api_key:
            st.success("Gemini API 키가 Streamlit secrets에서 로드되었습니다.")
        else:
            st.warning("GEMINI_API_KEY가 없으면 휴리스틱 분석으로 동작합니다.")
        st.caption(f"Model: `{model}`")
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
        if st.button("템플릿 분석하기", type="primary", disabled=analysis_template is None):
            with st.spinner("템플릿 구조를 분석하는 중입니다..."):
                st.session_state.analysis_definitions = analyze_definitions(
                    analysis_template.getvalue(),
                    dictionary,
                    schema,
                    use_gemini,
                    api_key,
                    model,
                )

        analysis_definitions = st.session_state.get("analysis_definitions")
        if analysis_definitions:
            st.success(f"표 정의 {len(analysis_definitions)}개를 찾았습니다.")
            with st.expander("템플릿 분석 결과 JSON", expanded=True):
                st.code(json.dumps(analysis_definitions, ensure_ascii=False, indent=2), language="json")
        else:
            st.info("템플릿만 먼저 분석해 표 위치와 구성요소를 확인할 수 있습니다.")

    with st.container(border=True):
        st.subheader("2. 보고서 생성")
        report_template_file = st.file_uploader(
            "보고서 생성에 사용할 엑셀 템플릿",
            type=["xlsx"],
            key="report_template_file",
        )
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

        ready_for_review = report_template_file is not None and source_file is not None
        report_signature = None
        if ready_for_review:
            report_signature = (
                report_template_file.name,
                getattr(report_template_file, "size", None),
                source_file.name,
                getattr(source_file, "size", None),
            )
            if st.session_state.get("report_signature") != report_signature:
                st.session_state.pop("report_definitions", None)

        if st.button("표 정의 검수 시작", type="primary", disabled=not ready_for_review):
            with st.spinner("보고서용 템플릿을 분석하는 중입니다..."):
                st.session_state.report_definitions = analyze_definitions(
                    report_template_file.getvalue(),
                    dictionary,
                    schema,
                    use_gemini,
                    api_key,
                    model,
                )
                st.session_state.report_signature = report_signature

        if not ready_for_review:
            st.info("템플릿과 원본 데이터를 모두 업로드하면 표 정의 검수와 보고서 생성이 가능합니다.")
            show_library(table_config)
            return

        template_bytes = report_template_file.getvalue()
        source_raw = apply_uploaded_media_name(read_source_file(source_file), media_name, dictionary)
        source_mapping_schema = build_source_mapping_schema(source_raw, dictionary)
        source_df = normalize_source_dataframe(source_raw, dictionary)

        st.subheader("원본 매핑 스키마")
        mapping_df = pd.DataFrame(source_mapping_schema)
        st.dataframe(mapping_df, use_container_width=True)
        with st.expander("원본 매핑 스키마 JSON"):
            st.code(json.dumps(source_mapping_schema, ensure_ascii=False, indent=2), language="json")

        st.subheader("정규화된 원본 데이터 미리보기")
        st.dataframe(source_df.head(20), use_container_width=True)

        report_definitions = st.session_state.get("report_definitions")
        if not report_definitions:
            st.warning("먼저 `표 정의 검수 시작` 버튼을 눌러 템플릿 표 정의를 추출하세요.")
            return

        st.subheader("템플릿 분석 결과 스키마")
        with st.expander("템플릿 분석 결과 JSON", expanded=False):
            st.code(json.dumps(report_definitions, ensure_ascii=False, indent=2), language="json")

        st.subheader("표 정의 검수")
        edited_definitions = []
        all_errors = []
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

        if st.button("완성 보고서 생성", type="primary", disabled=bool(all_errors)):
            with st.spinner("템플릿 서식을 보존하며 데이터를 채우는 중입니다..."):
                output_bytes = fill_workbook(template_bytes, source_df, edited_definitions, dictionary)
            st.success("보고서 생성이 완료되었습니다.")
            st.download_button(
                "완성된 엑셀 보고서 다운로드",
                data=output_bytes,
                file_name="completed_ad_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


def show_library(table_config: dict[str, Any]) -> None:
    st.subheader("내장 표 정의 라이브러리")
    library = table_config.get("table_library", {})
    if library:
        st.code(yaml.safe_dump(library, allow_unicode=True, sort_keys=False), language="yaml")


if __name__ == "__main__":
    main()
