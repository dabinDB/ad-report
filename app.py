from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import yaml

from ad_report.aggregation import aggregate_table, normalize_source_dataframe
from ad_report.dictionary import StandardDictionary, load_yaml
from ad_report.template_analyzer import analyze_template, analyze_with_openai
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
        use_openai = st.toggle("OpenAI로 템플릿 분석", value=False)
        api_key = st.text_input("OpenAI API Key", type="password", value=st.secrets.get("OPENAI_API_KEY", ""))
        model = st.text_input("Model", value=st.secrets.get("OPENAI_MODEL", "gpt-4o-mini"))
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

    template_file = st.file_uploader("1. 평소 쓰던 엑셀 보고서 템플릿 업로드", type=["xlsx"])
    source_file = st.file_uploader("2. 원본 매체 데이터 업로드", type=["xlsx", "csv", "tsv"])

    if not template_file or not source_file:
        st.info("템플릿과 원본 데이터를 업로드하면 표 정의 추출 화면이 열립니다.")
        show_library(table_config)
        return

    template_bytes = template_file.getvalue()
    source_raw = read_source_file(source_file)
    source_df = normalize_source_dataframe(source_raw, dictionary)

    st.subheader("원본 데이터 미리보기")
    st.dataframe(source_df.head(20), use_container_width=True)

    if "definitions" not in st.session_state:
        with st.spinner("템플릿 구조를 분석하는 중입니다..."):
            if use_openai and api_key:
                st.session_state.definitions = analyze_with_openai(template_bytes, dictionary, schema, api_key, model)
            else:
                st.session_state.definitions = analyze_template(template_bytes, dictionary)

    st.subheader("AI 추출 표 정의 검수")
    if not st.session_state.definitions:
        st.warning("자동으로 찾은 표가 없습니다. 템플릿에 표 제목과 헤더 행이 있는지 확인하거나, 아래 예시를 복사해 직접 정의를 추가하세요.")
        show_library(table_config)
        return

    edited_definitions = []
    all_errors = []
    for idx, definition in enumerate(st.session_state.definitions):
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
