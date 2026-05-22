# AI Excel Ad Report Generator

사용자가 쓰던 엑셀 보고서 템플릿과 원본 광고 데이터를 업로드하면, 템플릿 안의 표 구조를 분석해 선언적 표 정의로 바꾸고, 범용 집계 엔진으로 데이터를 채워 완성 보고서를 생성하는 Streamlit 앱입니다.

## 핵심 구조

- 템플릿 분석: 엑셀의 제목/헤더/빈 행 패턴을 읽어 표 위치와 컬럼을 추출합니다.
- 표 정의: `group_by`, `sort`, `limit`, `metrics`, `location` 형태의 YAML/JSON 친화 구조로 관리합니다.
- 검수 UI: AI/휴리스틱이 추출한 표 정의를 사용자가 Streamlit 화면에서 수정할 수 있습니다.
- 범용 집계: 표 타입별 함수 없이 정의만으로 `groupby -> aggregate -> sort -> limit`을 실행합니다.
- 엑셀 작성: 원본 템플릿의 서식을 보존하고 값만 채운 `.xlsx`를 다운로드합니다.

## Streamlit Cloud 배포

1. 이 저장소를 GitHub에 push합니다.
2. Streamlit Cloud에서 저장소를 연결합니다.
3. Main file path를 `app.py`로 지정합니다.
4. OpenAI 기반 분석을 쓰려면 Streamlit secrets에 아래 값을 추가합니다.

```toml
OPENAI_API_KEY = "sk-..."
OPENAI_MODEL = "gpt-4o-mini"
```

API 키가 없어도 제목/헤더 기반 분석으로 동작합니다.

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```
