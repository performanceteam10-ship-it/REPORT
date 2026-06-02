# DuckDB + Parquet 기반 리포트 로딩 개선 설계

작성일: 2026-06-02
대상: SharkNinja 데일리 리포트 Streamlit 대시보드 (`app.py`)

## 1. 문제 정의

매일 생성되는 `Madup_Sharkninja_Daily Report_YYMMDD.xlsx`는 **과거 전체 누적 데이터를 매번 통째로 다시 떠서** 올리는 구조다. 파일이 날마다 커지고(현재 50MB+ / 수백만 행), Streamlit 앱이 열릴 때마다 이 큰 xlsx를 `pd.read_excel(engine="openpyxl")`로 통째로 파싱한다.

`openpyxl`은 대용량 xlsx 파싱이 매우 느려, **앱 로딩(콜드 스타트·새로고침) 시간의 대부분이 이 다운로드 + 파싱 단계**에서 소비된다. Streamlit Community Cloud의 메모리(약 1GB대)에서 수백만 행을 pandas로 통째 적재하는 것도 부담이다.

## 2. 목표 / 비목표

**목표**
- 앱 로딩 시 `openpyxl` xlsx 파싱 제거
- 다운로드 용량 대폭 축소 (xlsx 50MB+ → parquet 5~15MB 예상)
- Streamlit 메모리 사용 안정화 (필요한 행/컬럼만 적재)
- DuckDB를 앱의 데이터 접근/집계 엔진으로 도입
- 사용자(업로더)의 작업 루틴 변화 최소화: "더블클릭 1단계" 추가만

**비목표**
- "매일 전체 재덤프" 생성 파이프라인 자체의 증분화 (C안) — 이번 범위 제외
- 모든 pandas 집계 함수의 전면 SQL 재작성 — 무거운 전체기간 집계만 선별 이관
- 완전자동 watcher — 이번 범위 제외 (반자동 더블클릭만)

## 3. 전체 흐름

**업로더(PC)**
```
1. 평소처럼 ...YYMMDD.xlsx 를 07.리포트 폴더(Dropbox 동기화)에 저장
2. convert_report.bat 더블클릭
   └ 폴더의 최신 xlsx → 같은 폴더에 ...YYMMDD.parquet 생성
   └ Dropbox 자동 동기화 (= 자동 업로드)
```

**앱(자동)**
```
3. 사용자가 앱 열면 → parquet 우선 다운로드 → DuckDB로 읽기/집계 (openpyxl 미사용)
```

리포트 폴더 경로: `C:\Users\MADUP\주식회사매드업 Dropbox\광고사업부\4. 광고주\샤크닌자\07. 리포트`
(= `app.py`의 `DEFAULT_REPORT_DIR`, Dropbox 동기화 폴더)

## 4. 구성 요소

### 4.1 변환기 (신규)
- `convert_report.py`
  - 인자로 폴더(기본값: `DEFAULT_REPORT_DIR`)를 받아 최신 `...YYMMDD.xlsx` 탐색
  - `raw` 시트를 읽어 `날짜`를 datetime으로 정규화한 뒤 같은 이름의 `.parquet`로 저장
  - 동일 날짜 parquet가 이미 있고 원본 xlsx가 더 최신이 아니면 스킵(재변환 방지)
  - 변환 결과(파일명, 행 수, 크기)를 콘솔에 출력
- `convert_report.bat`
  - `convert_report.py`를 더블클릭으로 실행하는 래퍼 (실행 후 결과 표시, 창 유지)

### 4.2 파일 탐색 (수정)
- `app.py` `REPORT_FILE_RE`, `drive_report.py` `REPORT_FILE_RE`: 확장자 `xlsx|parquet` 모두 허용
- `madup_api.py` `madup_report_filenames`: `.parquet` 파일명 변형도 후보에 추가
- 우선순위: **같은 날짜의 parquet가 있으면 parquet, 없으면 xlsx 폴백**

### 4.3 데이터 로딩 (수정: DuckDB 도입)
- `load_raw(path)`:
  - 확장자가 `.parquet` → DuckDB로 처리 (아래 데이터 접근 계층 사용)
  - 확장자가 `.xlsx` → 기존 `openpyxl` 경로 유지 (폴백, 하위호환)
- DuckDB 접근 계층(신규, 예: `data_access.py` 또는 `app.py` 내 함수군):
  - 받은 parquet 바이트는 기존처럼 임시파일로 기록 후 `read_parquet(temp)`로 연결
  - **전체기간 무거운 집계는 DuckDB SQL로 처리해 작은 결과만 반환**
    - 상품별: "전체 기간 네이버 브랜드스토어 상품 매출 TOP10", 매체상세별 누적 TOP10
    - 매체별: "채널-매체상세 누적 성과", 캠페인/그룹 누적 성과
  - **최근 구간 분석은 날짜 범위로 필터링한 슬라이스만 pandas로 로드**해 기존 함수 재사용
    - 이번 달 누계, 최근 7일 차트, 전일 대비(`agg_range`, `two_day_compare`, `two_day_compare_ss`) 등
  - `@st.cache_data`로 슬라이스/집계 결과 캐싱 유지 (ttl 기존값 준수)

### 4.4 의존성 (수정)
- `requirements.txt`에 추가: `duckdb`, `pyarrow`

### 4.5 문서 (수정)
- `README.md`: "엑셀 저장 후 convert_report.bat 더블클릭" 변환 단계, parquet 동작 방식, 폴백 동작 설명 추가

## 5. 동작/하위호환 규칙
- parquet 부재 시 기존 xlsx 경로로 무중단 폴백 → 점진 전환 가능
- Madup/구글드라이브/로컬 세 소스 모두 parquet 우선·xlsx 폴백 동일 적용
- 컬럼명/지표 정의(ROAS 등)는 기존과 동일하게 유지

## 6. 검증 계획
- 변환 정확성: 동일 날짜의 xlsx vs parquet 로드 결과의 행 수·주요 합계(비용/SS매출/DB매출/구매) 일치
- 지표 일치: 변환 전(xlsx)·후(parquet+DuckDB)로 아래 값이 동일한지 비교
  - 월 누계 TOTAL(노출/클릭/비용/매출/ROAS, SS·DB 모두)
  - KPI 달성률(네이버/쿠팡/G마켓)
  - 상품 전체기간 TOP10, 매체 누적 성과 상위 행
  - 전일 대비 데일리 코멘트 핵심 수치
- 성능: 동일 파일에 대해 앱 첫 로딩 시간 측정(개선 전/후), parquet 파일 크기 확인
- 폴백: parquet 없는 날짜에서 xlsx로 정상 동작 확인

## 7. 리스크
- DuckDB SQL 이관 구간에서 NULL/타입/한글 컬럼명 처리 차이로 수치 불일치 가능 → 검증 단계에서 합계 대조로 차단
- Streamlit Cloud 임시파일·메모리 환경 차이 → 임시파일 경로(`tempfile.gettempdir()`) 기존 방식 재사용으로 위험 최소화
- 사용자가 변환(bat) 실행을 누락 → parquet 없으면 xlsx 폴백으로 동작은 유지(느릴 뿐) + README 안내

## 8. 범위 밖 (후속 후보)
- C안: 누적 DuckDB DB 증분 관리로 생성/업로드 파이프라인 자체 경량화
- 완전자동 watcher
- 모든 pandas 집계의 전면 SQL 이관
