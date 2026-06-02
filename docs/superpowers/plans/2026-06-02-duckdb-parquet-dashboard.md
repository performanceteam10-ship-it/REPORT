# DuckDB + Parquet 리포트 로딩 개선 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 대용량 xlsx의 openpyxl 파싱 병목을 제거하기 위해, 업로드 직전 parquet 변환기를 추가하고 Streamlit 앱이 DuckDB로 parquet를 읽도록 바꾼다(xlsx 폴백 유지).

**Architecture:** 업로더는 `convert_report.bat`로 최신 xlsx를 같은 Dropbox 폴더에 parquet로 변환한다. 앱은 같은 날짜의 parquet가 있으면 그것을, 없으면 xlsx를 사용한다. parquet는 신규 `data_access.py`의 DuckDB 기반 로더로 읽고, xlsx는 기존 openpyxl 경로로 폴백한다. 기존 pandas 집계 로직은 그대로 두어 수치 동등성을 보장한다.

**Tech Stack:** Python, Streamlit, pandas, DuckDB, pyarrow(parquet), openpyxl(폴백), pytest(검증)

---

## File Structure

- Create: `convert_report.py` — 최신 xlsx 탐색 + raw 시트 → parquet 변환 (CLI)
- Create: `convert_report.bat` — `convert_report.py` 더블클릭 실행 래퍼
- Create: `data_access.py` — DuckDB 기반 parquet 로더 (`read_full`)
- Create: `tests/conftest.py` — 합성 리포트 fixture (xlsx + parquet)
- Create: `tests/test_convert_report.py` — 변환기 테스트
- Create: `tests/test_data_access.py` — DuckDB 로더 테스트
- Create: `tests/test_file_discovery.py` — parquet 우선 탐색 테스트
- Modify: `requirements.txt` — `duckdb`, `pyarrow` 추가
- Modify: `app.py` — 파일 탐색 정규식/우선순위, `load_raw` parquet 분기, 드라이브 임시파일 확장자
- Modify: `madup_api.py` — parquet 파일명 후보 추가(우선)
- Modify: `drive_report.py` — 정규식에 parquet 포함
- Modify: `README.md` — 변환 단계/동작 설명

**메모리 최적화 범위 주의:** Summary·상품별·매체별 페이지 모두 과거 전체 기간을 UI(기간 비교, 전체기간 TOP10, 누적 성과)에서 참조하므로, 본 계획은 "전체 로드를 DuckDB로 빠르게"까지 수행한다. 행 슬라이싱/집계 SQL 이관으로 RAM 추가 절감은 UX(과거 비교 범위) 변경을 동반하므로 후속 작업으로 분리한다(스펙 §8).

---

## Task 1: 의존성 추가

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: requirements.txt에 duckdb, pyarrow 추가**

`requirements.txt` 전체를 아래로 교체한다:

```
streamlit>=1.28.0
requests>=2.28.0
pandas>=2.0.0
openpyxl>=3.1.0
plotly>=5.0.0
google-api-python-client>=2.100.0
google-auth>=2.23.0
duckdb>=1.0.0
pyarrow>=15.0.0
pytest>=8.0.0
```

- [ ] **Step 2: 설치 및 임포트 확인**

Run: `pip install -r requirements.txt && python -c "import duckdb, pyarrow, pytest; print(duckdb.__version__)"`
Expected: 버전 문자열 출력, 에러 없음

- [ ] **Step 3: Commit**

```bash
git add requirements.txt
git commit -m "build: add duckdb, pyarrow, pytest dependencies"
```

---

## Task 2: 테스트 fixture (합성 리포트)

**Files:**
- Create: `tests/conftest.py`

- [ ] **Step 1: conftest.py 작성**

`tests/conftest.py` 생성:

```python
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

COL_IMP = "노출"
COL_CLICK = "클릭"
COL_COST = "비용(markup+)"
COL_REV_SS = "(스마트스토어센터기준) 전체 매출"
COL_REV_DB = "(대시보드기준) 전체 매출"
COL_BUY_SS = "(스마트스토어센터기준) 전체 구매"
COL_BUY_DB = "(대시보드기준) 전체 구매"
COL_COUP_REV = "쿠팡_매출"
COL_GM_REV = "g마켓_매출"
COL_PRODUCT = "실 전환 발생 상품"
COL_PROMO = "프로모션명"
COL_SS_KW = "스마트스토어센터_키워드"


def _sample_frame() -> pd.DataFrame:
    rows = []
    dates = pd.date_range("2026-04-01", periods=5, freq="D")
    medias = ["네이버 파워링크", "쿠팡", "G마켓"]
    for d in dates:
        for i, media in enumerate(medias):
            rows.append(
                {
                    "날짜": d,
                    "채널": "네이버브랜드스토어" if "네이버" in media else media,
                    "매체상세": media,
                    "캠페인": f"camp_{i}",
                    "그룹": f"grp_{i}",
                    "소재/키워드": f"kw_{i}",
                    COL_SS_KW: f"sskw_{i}",
                    "샤크/닌자구분": "샤크" if i % 2 == 0 else "닌자",
                    COL_PRODUCT: f"상품_{i}",
                    COL_PROMO: "상시" if i == 0 else "프로모션A",
                    COL_IMP: 1000 + i,
                    COL_CLICK: 100 + i,
                    COL_COST: 10000 + i * 1000,
                    COL_REV_SS: 50000 + i * 2000,
                    COL_REV_DB: 48000 + i * 2000,
                    COL_BUY_SS: 5 + i,
                    COL_BUY_DB: 4 + i,
                    COL_COUP_REV: 3000 if media == "쿠팡" else 0,
                    COL_GM_REV: 2000 if media == "G마켓" else 0,
                }
            )
    return pd.DataFrame(rows)


@pytest.fixture()
def sample_df() -> pd.DataFrame:
    return _sample_frame()


@pytest.fixture()
def report_dir(tmp_path: Path, sample_df: pd.DataFrame) -> Path:
    folder = tmp_path / "07. 리포트"
    folder.mkdir()
    xlsx = folder / "Madup_Sharkninja_Daily Report_260405.xlsx"
    sample_df.to_excel(xlsx, sheet_name="raw", index=False, engine="openpyxl")
    return folder
```

- [ ] **Step 2: fixture import 확인**

Run: `python -c "import pandas; pandas.DataFrame().to_excel"`
Expected: 에러 없음 (실제 검증은 다음 태스크 테스트에서)

- [ ] **Step 3: Commit**

```bash
git add tests/conftest.py
git commit -m "test: add synthetic SharkNinja report fixture"
```

---

## Task 3: 변환기 convert_report.py

**Files:**
- Create: `convert_report.py`
- Test: `tests/test_convert_report.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_convert_report.py` 생성:

```python
from __future__ import annotations

from pathlib import Path

import pandas as pd

from convert_report import find_latest_xlsx, xlsx_to_parquet


def test_find_latest_xlsx_picks_highest_tag(report_dir: Path):
    older = report_dir / "Madup_Sharkninja_Daily Report_260401.xlsx"
    pd.DataFrame({"날짜": [pd.Timestamp("2026-04-01")]}).to_excel(
        older, sheet_name="raw", index=False, engine="openpyxl"
    )
    latest = find_latest_xlsx(report_dir)
    assert latest is not None
    assert latest.name == "Madup_Sharkninja_Daily Report_260405.xlsx"


def test_xlsx_to_parquet_roundtrip(report_dir: Path):
    xlsx = report_dir / "Madup_Sharkninja_Daily Report_260405.xlsx"
    out = xlsx_to_parquet(xlsx)
    assert out.exists()
    assert out.name == "Madup_Sharkninja_Daily Report_260405.parquet"
    src = pd.read_excel(xlsx, sheet_name="raw", engine="openpyxl")
    got = pd.read_parquet(out)
    assert len(got) == len(src)
    assert float(got["비용(markup+)"].sum()) == float(src["비용(markup+)"].sum())
    assert pd.api.types.is_datetime64_any_dtype(got["날짜"])
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `pytest tests/test_convert_report.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'convert_report'`

- [ ] **Step 3: convert_report.py 구현**

`convert_report.py` 생성:

```python
"""최신 Madup_Sharkninja_Daily Report_YYMMDD.xlsx 의 raw 시트를 같은 폴더에 .parquet 로 변환."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

DEFAULT_REPORT_DIR = Path(
    r"C:\Users\MADUP\주식회사매드업 Dropbox\광고사업부\4. 광고주\샤크닌자\07. 리포트"
)
XLSX_RE = re.compile(r"^Madup_Sharkninja_Daily[- ]Report_(\d{6})\.xlsx$", re.IGNORECASE)


def find_latest_xlsx(folder: Path) -> Path | None:
    best: tuple[str, Path] | None = None
    if not folder.is_dir():
        return None
    for p in folder.iterdir():
        if not p.is_file():
            continue
        m = XLSX_RE.match(p.name)
        if m and (best is None or m.group(1) > best[0]):
            best = (m.group(1), p)
    return best[1] if best else None


def xlsx_to_parquet(xlsx_path: Path, out_path: Path | None = None) -> Path:
    out = out_path or xlsx_path.with_suffix(".parquet")
    df = pd.read_excel(xlsx_path, sheet_name="raw", engine="openpyxl")
    if "날짜" in df.columns:
        df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    df.to_parquet(out, engine="pyarrow", index=False)
    return out


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    folder = Path(argv[0]) if argv else DEFAULT_REPORT_DIR
    print(f"리포트 폴더: {folder}")
    latest = find_latest_xlsx(folder)
    if latest is None:
        print("변환할 xlsx를 찾지 못했습니다.")
        return 1
    out = latest.with_suffix(".parquet")
    if out.exists() and out.stat().st_mtime >= latest.stat().st_mtime:
        print(f"이미 최신 parquet 존재: {out.name} (스킵)")
        return 0
    result = xlsx_to_parquet(latest, out)
    rows = len(pd.read_parquet(result))
    size_mb = result.stat().st_size / (1024 * 1024)
    print(f"변환 완료: {result.name}  ({rows:,}행, {size_mb:.1f}MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `pytest tests/test_convert_report.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add convert_report.py tests/test_convert_report.py
git commit -m "feat: add xlsx->parquet converter for daily report"
```

---

## Task 4: 더블클릭 래퍼 convert_report.bat

**Files:**
- Create: `convert_report.bat`

- [ ] **Step 1: convert_report.bat 작성**

`convert_report.bat` 생성:

```bat
@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === SharkNinja 리포트 변환 (xlsx -> parquet) ===
python convert_report.py
echo.
echo 변환이 끝나면 Dropbox 가 자동 동기화합니다. 이 창은 닫아도 됩니다.
pause
```

- [ ] **Step 2: 수동 실행 확인 (선택)**

Run: `python convert_report.py "tests"`  (또는 실제 리포트 폴더 경로)
Expected: "변환할 xlsx를 찾지 못했습니다." 또는 정상 변환 메시지

- [ ] **Step 3: Commit**

```bash
git add convert_report.bat
git commit -m "feat: add double-click bat wrapper for report conversion"
```

---

## Task 5: DuckDB 로더 data_access.py

**Files:**
- Create: `data_access.py`
- Test: `tests/test_data_access.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_data_access.py` 생성:

```python
from __future__ import annotations

from pathlib import Path

import pandas as pd

from data_access import read_full


def test_read_full_matches_parquet(tmp_path: Path, sample_df: pd.DataFrame):
    pq = tmp_path / "r.parquet"
    sample_df.to_parquet(pq, engine="pyarrow", index=False)
    got = read_full(str(pq))
    assert len(got) == len(sample_df)
    assert pd.api.types.is_datetime64_any_dtype(got["날짜"])
    assert float(got["(스마트스토어센터기준) 전체 매출"].sum()) == float(
        sample_df["(스마트스토어센터기준) 전체 매출"].sum()
    )


def test_read_full_preserves_korean_columns(tmp_path: Path, sample_df: pd.DataFrame):
    pq = tmp_path / "r.parquet"
    sample_df.to_parquet(pq, engine="pyarrow", index=False)
    got = read_full(str(pq))
    for col in ["비용(markup+)", "실 전환 발생 상품", "샤크/닌자구분"]:
        assert col in got.columns
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `pytest tests/test_data_access.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'data_access'`

- [ ] **Step 3: data_access.py 구현**

`data_access.py` 생성:

```python
"""DuckDB 기반 parquet 로더. 앱의 parquet 읽기 경로에서 사용."""
from __future__ import annotations

import duckdb
import pandas as pd

DATE_COL = "날짜"


def read_full(path: str) -> pd.DataFrame:
    """parquet 전체를 DataFrame 으로 읽어 날짜 컬럼을 datetime 으로 정규화."""
    con = duckdb.connect(database=":memory:")
    try:
        df = con.execute(
            "SELECT * FROM read_parquet(?)", [str(path)]
        ).fetch_df()
    finally:
        con.close()
    if DATE_COL in df.columns:
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    return df
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `pytest tests/test_data_access.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add data_access.py tests/test_data_access.py
git commit -m "feat: add DuckDB-based parquet loader"
```

---

## Task 6: parquet 우선 파일 탐색

**Files:**
- Modify: `app.py:29-31` (REPORT_FILE_RE), `app.py:56-68` (find_latest_report)
- Modify: `drive_report.py:15-17` (REPORT_FILE_RE)
- Modify: `madup_api.py:14-19` (madup_report_filenames)
- Test: `tests/test_file_discovery.py`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_file_discovery.py` 생성:

```python
from __future__ import annotations

from pathlib import Path

import pandas as pd

import app as app_mod
from madup_api import madup_report_filenames


def test_report_re_matches_parquet():
    assert app_mod.REPORT_FILE_RE.match("Madup_Sharkninja_Daily Report_260405.parquet")
    assert app_mod.REPORT_FILE_RE.match("Madup_Sharkninja_Daily Report_260405.xlsx")


def test_find_latest_prefers_parquet(tmp_path: Path):
    folder = tmp_path
    df = pd.DataFrame({"날짜": [pd.Timestamp("2026-04-05")]})
    (folder / "Madup_Sharkninja_Daily Report_260405.xlsx").write_bytes(b"x")
    df.to_parquet(folder / "Madup_Sharkninja_Daily Report_260405.parquet", index=False)
    latest = app_mod.find_latest_report(folder)
    assert latest is not None
    assert latest.suffix == ".parquet"


def test_madup_filenames_parquet_first():
    names = madup_report_filenames("260405")
    assert names[0].endswith(".parquet")
    assert any(n.endswith(".xlsx") for n in names)
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `pytest tests/test_file_discovery.py -v`
Expected: FAIL (parquet 미매칭 / .xlsx 선호 / parquet 후보 없음)

- [ ] **Step 3: app.py REPORT_FILE_RE 수정**

`app.py` 의 `REPORT_FILE_RE` 정의(라인 29-31)를 아래로 교체:

```python
REPORT_FILE_RE = re.compile(
    r"^Madup_Sharkninja_Daily[- ]Report_(\d{6})\.(?:xlsx|parquet)$", re.IGNORECASE
)
```

- [ ] **Step 4: app.py find_latest_report 수정 (parquet 우선)**

`app.py` 의 `find_latest_report` 함수(라인 56-68)를 아래로 교체:

```python
def find_latest_report(folder: Path) -> Path | None:
    best: tuple[str, int, Path] | None = None  # (date_tag, ext_rank, path)
    if not folder.is_dir():
        return None
    for p in folder.iterdir():
        if not p.is_file():
            continue
        m = REPORT_FILE_RE.match(p.name)
        if not m:
            continue
        tag = m.group(1)
        ext_rank = 1 if p.suffix.lower() == ".parquet" else 0
        cand = (tag, ext_rank, p)
        if best is None or (tag, ext_rank) > (best[0], best[1]):
            best = cand
    return best[2] if best else None
```

- [ ] **Step 5: drive_report.py REPORT_FILE_RE 수정**

`drive_report.py` 의 `REPORT_FILE_RE`(라인 15-17)를 아래로 교체:

```python
REPORT_FILE_RE = re.compile(
    r"^Madup_Sharkninja_Daily[- ]Report_(\d{6})\.(?:xlsx|parquet)$", re.IGNORECASE
)
```

- [ ] **Step 6: madup_api.py madup_report_filenames 수정 (parquet 우선)**

`madup_api.py` 의 `madup_report_filenames`(라인 14-19)를 아래로 교체:

```python
def madup_report_filenames(tag: str) -> tuple[str, ...]:
    """Dropbox 파일명 변형. parquet 를 우선 시도하고 xlsx 로 폴백."""
    return (
        f"Madup_Sharkninja_Daily Report_{tag}.parquet",
        f"Madup_Sharkninja_Daily-Report_{tag}.parquet",
        f"Madup_Sharkninja_Daily Report_{tag}.xlsx",
        f"Madup_Sharkninja_Daily-Report_{tag}.xlsx",
    )
```

- [ ] **Step 7: 테스트 통과 확인**

Run: `pytest tests/test_file_discovery.py -v`
Expected: PASS (3 passed)

- [ ] **Step 8: Commit**

```bash
git add app.py drive_report.py madup_api.py tests/test_file_discovery.py
git commit -m "feat: prefer parquet over xlsx in report file discovery"
```

---

## Task 7: load_raw parquet 분기 + 드라이브 임시파일 확장자

**Files:**
- Modify: `app.py:71-75` (load_raw)
- Modify: `app.py:826-828` (drive temp 파일명)
- Modify: `app.py:18` 부근 (data_access import)

- [ ] **Step 1: data_access import 추가**

`app.py` 의 `from drive_report import ...`(라인 18) 위/아래에 추가:

```python
from data_access import read_full as _read_parquet_full
```

- [ ] **Step 2: load_raw 를 확장자 분기로 교체**

`app.py` 의 `load_raw` 함수(라인 71-75)를 아래로 교체:

```python
@st.cache_data(ttl=120, show_spinner=False)
def load_raw(report_path: str) -> pd.DataFrame:
    p = Path(report_path)
    if p.suffix.lower() == ".parquet":
        df = _read_parquet_full(str(p))
    else:
        df = pd.read_excel(p, sheet_name="raw", engine="openpyxl")
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    return df
```

- [ ] **Step 3: 드라이브 임시파일 확장자 보존**

`app.py` 의 드라이브 분기에서 임시파일을 만드는 부분(라인 826-828):

```python
            data = _drive_cached_bytes(file_id, sa_json)
            path = Path(tempfile.gettempdir()) / f"sn_drive_{file_id}.xlsx"
            path.write_bytes(data)
```

을 아래로 교체 (parquet 파일명이면 .parquet 임시파일로 저장):

```python
            data = _drive_cached_bytes(file_id, sa_json)
            ext = ".parquet" if str(name).lower().endswith(".parquet") else ".xlsx"
            path = Path(tempfile.gettempdir()) / f"sn_drive_{file_id}{ext}"
            path.write_bytes(data)
```

- [ ] **Step 4: Madup 임시파일 확장자 보존 (최신 자동)**

`app.py` Madup "최신 자동" 분기(라인 762-764):

```python
                    path = Path(tempfile.gettempdir()) / "sn_madup_latest.xlsx"
                    path.write_bytes(data)
                    st.caption(f"열림: `{dp}`")
```

을 아래로 교체:

```python
                    ext = ".parquet" if str(dp).lower().endswith(".parquet") else ".xlsx"
                    path = Path(tempfile.gettempdir()) / f"sn_madup_latest{ext}"
                    path.write_bytes(data)
                    st.caption(f"열림: `{dp}`")
```

- [ ] **Step 5: Madup 임시파일 확장자 보존 (날짜 선택)**

`app.py` Madup "날짜 선택" 분기(라인 788-790):

```python
                    path = Path(tempfile.gettempdir()) / f"sn_madup_{pick}.xlsx"
                    path.write_bytes(data)
                    st.caption(f"열림: `{dp}`")
```

을 아래로 교체:

```python
                    ext = ".parquet" if str(dp).lower().endswith(".parquet") else ".xlsx"
                    path = Path(tempfile.gettempdir()) / f"sn_madup_{pick}{ext}"
                    path.write_bytes(data)
                    st.caption(f"열림: `{dp}`")
```

- [ ] **Step 6: Madup 단일 경로 임시파일 확장자 보존**

`app.py` Madup 단일 파일 분기(라인 797-799):

```python
                path = Path(tempfile.gettempdir()) / "sn_madup_single.xlsx"
                path.write_bytes(data)
                st.caption(f"열림: `{single_path}`")
```

을 아래로 교체:

```python
                ext = ".parquet" if single_path.lower().endswith(".parquet") else ".xlsx"
                path = Path(tempfile.gettempdir()) / f"sn_madup_single{ext}"
                path.write_bytes(data)
                st.caption(f"열림: `{single_path}`")
```

- [ ] **Step 7: 앱 import/구문 정상성 확인**

Run: `python -c "import ast; ast.parse(open('app.py', encoding='utf-8').read()); print('ok')"`
Expected: `ok`

- [ ] **Step 8: 전체 테스트 재실행**

Run: `pytest -v`
Expected: 모든 테스트 PASS

- [ ] **Step 9: Commit**

```bash
git add app.py
git commit -m "feat: load parquet via DuckDB with xlsx fallback across all sources"
```

---

## Task 8: 변환 전후 수치 동등성 검증 스크립트

**Files:**
- Create: `tests/test_parity.py`

- [ ] **Step 1: 동등성 테스트 작성**

`tests/test_parity.py` 생성 (xlsx 경로와 parquet 경로의 핵심 집계가 동일한지 검증):

```python
from __future__ import annotations

from pathlib import Path

import pandas as pd

from convert_report import xlsx_to_parquet
from data_access import read_full


def _load_xlsx(p: Path) -> pd.DataFrame:
    df = pd.read_excel(p, sheet_name="raw", engine="openpyxl")
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    return df


def test_aggregates_match_between_xlsx_and_parquet(report_dir: Path):
    xlsx = report_dir / "Madup_Sharkninja_Daily Report_260405.xlsx"
    pq = xlsx_to_parquet(xlsx)

    a = _load_xlsx(xlsx)
    b = read_full(str(pq))

    assert len(a) == len(b)
    for col in [
        "비용(markup+)",
        "(스마트스토어센터기준) 전체 매출",
        "(대시보드기준) 전체 매출",
        "(스마트스토어센터기준) 전체 구매",
    ]:
        assert abs(float(a[col].sum()) - float(b[col].sum())) < 1e-6

    ga = a.groupby("매체상세")["(스마트스토어센터기준) 전체 매출"].sum().sort_index()
    gb = b.groupby("매체상세")["(스마트스토어센터기준) 전체 매출"].sum().sort_index()
    pd.testing.assert_series_equal(ga, gb, check_names=False)
```

- [ ] **Step 2: 테스트 실행**

Run: `pytest tests/test_parity.py -v`
Expected: PASS (1 passed)

- [ ] **Step 3: Commit**

```bash
git add tests/test_parity.py
git commit -m "test: verify xlsx vs parquet aggregate parity"
```

---

## Task 9: README 업데이트

**Files:**
- Modify: `README.md`

- [ ] **Step 1: README에 변환 단계 안내 추가**

`README.md` 상단의 `# SharkNinja Daily Report 대시보드` 제목 바로 아래에 다음 섹션을 삽입:

```markdown
## ⚡ 빠른 로딩: 엑셀 → Parquet 변환 (권장)

대용량 xlsx 는 앱이 열릴 때 파싱이 느립니다. 업로드 전에 **parquet 로 한 번 변환**하면
앱은 DuckDB 로 훨씬 빠르게 읽습니다.

1. 평소처럼 `Madup_Sharkninja_Daily Report_YYMMDD.xlsx` 를 리포트 폴더에 저장
2. **`convert_report.bat` 더블클릭** → 같은 폴더에 `...YYMMDD.parquet` 생성
3. Dropbox 가 자동 동기화 → 앱이 자동으로 parquet 사용

- parquet 가 있으면 우선 사용하고, 없으면 기존 xlsx 로 폴백합니다(무중단).
- 변환기는 폴더의 **최신 xlsx 1개**만 변환합니다. 다른 폴더를 쓰려면
  `python convert_report.py "폴더경로"` 로 실행하세요.
```

- [ ] **Step 2: 표시 확인**

Run: `python -c "print(open('README.md', encoding='utf-8').read()[:200])"`
Expected: 제목 + 새 섹션 일부 출력

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: document parquet conversion workflow"
```

---

## Task 10: 최종 통합 점검

**Files:** (없음 — 검증만)

- [ ] **Step 1: 전체 테스트**

Run: `pytest -v`
Expected: 전체 PASS

- [ ] **Step 2: 실제 폴더로 변환 스모크 (선택, 실데이터 있을 때)**

Run: `python convert_report.py`
Expected: "변환 완료: ...parquet (N행, X.XMB)" 출력, 같은 폴더에 parquet 생성, xlsx 대비 용량 대폭 감소

- [ ] **Step 3: 앱 구동 스모크**

Run: `streamlit run app.py`
확인: 사이드바에서 parquet 가 인식되어 "열림: ...parquet" 표시되고 Summary/상품별/매체별이 정상 렌더링되는지. (parquet 없으면 xlsx 로 정상 폴백되는지)

- [ ] **Step 4: 최종 상태 확인**

Run: `git status --short && git log --oneline -10`
Expected: 작업 트리 깨끗, 커밋 이력에 본 계획 태스크들 반영

---

## Self-Review 결과

- **스펙 커버리지:** 변환기(§4.1)=Task3·4, 파일탐색(§4.2)=Task6, DuckDB 로딩(§4.3 핵심)=Task5·7, 의존성(§4.4)=Task1, 문서(§4.5)=Task9, 검증(§6)=Task8·10. 폴백(§5)=Task6·7. 메모리 SQL 이관(§4.3 일부)은 UX 제약으로 후속 분리(§8) — 본문 상단 주의에 명시.
- **플레이스홀더:** 없음(모든 코드 단계에 실제 코드/명령/기대결과 포함).
- **타입 일관성:** `find_latest_xlsx`/`xlsx_to_parquet`(convert_report), `read_full`(data_access), `REPORT_FILE_RE`/`find_latest_report`(app), `madup_report_filenames`(madup_api) 시그니처가 태스크 간 일치.
