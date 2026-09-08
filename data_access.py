"""DuckDB 기반 parquet 로더. 앱의 parquet 읽기 경로에서 사용."""
from __future__ import annotations

import duckdb
import pandas as pd

DATE_COL = "날짜"


def read_full(path: str) -> pd.DataFrame:
    """parquet 전체를 DataFrame 으로 읽어 날짜 컬럼을 datetime 으로 정규화.

    문자열 라벨 컬럼(채널·매체·캠페인·상품·브랜드구분 등)은 category 로 다운캐스트한다.
    리포트 기간 확장으로 행수가 100만+로 늘면서 문자열 컬럼이 메모리 대부분을 차지해
    (약 485MB) Streamlit Cloud 무료 티어에서 OOM 크래시가 났다. category 변환으로
    전체 행/값을 그대로 보존하면서 메모리를 ~87MB(-82%)로 낮춰 OOM 을 방지한다.
    """
    con = duckdb.connect(database=":memory:")
    try:
        try:
            con.execute("SET enable_progress_bar=false")
        except Exception:
            pass
        df = con.execute("SELECT * FROM read_parquet(?)", [str(path)]).fetch_df()
    finally:
        con.close()
    if DATE_COL in df.columns:
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    for col in df.columns:
        if col == DATE_COL:
            continue
        if df[col].dtype == object or str(df[col].dtype) in ("string", "str"):
            df[col] = df[col].astype("category")
    return df
