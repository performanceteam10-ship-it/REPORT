"""DuckDB 기반 parquet 로더. 앱의 parquet 읽기 경로에서 사용."""
from __future__ import annotations

import duckdb
import pandas as pd

DATE_COL = "날짜"


def read_full(path: str) -> pd.DataFrame:
    """parquet 전체를 DataFrame 으로 읽어 날짜 컬럼을 datetime 으로 정규화."""
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
    return df
