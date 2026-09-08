"""parquet 로더. 앱의 parquet 읽기 경로에서 사용."""
from __future__ import annotations

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

DATE_COL = "날짜"
_BATCH_ROWS = 100_000


def _is_str(t: pa.DataType) -> bool:
    return pa.types.is_string(t) or pa.types.is_large_string(t)


def read_full(path: str) -> pd.DataFrame:
    """parquet 전체를 DataFrame 으로 읽어 날짜 컬럼을 datetime 으로 정규화.

    문자열 라벨 컬럼(채널·매체·캠페인·상품·브랜드구분 등)은 Arrow dictionary 로
    읽어 pandas category 로 받는다(전체 행/값 보존, resident df ~91MB).

    ★ OOM 방지 핵심 (2026-09-08): 리포트 기간이 8/1~ → 7/1~ 로 늘며 행수가
    100만+로 2배가 됐고, Streamlit Cloud 무료 티어(~1GB)에서 로드 중 크래시했다.
    - `SELECT *` (duckdb.fetch_df) 방식: 로드 RSS ~1.4GB → OOM.
    - 통째 `pq.read_table` 후 to_pandas: category 로 줄여도 로드 RSS ~980MB → 여전히 OOM
      (Arrow 가 parquet 전체를 압축해제하며 큰 C++ 버퍼를 잡음).
    - 아래처럼 row-group 을 배치로 스트리밍하며 배치별로 dictionary 캐스팅하고
      `to_pandas(self_destruct=True, split_blocks=True)` + 메모리풀 해제까지 하면
      로드 RSS ~386MB 로 떨어져 무료 티어에서 안전하다(실측).
    """
    pf = pq.ParquetFile(str(path))
    names = pf.schema_arrow.names
    batches: list[pa.RecordBatch] = []
    for batch in pf.iter_batches(batch_size=_BATCH_ROWS):
        arrays = []
        for i, name in enumerate(names):
            col = batch.column(i)
            if name != DATE_COL and _is_str(col.type):
                col = col.cast(pa.dictionary(pa.int32(), pa.string()))
            arrays.append(col)
        batches.append(pa.record_batch(arrays, names=names))
    table = pa.Table.from_batches(batches, schema=batches[0].schema) if batches else pf.read()
    del batches
    df = table.to_pandas(self_destruct=True, split_blocks=True)
    del table
    try:
        pa.default_memory_pool().release_unused()
    except Exception:
        pass
    if DATE_COL in df.columns:
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    return df
