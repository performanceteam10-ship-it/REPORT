"""parquet 로더. 앱의 parquet 읽기 경로에서 사용."""
from __future__ import annotations

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

DATE_COL = "날짜"


def read_full(path: str) -> pd.DataFrame:
    """parquet 전체를 DataFrame 으로 읽어 날짜 컬럼을 datetime 으로 정규화.

    문자열 라벨 컬럼(채널·매체·캠페인·상품·브랜드구분 등)은 Arrow dictionary 로
    읽어 pandas category 로 받는다. 리포트 기간 확장으로 행수가 100만+로 늘면서
    문자열 컬럼이 메모리 대부분을 차지, Streamlit Cloud 무료 티어에서 OOM 크래시가 났다.

    ★ 핵심: category 변환을 '읽은 뒤 astype' 으로 하면 전체 문자열 df 를 먼저
    만들어 로드 피크가 ~1.4GB 까지 치솟아 여전히 OOM 이었다. Arrow 단계에서
    dictionary 로 캐스팅한 뒤 to_pandas 하면 문자열이 pandas 에 통째로
    올라오지 않아 로드 피크가 ~70MB 로 떨어진다(전체 행/값은 그대로 보존).
    """
    table = pq.read_table(str(path))
    fields = []
    columns = []
    for name in table.column_names:
        col = table.column(name)
        t = col.type
        if name != DATE_COL and (pa.types.is_string(t) or pa.types.is_large_string(t)):
            col = col.cast(pa.dictionary(pa.int32(), pa.string()))
        fields.append(name)
        columns.append(col)
    table = pa.table(columns, names=fields)
    df = table.to_pandas()  # dictionary -> category, 나머지는 원형 dtype 유지
    if DATE_COL in df.columns:
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    return df
