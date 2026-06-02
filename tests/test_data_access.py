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
