from __future__ import annotations

from datetime import timedelta
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


def test_xlsx_to_parquet_handles_mixed_type_object_column(tmp_path: Path):
    """timedelta/문자열이 섞인 object 컬럼도 직렬화 실패 없이 변환되어야 한다."""
    xlsx = tmp_path / "Madup_Sharkninja_Daily Report_260406.xlsx"
    df = pd.DataFrame(
        {
            "날짜": pd.to_datetime(["2026-04-06", "2026-04-06", "2026-04-06"]),
            "노출/시간": [timedelta(minutes=30), "문자열", None],
            "비용(markup+)": [100, 200, 300],
        }
    )
    df.to_excel(xlsx, sheet_name="raw", index=False, engine="openpyxl")

    out = xlsx_to_parquet(xlsx)
    got = pd.read_parquet(out)
    assert len(got) == 3
    assert float(got["비용(markup+)"].sum()) == 600.0
    assert got["노출/시간"].isna().sum() == 1
