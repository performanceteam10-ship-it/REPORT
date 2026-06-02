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
