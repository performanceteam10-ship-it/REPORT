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
