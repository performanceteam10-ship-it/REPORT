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
