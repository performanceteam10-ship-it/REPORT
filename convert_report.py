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


def _to_str_keep_na(v: object) -> object:
    """결측은 그대로 두고 나머지는 문자열로. (timedelta/time/bytes 등 혼합 타입 object 컬럼 대비)"""
    try:
        if pd.isna(v):
            return v
    except (TypeError, ValueError):
        pass
    return str(v)


def make_parquet_safe(df: pd.DataFrame) -> pd.DataFrame:
    """parquet 직렬화가 깨지는 혼합 타입 object 컬럼을 문자열로 정규화(결측 보존).

    앱은 object 컬럼을 모두 문자열 라벨로만 사용하고 숫자 지표는 별도 숫자형 컬럼이므로 안전.
    `날짜` 는 호출 전에 datetime 으로 변환되어 object 가 아니므로 영향받지 않는다.
    """
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(_to_str_keep_na)
    return df


def xlsx_to_parquet(xlsx_path: Path, out_path: Path | None = None) -> Path:
    out = out_path or xlsx_path.with_suffix(".parquet")
    df = pd.read_excel(xlsx_path, sheet_name="raw", engine="openpyxl")
    if "날짜" in df.columns:
        df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    df = make_parquet_safe(df)
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
