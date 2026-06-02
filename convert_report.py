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


def xlsx_to_parquet(xlsx_path: Path, out_path: Path | None = None) -> Path:
    out = out_path or xlsx_path.with_suffix(".parquet")
    df = pd.read_excel(xlsx_path, sheet_name="raw", engine="openpyxl")
    if "날짜" in df.columns:
        df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
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
