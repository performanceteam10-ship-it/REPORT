"""
Madup Dropbox HTTP API — 파일 다운로드.
POST {base}/api/dropbox/download  JSON {"path": "/..."}  Header X-API-Key
"""
from __future__ import annotations

from datetime import datetime, timedelta

import requests

DEFAULT_BASE = "https://api-auth.madup-dct.site"


def madup_report_filenames(tag: str) -> tuple[str, ...]:
    """Dropbox 파일명 변형. parquet 를 우선 시도하고 xlsx 로 폴백."""
    return (
        f"Madup_Sharkninja_Daily Report_{tag}.parquet",
        f"Madup_Sharkninja_Daily-Report_{tag}.parquet",
        f"Madup_Sharkninja_Daily Report_{tag}.xlsx",
        f"Madup_Sharkninja_Daily-Report_{tag}.xlsx",
    )


def download_madup_file(
    api_key: str,
    path: str,
    base_url: str = DEFAULT_BASE,
    timeout: int = 120,
) -> bytes:
    url = f"{base_url.rstrip('/')}/api/dropbox/download"
    r = requests.post(
        url,
        headers={"X-API-Key": api_key, "Content-Type": "application/json"},
        json={"path": path},
        timeout=timeout,
    )
    if r.status_code != 200:
        raise requests.HTTPError(f"HTTP {r.status_code}: {r.text[:500]}")
    return r.content


def find_latest_madup_path(
    api_key: str,
    folder: str,
    base_url: str = DEFAULT_BASE,
    max_days: int = 120,
) -> tuple[str, bytes]:
    """폴더 안에서 날짜 역순으로 파일명 시도 후 최초 성공분 반환 (path, bytes)."""
    folder = folder.rstrip("/")
    last_err: Exception | None = None
    for i in range(max_days):
        d = datetime.now() - timedelta(days=i)
        tag = d.strftime("%y%m%d")
        for fname in madup_report_filenames(tag):
            path = f"{folder}/{fname}"
            try:
                b = download_madup_file(api_key, path, base_url)
                if len(b) > 200:
                    return path, b
            except Exception as e:
                last_err = e
                continue
    msg = str(last_err) if last_err else "파일 없음"
    raise FileNotFoundError(f"최근 {max_days}일 안에 유효한 리포트 xlsx를 찾지 못했습니다. ({msg})")
