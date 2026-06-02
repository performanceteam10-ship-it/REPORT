"""
구글 드라이브 폴더에서 Madup_Sharkninja_Daily Report_YYMMDD.xlsx 목록/다운로드.
Streamlit Secrets: GOOGLE_DRIVE_FOLDER_ID, GOOGLE_SERVICE_ACCOUNT_JSON (전체 JSON 문자열)
"""
from __future__ import annotations

import io
import json
import re

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

REPORT_FILE_RE = re.compile(
    r"^Madup_Sharkninja_Daily[- ]Report_(\d{6})\.(?:xlsx|parquet)$", re.IGNORECASE
)
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def _parse_sa(sa: str | dict) -> dict:
    if isinstance(sa, dict):
        return sa
    return json.loads(sa)


def _service_from_json(sa_json: str | dict):
    info = _parse_sa(sa_json)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_drive_reports(folder_id: str, sa_json: str | dict) -> list[tuple[str, str, str]]:
    """(date_tag, file_id, filename) 목록, date_tag 내림차순 정렬."""
    svc = _service_from_json(sa_json)
    q = f"'{folder_id}' in parents and trashed=false and mimeType!='application/vnd.google-apps.folder'"
    out: list[tuple[str, str, str]] = []
    page_token = None
    while True:
        resp = (
            svc.files()
            .list(
                q=q,
                spaces="drive",
                fields="nextPageToken, files(id, name)",
                pageToken=page_token,
                pageSize=100,
            )
            .execute()
        )
        for f in resp.get("files", []):
            name = f.get("name", "")
            m = REPORT_FILE_RE.match(name)
            if m:
                out.append((m.group(1), f["id"], name))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    out.sort(key=lambda x: x[0], reverse=True)
    return out


def download_drive_file(file_id: str, sa_json: str | dict) -> bytes:
    svc = _service_from_json(sa_json)
    req = svc.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue()
