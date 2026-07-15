"""
Dropbox 리포트 폴더 감시 → xlsx 저장 시 자동 parquet 변환.

실행: python watch_and_convert.py
종료: Ctrl+C
"""
from __future__ import annotations

import time
from pathlib import Path

from watchdog.events import FileSystemEventHandler, FileModifiedEvent, FileCreatedEvent
from watchdog.observers import Observer

from convert_report import XLSX_RE, xlsx_to_parquet

WATCH_DIR = Path(
    r"C:\Users\MADUP\주식회사매드업 Dropbox\광고사업부\4. 광고주\샤크닌자\07. 리포트"
)


class ReportHandler(FileSystemEventHandler):
    def _handle(self, path: str) -> None:
        p = Path(path)
        if not XLSX_RE.match(p.name):
            return
        # Excel이 저장 중일 때 임시파일(~$...)이 생기므로 잠시 대기
        time.sleep(2)
        if not p.exists():
            return
        out = p.with_suffix(".parquet")
        print(f"[변환 시작] {p.name}")
        try:
            xlsx_to_parquet(p, out)
            size_mb = out.stat().st_size / (1024 * 1024)
            print(f"[완료] {out.name}  ({size_mb:.1f}MB)")
        except Exception as e:
            print(f"[오류] {e}")

    def on_created(self, event: FileCreatedEvent) -> None:
        if not event.is_directory:
            self._handle(event.src_path)

    def on_modified(self, event: FileModifiedEvent) -> None:
        if not event.is_directory:
            self._handle(event.src_path)


def main() -> None:
    if not WATCH_DIR.is_dir():
        print(f"폴더 없음: {WATCH_DIR}")
        return

    print(f"감시 시작: {WATCH_DIR}")
    print("xlsx 저장 시 자동으로 parquet 변환합니다. 종료: Ctrl+C\n")

    observer = Observer()
    observer.schedule(ReportHandler(), str(WATCH_DIR), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    print("\n감시 종료")


if __name__ == "__main__":
    main()
