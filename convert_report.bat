@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === SharkNinja 리포트 변환 (xlsx -^> parquet) ===
python convert_report.py
echo.
echo 변환이 끝나면 Dropbox 가 자동 동기화합니다. 이 창은 닫아도 됩니다.
pause
