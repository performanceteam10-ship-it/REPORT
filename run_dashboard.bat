@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo SharkNinja Daily 대시보드 실행 중...
echo 브라우저가 열리면 팀원에게는 아래 주소를 알려주세요 (같은 네트워크):
echo   http://[이 PC의 IP]:8501
echo.
streamlit run app.py
pause
