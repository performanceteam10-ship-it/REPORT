@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo  SharkNinja Daily 대시보드
echo ========================================
echo.

REM 더블클릭 시 PATH에 streamlit 이 없을 수 있어서 python -m 으로 실행
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -m streamlit run app.py
    goto :after
)
where python >nul 2>&1
if %ERRORLEVEL%==0 (
    python -m streamlit run app.py
    goto :after
)

echo [오류] Python 을 찾을 수 없습니다.
echo Python 3 을 설치한 뒤, 이 폴더에서 다음을 한 번 실행하세요:
echo   pip install -r requirements.txt
goto :end

:after
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [오류] 실행 실패. 아래를 확인하세요:
    echo   1^) pip install -r requirements.txt
    echo   2^) 브라우저에서 직접 열기: http://localhost:8501
)

:end
echo.
echo 같은 네트워크 팀원에게 알려줄 주소 예: http://[이 PC IP]:8501
echo.
pause
