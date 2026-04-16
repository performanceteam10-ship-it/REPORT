@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo  SharkNinja Daily 대시보드
echo ========================================
echo.

where py >nul 2>&1
if %ERRORLEVEL%==0 goto :run_py
where python >nul 2>&1
if %ERRORLEVEL%==0 goto :run_python

echo [오류] Python 을 찾을 수 없습니다.
echo Python 3 설치 후: pip install -r requirements.txt
goto :end

:run_py
py -m streamlit run app.py --server.port 8501
if %ERRORLEVEL% NEQ 0 goto :retry8502
goto :done

:run_python
python -m streamlit run app.py --server.port 8501
if %ERRORLEVEL% NEQ 0 goto :retry8502
goto :done

:retry8502
echo.
echo [안내] 포트 8501 이 사용 중입니다. 8502 로 다시 실행합니다.
echo 브라우저: http://localhost:8502
echo.
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -m streamlit run app.py --server.port 8502
    goto :done
)
python -m streamlit run app.py --server.port 8502
goto :done

:done
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [오류] 그래도 안 되면: pip install -r requirements.txt
    echo 이미 예전 대시보드가 켜져 있을 수 있습니다. 브라우저에서
    echo   http://localhost:8501
    echo 을 직접 열어보세요.
)

:end
echo.
echo 팀원(같은 Wi-Fi): http://[이 PC IP]:8501 또는 :8502
echo.
pause
