@echo off
title Ck's Threat Hunter App Launcher
echo ========================================================
echo          Ck's Threat Hunter Application Launcher
echo ========================================================
echo.
echo [1/2] Changing directory to threat-hunter...
cd /d "%~dp0"

echo [2/2] Launching Streamlit Server and opening browser...
echo.
:: Start a parallel process to open browser after 3 seconds once server is ready
start /b cmd /c "timeout /t 3 >nul && start http://localhost:8501"

:: Run Streamlit server in headless mode (so we manage the browser open cleanly ourselves)
..\.venv\Scripts\python.exe -m streamlit run app.py --server.headless true

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Failed to start Streamlit. Please check if python virtual environment is setup.
    pause
)
