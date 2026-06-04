@echo off
title Voucher Filtration — Streamlit App
color 0A

echo.
echo ============================================================
echo   VOUCHER FILTRATION — Starting Web App
echo ============================================================
echo.

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERROR: Python not found. Install from python.org
    pause & exit /b
)

echo Installing / checking dependencies...
pip install -r requirements.txt -q
echo Done.
echo.

echo Starting Streamlit app...
echo.
echo ============================================================
echo   App will open in your browser automatically.
echo   URL: http://localhost:8501
echo   Press Ctrl+C in this window to stop the app.
echo ============================================================
echo.

streamlit run app.py
pause
