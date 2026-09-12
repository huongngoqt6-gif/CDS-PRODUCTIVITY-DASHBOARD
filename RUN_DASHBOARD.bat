@echo off
chcp 65001 >nul
title Dashboard Streamlit
cd /d "%~dp0"
python -m streamlit run app.py
if errorlevel 1 (
    echo.
    echo Không chạy được dashboard. Hãy chạy INSTALL_LIBRARIES.bat trước.
    pause
)
