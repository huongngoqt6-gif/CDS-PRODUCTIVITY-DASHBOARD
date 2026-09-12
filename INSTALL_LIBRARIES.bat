@echo off
chcp 65001 >nul
title Cài thư viện Dashboard Streamlit
echo ============================================================
echo CÀI THƯ VIỆN STREAMLIT - CS HAD
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Chưa tìm thấy Python.
    echo Hãy cài Python trước và chọn "Add Python to PATH".
    pause
    exit /b 1
)

echo [1/4] Kiểm tra Python...
python --version
if errorlevel 1 goto :error

echo [2/4] Nâng cấp pip...
python -m pip install --upgrade pip
if errorlevel 1 goto :error

echo [3/4] Cài thư viện từ requirements.txt...
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [4/4] Kiểm tra Streamlit...
python -m streamlit --version
if errorlevel 1 goto :error

echo.
echo CÀI ĐẶT THÀNH CÔNG.
echo Chạy dashboard bằng file RUN_DASHBOARD.bat
pause
exit /b 0

:error
echo.
echo CÀI ĐẶT KHÔNG THÀNH CÔNG.
echo Hãy chụp màn hình lỗi để kiểm tra.
pause
exit /b 1
