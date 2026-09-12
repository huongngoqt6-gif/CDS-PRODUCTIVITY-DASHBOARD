@echo off
chcp 65001 >nul
title Kiểm tra môi trường
echo PYTHON:
python --version
echo.
echo PIP:
python -m pip --version
echo.
echo STREAMLIT:
python -m streamlit --version
echo.
echo THƯ MỤC HIỆN TẠI:
cd
echo.
pause
