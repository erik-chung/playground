@echo off
echo ================================================================
echo          QQ Group Member Export Tool
echo ================================================================
echo.

echo [1/2] Checking Python...
python --version
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python first.
    pause
    exit /b 1
)

echo [2/2] Checking dependencies...
python -c "import flask; import flask_cors; import openpyxl; import requests"
if errorlevel 1 (
    echo Installing dependencies...
    pip install flask flask-cors openpyxl requests -i https://pypi.tuna.tsinghua.edu.cn/simple
)

echo.
echo ================================================================
echo Starting...
echo Make sure NapCatQQ is running and QQ is logged in.
echo Open browser: http://127.0.0.1:8080
echo ================================================================
echo.

python web_exporter.py

pause
