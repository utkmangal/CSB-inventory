@echo off
echo ===================================================
echo CSB Equipment Inventory Dashboard Data Updater
echo ===================================================
echo.
echo 1. Reading Excel file (담당자.xlsx)...
python scripts/sync_data.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to update data. Please check if python and openpyxl are installed.
    pause
    exit /b %errorlevel%
)
echo.
echo 2. Running repository integrity check...
python scripts/pre_push_check.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Integrity check failed. HTML file may be invalid.
    pause
    exit /b %errorlevel%
)
echo.
echo [SUCCESS] Dashboard has been successfully updated!
echo You can open index.html in your browser to verify the changes.
echo.
pause
