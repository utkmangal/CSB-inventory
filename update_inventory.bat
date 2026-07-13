@echo off
cd /d "%~dp0"
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
echo 2. Extracting equipment images from 담당자.xlsx...
python scripts/extract_equipment_images.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to extract equipment images. Please check the workbook and Pillow installation.
    pause
    exit /b %errorlevel%
)
echo.
echo 3. Extracting animal-research-lab images...
python scripts/extract_animal_lab_images.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to extract animal-research-lab images. Please check the workbook and Pillow installation.
    pause
    exit /b %errorlevel%
)
echo.
echo 4. Running repository integrity check...
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
