@echo off
setlocal
cd /d "%~dp0"

echo ===================================================
echo CSB Equipment Inventory Commit and Push
echo ===================================================
echo.

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo [ERROR] This folder is not a Git repository.
    pause
    exit /b 1
)

git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No 'origin' remote is configured.
    pause
    exit /b 1
)

echo 1. Staging updated dashboard files and images...
git add index.html images scripts update_inventory.bat
if errorlevel 1 (
    echo [ERROR] Failed to stage the inventory updates.
    pause
    exit /b 1
)

git diff --cached --quiet
if not errorlevel 1 (
    echo.
    echo No inventory changes to commit.
    pause
    exit /b 0
)

set "COMMIT_MESSAGE=Update equipment inventory"
if not "%~1"=="" set "COMMIT_MESSAGE=%~1"

echo.
echo 2. Creating commit: %COMMIT_MESSAGE%
git commit -m "%COMMIT_MESSAGE%"
if errorlevel 1 (
    echo [ERROR] Commit failed.
    pause
    exit /b 1
)

for /f "delims=" %%B in ('git branch --show-current') do set "BRANCH=%%B"
if "%BRANCH%"=="" (
    echo [ERROR] Could not determine the current branch.
    pause
    exit /b 1
)

echo.
echo 3. Pushing %BRANCH% to origin...
git push origin "%BRANCH%"
if errorlevel 1 (
    echo [ERROR] Push failed. Your commit is local and can be pushed later.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Inventory changes have been committed and pushed.
pause
