@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

echo.
echo  ================================================
echo    Planing  ^|  Publish to GitHub
echo  ================================================
echo.
echo  Commits the changes in this folder and pushes them to GitHub.
echo  Files ignored by .gitignore are never included.
echo  Options:  --dry-run  (preview only)   --yes  (no questions)
echo            --no-push  (commit only)    -m "subject"
echo.

:: -- Check Python (standard library only, no venv needed) ----------------
set "PY=python"
python --version >nul 2>&1
if errorlevel 1 (
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo  [ERROR] Python is not installed.
        echo          Install Python 3.10 or newer from: https://python.org
        if not defined PLANING_NO_PAUSE pause
        exit /b 1
    )
    set "PY=py -3"
)
%PY% -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo  [ERROR] Python 3.10 or newer is required.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)

:: -- Run the publisher -----------------------------------------------------
%PY% publish.py %*
set "RESULT=%ERRORLEVEL%"

echo.
if "%RESULT%"=="0" (
    echo  ================================================
    echo    Finished.
    echo  ================================================
) else (
    echo  ================================================
    echo    Not finished - read the messages above.
    echo  ================================================
)
echo.
if not defined PLANING_NO_PAUSE pause
exit /b %RESULT%
