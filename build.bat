@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo  ================================================
echo    Planing  ^|  Build Script
echo  ================================================
echo.

:: -- Check Python --------------------------------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python is not installed.
    echo          Install Python 3.10 or newer from: https://python.org
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)
python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo  [ERROR] Python 3.10 or newer is required.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo  [OK] %%v detected
echo.

:: -- Create virtual environment -------------------------------------------
echo  [1/5]  Creating virtual environment (venv)...
if exist venv (
    echo         Removing old venv...
    rmdir /s /q venv
)
python -m venv venv
if errorlevel 1 (
    echo  [ERROR] Failed to create venv.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)
echo         Done.
echo.

:: -- Activate venv ---------------------------------------------------------
echo  [2/5]  Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo  [ERROR] Could not activate venv.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)
echo         Done.
echo.

:: -- Install dependencies ---------------------------------------------------
echo  [3/5]  Installing dependencies from requirements.txt...
python -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo  [ERROR] Could not upgrade pip.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [ERROR] pip install requirements.txt failed.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)
echo         Done.
echo.

:: -- Clean old build / dist ---------------------------------------------------
echo  [4/5]  Cleaning old build and dist folders...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist
echo         Done.
echo.

:: -- Build with PyInstaller ---------------------------------------------------
echo  [5/5]  Building executable with PyInstaller...
echo.
python -m PyInstaller main.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo  [ERROR] PyInstaller build FAILED.
    echo          Check the output above for details.
    if not defined PLANING_NO_PAUSE pause
    exit /b 1
)

echo.
echo  ================================================
echo    Build Complete!
echo    Output folder: dist\Planing\
echo  ================================================
echo.
echo  Next step - create the installer:
echo    1. Install Inno Setup 6 from: https://jrsoftware.org/isdl.php
echo    2. Open installer.iss with Inno Setup Compiler
echo    3. Press Ctrl+F9 to compile
echo    4. The setup file will appear in: installer_output\
echo.
if not defined PLANING_NO_PAUSE pause
