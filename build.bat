@echo off
setlocal enabledelayedexpansion

echo.
echo  ================================================
echo    Planing  ^|  Build Script
echo  ================================================
echo.

:: -- Check Python --------------------------------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python is not installed.
    echo          Install Python 3.9 or newer from: https://python.org
    pause
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
    pause
    exit /b 1
)
echo         Done.
echo.

:: -- Activate venv ---------------------------------------------------------
echo  [2/5]  Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo  [ERROR] Could not activate venv.
    pause
    exit /b 1
)
echo         Done.
echo.

:: -- Install dependencies ---------------------------------------------------
echo  [3/5]  Installing dependencies from requirements.txt...
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [ERROR] pip install requirements.txt failed.
    pause
    exit /b 1
)
echo         Installing PyInstaller...
pip install pyinstaller --quiet
if errorlevel 1 (
    echo  [ERROR] pip install pyinstaller failed.
    pause
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
pyinstaller main.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo  [ERROR] PyInstaller build FAILED.
    echo          Check the output above for details.
    pause
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
pause
