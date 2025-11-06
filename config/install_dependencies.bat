@echo off
REM Installation script for DaVinci Resolve VFX Pipeline Scripts
REM This script installs required Python dependencies using your system Python

echo ================================================
echo DaVinci Resolve VFX Editor Scripts - Dependency Installer
echo ================================================
echo.

REM Check if python is available
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Error: Python not found in PATH
    echo.
    echo Please install Python 3.6 or later and try again.
    echo Visit: https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

set "PYTHON_CMD=python"

echo Found Python at:
where python
echo.

REM Check Python version
echo Python version:
%PYTHON_CMD% --version
echo.

REM Ensure pip is installed
echo Ensuring pip is installed...
%PYTHON_CMD% -m ensurepip --default-pip 2>nul

REM Upgrade pip
echo Upgrading pip...
%PYTHON_CMD% -m pip install --upgrade pip

echo.
echo Installing dependencies from requirements.txt...
echo.

REM Install requirements
%PYTHON_CMD% -m pip install -r requirements.txt

if %errorlevel% equ 0 (
    echo.
    echo ================================================
    echo √ Installation complete!
    echo ================================================
    echo.
    echo You can now run the scripts. For example:
    echo   python clip-inventory.py
    echo   python shot-list.py --output my_shot_list.xlsx
    echo.
) else (
    echo.
    echo ================================================
    echo X Installation failed
    echo ================================================
    echo.
    echo Please try installing manually:
    echo   python -m pip install -r requirements.txt
    echo.
)

pause