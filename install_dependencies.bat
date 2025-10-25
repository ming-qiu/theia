@echo off
REM Installation script for DaVinci Resolve VFX Pipeline Scripts
REM This script installs required Python dependencies using Resolve's Python

echo ================================================
echo DaVinci Resolve VFX Scripts - Dependency Installer
echo ================================================
echo.

REM Common Resolve Python locations
set "RESOLVE_PYTHON=C:\Program Files\Blackmagic Design\DaVinci Resolve\python.exe"
set "RESOLVE_PYTHON_ALT=C:\Program Files\Blackmagic Design\DaVinci Resolve Studio\python.exe"

REM Check if Resolve Python exists
if exist "%RESOLVE_PYTHON%" (
    goto :found_python
)

if exist "%RESOLVE_PYTHON_ALT%" (
    set "RESOLVE_PYTHON=%RESOLVE_PYTHON_ALT%"
    goto :found_python
)

echo Error: Could not find DaVinci Resolve Python at:
echo   %RESOLVE_PYTHON%
echo   %RESOLVE_PYTHON_ALT%
echo.
echo Please locate your Resolve Python installation and run:
echo   "C:\Path\To\Resolve\python.exe" -m pip install -r requirements.txt
pause
exit /b 1

:found_python
echo Found DaVinci Resolve Python at:
echo   %RESOLVE_PYTHON%
echo.

REM Check Python version
echo Python version:
"%RESOLVE_PYTHON%" --version
echo.

REM Ensure pip is installed
echo Ensuring pip is installed...
"%RESOLVE_PYTHON%" -m ensurepip --default-pip 2>nul

REM Upgrade pip
echo Upgrading pip...
"%RESOLVE_PYTHON%" -m pip install --upgrade pip

echo.
echo Installing dependencies from requirements.txt...
echo.

REM Install requirements
"%RESOLVE_PYTHON%" -m pip install -r requirements.txt

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
    echo   "%RESOLVE_PYTHON%" -m pip install -r requirements.txt
    echo.
)

pause
