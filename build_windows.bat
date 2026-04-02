@echo off
REM Windows build script for Excel Compare Tool
REM This script creates a Windows executable using PyInstaller

echo ============================================================
echo   Excel Compare Tool - Windows Build Script
echo ============================================================
echo.

REM Check if uv is installed
where uv >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: uv is not installed!
    echo Please install uv first: pip install uv
    echo Or download from: https://github.com/astral-sh/uv
    pause
    exit /b 1
)

echo Step 1: Installing dependencies...
uv sync
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo Step 2: Building Windows executable...
echo.

REM Build using PyInstaller
uv run pyinstaller ^
    --onefile ^
    --windowed ^
    --name ExcelCompare ^
    --add-data "src\excel_compare;excel_compare" ^
    --hidden-import "tkinterdnd2" ^
    --clean ^
    --noconfirm ^
    src\excel_compare\main.py

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Build completed successfully!
echo ============================================================
echo.
if exist "dist\ExcelCompare.exe" (
    echo Executable: dist\ExcelCompare.exe
    echo.
    for %%A in ("dist\ExcelCompare.exe") do echo File size: %%~zA bytes
    echo.
    echo You can now run the executable directly!
    echo No Python installation required on target machine.
) else (
    echo WARNING: ExcelCompare.exe not found in dist folder
    echo.
    echo Please check the build output above for errors.
)

echo.
pause
