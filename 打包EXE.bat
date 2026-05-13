@echo off
setlocal
cd /d %~dp0

where py >nul 2>nul
if errorlevel 1 (
    echo Python launcher "py" was not found. Please install Python 3 first.
    pause
    exit /b 1
)

echo Installing or updating build dependencies...
py -m pip install pyinstaller pillow
if errorlevel 1 (
    echo Failed to install build dependencies.
    pause
    exit /b 1
)

echo Building EXE...
py -m PyInstaller --clean excel_check_tool.spec
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build completed.
echo EXE path: dist\表格核对工具.exe
pause
