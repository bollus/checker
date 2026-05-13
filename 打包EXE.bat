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

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set BUILD_DATE=%%i
for /f %%i in ('git rev-parse --short HEAD 2^>nul') do set GIT_SHA=%%i
if "%GIT_SHA%"=="" set GIT_SHA=local
set VERSIONED_EXE=表格核对工具-%BUILD_DATE%-%GIT_SHA%.exe
copy /Y "dist\表格核对工具.exe" "dist\%VERSIONED_EXE%" >nul
if errorlevel 1 (
    echo Failed to create versioned EXE.
    pause
    exit /b 1
)

echo.
echo Build completed.
echo EXE path: dist\表格核对工具.exe
echo Versioned EXE path: dist\%VERSIONED_EXE%
pause
