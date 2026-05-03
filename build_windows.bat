@echo off
REM SheetPic Windows 打包脚本
REM 用法:
REM   build_windows.bat          - 打包 x86_64 版本
REM   build_windows.bat arm64    - 打包 ARM64 版本
REM
REM 要求: Python 3.10+, PyInstaller, Pillow, openpyxl, pandas, requests

echo ========================================
echo SheetPic Windows 打包
echo ========================================
echo.

set APP_NAME=SheetPic
set MAIN_SCRIPT=sheetpic.py
set ICON=icon.ico
set ARCH=x86_64

if "%1"=="arm64" set ARCH=ARM64

echo 平台: Windows (%ARCH%)
echo Python:
python --version
echo.

REM 检查依赖
echo 检查依赖...
python -c "import PyInstaller" 2>nul || (echo ERROR: 缺少 PyInstaller，请运行: pip install pyinstaller && exit /b 1)
python -c "import PIL" 2>nul || (echo ERROR: 缺少 Pillow，请运行: pip install Pillow && exit /b 1)
python -c "import openpyxl" 2>nul || (echo ERROR: 缺少 openpyxl，请运行: pip install openpyxl && exit /b 1)
python -c "import pandas" 2>nul || (echo ERROR: 缺少 pandas，请运行: pip install pandas && exit /b 1)
python -c "import requests" 2>nul || (echo ERROR: 缺少 requests，请运行: pip install requests && exit /b 1)
echo 依赖检查通过
echo.

REM 清理旧构建
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM PyInstaller 打包
echo ========================================
echo 步骤 1: PyInstaller 打包
echo ========================================

set ICON_ARG=
if exist %ICON% set ICON_ARG=--icon=%ICON%

python -m PyInstaller --windowed --onefile --noconfirm --clean --name=%APP_NAME% %ICON_ARG% %MAIN_SCRIPT%

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller 打包失败
    exit /b 1
)

REM 输出结果
echo.
echo ========================================
echo 打包完成
echo ========================================

if exist dist\%APP_NAME%.exe (
    echo   输出: dist\%APP_NAME%.exe
    echo   架构: %ARCH%
    echo   签名: 无（SmartScreen 会弹警告，点仍要运行即可）
) else (
    echo   ERROR: 未找到 dist\%APP_NAME%.exe
    exit /b 1
)

echo.
pause
