@echo off
title QA Screenshot Tool — Build
echo ============================================
echo  QA Screenshot Tool  ^|  Build portable .exe
echo ============================================
echo.

REM Check Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found.
    echo Download from https://python.org and make sure to tick "Add to PATH".
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: pip install failed.
    pause
    exit /b 1
)

echo.
echo [2/3] Building portable executable...
pyinstaller ^
    --onefile ^
    --noconsole ^
    --name "QA_Screenshot_Tool" ^
    --clean ^
    qa_tool.py

if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo [3/3] Cleaning up build artefacts...
if exist build   rmdir /s /q build
if exist *.spec  del /q *.spec

echo.
echo ============================================
echo  BUILD SUCCESSFUL!
echo  Executable: dist\QA_Screenshot_Tool.exe
echo.
echo  The .exe is fully self-contained.
echo  Copy it to any Windows PC or USB drive
echo  and run it — no installation needed.
echo ============================================
pause
