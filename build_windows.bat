@echo off
setlocal
cd /d "%~dp0"

rem Usage: build_windows.bat [/nopause]   (/nopause = no "Press any key" prompts, for scripts/CI)
if /i "%~1"=="/nopause" set NOPAUSE=1

where py >nul 2>nul
if errorlevel 1 (
    echo Python Launcher was not found. Install 64-bit Python 3.10 or newer for Windows from python.org,
    echo and select the option to install the Python Launcher.
    if not defined NOPAUSE pause
    exit /b 1
)

py -3 -c "import sys, struct; sys.exit(0 if sys.version_info >= (3, 10) and struct.calcsize('P') == 8 else 1)" >nul 2>nul
if errorlevel 1 (
    echo 64-bit Python 3.10 or newer was not found. Install it for Windows, then run this file again.
    if not defined NOPAUSE pause
    exit /b 1
)

if not exist ".venv-winbuild\Scripts\python.exe" (
    py -3 -m venv .venv-winbuild
    if errorlevel 1 goto failed
)

call ".venv-winbuild\Scripts\activate.bat"
python -m pip install --upgrade pip
if errorlevel 1 goto failed
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto failed

python -m PyInstaller --clean --noconfirm --onefile --windowed ^
    --name MarkdownConverter ^
    --hidden-import docx ^
    --hidden-import markdownify ^
    --hidden-import bs4 ^
    --hidden-import openpyxl ^
    --hidden-import pypdf ^
    --hidden-import cryptography ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    md_converter_app.py
if errorlevel 1 goto failed

echo.
echo Build complete: dist\MarkdownConverter.exe
echo Copy that EXE to the Windows PC where it will be used.
if not defined NOPAUSE pause
exit /b 0

:failed
echo.
echo Build failed. Read the error above, check your internet connection, and try again.
if not defined NOPAUSE pause
exit /b 1
