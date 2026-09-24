@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo Python Launcher was not found. Install Python 3.12 for Windows from python.org,
    echo and select the option to install the Python Launcher.
    pause
    exit /b 1
)

py -3.12 --version >nul 2>nul
if errorlevel 1 (
    echo Python 3.12 was not found. Install Python 3.12 for Windows, then run this file again.
    pause
    exit /b 1
)

if not exist ".venv-winbuild\Scripts\python.exe" (
    py -3.12 -m venv .venv-winbuild
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
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import pypdf ^
    --hidden-import PyPDF2 ^
    md_converter_app.py
if errorlevel 1 goto failed

echo.
echo Build complete: dist\MarkdownConverter.exe
echo Copy that EXE to the Windows PC where it will be used.
pause
exit /b 0

:failed
echo.
echo Build failed. Read the error above, check your internet connection, and try again.
pause
exit /b 1
