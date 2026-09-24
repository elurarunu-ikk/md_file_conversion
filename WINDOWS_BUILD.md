# Build the Windows app

The `.exe` must be built on a Windows computer. This folder includes a one-click build script; the resulting app does not require Python on the computer where it will be used.

## Build it once on Windows

1. Copy the project folder to the Windows computer. Include `md_converter_app.py`, `requirements.txt`, and `build_windows.bat`.
2. Install **Python 3.12 for Windows** from the [official Windows downloads page](https://www.python.org/downloads/windows/). In the installer, enable **Install launcher for all users** (or **Add Python to PATH**).
3. Double-click `build_windows.bat`. The first run downloads the required packages and can take several minutes. Keep the computer connected to the internet.
4. When it says the build is complete, find `dist\MarkdownConverter.exe` in the project folder.
5. Copy that `.exe` to the Windows computer where it will be used and double-click it to open the converter.

The EXE is built for the Windows computer's CPU architecture. Create it on a Windows PC that matches the target computer (most current PCs are 64-bit Intel/AMD).

## Notes

- Do not use the macOS `md_converter_app.spec` to build the Windows version; it defines a macOS app bundle.
- Windows SmartScreen may show a warning because the app is not digitally signed. The user can choose **More info** and **Run anyway** if they trust the copy they received.
- The converter reads DOCX, XLSX, HTML, and PDF files. Scanned PDFs that contain only page images do not have extractable text.
