#!/bin/bash
set -e

cd "$(dirname "$0")"

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller
python3 -m PyInstaller --clean md_converter_app.spec

echo "Build complete. Output is in the dist/ folder."
