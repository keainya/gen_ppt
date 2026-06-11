@echo off
cd /d "%~dp0"

if not exist ".venv\" (
    echo [*] Creating Python virtual environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -q python-pptx

echo [*] Environment ready
