@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo [*] .venv incomplete or missing, recreating...
    if exist ".venv\" rmdir /s /q ".venv"
    python -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -q python-pptx

echo [*] Environment ready
