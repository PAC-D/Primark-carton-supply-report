@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install it from https://www.python.org/downloads/ and tick "Add to PATH".
    pause
    exit /b 1
)
python -m pip install -r requirements.txt --quiet
python -m streamlit run app.py
