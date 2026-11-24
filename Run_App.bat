@echo off
chcp 65001 >nul
title Translation Tool
cd /d "%~dp0"

rem Dung venv neu co, neu khong thi dung Python he thong
if exist ".venv\Scripts\activate.bat" call ".venv\Scripts\activate.bat"
if exist "venv\Scripts\activate.bat" call "venv\Scripts\activate.bat"

python -c "import streamlit" 2>nul
if errorlevel 1 (
    echo Chua cai thu vien. Dang cai dat tu requirements.txt...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Cai dat that bai. Kiem tra lai Python/pip.
        pause
        exit /b 1
    )
)

echo Dang khoi dong Translation Tool... (dong cua so nay de tat ung dung)
python -m streamlit run app.py

pause
