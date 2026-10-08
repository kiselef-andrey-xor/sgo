@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python не найден в PATH.
    echo Установите Python 3.9+ с сайта python.org и отметьте пункт "Add python.exe to PATH".
    pause
    exit /b 1
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Первый запуск: устанавливаю библиотеку openpyxl...
    python -m pip install --user openpyxl
)

start "" python gui.py
exit /b 0
