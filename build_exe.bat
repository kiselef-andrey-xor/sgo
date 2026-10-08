@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python не найден в PATH. Установите Python 3.9+ и отметьте "Add python.exe to PATH".
    pause
    exit /b 1
)

echo Устанавливаю зависимости (openpyxl, pyinstaller)...
python -m pip install --upgrade openpyxl pyinstaller
if errorlevel 1 (
    echo Не удалось установить зависимости. Проверьте доступ в интернет.
    pause
    exit /b 1
)

echo Собираю однофайловое приложение...
pyinstaller --noconfirm --onefile --windowed --name "Конвертер расписания" gui.py
if errorlevel 1 (
    echo Сборка не удалась. Подробности выше.
    pause
    exit /b 1
)

echo.
echo Готово: dist\Конвертер расписания.exe
echo Рядом с exe при первом запуске создадутся aliases.json и settings.json.
pause
exit /b 0
