@echo off
chcp 65001 >nul
echo ===================================
echo   Сверка документов — запуск
echo ===================================
echo.
echo [1/2] Установка зависимостей...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo ОШИБКА: pip не найден.
    echo Установите Python 3.9+ и добавьте его в PATH.
    pause
    exit /b 1
)
echo.
echo [2/2] Запуск сервера...
echo.
echo  Откройте браузер:  http://localhost:8000
echo  Остановка:         Ctrl+C
echo.
python main.py
pause
