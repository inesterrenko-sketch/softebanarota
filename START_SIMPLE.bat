@echo off
chcp 65001 >nul
title Contract Filler App - Быстрый запуск
cls

echo ========================================
echo   Contract Filler App
echo   Быстрый запуск (без venv)
echo ========================================
echo.

REM Переход в папку приложения
cd /d "%~dp0"

REM Установка зависимостей напрямую
echo [ИНФО] Устанавливаю зависимости...
python -m pip install --upgrade pip --quiet --break-system-packages 2>nul
python -m pip install -r requirements.txt --quiet --break-system-packages 2>nul

echo [OK] Зависимости установлены
echo.

REM Запуск приложения
echo ========================================
echo   Запускаю приложение...
echo   Откройте браузер: http://127.0.0.1:8000
echo   Для остановки нажмите Ctrl+C
echo ========================================
echo.

python -m uvicorn app.main:app --host 127.0.0.1 --port 8000

pause
