@echo off
chcp 65001 >nul
title Contract Filler App - Автозапуск
cls

echo ========================================
echo   Contract Filler App
echo   Автоматический запуск приложения
echo ========================================
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден! Установите Python 3.8+
    pause
    exit /b 1
)

echo [OK] Python установлен
echo.

REM Переход в папку приложения
cd /d "%~dp0"

REM Проверка виртуального окружения
if not exist ".venv" (
    echo [ИНФО] Создаю виртуальное окружение...
    python -m venv .venv
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось создать виртуальное окружение
        pause
        exit /b 1
    )
    echo [OK] Виртуальное окружение создано
    echo.
)

REM Активация виртуального окружения
echo [ИНФО] Активирую виртуальное окружение...
call .venv\Scripts\activate.bat

REM Установка зависимостей
echo [ИНФО] Устанавливаю зависимости...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet

echo [OK] Зависимости установлены
echo.

REM Запуск приложения
echo ========================================
echo   Запускаю приложение...
echo   Откройте браузер: http://127.0.0.1:8000
echo   Для остановки нажмите Ctrl+C
echo ========================================
echo.

REM Используем python -m uvicorn вместо просто uvicorn
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000

pause
