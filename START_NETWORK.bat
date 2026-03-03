@echo off
chcp 65001 >nul
title ContractFill - Запуск в сети

echo.
echo ╔═══════════════════════════════════════════════════════════╗
echo ║     CONTRACTFILL - ЗАПУСК В ЛОКАЛЬНОЙ СЕТИ               ║
echo ╚═══════════════════════════════════════════════════════════╝
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python не найден! Установите Python 3.11 или выше.
    echo    Скачать: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✓ Python найден
echo.

REM Проверка виртуального окружения
if not exist "venv\Scripts\activate.bat" (
    echo 📦 Создание виртуального окружения...
    python -m venv venv
    if errorlevel 1 (
        echo ❌ Ошибка создания виртуального окружения
        pause
        exit /b 1
    )
    echo ✓ Виртуальное окружение создано
    echo.
)

REM Активация виртуального окружения
echo 🔄 Активация окружения...
call venv\Scripts\activate.bat
echo ✓ Окружение активировано
echo.

REM Установка/обновление зависимостей
echo 📥 Проверка зависимостей...
pip install -q --upgrade pip
pip install -q -r requirements.txt
if errorlevel 1 (
    echo ❌ Ошибка установки зависимостей
    pause
    exit /b 1
)
echo ✓ Зависимости установлены
echo.

REM Получение IP адреса
echo 🌐 Определение IP адреса...
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
    set IP=%%a
    goto :found_ip
)
:found_ip
set IP=%IP: =%
echo ✓ Ваш IP адрес: %IP%
echo.

echo ╔═══════════════════════════════════════════════════════════╗
echo ║  ПРИЛОЖЕНИЕ ЗАПУСКАЕТСЯ В РЕЖИМЕ СЕТИ                    ║
echo ║                                                           ║
echo ║  Доступ с этого компьютера:                              ║
echo ║  → http://localhost:8000                                  ║
echo ║                                                           ║
echo ║  Доступ с других компьютеров в сети:                     ║
echo ║  → http://%IP%:8000                          ║
echo ║                                                           ║
echo ║  Для остановки: Ctrl+C                                   ║
echo ╚═══════════════════════════════════════════════════════════╝
echo.
echo 🚀 Запуск сервера...
echo.

REM Запуск с доступом из сети (0.0.0.0)
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload

pause
