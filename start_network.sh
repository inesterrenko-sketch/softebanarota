#!/bin/bash

# Цвета для вывода
RED='\033[0;31m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo ""
echo "╔═══════════════════════════════════════════════════════════╗"
echo "║     CONTRACTFILL - ЗАПУСК В ЛОКАЛЬНОЙ СЕТИ               ║"
echo "╚═══════════════════════════════════════════════════════════╝"
echo ""

# Проверка Python
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Python3 не найден!${NC}"
    echo "   Установите Python 3.11 или выше"
    exit 1
fi

echo -e "${GREEN}✓ Python найден${NC}"
echo ""

# Создание виртуального окружения
if [ ! -d "venv" ]; then
    echo "📦 Создание виртуального окружения..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo -e "${RED}❌ Ошибка создания виртуального окружения${NC}"
        exit 1
    fi
    echo -e "${GREEN}✓ Виртуальное окружение создано${NC}"
    echo ""
fi

# Активация окружения
echo "🔄 Активация окружения..."
source venv/bin/activate
echo -e "${GREEN}✓ Окружение активировано${NC}"
echo ""

# Установка зависимостей
echo "📥 Проверка зависимостей..."
pip install -q --upgrade pip
pip install -q -r requirements.txt
if [ $? -ne 0 ]; then
    echo -e "${RED}❌ Ошибка установки зависимостей${NC}"
    exit 1
fi
echo -e "${GREEN}✓ Зависимости установлены${NC}"
echo ""

# Получение IP адреса
echo "🌐 Определение IP адреса..."
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS
    IP=$(ipconfig getifaddr en0)
    if [ -z "$IP" ]; then
        IP=$(ipconfig getifaddr en1)
    fi
else
    # Linux
    IP=$(hostname -I | awk '{print $1}')
fi

if [ -z "$IP" ]; then
    IP="localhost"
fi

echo -e "${GREEN}✓ Ваш IP адрес: $IP${NC}"
echo ""

echo "╔═══════════════════════════════════════════════════════════╗"
echo "║  ПРИЛОЖЕНИЕ ЗАПУСКАЕТСЯ В РЕЖИМЕ СЕТИ                    ║"
echo "║                                                           ║"
echo "║  Доступ с этого компьютера:                              ║"
echo "║  → http://localhost:8000                                  ║"
echo "║                                                           ║"
echo "║  Доступ с других компьютеров в сети:                     ║"
echo "║  → http://$IP:8000                                   ║"
echo "║                                                           ║"
echo "║  Для остановки: Ctrl+C                                   ║"
echo "╚═══════════════════════════════════════════════════════════╝"
echo ""
echo "🚀 Запуск сервера..."
echo ""

# Запуск с доступом из сети
python3 -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
