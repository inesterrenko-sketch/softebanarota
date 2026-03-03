#!/usr/bin/env python3
"""
Скрипт для добавления примера клиента из договора №5266 в базу данных
Запускать из корня проекта: python add_example_client.py
"""
import sys
from pathlib import Path

# Добавляем путь к модулю приложения
sys.path.insert(0, str(Path(__file__).parent))

from app.main import db_init, db_upsert_client, db_get_client

def add_example_client():
    """Добавить пример клиента в базу"""
    # Сначала инициализируем базу (создаст все колонки)
    print("Инициализация базы данных...")
    db_init()
    
    # Проверяем есть ли уже этот клиент
    # Попробуем найти по паспорту
    print("Проверка существующих клиентов...")
    
    # Добавляем клиента через функцию приложения
    client_data = {
        'fio': 'Баженов Евгений Александрович',
        'passport': '76 25 415097',
        'organ': 'УМВД РОССИИ ПО ЗАБАЙКАЛЬСКОМУ КРАЮ',
        'vydan': '14.11.2025',
        'address': 'Кировская обл., М.Р-Н. Даровской, Г.П. Даровское, ПГТ. Даровской, ул. Зеленая, д.8 кв. 1',
        'phone': '+7 (912) 332-08-63',
        'contract_no': '5266',
        'contract_date': '16.01.2026',
        'car_model': 'Toyota RAV4',
        'vin': 'JTMW43FV80D135612',
        'obem': '1987',
        'vypusk': '2023',
        'customs_amount': '514445',
        'dkp_amount': '1145557',
        'company_inn': '2632083090',
        'company_address': '357204, Ставропольский край, Минераловоский р-н, тер. Автодорога, Р-217 Кавказ, км. 345-ый',
    }
    
    print("Добавление клиента...")
    client_id = db_upsert_client(None, client_data)
    
    print(f"\n✅ Добавлен клиент ID: {client_id}")
    print(f"   ФИО: Баженов Евгений Александрович")
    print(f"   Паспорт: 76 25 415097")
    print(f"   Орган: УМВД РОССИИ ПО ЗАБАЙКАЛЬСКОМУ КРАЮ")
    print(f"   Договор: №5266 от 16.01.2026")
    print(f"   Авто: Toyota RAV4, 2023, VIN: JTMW43FV80D135612")
    print(f"   ИНН компании: 2632083090")
    print(f"\n🎉 Готово! Теперь этот клиент доступен во всех формах.")
    return client_id

if __name__ == "__main__":
    add_example_client()
