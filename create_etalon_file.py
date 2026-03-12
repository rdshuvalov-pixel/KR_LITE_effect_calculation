#!/usr/bin/env python3
"""
Создаёт эталонный Excel-файл для проверки работы KeepRise Lite.
Запуск: python3 create_etalon_file.py
Результат: etalon_check.xlsx в текущей директории.
"""
import pandas as pd
from io import BytesIO
from pathlib import Path
from datetime import datetime, timedelta

OUTPUT = Path(__file__).parent / "etalon_check.xlsx"

# Тестовые товары (будут в переоценке)
TEST_PIDS = [101, 102]
# Контрольные товары
CONTROL_PIDS = [201, 202, 203, 204, 205]
ALL_PIDS = TEST_PIDS + CONTROL_PIDS

# Дата старта новой цены (вторник — тест начнётся со следующей недели)
PRICE_START = datetime(2025, 12, 2)  # 02.12.2025
# Диапазон данных
START_DATE = datetime(2025, 11, 1)
END_DATE = datetime(2025, 12, 28)


def week_start(d):
    """Понедельник недели."""
    return d - timedelta(days=d.weekday())


def generate_sales():
    """Тестовые товары: до PRICE_START — цена 120 (Current), после — 150 (New_Price в допуске активации)."""
    rows = []
    for pid in ALL_PIDS:
        name = f"Товар {pid} ({'тест' if pid in TEST_PIDS else 'контроль'})"
        d = START_DATE
        while d <= END_DATE:
            if pid in TEST_PIDS and d >= PRICE_START:
                price = 150.0  # New_Price — в допуске активации (1%)
            elif pid in TEST_PIDS:
                price = 120.0  # Current_Price
            else:
                price = 100 + (pid % 10) * 10  # контроль
            for _ in range(2):
                rows.append({
                    'product_id': pid,
                    'recorded_on': d,
                    'price': price,
                    'quantity': 1 + (pid % 3),
                    'name_full': name
                })
            d += timedelta(days=1)
    return pd.DataFrame(rows)


def generate_test_prices():
    rows = []
    for pid in TEST_PIDS:
        rows.append({
            'product_id': pid,
            'New_Price_Start': PRICE_START,
            'New_Price': 150,
            'Current_Price': 120
        })
    return pd.DataFrame(rows)


def generate_costs():
    rows = []
    for pid in ALL_PIDS:
        d = START_DATE
        while d <= END_DATE:
            rows.append({
                'product_id': pid,
                'date': d,
                'cost': 80,
                'stock': 5
            })
            d += timedelta(days=1)
    return pd.DataFrame(rows)


def get_etalon_bytes() -> bytes:
    """Генерирует эталон в памяти и возвращает байты для скачивания."""
    sales = generate_sales()
    test_prices = generate_test_prices()
    costs = generate_costs()
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        sales.to_excel(w, sheet_name='Продажи', index=False)
        test_prices.to_excel(w, sheet_name='Тестовые цены', index=False)
        costs.to_excel(w, sheet_name='Себестоимость', index=False)
    return buf.getvalue()


def main():
    sales = generate_sales()
    test_prices = generate_test_prices()
    costs = generate_costs()

    with pd.ExcelWriter(OUTPUT, engine='openpyxl') as w:
        sales.to_excel(w, sheet_name='Продажи', index=False)
        test_prices.to_excel(w, sheet_name='Тестовые цены', index=False)
        costs.to_excel(w, sheet_name='Себестоимость', index=False)

    print(f"Создан: {OUTPUT}")
    print(f"  Продажи: {len(sales)} строк, товаров {len(ALL_PIDS)}")
    print(f"  Тестовые цены: {len(TEST_PIDS)} товаров, старт {PRICE_START.date()}")
    print(f"  Себестоимость: {len(costs)} строк")
    print("\nЗагрузите etalon_check.xlsx в приложение для проверки.")

if __name__ == "__main__":
    main()
