"""
Скрипт для восстановления данных себестоимости на основе данных продаж.
Для каждого уникального product_id из продаж создаются записи на каждую дату
между минимальной и максимальной датой продаж со значениями stock=1 и cost=1.
"""

import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import os

# Константы
INPUT_FILE = 'Энтерсайт_Литовск_анализ_эффекта_10012026.xlsx'
SALES_SHEET = 'Продажи'
COST_SHEET = 'Себестоимость'


def restore_cost_from_sales():
    """
    Восстанавливает данные себестоимости на основе данных продаж.
    """
    print(f"Чтение данных из файла {INPUT_FILE}...")
    
    # 1. Чтение данных из "Продажи"
    print(f"Загрузка данных из вкладки '{SALES_SHEET}'...")
    sales_df = pd.read_excel(INPUT_FILE, sheet_name=SALES_SHEET)
    
    # Преобразуем дату в datetime
    sales_df['recorded_on'] = pd.to_datetime(sales_df['recorded_on'])
    
    # Извлекаем уникальные product_id
    unique_products = sales_df['product_id'].unique()
    print(f"Найдено уникальных товаров: {len(unique_products)}")
    
    # Находим минимальную и максимальную дату
    min_date = sales_df['recorded_on'].min()
    max_date = sales_df['recorded_on'].max()
    print(f"Диапазон дат продаж: {min_date.date()} - {max_date.date()}")
    
    # 2. Генерация восстановленных данных
    print("\nГенерация восстановленных данных...")
    restored_rows = []
    
    for product_id in unique_products:
        # Создаем диапазон дат от минимальной до максимальной (включительно)
        date_range = pd.date_range(start=min_date, end=max_date, freq='D')
        
        # Для каждой даты создаем запись
        for date in date_range:
            restored_rows.append({
                'product_id': product_id,
                'stock': 1,
                'cost': 1,
                'date': date
            })
    
    restored_df = pd.DataFrame(restored_rows)
    print(f"Сгенерировано {len(restored_df)} строк восстановленных данных")
    
    # 3. Объединение с существующими данными
    print(f"\nЗагрузка существующих данных из вкладки '{COST_SHEET}'...")
    existing_costs_df = pd.read_excel(INPUT_FILE, sheet_name=COST_SHEET)
    
    # Преобразуем дату в datetime для существующих данных
    if not existing_costs_df.empty and 'date' in existing_costs_df.columns:
        existing_costs_df['date'] = pd.to_datetime(existing_costs_df['date'])
        print(f"Загружено {len(existing_costs_df)} существующих строк")
        
        # Объединяем данные
        combined_df = pd.concat([existing_costs_df, restored_df], ignore_index=True)
    else:
        combined_df = restored_df
        print("Существующих данных не найдено")
    
    # Удаляем дубликаты (если для product_id+date уже есть запись)
    print("\nУдаление дубликатов...")
    before_dedup = len(combined_df)
    combined_df = combined_df.drop_duplicates(subset=['product_id', 'date'], keep='first')
    after_dedup = len(combined_df)
    print(f"Удалено дубликатов: {before_dedup - after_dedup}")
    
    # Сортируем по product_id и date
    combined_df = combined_df.sort_values(['product_id', 'date']).reset_index(drop=True)
    
    # 4. Сохранение результатов
    print(f"\nСохранение данных в вкладку '{COST_SHEET}'...")
    
    # Загружаем существующий файл
    book = load_workbook(INPUT_FILE)
    
    # Записываем данные обратно на вкладку "Себестоимость"
    with pd.ExcelWriter(INPUT_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        combined_df.to_excel(writer, sheet_name=COST_SHEET, index=False)
    
    print(f"Данные успешно сохранены в вкладку '{COST_SHEET}'")
    print(f"\nСтатистика:")
    print(f"  - Всего строк: {len(combined_df)}")
    print(f"  - Уникальных товаров: {combined_df['product_id'].nunique()}")
    print(f"  - Диапазон дат: {combined_df['date'].min().date()} - {combined_df['date'].max().date()}")


if __name__ == '__main__':
    restore_cost_from_sales()
