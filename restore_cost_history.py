"""
Скрипт для восстановления исторических данных себестоимости.
Для каждого товара берется самая ранняя известная дата, и генерируются строки
с датами от 22.09.2025 до этой даты с теми же значениями остатка и себестоимости.
"""

import pandas as pd
from datetime import timedelta
import os

SHEET_NAME = 'Себестоимость'
NEW_SHEET_NAME = 'Восстановленные данные'
START_DATE = pd.to_datetime('2025-09-22')


def find_input_file():
    """Находит файл с new_cost в имени (для CLI)."""
    for file in os.listdir('.'):
        if 'new_cost' in file and file.endswith('.xlsx'):
            return file
    raise FileNotFoundError("Не найден файл с 'new_cost' в имени")


def restore_cost_history(input_file):
    """
    Восстанавливает исторические данные себестоимости.
    input_file: путь к xlsx.
    Добавляет вкладку «Восстановленные данные» с историей с 22.09.2025.
    """
    df = pd.read_excel(input_file, sheet_name=SHEET_NAME)
    
    # Преобразуем дату в datetime
    df['date'] = pd.to_datetime(df['date'])
    
    earliest_dates = df.groupby('product_id')['date'].min()
    
    restored_rows = []
    
    for product_id, min_date in earliest_dates.items():
        # Получаем строку с минимальной датой для этого товара
        product_data = df[(df['product_id'] == product_id) & (df['date'] == min_date)].iloc[0]
        stock = product_data['stock']
        cost = product_data['cost']
        
        # Если минимальная дата больше 22.09.2025, генерируем исторические данные
        if min_date > START_DATE:
            # Генерируем даты от START_DATE до (min_date - 1 день)
            date_range = pd.date_range(start=START_DATE, end=min_date - timedelta(days=1), freq='D')
            
            # Создаем строки для каждой даты
            for date in date_range:
                restored_rows.append({
                    'product_id': product_id,
                    'stock': stock,
                    'cost': cost,
                    'date': date
                })
    
    restored_df = pd.DataFrame(restored_rows)
    
    if len(restored_df) == 0:
        return {'rows': 0, 'products': 0, 'message': 'Нет данных для восстановления (все товары имеют даты <= 22.09.2025)'}
    
    # Сортируем
    restored_df = restored_df.sort_values(['product_id', 'date']).reset_index(drop=True)
    
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        restored_df.to_excel(writer, sheet_name=NEW_SHEET_NAME, index=False)

    return {
        'rows': len(restored_df),
        'products': restored_df['product_id'].nunique(),
        'date_min': restored_df['date'].min(),
        'date_max': restored_df['date'].max()
    }


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else find_input_file()
    r = restore_cost_history(path)
    if r:
        print(f"Готово: {r['rows']} строк, {r['products']} товаров")
