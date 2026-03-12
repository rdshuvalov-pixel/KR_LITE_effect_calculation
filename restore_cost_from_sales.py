"""
Скрипт для восстановления данных себестоимости на основе данных продаж.
Для каждого уникального product_id из продаж создаются записи на каждую дату
между минимальной и максимальной датой продаж со значениями stock=1 и cost=1.
"""

import pandas as pd

# Константы
SALES_SHEET = 'Продажи'
COST_SHEET = 'Себестоимость'


def restore_cost_from_sales(input_file):
    """
    Восстанавливает данные себестоимости на основе данных продаж.
    input_file: путь к xlsx или file-like object.
    """
    # 1. Чтение данных из "Продажи"
    sales_df = pd.read_excel(input_file, sheet_name=SALES_SHEET)
    
    # Преобразуем дату в datetime
    sales_df['recorded_on'] = pd.to_datetime(sales_df['recorded_on'])
    
    # Извлекаем уникальные product_id
    unique_products = sales_df['product_id'].unique()
    
    min_date = sales_df['recorded_on'].min()
    max_date = sales_df['recorded_on'].max()
    
    # 2. Генерация восстановленных данных
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
    
    # 3. Объединение с существующими данными
    existing_costs_df = pd.read_excel(input_file, sheet_name=COST_SHEET)
    
    # Преобразуем дату в datetime для существующих данных
    if not existing_costs_df.empty and 'date' in existing_costs_df.columns:
        existing_costs_df['date'] = pd.to_datetime(existing_costs_df['date'])
        combined_df = pd.concat([existing_costs_df, restored_df], ignore_index=True)
    else:
        combined_df = restored_df
    
    before_dedup = len(combined_df)
    combined_df = combined_df.drop_duplicates(subset=['product_id', 'date'], keep='first')
    
    # Сортируем по product_id и date
    combined_df = combined_df.sort_values(['product_id', 'date']).reset_index(drop=True)
    
    # 4. Сохранение результатов
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        combined_df.to_excel(writer, sheet_name=COST_SHEET, index=False)

    return {
        'rows': len(combined_df),
        'products': combined_df['product_id'].nunique(),
        'date_min': combined_df['date'].min(),
        'date_max': combined_df['date'].max()
    }


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Энтерсайт_Литовск_анализ_эффекта_10012026.xlsx'
    r = restore_cost_from_sales(path)
    print(f"Готово: {r['rows']} строк, {r['products']} товаров")
