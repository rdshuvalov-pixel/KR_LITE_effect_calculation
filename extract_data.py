#!/usr/bin/env python3
"""
Извлечение данных для презентации по позиции 04871
"""
import pandas as pd
import numpy as np
import json
from pathlib import Path

# Параметры
PRODUCT_ID = 4871  # "КИНДЕР БУЭНО" БАТОНЧИК МОЛОЧНЫЙ 43 ГР
PRE_TEST_WEEKS = 4
STOCK_THRESHOLD_PCT = 51  # Допустимое время на остатке >= 51%
CONTIGUOUS = True  # Непрерывный дотестовый период

# Загрузка данных
excel_path = Path(__file__).parent / "S-market_эффект_18012026.xlsx"
xl = pd.ExcelFile(excel_path)

test_prices = pd.read_excel(xl, 'Тестовые цены')
sales = pd.read_excel(xl, 'Продажи')
costs = pd.read_excel(xl, 'Себестоимость')

# Предобработка дат
test_prices['New_Price_Start'] = pd.to_datetime(test_prices['New_Price_Start'])
sales['recorded_on'] = pd.to_datetime(sales['recorded_on'])
costs['date'] = pd.to_datetime(costs['date'])

# Выручка
sales['revenue'] = sales['price'] * sales['quantity']
sales['week_start'] = sales['recorded_on'].apply(lambda x: x - pd.Timedelta(days=x.weekday())).dt.normalize()
costs['week_start'] = costs['date'].apply(lambda x: x - pd.Timedelta(days=x.weekday())).dt.normalize()

# Используем заданный product_id
pid = PRODUCT_ID
product_match = sales[sales['product_id'] == pid]
if product_match.empty:
    print(f"Позиция {pid} не найдена в продажах!")
    exit(1)

product_name = product_match['name_full'].iloc[0] if 'name_full' in product_match.columns else f"Товар {pid}"
print(f"Найден product_id: {pid}, название: {product_name}")

# Проверяем, что это тестовый товар
if pid not in test_prices['product_id'].values:
    print(f"ВНИМАНИЕ: Позиция {pid} не является тестовой!")
    exit(1)

# Определяем дату начала теста
product_test_info = test_prices[test_prices['product_id'] == pid]
test_start_date = product_test_info['New_Price_Start'].min()
test_start_week = pd.to_datetime(test_start_date - pd.Timedelta(days=test_start_date.weekday())).normalize()
print(f"Дата начала теста: {test_start_date}, неделя: {test_start_week}")

# Определяем контрольную группу
all_product_ids = set(sales['product_id'].unique())
test_product_ids = set(test_prices['product_id'].unique())
control_product_ids = all_product_ids - test_product_ids
print(f"Всего товаров: {len(all_product_ids)}, тестовых: {len(test_product_ids)}, контрольных: {len(control_product_ids)}")

# Агрегация остатков по неделям
costs['has_stock'] = costs['stock'] > 0
weekly_stock_days = costs.groupby(['product_id', 'week_start'])['has_stock'].sum().reset_index()
stock_days_lookup = dict(zip(zip(weekly_stock_days['product_id'], weekly_stock_days['week_start']), weekly_stock_days['has_stock']))

# Поиск дотестового периода
min_days_stock = 7 * (STOCK_THRESHOLD_PCT / 100.0)
print(f"Минимум дней на остатке: {min_days_stock:.1f}")

def find_pre_test_period(pid, end_week, n_weeks, min_days):
    """Поиск непрерывного дотестового периода"""
    max_lookback = 26
    current_end = end_week
    
    for _ in range(max_lookback):
        current_start = current_end - pd.Timedelta(weeks=n_weeks-1)
        window_weeks = pd.date_range(start=current_start, end=current_end, freq='W-MON')
        
        is_valid = True
        for w in window_weeks:
            stock_days = stock_days_lookup.get((pid, w), 0)
            if stock_days < min_days:
                is_valid = False
                break
        
        if is_valid:
            return list(window_weeks)
        
        current_end = current_end - pd.Timedelta(weeks=1)
    
    return None

# Ищем дотестовый период (начиная с недели до теста)
pre_test_search_end = test_start_week - pd.Timedelta(weeks=1)
pre_test_weeks = find_pre_test_period(pid, pre_test_search_end, PRE_TEST_WEEKS, min_days_stock)

if pre_test_weeks is None:
    print("Не удалось найти валидный дотестовый период!")
    # Берём просто 4 недели до теста
    pre_test_weeks = pd.date_range(end=pre_test_search_end, periods=PRE_TEST_WEEKS, freq='W-MON').tolist()

print(f"Дотестовый период: {[w.strftime('%Y-%m-%d') for w in pre_test_weeks]}")

# Определяем тестовый период (4 недели начиная с test_start_week)
last_data_week = sales['week_start'].max()
test_weeks_list = pd.date_range(start=test_start_week, periods=4, freq='W-MON')
test_weeks_list = [w for w in test_weeks_list if w <= last_data_week]
print(f"Тестовый период: {[w.strftime('%Y-%m-%d') for w in test_weeks_list]}")

# Агрегация выручки по неделям
weekly_sales = sales.groupby(['product_id', 'week_start']).agg({
    'revenue': 'sum',
    'quantity': 'sum'
}).reset_index()

# Выручка по позиции
product_weekly = weekly_sales[weekly_sales['product_id'] == pid].set_index('week_start')['revenue']

# Выручка контрольной группы
control_sales = sales[sales['product_id'].isin(control_product_ids)]
control_weekly = control_sales.groupby('week_start')['revenue'].sum()

# Собираем данные для графика
all_weeks = sorted(list(pre_test_weeks) + list(test_weeks_list))
weeks_str = [w.strftime('%Y-%m-%d') for w in all_weeks]
revenue_position = [int(product_weekly.get(w, 0)) for w in all_weeks]
revenue_control = [int(control_weekly.get(w, 0)) for w in all_weeks]

# Средние значения
pre_test_pos_avg = np.mean([product_weekly.get(w, 0) for w in pre_test_weeks])
test_pos_avg = np.mean([product_weekly.get(w, 0) for w in test_weeks_list])
pre_test_ctrl_avg = np.mean([control_weekly.get(w, 0) for w in pre_test_weeks])
test_ctrl_avg = np.mean([control_weekly.get(w, 0) for w in test_weeks_list])

print("\n=== РЕЗУЛЬТАТЫ ===")
print(f"Позиция: {product_name} (id {pid})")
print(f"Дотестовый период: {pre_test_weeks[0].strftime('%d.%m.%Y')} - {pre_test_weeks[-1].strftime('%d.%m.%Y')}")
print(f"Тестовый период: {test_weeks_list[0].strftime('%d.%m.%Y')} - {test_weeks_list[-1].strftime('%d.%m.%Y')}")
print(f"\nСредняя выручка позиции:")
print(f"  Дотест: {pre_test_pos_avg:,.0f} ₽")
print(f"  Тест: {test_pos_avg:,.0f} ₽")
print(f"\nСредняя выручка контрольной группы:")
print(f"  Дотест: {pre_test_ctrl_avg:,.0f} ₽")
print(f"  Тест: {test_ctrl_avg:,.0f} ₽")

# Расчёт эффекта
pos_growth = (test_pos_avg / pre_test_pos_avg - 1) * 100 if pre_test_pos_avg > 0 else 0
ctrl_growth = (test_ctrl_avg / pre_test_ctrl_avg - 1) * 100 if pre_test_ctrl_avg > 0 else 0
effect = pos_growth - ctrl_growth

print(f"\nПрирост позиции: {pos_growth:+.1f}%")
print(f"Прирост контрольной группы: {ctrl_growth:+.1f}%")
print(f"ЭФФЕКТ ПЕРЕОЦЕНКИ: {effect:+.1f}%")

# Формируем данные для презентации
presentation_data = {
    "productName": str(product_name),
    "productId": PRODUCT_ID,
    "weeks": weeks_str,
    "revenuePosition": revenue_position,
    "revenueControl": revenue_control,
    "preTestWeeks": len(pre_test_weeks),
    "testWeeks": len(test_weeks_list),
    "preTestPosAvg": round(pre_test_pos_avg),
    "testPosAvg": round(test_pos_avg),
    "preTestCtrlAvg": round(pre_test_ctrl_avg),
    "testCtrlAvg": round(test_ctrl_avg),
    "posGrowth": round(pos_growth, 1),
    "ctrlGrowth": round(ctrl_growth, 1),
    "effect": round(effect, 1)
}

print(f"\n=== JSON для презентации ===")
print(json.dumps(presentation_data, ensure_ascii=False, indent=2))

# Сохраняем в файл
output_path = Path(__file__).parent / "presentation_data.json"
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(presentation_data, f, ensure_ascii=False, indent=2)
print(f"\nДанные сохранены в {output_path}")
