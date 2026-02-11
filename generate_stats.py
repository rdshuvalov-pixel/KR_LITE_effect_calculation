#!/usr/bin/env python3
"""
Генерация полной статистики для презентации с параметрами сервиса.
"""
import pandas as pd
import numpy as np
import json
from pathlib import Path
from datetime import datetime

from calculator import EffectCalculator

# Параметры сервиса (как указал пользователь)
PARAMS = {
    # Дотестовый период
    'pre_test_weeks_count': 4,
    'contiguous_pre_test': True,  # Непрерывный
    'pre_test_stock_threshold': 51,  # Допуск остатка 51%
    
    # Тестовый период
    'use_stock_filter': True,  # Фильтрация недель включена
    'stock_threshold_pct': 30,  # Допуск остатка 30% от среднедневного в дотестовый
    
    # Активация цены
    'activation_threshold': 1,  # Порог отклонения 1%
    'activation_use_rounding': True,  # Округление включено
    'activation_round_value': 90,  # До 90 копеек
    'activation_round_direction': 'up',  # Вверх
    'activation_wap_from_change_date': True,
    'activation_min_days_threshold': 4,  # Мин. дней WAP 4
    
    # Режим расчёта
    'test_use_week_values': False,  # Среднее по тестовым неделям
}

print("=" * 60)
print("РАСЧЁТ СТАТИСТИКИ ПЕРЕОЦЕНКИ")
print("=" * 60)
print("\nПараметры сервиса:")
print(f"  Дотестовый период: {PARAMS['pre_test_weeks_count']} недель, {'непрерывный' if PARAMS['contiguous_pre_test'] else 'не непрерывный'}")
print(f"  Допуск остатка (дотест): {PARAMS['pre_test_stock_threshold']}%")
print(f"  Допуск остатка (тест): {PARAMS['stock_threshold_pct']}% от базового")
print(f"  Порог отклонения цены: {PARAMS['activation_threshold']}%")
print(f"  Округление: {'вверх' if PARAMS['activation_round_direction'] == 'up' else 'вниз'} до {PARAMS['activation_round_value']} коп.")
print(f"  Мин. дней WAP: {PARAMS['activation_min_days_threshold']}")

# Загрузка данных
excel_path = Path(__file__).parent / "S-market_эффект_18012026.xlsx"
print(f"\nЗагрузка данных из {excel_path.name}...")

calc = EffectCalculator(excel_path)
calc.preprocess()

print(f"  Всего товаров: {len(calc.all_products)}")
print(f"  Тестовых товаров: {len(calc.test_product_ids)}")
print(f"  Контрольная группа: {len(calc.control_product_ids)}")

# Расчёт
print("\nВыполняется расчёт...")
results = calc.calculate(
    use_stock_filter=PARAMS['use_stock_filter'],
    stock_threshold_pct=PARAMS['stock_threshold_pct'],
    pre_test_weeks_count=PARAMS['pre_test_weeks_count'],
    pre_test_stock_threshold=PARAMS['pre_test_stock_threshold'],
    contiguous_pre_test=PARAMS['contiguous_pre_test'],
    activation_threshold=PARAMS['activation_threshold'],
    activation_use_rounding=PARAMS['activation_use_rounding'],
    activation_round_value=PARAMS['activation_round_value'],
    activation_round_direction=PARAMS['activation_round_direction'],
    activation_wap_from_change_date=PARAMS['activation_wap_from_change_date'],
    activation_min_days_threshold=PARAMS['activation_min_days_threshold'],
    test_use_week_values=PARAMS['test_use_week_values'],
)

summary = calc.get_summary()

if summary is None:
    print("Ошибка: не удалось получить статистику!")
    exit(1)

print("\n" + "=" * 60)
print("РЕЗУЛЬТАТЫ")
print("=" * 60)

# 1. Список переоценок с датами и количеством позиций (отправлено / встало на полку / принято для анализа)
print("\n1. СПИСОК ПЕРЕОЦЕНОК:")
price_changes = []

# Получаем данные активации для определения "встало на полку"
activation_df = calc.get_activation_details(
    PARAMS['activation_threshold'],
    use_rounding=PARAMS['activation_use_rounding'],
    round_value=PARAMS['activation_round_value'],
    round_direction=PARAMS['activation_round_direction'],
    wap_from_change_date=PARAMS['activation_wap_from_change_date'],
    min_days_threshold=PARAMS['activation_min_days_threshold']
)

# "Встало на полку" = позиции где хотя бы раз была наша цена
activated_pids = set()
if not activation_df.empty:
    # Фильтруем записи где статус НЕ начинается с "не та цена"
    activated_rows = activation_df[~activation_df['Status'].str.startswith('не та цена', na=False)]
    activated_pids = set(activated_rows['product_id'].unique())

# Группируем по дате начала теста
activated_by_date = {}
for pid in activated_pids:
    test_info = calc.test_prices[calc.test_prices['product_id'] == pid]
    if not test_info.empty:
        start_date = test_info['New_Price_Start'].min()
        if start_date not in activated_by_date:
            activated_by_date[start_date] = 0
        activated_by_date[start_date] += 1

# Принятые для анализа (не исключённые)
valid_results = results[results['Is_Excluded'] == False]
accepted_by_date = valid_results.groupby('Test_Start_Date')['product_id'].nunique()

print(f"   {'Дата':<12} {'Отправлено':>10} {'На полку':>10} {'В анализ':>10}")
print("   " + "-" * 45)
total_sent, total_activated, total_accepted = 0, 0, 0

for date, sent_count in summary['products_per_change'].items():
    date_str = date.strftime('%d.%m.%Y') if hasattr(date, 'strftime') else str(date)
    activated_count = int(activated_by_date.get(date, 0))
    accepted_count = int(accepted_by_date.get(date, 0))
    print(f"   {date_str:<12} {sent_count:>10} {activated_count:>10} {accepted_count:>10}")
    total_sent += sent_count
    total_activated += activated_count
    total_accepted += accepted_count
    price_changes.append({
        'date': date_str, 
        'sent': int(sent_count),
        'activated': activated_count,
        'accepted': accepted_count
    })

print("   " + "-" * 45)
print(f"   {'Итого':<12} {total_sent:>10} {total_activated:>10} {total_accepted:>10}")

# 2. Суммарный эффект
print("\n2. СУММАРНЫЙ ЭФФЕКТ:")
print(f"   Эффект по выручке: {summary['effect_revenue_pct']:+.2f}%")
print(f"   Эффект по прибыли: {summary['effect_profit_pct']:+.2f}%")
print(f"   Абс. эффект по выручке: {summary['total_abs_effect_revenue']:+,.0f} ₽")
print(f"   Абс. эффект по прибыли: {summary['total_abs_effect_profit']:+,.0f} ₽")

# 3. Выручка тестового ассортимента
print("\n3. ВЫРУЧКА ТЕСТОВОГО АССОРТИМЕНТА:")
print(f"   Без эффекта: {summary['revenue_without_effect']:,.0f} ₽")
print(f"   С эффектом (факт): {summary['total_fact_revenue']:,.0f} ₽")
print(f"   Прибыль без эффекта: {summary['profit_without_effect']:,.0f} ₽")
print(f"   Прибыль с эффектом (факт): {summary['total_fact_profit']:,.0f} ₽")

# 4. Выручка магазина за тестовый период
print("\n4. ВЫРУЧКА МАГАЗИНА ЗА ТЕСТОВЫЙ ПЕРИОД:")
print(f"   Общая выручка: {summary['global_revenue']:,.0f} ₽")
print(f"   Общая прибыль: {summary['global_profit']:,.0f} ₽")
print(f"   Доля тестового ассортимента в выручке: {summary['test_share_revenue']:.2f}%")
print(f"   Доля тестового ассортимента в прибыли: {summary['test_share_profit']:.2f}%")

# 5. Дополнительная статистика
print("\n5. ДОПОЛНИТЕЛЬНАЯ СТАТИСТИКА:")
print(f"   Длительность теста: {summary['test_duration_weeks']:.0f} недель")
print(f"   Количество переоценок: {summary['price_changes_count']}")
print(f"   Тестовых позиций всего: {summary['tested_count']}")
print(f"   Позиций с ростом: {summary['growth_stats']['count']} (эффект выр.: {summary['growth_stats']['revenue_effect']:+,.0f} ₽)")
print(f"   Позиций со снижением: {summary['decline_stats']['count']} (эффект выр.: {summary['decline_stats']['revenue_effect']:+,.0f} ₽)")

# 6. СЦЕНАРИИ
print("\n6. МОДЕЛЬНЫЕ СЦЕНАРИИ:")

global_revenue = summary['global_revenue']
global_profit = summary['global_profit']
test_share_revenue = summary['test_share_revenue'] / 100  # В долях
effect_revenue_pct = summary['effect_revenue_pct'] / 100  # В долях
effect_profit_pct = summary['effect_profit_pct'] / 100

# Сценарий 1: Если прирост по магазину был бы +3%
scenario1_target = 0.03  # +3%
scenario1_revenue = global_revenue * scenario1_target
scenario1_profit = global_profit * scenario1_target
print(f"   Сценарий 1: Целевой прирост магазина +3%")
print(f"     → Нужно получить: +{scenario1_revenue:,.0f} ₽ выручки, +{scenario1_profit:,.0f} ₽ прибыли")

# Сценарий 2: Если затронули в 2 раза больше товаров (по выручке)
scenario2_share = test_share_revenue * 2
scenario2_effect_revenue = global_revenue * scenario2_share * effect_revenue_pct
scenario2_effect_profit = global_profit * scenario2_share * effect_profit_pct
print(f"   Сценарий 2: Доля тестируемых товаров x2 ({scenario2_share*100:.2f}%)")
print(f"     → Эффект: +{scenario2_effect_revenue:,.0f} ₽ выручки, +{scenario2_effect_profit:,.0f} ₽ прибыли")

# Сценарий 3: Весь ассортимент, но эффект в 5 раз меньше
scenario3_effect_rev = effect_revenue_pct / 5
scenario3_effect_prof = effect_profit_pct / 5
scenario3_result_revenue = global_revenue * scenario3_effect_rev
scenario3_result_profit = global_profit * scenario3_effect_prof
print(f"   Сценарий 3: Весь ассортимент, эффект в 5 раз меньше ({scenario3_effect_rev*100:.1f}%)")
print(f"     → Эффект: +{scenario3_result_revenue:,.0f} ₽ выручки, +{scenario3_result_profit:,.0f} ₽ прибыли")

# Формируем данные для презентации
stats_data = {
    'priceChanges': price_changes,
    'effectRevenuePct': round(summary['effect_revenue_pct'], 2),
    'effectProfitPct': round(summary['effect_profit_pct'], 2),
    'absEffectRevenue': round(summary['total_abs_effect_revenue']),
    'absEffectProfit': round(summary['total_abs_effect_profit']),
    'revenueWithoutEffect': round(summary['revenue_without_effect']),
    'revenueWithEffect': round(summary['total_fact_revenue']),
    'profitWithoutEffect': round(summary['profit_without_effect']),
    'profitWithEffect': round(summary['total_fact_profit']),
    'globalRevenue': round(summary['global_revenue']),
    'globalProfit': round(summary['global_profit']),
    'testShareRevenue': round(summary['test_share_revenue'], 2),
    'testShareProfit': round(summary['test_share_profit'], 2),
    'testDurationWeeks': int(summary['test_duration_weeks']),
    'testedCount': summary['tested_count'],
    'growthCount': summary['growth_stats']['count'],
    'growthEffectRevenue': round(summary['growth_stats']['revenue_effect']),
    'declineCount': summary['decline_stats']['count'],
    'declineEffectRevenue': round(summary['decline_stats']['revenue_effect']),
    # Метаданные для презентации
    'sourceFileName': excel_path.name,
    'generatedAt': datetime.now().isoformat(),
    # Сценарии
    'scenario1': {
        'targetPct': 3,
        'targetRevenue': round(scenario1_revenue),
        'targetProfit': round(scenario1_profit),
    },
    'scenario2': {
        'sharePct': round(scenario2_share * 100, 2),
        'effectRevenue': round(scenario2_effect_revenue),
        'effectProfit': round(scenario2_effect_profit),
    },
    'scenario3': {
        'effectPct': round(scenario3_effect_rev * 100, 2),
        'effectRevenue': round(scenario3_result_revenue),
        'effectProfit': round(scenario3_result_profit),
    },
}

# Сохраняем в JSON
output_path = Path(__file__).parent / "stats_data.json"
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(stats_data, f, ensure_ascii=False, indent=2)
print(f"\nДанные сохранены в {output_path}")

print("\n" + "=" * 60)
print("JSON для презентации:")
print("=" * 60)
print(json.dumps(stats_data, ensure_ascii=False, indent=2))
