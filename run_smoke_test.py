#!/usr/bin/env python3
"""
Дымовой тест KeepRise Lite.
Запуск: python run_smoke_test.py [путь_к_файлу.xlsx]
Без аргумента ищет .xlsx в текущей директории.
Выводит [PASS]/[FAIL] по каждому шагу — инструмент корректен, если все PASS.
"""
import sys
from pathlib import Path

# Корень проекта
ROOT = Path(__file__).parent


def find_xlsx():
    """Ищем любой xlsx в директории."""
    for f in sorted(ROOT.glob("*.xlsx")):
        if "~$" not in f.name:  # исключаем временные
            return f
    return None


def run_smoke_test(file_path: Path) -> bool:
    """Выполняет дымовой тест. Возвращает True если все проверки пройдены."""
    all_ok = True

    # 1. Импорт
    try:
        from calculator import EffectCalculator
        print("[PASS] Импорт calculator")
    except Exception as e:
        print(f"[FAIL] Импорт calculator: {e}")
        return False

    # 2. Загрузка
    try:
        calc = EffectCalculator(file_path)
        print(f"[PASS] Загрузка файла: {file_path.name}")
    except Exception as e:
        print(f"[FAIL] Загрузка: {e}")
        return False

    # 3. Preprocess
    try:
        calc.preprocess()
        print("[PASS] Preprocess")
    except Exception as e:
        print(f"[FAIL] Preprocess: {e}")
        return False

    # 4. Проверка данных
    try:
        n_test = len(calc.test_product_ids)
        n_control = len(calc.control_product_ids)
        if n_test == 0:
            print("[WARN] Тестовых товаров: 0 (нет в Тестовых ценах или нет пересечения с Продажами)")
        else:
            print(f"[PASS] Тестовых товаров: {n_test}, контрольных: {n_control}")
    except Exception as e:
        print(f"[FAIL] Проверка данных: {e}")
        all_ok = False

    # 5. Расчёт
    try:
        results = calc.calculate(
            use_stock_filter=False,
            stock_threshold_pct=30,
            pre_test_weeks_count=2,
            pre_test_stock_threshold=10,
            contiguous_pre_test=True,
            activation_threshold=1,
            activation_use_rounding=True,
            activation_round_value=90,
            activation_round_direction="up",
            activation_wap_from_change_date=True,
            activation_min_days_threshold=2,
            test_use_week_values=True
        )
        n_rows = len(results) if hasattr(results, '__len__') else 0
        if n_rows == 0 and n_test > 0:
            print("[WARN] Расчёт выполнен, но results пуст (все товары исключены или нет pre-test)")
        else:
            print(f"[PASS] Расчёт: {n_rows} строк в results")
    except Exception as e:
        print(f"[FAIL] Расчёт: {e}")
        return False

    # 6. Summary
    try:
        summary = calc.get_summary()
        if summary is None:
            print("[WARN] get_summary вернул None (нет валидных недель)")
        else:
            required = ['total_abs_effect_revenue', 'effect_revenue_pct', 'growth_stats', 'decline_stats', 'unchanged_stats']
            missing = [k for k in required if k not in summary]
            if missing:
                print(f"[FAIL] В summary отсутствуют ключи: {missing}")
                all_ok = False
            else:
                print("[PASS] get_summary содержит все ключи")
    except Exception as e:
        print(f"[FAIL] get_summary: {e}")
        all_ok = False

    # 7. Структура summary
    if summary:
        try:
            g = summary.get('growth_stats', {})
            d = summary.get('decline_stats', {})
            u = summary.get('unchanged_stats', {})
            has_product_ids = 'product_ids' in g and 'product_ids' in d and 'product_ids' in u
            if has_product_ids:
                print("[PASS] growth/decline/unchanged содержат product_ids")
            else:
                print("[WARN] product_ids отсутствуют в stats (возможна старая версия)")
        except Exception as e:
            print(f"[FAIL] Проверка stats: {e}")

    # 8. report_generator (опционально)
    try:
        from report_generator import WordReportGenerator
        pid = list(calc.test_product_ids)[0] if calc.test_product_ids else None
        if pid and summary:
            params = {'pre_test_weeks_count': 2, 'pre_test_stock_threshold': 10}
            act_params = {'threshold_pct': 1, 'use_rounding': True, 'round_value': 90}
            gen = WordReportGenerator(calc, pid, params, results_summary=summary, activation_params=act_params)
            doc = gen.generate()
            print("[PASS] WordReportGenerator генерирует документ")
        else:
            print("[SKIP] WordReportGenerator: нет тестового товара или summary")
    except Exception as e:
        print(f"[WARN] WordReportGenerator: {e}")

    return all_ok


def main():
    if len(sys.argv) >= 2:
        file_path = Path(sys.argv[1])
        if not file_path.exists():
            print(f"[FAIL] Файл не найден: {file_path}")
            sys.exit(1)
    else:
        file_path = find_xlsx()
        if not file_path:
            print("[FAIL] Нет .xlsx в директории. Укажите путь: python run_smoke_test.py файл.xlsx")
            print("       Или поместите S-market_эффект_18012026.xlsx / САБИ_Челябинск_анализ_эффекта_26012026.xlsx в каталог.")
            sys.exit(1)

    print("=" * 50)
    print("KeepRise Lite — дымовой тест")
    print("=" * 50)

    ok = run_smoke_test(file_path)

    print("=" * 50)
    if ok:
        print("ИТОГ: [PASS] Инструмент работает корректно.")
    else:
        print("ИТОГ: [FAIL] Обнаружены проблемы. См. DEBUG_CHECKLIST.md")
    print("=" * 50)

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
