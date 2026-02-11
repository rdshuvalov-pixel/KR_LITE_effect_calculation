"""
Генерация HTML-презентации по результатам расчёта.
Формирует stats_data и presentation_data из calc/summary/results, затем HTML с встроенными данными.
"""
import json
from pathlib import Path
from datetime import datetime


def _serialize(obj):
    """Сериализация для JSON (даты, numpy)."""
    if hasattr(obj, 'isoformat'):
        return obj.isoformat()
    if hasattr(obj, 'item'):
        return obj.item()
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")


def build_stats_data(calc, results, summary, activation_params, source_file_name, calc_params=None):
    """Формирует stats_data по образцу generate_stats.py."""
    calc_params = calc_params or {}
    activation_df = calc.get_activation_details(
        threshold_pct=activation_params.get('threshold_pct', 1),
        use_rounding=activation_params.get('use_rounding', True),
        round_value=activation_params.get('round_value', 90),
        round_direction=activation_params.get('round_direction', 'up'),
        wap_from_change_date=activation_params.get('wap_from_change_date', True),
        min_days_threshold=activation_params.get('min_days_threshold', 2),
    )

    activated_pids = set()
    if not activation_df.empty:
        activated_rows = activation_df[~activation_df['Status'].str.startswith('не та цена', na=False)]
        activated_pids = set(activated_rows['product_id'].unique())

    activated_by_date = {}
    for pid in activated_pids:
        test_info = calc.test_prices[calc.test_prices['product_id'] == pid]
        if not test_info.empty:
            start_date = test_info['New_Price_Start'].min()
            if start_date not in activated_by_date:
                activated_by_date[start_date] = 0
            activated_by_date[start_date] += 1

    valid_results = results[results['Is_Excluded'] == False]
    accepted_by_date = valid_results.groupby('Test_Start_Date')['product_id'].nunique()

    price_changes = []
    for date, sent_count in summary.get('products_per_change', {}).items():
        date_str = date.strftime('%d.%m.%Y') if hasattr(date, 'strftime') else str(date)
        activated_count = int(activated_by_date.get(date, 0))
        accepted_count = int(accepted_by_date.get(date, 0))
        price_changes.append({
            'date': date_str,
            'sent': int(sent_count),
            'activated': activated_count,
            'accepted': accepted_count
        })

    global_revenue = summary['global_revenue']
    global_profit = summary['global_profit']
    test_share_revenue = summary['test_share_revenue'] / 100
    test_share_profit = summary['test_share_profit'] / 100
    effect_revenue_pct = summary['effect_revenue_pct'] / 100
    effect_profit_pct = summary['effect_profit_pct'] / 100

    scenario1_revenue = global_revenue * 0.03
    scenario1_profit = global_profit * 0.03
    scenario2_share = test_share_revenue * 2
    scenario2_effect_revenue = global_revenue * scenario2_share * effect_revenue_pct
    scenario2_effect_profit = global_profit * scenario2_share * effect_profit_pct
    scenario3_effect_rev = effect_revenue_pct / 5
    scenario3_effect_prof = effect_profit_pct / 5
    scenario3_result_revenue = global_revenue * scenario3_effect_rev
    scenario3_result_profit = global_profit * scenario3_effect_prof

    return {
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
        'sourceFileName': source_file_name,
        'generatedAt': datetime.now().isoformat(),
        'usedParams': {
            'preTestWeeks': calc_params.get('pre_test_weeks', 2),
            'preTestStockPct': calc_params.get('pre_test_threshold', 10),
            'contiguousPreTest': calc_params.get('contiguous_pre_test', True),
            'useStockFilter': calc_params.get('use_stock_filter', False),
            'stockThresholdPct': calc_params.get('stock_threshold_pct', 30),
            'activationThreshold': activation_params.get('threshold_pct', 1),
            'useRounding': activation_params.get('use_rounding', True),
            'roundValue': activation_params.get('round_value', 90),
            'roundDirection': activation_params.get('round_direction', 'up'),
            'minDaysWap': activation_params.get('min_days_threshold', 2),
            'testUseWeekValues': calc_params.get('test_use_week_values', True),
        },
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


def build_presentation_data(calc, results, pid, activation_params, use_week_values=True):
    """Формирует presentation_data для примера одной позиции."""
    prod_res = results[results['product_id'] == pid]
    if prod_res.empty:
        return None

    row = prod_res.iloc[0]
    pre_test_weeks = row.get('PreTest_Weeks', [])
    if not isinstance(pre_test_weeks, list):
        pre_test_weeks = []

    valid_test = prod_res[prod_res['Is_Excluded'] == False]
    test_weeks = sorted(valid_test['week_start'].unique().tolist())

    all_weeks = []
    for w in pre_test_weeks:
        if hasattr(w, 'strftime'):
            all_weeks.append(w)
    for w in test_weeks:
        if w not in all_weeks:
            all_weeks.append(w)
    all_weeks = sorted(all_weeks)

    prod_sales = calc.weekly_sales[calc.weekly_sales['product_id'] == pid].set_index('week_start')['revenue']
    control_data = calc.control_weekly_sales.set_index('week_start')['control_revenue']

    revenue_position = [int(prod_sales.get(w, 0)) for w in all_weeks]
    revenue_control = [int(control_data.get(w, 0)) for w in all_weeks]

    pre_n = len(pre_test_weeks)
    test_n = len(test_weeks)

    if pre_n > 0:
        pre_pos = [prod_sales.get(w, 0) for w in pre_test_weeks]
        pre_ctrl = [control_data.get(w, 0) for w in pre_test_weeks]
        pre_test_pos_avg = sum(pre_pos) / len(pre_pos)
        pre_test_ctrl_avg = sum(pre_ctrl) / len(pre_ctrl)
    else:
        pre_test_pos_avg = 0
        pre_test_ctrl_avg = 0

    if test_n > 0:
        test_pos = [prod_sales.get(w, 0) for w in test_weeks]
        test_ctrl = [control_data.get(w, 0) for w in test_weeks]
        test_pos_avg = sum(test_pos) / len(test_pos)
        test_ctrl_avg = sum(test_ctrl) / len(test_ctrl)
    else:
        test_pos_avg = 0
        test_ctrl_avg = 0

    pos_growth = (test_pos_avg / pre_test_pos_avg - 1) * 100 if pre_test_pos_avg > 0 else 0
    ctrl_growth = (test_ctrl_avg / pre_test_ctrl_avg - 1) * 100 if pre_test_ctrl_avg > 0 else 0
    effect = pos_growth - ctrl_growth

    product_name = calc.product_names.get(pid, f"Товар {pid}")
    if hasattr(product_name, 'iloc'):
        product_name = str(product_name.iloc[0]) if not product_name.empty else f"Товар {pid}"

    weeks_str = [w.strftime('%Y-%m-%d') if hasattr(w, 'strftime') else str(w) for w in all_weeks]

    return {
        'productName': str(product_name),
        'productId': int(pid),
        'weeks': weeks_str,
        'revenuePosition': revenue_position,
        'revenueControl': revenue_control,
        'preTestWeeks': pre_n,
        'testWeeks': test_n,
        'preTestPosAvg': round(pre_test_pos_avg),
        'testPosAvg': round(test_pos_avg),
        'preTestCtrlAvg': round(pre_test_ctrl_avg),
        'testCtrlAvg': round(test_ctrl_avg),
        'posGrowth': round(pos_growth, 1),
        'ctrlGrowth': round(ctrl_growth, 1),
        'effect': round(effect, 1),
    }


def generate_html(stats_data, presentation_data, template_path=None):
    """Генерирует самодостаточный HTML с встроенными данными."""
    base = Path(__file__).parent
    tpl_path = template_path or (base / "presentation.html")
    html = tpl_path.read_text(encoding='utf-8')

    pres_json = json.dumps(presentation_data or {}, ensure_ascii=False, default=_serialize)
    stats_json = json.dumps(stats_data, ensure_ascii=False, default=_serialize)
    inject_code = (
        "window.PRESENTATION_DATA=" + pres_json + ";\n"
        "    window.STATS_DATA=" + stats_json + ";\n\n    "
    )
    html = html.replace(
        "  <script>\n    function fmtWeek",
        "  <script>\n    " + inject_code + "function fmtWeek",
        1
    )
    return html


def save_presentation_and_manage_history(html_content, base_dir, max_history=3):
    """
    Сохраняет презентацию в static и ведёт FIFO-историю.
    Возвращает имя файла для ссылки.
    """
    base = Path(base_dir)
    static_dir = base / "static"
    static_dir.mkdir(exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"presentation_{ts}.html"
    filepath = static_dir / filename

    history_path = static_dir / "presentation_history.json"
    history = []
    if history_path.exists():
        try:
            history = json.loads(history_path.read_text(encoding='utf-8'))
        except Exception:
            history = []

    history.append(filename)
    while len(history) > max_history:
        old = history.pop(0)
        old_path = static_dir / old
        if old_path.exists():
            try:
                old_path.unlink()
            except OSError:
                pass

    filepath.write_text(html_content, encoding='utf-8')
    history_path.write_text(json.dumps(history, ensure_ascii=False), encoding='utf-8')

    return filename
