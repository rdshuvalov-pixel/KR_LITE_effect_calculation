import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from calculator import EffectCalculator
from report_generator import WordReportGenerator
from presentation_builder import (
    build_stats_data,
    build_presentation_data,
    generate_html,
    save_presentation_and_manage_history,
)
from pdf_generator import export_html_to_pdf
import io
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
from pathlib import Path

st.set_page_config(page_title="Калькулятор Эффективности Ценообразования", layout="wide")

st.title("KeepRise Lite: A/B Test Calculator")

with st.expander("ℹ️ Методология (Читать)"):
    try:
        with open("METHODOLOGY.md", "r") as f:
            st.markdown(f.read())
    except FileNotFoundError:
        st.error("Файл METHODOLOGY.md не найден.")

# --- SIDEBAR CONFIGURATION ---
with st.sidebar:
    st.header("Настройки")

    with st.form("settings_form"):
        with st.expander("Параметры Pre-Test", expanded=False):
            pre_test_weeks = st.number_input("Длительность (недель)", min_value=1, value=2, step=1)
            
            pre_test_threshold = st.number_input(
                "Допуск доступности стока (%)", 
                min_value=0, max_value=100, value=10, step=1,
                help="Процент дней в периоде, когда товар был на остатке."
            )
            
            contiguous_pre_test = st.checkbox(
                "Только непрерывный период", value=True,
                help="Если включено, ищутся 3 недели подряд. Если выключено, набираются ближайшие 3 подходящие недели."
            )

        with st.expander("Параметры Test", expanded=False):
            use_stock_filter = st.checkbox("Фильтр недель по остаткам", value=False)
            
            threshold_pct = st.number_input(
                "Порог отклонения остатка (%)", 
                min_value=0, max_value=100, value=30, step=1,
                help="Если остаток недели < (Базовый * (1 - Порог%)), неделя исключается."
            )
            
            test_calc_mode = st.radio(
                "Расчет тестового периода",
                ["Значение текущей недели", "Среднее по тестовым неделям"],
                index=0,
                help="Определяет, как считать R_tt и R_ct в тестовом периоде."
            )
    
        with st.expander("Настройки Активации Цен", expanded=False):
            activation_threshold = st.number_input(
                "Порог отклонения цены (%)", 
                min_value=0, max_value=100, value=1, step=1,
                help="Неделя считается неактивированной (Mismatch), если фактическая средняя цена отклоняется от плановой более чем на этот процент.",
                key="activation_threshold"
            )

            use_rounding = st.checkbox("Использовать округление плановой цены", value=True, key="activation_use_rounding")
            round_direction_label = st.selectbox(
                "Направление округления", 
                ["Вверх до значения", "Вниз до значения", "К ближайшему"], 
                index=0,
                key="activation_round_direction_label"
            )
            round_value = st.number_input(
                "Копейки для округления", 
                min_value=0, max_value=99, value=90, step=1,
                key="activation_round_value"
            )
            
            activation_wap_from_change_date = st.checkbox(
                "Умный расчёт WAP",
                value=True,
                help="Если включено, для недель с переоценкой WAP считается с момента первого реального вхождения цены.",
                key="activation_wap_from_change_date"
            )
            
            min_days_threshold = st.number_input(
                "Мин. дней для WAP",
                min_value=1, max_value=7, value=2, step=1,
                help="Минимальное количество дней работы цены до конца недели для учета.",
                key="activation_min_days_threshold"
            )
        
        # Form Submit Button
        submit_button = st.form_submit_button("Применить настройки")
    
    # Process inputs outside form to ensure variables exist even if not submitted yet (using defaults or session state)
    if round_direction_label.startswith("Вверх"):
        activation_round_direction = "up"
    elif round_direction_label.startswith("Вниз"):
        activation_round_direction = "down"
    else:
        activation_round_direction = "nearest"

uploaded_file = st.file_uploader("Загрузите файл Excel (S-Market.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        calc = EffectCalculator(uploaded_file)
        summary = None
        
        # Run on first upload, on file change, or after submit
        file_changed = (
            st.session_state.get('uploaded_file_name') != uploaded_file.name
            or st.session_state.get('uploaded_file_size') != uploaded_file.size
        )
        should_run = submit_button or file_changed or ('summary' not in st.session_state)
        
        if should_run:
            with st.spinner('Обработка данных и расчет...'):
                # Preprocessing
                calc.preprocess()
                
                # Calculation
                results = calc.calculate(
                    use_stock_filter=use_stock_filter, 
                    stock_threshold_pct=threshold_pct,
                    pre_test_weeks_count=pre_test_weeks,
                    pre_test_stock_threshold=pre_test_threshold,
                    contiguous_pre_test=contiguous_pre_test,
                    test_use_week_values=(test_calc_mode == "Значение текущей недели"),
                    activation_threshold=activation_threshold,
                    activation_use_rounding=use_rounding,
                    activation_round_value=round_value,
                    activation_round_direction=activation_round_direction,
                    activation_wap_from_change_date=activation_wap_from_change_date,
                    activation_min_days_threshold=min_days_threshold
                )
                summary = calc.get_summary()
                
                # Store results in session state to persist after other interactions
                st.session_state['results'] = results
                st.session_state['summary'] = summary
                st.session_state['calc_instance'] = calc  # Need instance for details
                st.session_state['uploaded_file_name'] = uploaded_file.name
                st.session_state['uploaded_file_size'] = uploaded_file.size

                # Генерация презентации под текущий расчёт (FIFO, последние 3)
                try:
                    _act_round = st.session_state.get("activation_round_direction_label", "Вверх до значения")
                    _act_dir = "up" if _act_round.startswith("Вверх") else ("down" if _act_round.startswith("Вниз") else "nearest")
                    _act_params = {
                        "threshold_pct": st.session_state.get("activation_threshold", 1),
                        "use_rounding": st.session_state.get("activation_use_rounding", True),
                        "round_value": st.session_state.get("activation_round_value", 90),
                        "round_direction": _act_dir,
                        "wap_from_change_date": st.session_state.get("activation_wap_from_change_date", True),
                        "min_days_threshold": st.session_state.get("activation_min_days_threshold", 2),
                    }
                    _calc_params = {
                        "pre_test_weeks": pre_test_weeks,
                        "pre_test_threshold": pre_test_threshold,
                        "contiguous_pre_test": contiguous_pre_test,
                        "use_stock_filter": use_stock_filter,
                        "stock_threshold_pct": threshold_pct,
                        "test_use_week_values": (test_calc_mode == "Значение текущей недели"),
                    }
                    stats_data = build_stats_data(
                        calc, results, summary, _act_params, uploaded_file.name, _calc_params
                    )
                    valid_pids = results[results["Is_Excluded"] == False]["product_id"].unique()
                    pid = int(valid_pids[0]) if len(valid_pids) > 0 else int(list(calc.test_product_ids)[0])
                    pres_data = build_presentation_data(
                        calc, results, pid, _act_params,
                        use_week_values=(test_calc_mode == "Значение текущей недели"),
                    )
                    html = generate_html(stats_data, pres_data)
                    base_dir = Path(__file__).parent
                    fname = save_presentation_and_manage_history(html, base_dir, max_history=3)
                    st.session_state["presentation_filename"] = fname
                except Exception as _e:
                    st.session_state["presentation_filename"] = None
        
        # Retrieve from session state if available and not just recalculated
        elif 'results' in st.session_state and 'summary' in st.session_state:
            results = st.session_state['results']
            summary = st.session_state['summary']
            # Update calc instance reference if needed for helper methods
            if 'calc_instance' in st.session_state:
                cached_calc = st.session_state['calc_instance']
                # Check if cached instance is outdated (missing new methods)
                if hasattr(cached_calc, 'get_wap_calculation_example'):
                    calc = cached_calc
                else:
                    # Cached instance is old, refresh it
                    with st.spinner('Обновление версии калькулятора...'):
                        calc.preprocess()
                        st.session_state['calc_instance'] = calc 
            
        if summary: 
            if submit_button:
                st.success("Расчет завершен успешно!")
            else:
                st.info("Показаны результаты предыдущего расчета.")
            
            # --- SUMMARY TABS ---
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Общие Результаты", "🔍 Анализ Товара", "💰 Анализ активации цен", "📋 Контрольная Группа", "📖 Отчет-история"])
            
            with tab1:
                st.subheader("Ключевые показатели эффективности")
                
                # --- Row 1: Revenue Metrics ---
                st.markdown("**Выручка (Revenue)**")
                r1, r2, r3, r4, r5 = st.columns(5)
                
                r1.metric("Эффект", f"{summary.get('total_abs_effect_revenue', 0):,.0f} ₽", 
                          delta=f"{summary.get('effect_revenue_pct', 0):.2f}%")
                
                r2.metric("Без эффекта (Test)", f"{summary.get('revenue_without_effect', 0):,.0f} ₽")
                r3.metric("С эффектом (Test)", f"{summary.get('total_fact_revenue', 0):,.0f} ₽")
                r4.metric("Полный оборот (Global)", f"{summary.get('global_revenue', 0):,.0f} ₽")
                r5.metric("Доля Test (%)", f"{summary.get('test_share_revenue', 0):.1f}%")
                
                st.divider()
                
                # --- Row 2: Profit Metrics ---
                st.markdown("**Прибыль (Profit)**")
                p1, p2, p3, p4, p5 = st.columns(5)
                
                p1.metric("Эффект", f"{summary.get('total_abs_effect_profit', 0):,.0f} ₽",
                          delta=f"{summary.get('effect_profit_pct', 0):.2f}%")
                
                p2.metric("Без эффекта (Test)", f"{summary.get('profit_without_effect', 0):,.0f} ₽")
                p3.metric("С эффектом (Test)", f"{summary.get('total_fact_profit', 0):,.0f} ₽")
                p4.metric("Полная прибыль (Global)", f"{summary.get('global_profit', 0):,.0f} ₽")
                p5.metric("Доля Test (%)", f"{summary.get('test_share_profit', 0):.1f}%")
                
                st.markdown("---")
                
                # --- Statistics Block ---
                st.subheader("Статистика теста")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Протестировано позиций", summary.get('tested_count', 0))
                c1.metric("Исключено позиций", summary.get('excluded_products_count', 0))
                
                c2.metric("Длительность теста", f"{summary.get('test_duration_weeks', 0):.0f} нед.")
                c2.metric("Исключено периодов (недель)", summary.get('excluded_weeks_count', 0))
                
                c3.metric("Количество переоценок", summary.get('price_changes_count', 0))
                
                # Expanders for details
                with st.expander("Детализация по переоценкам"):
                    changes_data = pd.DataFrame(
                        list(summary.get('products_per_change', {}).items()),
                        columns=['Дата старта цены', 'Кол-во товаров']
                    ).sort_values('Дата старта цены')
                    st.dataframe(changes_data, use_container_width=True)
                
                st.markdown("### Результаты по направлениям")
                g_stats = summary.get('growth_stats', {})
                d_stats = summary.get('decline_stats', {})
                
                col_g, col_d = st.columns(2)
                
                with col_g:
                    st.success(f"📈 РОСТ: {g_stats.get('count', 0)} позиций")
                    st.write(f"Эффект (Выручка): **{g_stats.get('revenue_effect', 0):,.0f} ₽**")
                    st.write(f"Эффект (Прибыль): **{g_stats.get('profit_effect', 0):,.0f} ₽**")
                    
                with col_d:
                    st.error(f"📉 ПАДЕНИЕ: {d_stats.get('count', 0)} позиций")
                    st.write(f"Эффект (Выручка): **{d_stats.get('revenue_effect', 0):,.0f} ₽**")
                    st.write(f"Эффект (Прибыль): **{d_stats.get('profit_effect', 0):,.0f} ₽**")
                
            with tab2:
                st.header("Детальный анализ товара")
                
                name_map = calc.product_names.to_dict()
                all_test_pids = sorted(list(calc.test_product_ids))
                
                effect_map = {}
                
                included_pids = set(results['product_id'].unique()) if not results.empty else set()
                excluded_pids = set(all_test_pids) - included_pids
                
                activation_threshold = st.session_state.get("activation_threshold", 10)
                activation_use_rounding = st.session_state.get("activation_use_rounding", False)
                activation_round_direction_label = st.session_state.get("activation_round_direction_label", "Вверх до значения")
                activation_round_value = st.session_state.get("activation_round_value", 90)
                activation_wap_from_change_date = st.session_state.get("activation_wap_from_change_date", True)
                activation_min_days_threshold = st.session_state.get("activation_min_days_threshold", 3)
                
                if activation_round_direction_label.startswith("Вверх"):
                    activation_round_direction = "up"
                elif activation_round_direction_label.startswith("Вниз"):
                    activation_round_direction = "down"
                else:
                    activation_round_direction = "nearest"
                
                activation_df = calc.get_activation_details(
                    activation_threshold,
                    use_rounding=activation_use_rounding,
                    round_value=activation_round_value,
                    round_direction=activation_round_direction,
                    wap_from_change_date=activation_wap_from_change_date,
                    min_days_threshold=activation_min_days_threshold
                )
                
                activation_status_map = {}
                activation_not_our_weeks = 0
                activation_not_our_products = 0
                if not activation_df.empty:
                    activation_status_map = {
                        (row['product_id'], row['week_start']): row['Status']
                        for _, row in activation_df.iterrows()
                    }
                    not_our_mask = activation_df['Status'].str.startswith('не та цена', na=False)
                    activation_not_our_weeks = int(not_our_mask.sum())
                    activation_not_our_products = activation_df[not_our_mask]['product_id'].nunique()
                
                # Logic for excluding "Not Our Price" weeks is already in calc.calculate()
                # Total_Product_Effect handles both calculation modes correctly
                
                if not results.empty:
                    if 'Total_Effect_Revenue' in results.columns:
                        # Use pre-calculated Total Effect
                        # Create mapping for both revenue and profit
                        effect_map_rev = results.drop_duplicates('product_id').set_index('product_id')['Total_Effect_Revenue'].to_dict()
                        effect_map_prof = results.drop_duplicates('product_id').set_index('product_id')['Total_Effect_Profit'].to_dict()
                    elif 'Total_Product_Effect' in results.columns:
                        effect_map_rev = results.drop_duplicates('product_id').set_index('product_id')['Total_Product_Effect'].to_dict()
                        effect_map_prof = {} # Not available
                    else:
                        # Fallback
                        col = 'Abs_Effect_Revenue' if 'Abs_Effect_Revenue' in results.columns else 'Abs_Effect'
                        effect_map_rev = results.groupby('product_id')[col].sum().to_dict()
                        effect_map_prof = {}
                        if 'Abs_Effect_Profit' in results.columns:
                            effect_map_prof = results.groupby('product_id')['Abs_Effect_Profit'].sum().to_dict()
                
                growth_pids = {pid for pid, eff in effect_map_rev.items() if eff > 0}
                decline_pids = {pid for pid, eff in effect_map_rev.items() if eff <= 0}
                
                filter_mode = st.radio("Фильтр списка:", ["Все", "Рост 📈", "Падение 📉", "Исключенные ❌"], horizontal=True)
                
                filtered_pids = all_test_pids
                if filter_mode == "Рост 📈":
                    filtered_pids = sorted(list(growth_pids))
                elif filter_mode == "Падение 📉":
                    filtered_pids = sorted(list(decline_pids))
                elif filter_mode == "Исключенные ❌":
                    filtered_pids = sorted(list(excluded_pids))
                
                if not filtered_pids:
                    st.warning("Нет товаров в выбранной категории.")
                    selected_pid = None
                else:
                    def format_func(pid):
                        name = name_map.get(pid, "Unknown")
                        if pid in included_pids:
                            eff_rev = effect_map_rev.get(pid, 0)
                            eff_prof = effect_map_prof.get(pid, 0)
                            icon = "📈" if eff_rev > 0 else "📉"
                            return f"{icon} [{pid}] {name} (Rev: {eff_rev:,.0f} ₽, Prof: {eff_prof:,.0f} ₽)"
                        else:
                            return f"❌ [{pid}] {name}"
                    
                    selected_pid = st.selectbox("Выберите товар:", filtered_pids, format_func=format_func)
                
                if selected_pid:
                    timeline = calc.get_product_timeline(selected_pid)
                    prod_name = calc.product_names.get(selected_pid, "Unknown")
                    
                    if not timeline.empty:
                        timeline['activation_status'] = timeline['week_start'].apply(
                            lambda w: activation_status_map.get((selected_pid, w), "")
                        )
                        not_our_mask = timeline['activation_status'].str.startswith('не та цена', na=False)
                        not_our_test_mask = (timeline['period_label'] == 'Test') & not_our_mask
                        timeline.loc[not_our_test_mask, 'period_label'] = 'NotOurPrice'
                        timeline.loc[not_our_test_mask, 'abs_effect_revenue'] = 0
                        timeline.loc[not_our_test_mask, 'abs_effect_profit'] = 0
                    
                    st.subheader(f"Товар: {prod_name} (ID: {selected_pid})")
                    
                    prod_res = pd.DataFrame()
                    if not results.empty:
                        prod_res = results[results['product_id'] == selected_pid]
                    
                    info_col1, info_col2, info_col3 = st.columns(3)
                    
                    if not prod_res.empty:
                        # Extract list of weeks for display
                        pre_weeks_list = prod_res['PreTest_Weeks'].iloc[0]
                        pre_start = min(pre_weeks_list)
                        pre_end = max(pre_weeks_list)
                        
                        test_start = prod_res['week_start'].min()
                        test_end = prod_res['week_start'].max()
                        
                        base_stock = prod_res['Baseline_Stock'].iloc[0]
                        
                        # Use Total_Effect_Revenue if available, else Total_Product_Effect or sum Abs_Effect
                        if 'Total_Effect_Revenue' in prod_res.columns:
                             total_effect_rev = prod_res['Total_Effect_Revenue'].iloc[0]
                             total_effect_prof = prod_res['Total_Effect_Profit'].iloc[0]
                        elif 'Total_Product_Effect' in prod_res.columns:
                            total_effect_rev = prod_res['Total_Product_Effect'].iloc[0]
                            total_effect_prof = 0 # Not available in old data
                        else:
                            total_effect_rev = prod_res['Abs_Effect'].sum()
                            total_effect_prof = 0
                        
                        # Show Period Dates (simplified: Start of first - End of last)
                        pre_end_display = pre_end + pd.Timedelta(days=6)
                        test_end_display = test_end + pd.Timedelta(days=6)
                        
                        info_col1.info(f"**Pre-Test (диапазон):**\n{pre_start.strftime('%d.%m')} - {pre_end_display.strftime('%d.%m.%Y')}")
                        info_col2.success(f"**Test:**\n{test_start.strftime('%d.%m')} - {test_end_display.strftime('%d.%m.%Y')}")
                        
                        # Aligned metrics
                        info_col3.metric("Базовый остаток", f"{base_stock:.1f}")
                        
                        m1, m2 = st.columns(2)
                        m1.metric("Суммарный эффект (Revenue)", f"{total_effect_rev:,.2f} ₽")
                        m2.metric("Суммарный эффект (Profit)", f"{total_effect_prof:,.2f} ₽")
                        
                    else:
                        info_col1.error("Товар исключен из расчета")
                        info_col2.write("Причины исключения: см. ниже")
                    
                    exclusion_reasons = []
                    if not timeline.empty:
                        if (timeline['period_label'] == 'LowStock_Before').any():
                            exclusion_reasons.append("LowStock_Before")
                        if (timeline['period_label'] == 'LowStock_Test').any():
                            exclusion_reasons.append("LowStock_Test")
                        if (timeline['period_label'] == 'NotOurPrice').any():
                            exclusion_reasons.append("не та цена")
                    
                    if exclusion_reasons:
                        st.info(f"Причины исключений: {', '.join(exclusion_reasons)}")
                    else:
                        st.info("Причины исключений: нет")

                    if not timeline.empty:
                        def highlight_timeline(row):
                            style = ['color: white'] * len(row) 
                            label = row['period_label']
                            if label == 'Pre-Test':
                                return ['background-color: #ffff99; color: black'] * len(row)
                            elif label == 'Test':
                                return ['background-color: #90ee90; color: black'] * len(row)
                            # LowStock_Before, LowStock_Test, NotOurPrice have no fill
                            return style

                        # Use Abs_Effect_Revenue if available, else Abs_Effect
                        y_col = 'abs_effect_revenue' if 'abs_effect_revenue' in timeline.columns else 'abs_effect'
                        
                        cols = ['week_formatted', 'period_label', 'activation_status', 
                                'product_revenue', 'Product_Profit', 
                                'control_revenue', 'Control_Profit',
                                'avg_stock', y_col, 'abs_effect_profit', 'has_sales']
                        
                        # Merge profits from results to timeline for display
                        if 'Product_Profit' not in timeline.columns:
                            prod_results = results[results['product_id'] == selected_pid]
                            if not prod_results.empty:
                                profit_map = prod_results.set_index('week_start')[['Product_Profit', 'Control_Profit']]
                                timeline['Product_Profit'] = timeline['week_start'].map(profit_map['Product_Profit']).fillna(0)
                                timeline['Control_Profit'] = timeline['week_start'].map(profit_map['Control_Profit']).fillna(0)
                            else:
                                timeline['Product_Profit'] = 0
                                timeline['Control_Profit'] = 0

                        # Ensure profit effect col exists
                        if 'abs_effect_profit' not in timeline.columns:
                            timeline['abs_effect_profit'] = 0
                            
                        # Format dict needs to match column names in `cols`
                        fmt_dict_tl = {
                            'product_revenue': '{:,.2f}', 
                            'Product_Profit': '{:,.2f}',
                            'control_revenue': '{:,.2f}',
                            'Control_Profit': '{:,.2f}',
                            'avg_stock': '{:,.2f}',
                        }
                        if y_col in cols:
                            fmt_dict_tl[y_col] = '{:,.2f}'
                        if 'abs_effect_profit' in cols:
                            fmt_dict_tl['abs_effect_profit'] = '{:,.2f}'
                            
                        # Rename columns for nicer display if needed, but here we just show raw col names
                        st.dataframe(timeline[cols].style.apply(highlight_timeline, axis=1).format(fmt_dict_tl))
                        
                        fig_tl = make_subplots(specs=[[{"secondary_y": True}]])
                        
                        fig_tl.add_trace(
                            go.Scatter(x=timeline['week_formatted'], y=timeline['product_revenue'], name="Выручка товара"),
                            secondary_y=False,
                        )
                        fig_tl.add_trace(
                            go.Scatter(x=timeline['week_formatted'], y=timeline['avg_stock'], name="Средний остаток", line=dict(dash='dot')),
                            secondary_y=True,
                        )
                        
                        # Mark LowStock weeks explicitly
                        excluded_points = timeline[timeline['period_label'].isin(['LowStock_Test'])]
                        if not excluded_points.empty:
                             fig_tl.add_trace(
                                go.Scatter(x=excluded_points['week_formatted'], y=excluded_points['avg_stock'], 
                                           mode='markers', marker=dict(color='red', size=10, symbol='x'),
                                           name="LowStock Test"),
                                secondary_y=True
                            )

                        fig_tl.update_layout(title_text="Динамика Выручки и Остатков")
                        fig_tl.update_yaxes(title_text="Выручка", secondary_y=False)
                        fig_tl.update_yaxes(title_text="Остаток (шт)", secondary_y=True)

                        st.plotly_chart(fig_tl, use_container_width=True)
                    else:
                        st.warning("Нет данных для отображения таймлайна.")

            with tab3:
                st.header("💰 Анализ активации цен")
                
                # Parameters moved to sidebar
                # Only display results here
                
                activation_df = calc.get_activation_details(
                    activation_threshold,
                    use_rounding=activation_use_rounding,
                    round_value=activation_round_value,
                    round_direction=activation_round_direction,
                    wap_from_change_date=activation_wap_from_change_date,
                    min_days_threshold=activation_min_days_threshold
                )
                
                if not activation_df.empty:
                    # Metrics
                    total_weeks = len(activation_df)
                    not_ok_weeks = int((~activation_df['Can_Use_In_Analysis']).sum())
                    not_ok_pct = (not_ok_weeks / total_weeks * 100) if total_weeks > 0 else 0
                    
                    products_with_not_ok = activation_df[~activation_df['Can_Use_In_Analysis']]['product_id'].nunique()

                    can_use_weeks = int(activation_df['Can_Use_In_Analysis'].sum())
                    can_use_pct = (can_use_weeks / total_weeks * 100) if total_weeks > 0 else 0
                    
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Всего проверено недель", total_weeks)
                    c2.metric("Недель не по правилам", f"{not_ok_weeks} ({not_ok_pct:.1f}%)")
                    c3.metric("Товаров с нарушениями", products_with_not_ok)
                    c4.metric("Можно брать в анализ", f"{can_use_weeks} ({can_use_pct:.1f}%)")
                    
                    st.subheader("Детализация активации цен")
                    st.info("""
                    **Легенда статусов:**
                    * **ОК (текущая, ...):** WAP попал в допуск к текущей плановой цене.
                    * **ОК (предыдущая, ...):** WAP попал в допуск к любой предыдущей плановой или последней примененной.
                    * **план. даты:** WAP считался с даты первого реального вхождения цены в допуске (умный режим).
                    * **полная нед.:** WAP считался по всей неделе (цена не найдена или мало дней).
                    * **не та цена (мимо плана):** WAP вне допуска всех плановых цен.
                    * **не та цена (нет продаж):** Продаж нет и ранее не было подтвержденной нашей цены.
                    * **не та цена (совпадение):** В первую неделю совпало с планом, но цена не менялась относительно последней фактической.
                    * **наша цена не было продаж:** Продаж нет, но ранее была продажа по нашей цене.
                    """)
                    
                    with st.expander("Методология анализа активации цен (скрыто)"):
                        st.markdown("""
                        **Как считаем WAP (факт. цену недели):**
                        1. Базово WAP считается как выручка / количество за неделю.
                        2. Если включен «умный расчет WAP» и неделя с переоценкой:
                           - ищем первое вхождение цены в допуске от плановой (начиная с плановой даты),
                           - если найдено и дней до конца недели >= порога — считаем WAP с этой даты,
                           - если цена не найдена или дней меньше порога — используем WAP за всю неделю.
                        3. Округление плановой цены применяется до сравнения, если включено.
                        
                        **Как определяем статус:**
                        - сравниваем WAP с текущей, предыдущими и последней примененной плановой ценой,
                        - используем допуск из настроек,
                        - если продаж нет, статус зависит от наличия подтвержденной нашей цены ранее.
                        """)
                    
                    # Column selection expander
                    with st.expander("Выбрать колонки для отображения"):
                        col1, col2, col3 = st.columns(3)
                        show_product_id = col1.checkbox("product_id", value=True, key="act_col_pid")
                        show_product_name = col1.checkbox("product_name", value=True, key="act_col_pname")
                        show_week = col1.checkbox("week_formatted", value=True, key="act_col_week")
                        
                        show_plan_price = col2.checkbox("Plan_Price", value=True, key="act_col_plan")
                        show_fact_price = col2.checkbox("Fact_Price", value=True, key="act_col_fact")
                        show_fact_cost = col2.checkbox("Fact_Cost", value=True, key="act_col_fact_cost")
                        show_deviation = col2.checkbox("Deviation", value=True, key="act_col_dev")
                        show_status = col2.checkbox("Status", value=True, key="act_col_status")
                        
                        show_plan_price_prev = col3.checkbox("Plan_Price_Prev", value=False, key="act_col_plan_prev")
                        show_deviation_prev = col3.checkbox("Deviation_Prev", value=False, key="act_col_dev_prev")
                        show_fact_price_prev = col3.checkbox("Fact_Price_Prev", value=True, key="act_col_fact_prev")
                        show_fact_change_pct = col3.checkbox("Fact_Change_Pct", value=True, key="act_col_fact_chg")
                        show_is_change = col3.checkbox("Is_Change", value=True, key="act_col_is_chg")
                        show_plan_unused = col3.checkbox("Plan_Price_Unused", value=True, key="act_col_unused_plan")
                        show_dev_unused = col3.checkbox("Deviation_Unused", value=True, key="act_col_unused_dev")
                    
                    def highlight_cells(x):
                        # Default: white text on dark theme background
                        df_styler = pd.DataFrame('color: white', index=x.index, columns=x.columns)
                        
                        # Row styling for weeks that cannot be used in analysis
                        mask_bad = activation_df['Can_Use_In_Analysis'] == False
                        df_styler.loc[mask_bad, :] = 'background-color: #ff6666; color: black'
                        
                        # Cell styling for Plan_Price on price change weeks (green background)
                        mask_change = activation_df['Is_Price_Change_Week'] == True
                        if 'Plan_Price' in df_styler.columns:
                            df_styler.loc[mask_change, 'Plan_Price'] = 'background-color: #66ff66; color: black; font-weight: bold'
                        
                        # Cell styling for Is_Change when fact price changed (green background)
                        mask_fact_change = activation_df['Is_Fact_Change'] == True
                        if 'Is_Change' in df_styler.columns:
                            df_styler.loc[mask_fact_change, 'Is_Change'] = 'background-color: #66ff66; color: black; font-weight: bold'

                        # Grey text for unused columns
                        for col in ['Plan_Price_Unused', 'Deviation_Unused']:
                            if col in df_styler.columns:
                                df_styler.loc[:, col] = 'color: #888888'
                        
                        return df_styler

                    # Rename columns for display
                    display_df = activation_df.copy()
                    
                    # Convert Is_Fact_Change to Yes/No
                    display_df['Is_Change'] = display_df['Is_Fact_Change'].apply(lambda x: 'Да' if x else 'Нет')
                    
                    display_df = display_df.rename(columns={
                        'Plan_Price_Current': 'Plan_Price',
                        'Plan_Price_Previous': 'Plan_Price_Prev',
                        'Deviation_From_Current_Pct': 'Deviation',
                        'Deviation_From_Previous_Pct': 'Deviation_Prev',
                        'Plan_Price_Unused': 'Plan_Price_Unused',
                        'Deviation_Unused_Pct': 'Deviation_Unused',
                        'Fact_Price_Change_Pct': 'Fact_Change_Pct'
                    })
                    
                    # Build display_cols based on checkboxes
                    display_cols = []
                    if show_product_id:
                        display_cols.append('product_id')
                    if show_product_name:
                        display_cols.append('product_name')
                    if show_week:
                        display_cols.append('week_formatted')
                    if show_plan_price:
                        display_cols.append('Plan_Price')
                    if show_fact_price:
                        display_cols.append('Fact_Price')
                    if show_deviation:
                        display_cols.append('Deviation')
                    if show_status:
                        display_cols.append('Status')
                    if show_plan_price_prev:
                        display_cols.append('Plan_Price_Prev')
                    if show_deviation_prev:
                        display_cols.append('Deviation_Prev')
                    if show_fact_price_prev:
                        display_cols.append('Fact_Price_Prev')
                    if show_fact_change_pct:
                        display_cols.append('Fact_Change_Pct')
                    if show_is_change:
                        display_cols.append('Is_Change')
                    if show_plan_unused:
                        display_cols.append('Plan_Price_Unused')
                    if show_dev_unused:
                        display_cols.append('Deviation_Unused')
                    if show_fact_cost:
                        display_cols.append('Fact_Cost')
                    
                    # Format with custom handling for None values
                    def format_optional(val):
                        if pd.isna(val) or val is None:
                            return "-"
                        return f"{val:.2f}%"
                    
                    def format_price_optional(val):
                        if pd.isna(val) or val is None:
                            return "-"
                        return f"{val:,.2f}"

                    st.dataframe(display_df[display_cols].style.apply(highlight_cells, axis=None).format({
                        'Plan_Price': '{:,.2f}',
                        'Plan_Price_Prev': format_price_optional,
                        'Fact_Price': '{:,.2f}',
                        'Fact_Price_Prev': format_price_optional,
                        'Deviation': format_optional,
                        'Deviation_Prev': format_optional,
                        'Fact_Change_Pct': format_optional,
                        'Plan_Price_Unused': format_price_optional,
                        'Deviation_Unused': format_optional,
                        'Fact_Cost': '{:,.2f}'
                    }))
                else:
                    st.info("Нет данных для анализа активации цен (нет тестовых товаров или продаж).")
                
                # --- Анализ по переоценкам ---
                st.divider()
                st.subheader("📊 Анализ активации по переоценкам")
                
                # Получаем данные анализа переоценок
                reval_summary, reval_detail = calc.analyze_revaluation_activation(
                    activation_threshold=activation_threshold,
                    use_rounding=activation_use_rounding,
                    round_value=activation_round_value,
                    round_direction=activation_round_direction,
                    wap_from_change_date=activation_wap_from_change_date,
                    min_days_threshold=activation_min_days_threshold
                )
                
                if not reval_summary.empty:
                    # Общая статистика по всем переоценкам
                    st.markdown("### Общая статистика")
                    total_revaluations = len(reval_summary)
                    total_proposed = reval_summary['total_proposed'].sum()
                    total_on_time = reval_summary['activated_on_time'].sum()
                    total_later = reval_summary['activated_later'].sum()
                    total_rejected = reval_summary['rejected'].sum()
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    col1.metric("Всего переоценок", total_revaluations)
                    col2.metric("Предложено цен", total_proposed)
                    col3.metric("Активировано вовремя", f"{total_on_time} ({total_on_time/total_proposed*100:.1f}%)" if total_proposed > 0 else "0")
                    col4.metric("Активировано позже", f"{total_later} ({total_later/total_proposed*100:.1f}%)" if total_proposed > 0 else "0")
                    col5.metric("Отклонено", f"{total_rejected} ({total_rejected/total_proposed*100:.1f}%)" if total_proposed > 0 else "0")
                    
                    # График динамики активации по переоценкам
                    st.markdown("### Динамика активации по переоценкам")
                    fig_reval = go.Figure()
                    
                    fig_reval.add_trace(go.Bar(
                        x=reval_summary['revaluation_date'],
                        y=reval_summary['activated_on_time'],
                        name='Активировано вовремя',
                        marker_color='#2ecc71'
                    ))
                    
                    fig_reval.add_trace(go.Bar(
                        x=reval_summary['revaluation_date'],
                        y=reval_summary['activated_later'],
                        name='Активировано позже',
                        marker_color='#f39c12'
                    ))
                    
                    fig_reval.add_trace(go.Bar(
                        x=reval_summary['revaluation_date'],
                        y=reval_summary['rejected'],
                        name='Отклонено',
                        marker_color='#e74c3c'
                    ))
                    
                    fig_reval.update_layout(
                        title='Распределение статусов активации по переоценкам',
                        xaxis_title='Дата переоценки',
                        yaxis_title='Количество товаров',
                        barmode='stack',
                        hovermode='x unified',
                        height=400
                    )
                    
                    st.plotly_chart(fig_reval, use_container_width=True)
                    
                    # График процента активации
                    fig_rate = go.Figure()
                    
                    fig_rate.add_trace(go.Scatter(
                        x=reval_summary['revaluation_date'],
                        y=reval_summary['activation_rate_on_time'],
                        mode='lines+markers',
                        name='% активации вовремя',
                        line=dict(color='#2ecc71', width=3),
                        marker=dict(size=8)
                    ))
                    
                    fig_rate.add_trace(go.Scatter(
                        x=reval_summary['revaluation_date'],
                        y=reval_summary['activation_rate_total'],
                        mode='lines+markers',
                        name='% активации всего',
                        line=dict(color='#3498db', width=3),
                        marker=dict(size=8)
                    ))
                    
                    fig_rate.update_layout(
                        title='Процент активации цен по переоценкам',
                        xaxis_title='Дата переоценки',
                        yaxis_title='Процент (%)',
                        hovermode='x unified',
                        height=400,
                        yaxis=dict(range=[0, 100])
                    )
                    
                    st.plotly_chart(fig_rate, use_container_width=True)
                    
                    # Выбор конкретной переоценки для детального анализа
                    st.markdown("### Детальный анализ переоценки")
                    
                    # Создаем список для выбора
                    reval_options = []
                    for i, (idx, row) in enumerate(reval_summary.iterrows()):
                        label = f"{row['revaluation_date'].strftime('%d.%m.%Y')} - {row['planned_week_formatted']} (Предложено: {row['total_proposed']})"
                        reval_options.append((i, label, row['revaluation_date']))
                    
                    if reval_options:
                        selected_reval_idx = st.selectbox(
                            "Выберите переоценку для детального анализа:",
                            options=range(len(reval_options)),
                            format_func=lambda x: reval_options[x][1]
                        )
                        
                        selected_reval = reval_summary.iloc[selected_reval_idx]
                        selected_date = reval_options[selected_reval_idx][2]
                        
                        # Фильтруем детальные данные по выбранной переоценке
                        detail_filtered = reval_detail[reval_detail['revaluation_date'] == selected_date].copy()
                        
                        if not detail_filtered.empty:
                            # Метрики для выбранной переоценки
                            st.markdown(f"#### Переоценка от {selected_date.strftime('%d.%m.%Y')}")
                            
                            col1, col2, col3, col4 = st.columns(4)
                            col1.metric("Предложено", int(selected_reval['total_proposed']))
                            col2.metric("Активировано вовремя", int(selected_reval['activated_on_time']))
                            col3.metric("Активировано позже", int(selected_reval['activated_later']))
                            col4.metric("Отклонено", int(selected_reval['rejected']))
                            
                            # Круговая диаграмма распределения
                            st.markdown("##### Распределение статусов")
                            fig_pie = go.Figure(data=[go.Pie(
                                labels=['Активировано вовремя', 'Активировано позже', 'Отклонено'],
                                values=[
                                    selected_reval['activated_on_time'],
                                    selected_reval['activated_later'],
                                    selected_reval['rejected']
                                ],
                                marker_colors=['#2ecc71', '#f39c12', '#e74c3c']
                            )])
                            
                            fig_pie.update_layout(height=400)
                            st.plotly_chart(fig_pie, use_container_width=True)
                            
                            # Детальная таблица по товарам
                            st.markdown("##### Детализация по товарам")
                            st.info("""
                            **Примечание:** Данные в таблице используют ту же логику анализа активации, что и основная таблица выше. 
                            Статусы и показатели рассчитываются на основе тех же критериев (порог отклонения, округление, умный WAP).
                            
                            **Колонки:**
                            - **Плановая цена / Фактическая цена** - значения с недели активации (или плановой недели для отклоненных)
                            - **Отклонение от плана, %** - процент отклонения фактической цены от плановой
                            - **Статус из анализа** - точный статус из анализа активации (совпадает с основной таблицей)
                            - Для товаров, активированных позже, также показываются данные с плановой недели для сравнения
                            """)
                            
                            # Переименовываем колонки для отображения
                            # Собираем все доступные колонки
                            available_cols = ['product_id', 'product_name', 'status', 'status_reason', 'activation_week']
                            price_cols = ['plan_price', 'fact_price', 'deviation_pct', 'activation_status', 'fact_cost']
                            planned_week_cols = ['plan_price_planned_week', 'fact_price_planned_week', 
                                                'deviation_pct_planned_week', 'activation_status_planned_week']
                            
                            # Проверяем, какие колонки есть в данных
                            cols_to_show = [col for col in available_cols + price_cols + planned_week_cols if col in detail_filtered.columns]
                            
                            display_detail = detail_filtered[cols_to_show].copy()
                            
                            # Переименовываем колонки
                            rename_map = {
                                'product_id': 'ID товара',
                                'product_name': 'Название товара',
                                'status': 'Статус активации',
                                'status_reason': 'Причина/Детали',
                                'activation_week': 'Неделя активации',
                                'plan_price': 'Плановая цена',
                                'fact_price': 'Фактическая цена',
                                'deviation_pct': 'Отклонение от плана, %',
                                'activation_status': 'Статус из анализа',
                                'fact_cost': 'Себестоимость',
                                'plan_price_planned_week': 'План. цена (план. нед.)',
                                'fact_price_planned_week': 'Факт. цена (план. нед.)',
                                'deviation_pct_planned_week': 'Отклонение (план. нед.), %',
                                'activation_status_planned_week': 'Статус (план. нед.)'
                            }
                            
                            display_detail = display_detail.rename(columns=rename_map)
                            
                            # Форматируем неделю активации
                            def format_week_display(week):
                                if pd.isna(week) or week is None:
                                    return "-"
                                week_str = pd.to_datetime(week)
                                week_end = week_str + pd.Timedelta(days=6)
                                return f"{week_str.strftime('%d.%m.%Y')} - {week_end.strftime('%d.%m.%Y')}"
                            
                            if 'Неделя активации' in display_detail.columns:
                                display_detail['Неделя активации'] = display_detail['Неделя активации'].apply(format_week_display)
                            
                            # Переводим статусы на русский
                            status_map = {
                                'activated_on_time': 'Активировано вовремя',
                                'activated_later': 'Активировано позже',
                                'rejected': 'Отклонено'
                            }
                            if 'Статус активации' in display_detail.columns:
                                display_detail['Статус активации'] = display_detail['Статус активации'].map(status_map)
                            
                            # Форматируем числовые значения
                            def format_price(val):
                                if pd.isna(val) or val is None:
                                    return "-"
                                return f"{val:,.2f}"
                            
                            def format_pct(val):
                                if pd.isna(val) or val is None:
                                    return "-"
                                return f"{val:.2f}%"
                            
                            # Применяем форматирование
                            price_cols_display = ['Плановая цена', 'Фактическая цена', 'Себестоимость', 
                                                  'План. цена (план. нед.)', 'Факт. цена (план. нед.)']
                            pct_cols_display = ['Отклонение от плана, %', 'Отклонение (план. нед.), %']
                            
                            # Переупорядочиваем колонки для лучшей читаемости
                            # Сначала основные, потом данные с плановой недели (если есть)
                            col_order = ['ID товара', 'Название товара', 'Статус активации', 'Неделя активации',
                                        'Плановая цена', 'Фактическая цена', 'Отклонение от плана, %', 
                                        'Статус из анализа', 'Себестоимость']
                            
                            # Добавляем колонки с плановой недели, если они есть
                            planned_week_cols = ['План. цена (план. нед.)', 'Факт. цена (план. нед.)', 
                                                'Отклонение (план. нед.), %', 'Статус (план. нед.)']
                            for col in planned_week_cols:
                                if col in display_detail.columns:
                                    col_order.append(col)
                            
                            # Добавляем оставшиеся колонки
                            for col in display_detail.columns:
                                if col not in col_order:
                                    col_order.append(col)
                            
                            # Переупорядочиваем только существующие колонки
                            col_order = [col for col in col_order if col in display_detail.columns]
                            if col_order:
                                display_detail = display_detail[col_order]
                            
                            for col in price_cols_display:
                                if col in display_detail.columns:
                                    display_detail[col] = display_detail[col].apply(format_price)
                            
                            for col in pct_cols_display:
                                if col in display_detail.columns:
                                    display_detail[col] = display_detail[col].apply(format_pct)
                            
                            # Применяем цветовую подсветку
                            def highlight_status_row(row):
                                colors = {
                                    'Активировано вовремя': 'background-color: #d4edda; color: black',
                                    'Активировано позже': 'background-color: #fff3cd; color: black',
                                    'Отклонено': 'background-color: #f8d7da; color: black'
                                }
                                status = row.get('Статус активации', '')
                                return [colors.get(status, '')] * len(row)
                            
                            # Определяем формат для числовых колонок
                            format_dict = {}
                            for col in price_cols_display:
                                if col in display_detail.columns:
                                    format_dict[col] = lambda x: x  # Уже отформатировано
                            for col in pct_cols_display:
                                if col in display_detail.columns:
                                    format_dict[col] = lambda x: x  # Уже отформатировано
                            
                            st.dataframe(
                                display_detail.style.apply(highlight_status_row, axis=1),
                                use_container_width=True,
                                hide_index=True
                            )
                        else:
                            st.warning("Нет детальных данных для выбранной переоценки.")
                else:
                    st.info("Нет данных о переоценках для анализа.")

            with tab4:
                st.header("Состав Контрольной Группы")
                control_df = calc.get_control_group_info()
                st.dataframe(control_df)

            with tab5:
                st.header("Отчет о тестировании")
                
                # --- 1. Select Product ---
                # Get list of processed products
                processed_pids = sorted(list(results['product_id'].unique())) if not results.empty else []
                
                if not processed_pids:
                    st.warning("Нет данных для анализа.")
                else:
                    # Helper to format dropdown options
                    def report_format_func(pid):
                        name = calc.product_names.get(pid, "Unknown")
                        return f"[{pid}] {name}"
                    
                    # Try to find a good default (positive effect)
                    default_idx = 0
                    if not results.empty:
                        col_eff = 'Total_Effect_Revenue' if 'Total_Effect_Revenue' in results.columns else 'Abs_Effect_Revenue'
                        if col_eff not in results.columns: col_eff = 'Abs_Effect'
                        best_pid = results.sort_values(col_eff, ascending=False).iloc[0]['product_id']
                        if best_pid in processed_pids:
                            default_idx = processed_pids.index(best_pid)

                    selected_report_pid = st.selectbox(
                        "Выберите товар для анализа:", 
                        processed_pids, 
                        index=default_idx,
                        format_func=report_format_func,
                        key="report_hero_select"
                    )
                    
                    if selected_report_pid:
                        hero_name = calc.product_names.get(selected_report_pid, f"ID {selected_report_pid}")
                        st.subheader(f"Анализируемый товар: {hero_name}")
                        
                        hero_results = results[results['product_id'] == selected_report_pid] if not results.empty else pd.DataFrame()
                        timeline = calc.get_product_timeline(selected_report_pid)
                        
                        # --- 2. Input Data: Price Changes ---
                        st.markdown("### 1. Вводные данные: Переоценки")
                        
                        # Get price changes from raw data (calculator has self.test_prices)
                        # We need to filter by product_id
                        hero_prices = calc.test_prices[calc.test_prices['product_id'] == selected_report_pid].copy()
                        
                        if not hero_prices.empty:
                            hero_prices = hero_prices.sort_values('New_Price_Start')
                            
                            # --- Analysis Metrics ---
                            changes_count = len(hero_prices)
                            
                            # Frequency (Avg days between changes)
                            freq_str = "-"
                            if changes_count > 1:
                                days_between = (hero_prices['New_Price_Start'].max() - hero_prices['New_Price_Start'].min()).days
                                freq = days_between / (changes_count - 1)
                                freq_str = f"{freq:.1f} дн."
                            
                            # Average Change % (relative to current price)
                            hero_prices['Change_Pct'] = 0.0
                            mask_valid = hero_prices['Current_Price'] > 0
                            hero_prices.loc[mask_valid, 'Change_Pct'] = (
                                (hero_prices.loc[mask_valid, 'New_Price'] - hero_prices.loc[mask_valid, 'Current_Price']) /
                                hero_prices.loc[mask_valid, 'Current_Price'] * 100
                            )
                            
                            avg_change_pct = hero_prices['Change_Pct'].abs().mean()
                            
                            # Avg Change Rub (relative to current price)
                            diffs_rub = (hero_prices['New_Price'] - hero_prices['Current_Price']).abs()
                            avg_change_rub = diffs_rub.mean()
                            
                            # Duration
                            first_date = hero_prices['New_Price_Start'].min()
                            # Find max date in sales for this product
                            last_sale_date = calc.sales[calc.sales['product_id'] == selected_report_pid]['recorded_on'].max()
                            
                            duration_days = 0
                            if pd.notnull(last_sale_date) and last_sale_date > first_date:
                                duration_days = (last_sale_date - first_date).days
                            
                            st.info(f"""
                            **Анализ переоценок:**
                            - Количество переоценок: **{changes_count}**
                            - Частота переоценок: **{freq_str}**
                            - Среднее изменение цены: **{avg_change_rub:.0f} ₽** ({avg_change_pct:.2f}%)
                            - Старт теста: **{first_date:%d.%m.%Y}** (продолжительность: **{duration_days}** дн.)
                            """)

                            st.markdown("Список запланированных изменений цен:")
                            
                            # Prepare display table
                            price_display = hero_prices[['New_Price_Start', 'New_Price', 'Current_Price', 'Change_Pct']].copy()
                            price_display.columns = ['Дата старта', 'Новая цена', 'Текущая цена', 'Изменение %']
                            
                            st.dataframe(
                                price_display.style
                                .set_properties(subset=['Новая цена'], **{'background-color': '#90ee90', 'color': 'black', 'font-weight': 'bold'})
                                .format({
                                    'Дата старта': '{:%d.%m.%Y}',
                                    'Новая цена': '{:,.2f}',
                                    'Текущая цена': '{:,.2f}',
                                    'Изменение %': '{:+.2f}%'
                                }),
                                use_container_width=True
                            )
                        else:
                            st.warning("Нет данных о переоценках для этого товара.")

                        # --- 2. Input Data: Raw Sales (moved before Weekly) ---
                        st.markdown("### 2. Вводные данные: Продажи (по дням)")
                        st.markdown("Детальные данные о продажах с расчетом выручки и прибыли.")
                        
                        # Get raw sales for this product
                        hero_sales = calc.sales[calc.sales['product_id'] == selected_report_pid].copy()
                        
                        if not hero_sales.empty:
                            hero_sales = hero_sales.sort_values('recorded_on')
                            
                            if 'cost_volume' not in hero_sales.columns:
                                hero_sales['cost_volume'] = 0 
                                
                            hero_sales['profit'] = hero_sales['revenue'] - hero_sales['cost_volume']
                            
                            start_date = hero_prices['New_Price_Start'].min() if not hero_prices.empty else None
                            
                            def highlight_sales_row(row):
                                if start_date and row['Дата'] >= start_date:
                                    return ['background-color: #90ee90; color: black'] * len(row)
                                return [''] * len(row)
                            
                            cols_show = ['recorded_on', 'price', 'quantity', 'revenue', 'cost_at_sale', 'profit']
                            cols_rename = {
                                'recorded_on': 'Дата',
                                'price': 'Цена продажи',
                                'quantity': 'Кол-во (шт)',
                                'revenue': 'Выручка',
                                'cost_at_sale': 'Себест. ед.',
                                'profit': 'Прибыль'
                            }
                            
                            display_sales = hero_sales[cols_show].rename(columns=cols_rename)
                            
                            st.dataframe(display_sales.style.apply(highlight_sales_row, axis=1).format({
                                'Дата': '{:%d.%m.%Y}',
                                'Цена продажи': '{:,.2f}',
                                'Кол-во (шт)': '{:.0f}',
                                'Выручка': '{:,.2f}',
                                'Себест. ед.': '{:,.2f}',
                                'Прибыль': '{:,.2f}'
                            }), use_container_width=True)
                            
                            st.caption("Зеленая заливка — период действия тестовых цен (начиная с первой даты переоценки).")
                        else:
                            st.warning("Продаж по данному товару не найдено.")

                        # --- 3. Input Data: Weekly Report Data ---
                        st.markdown("### 3. Вводные данные: Понедельные показатели")
                        
                        # Retrieve params from session state or defaults
                        activation_threshold = st.session_state.get("activation_threshold", 10)
                        activation_use_rounding = st.session_state.get("activation_use_rounding", False)
                        activation_round_direction_label = st.session_state.get("activation_round_direction_label", "Вверх до значения")
                        activation_round_value = st.session_state.get("activation_round_value", 90)
                        activation_wap_from_change_date = st.session_state.get("activation_wap_from_change_date", True)
                        activation_min_days_threshold = st.session_state.get("activation_min_days_threshold", 3)
                        
                        if activation_round_direction_label.startswith("Вверх"):
                            activation_round_direction = "up"
                        elif activation_round_direction_label.startswith("Вниз"):
                            activation_round_direction = "down"
                        else:
                            activation_round_direction = "nearest"

                        activation_params = {
                            "threshold_pct": activation_threshold,
                            "use_rounding": activation_use_rounding,
                            "round_value": activation_round_value,
                            "round_direction": activation_round_direction,
                            "wap_from_change_date": activation_wap_from_change_date,
                            "min_days_threshold": activation_min_days_threshold
                        }

                        weekly_df = calc.get_product_weekly_report_data(selected_report_pid, activation_params)
                        
                        if not weekly_df.empty:
                            # Formatting helper
                            def fmt_price(x):
                                return f"{x:,.2f}" if pd.notnull(x) else "-"

                            # Styling function
                            def highlight_weekly_row(row):
                                if row.get('is_test_period', False):
                                    return ['background-color: #90ee90; color: black'] * len(row)
                                return [''] * len(row)

                            cols_to_show = [
                                'product_id', 'product_name', 'week_formatted',
                                'Plan_price', 'Fact_price', 'Cost_price',
                                'Суммарная выручка', 'Суммарная прибыль'
                            ]
                            
                            # Prepare display dataframe with Russian names
                            display_df = weekly_df[cols_to_show].copy()
                            display_df.columns = [
                                'ID товара', 'Наименование', 'Неделя',
                                'План. цена', 'Факт. цена', 'Себест.',
                                'Выручка (нед.)', 'Прибыль (нед.)'
                            ]
                            # Add hidden column for styling
                            display_df['is_test_period'] = weekly_df['is_test_period'].values
                            
                            def highlight_display_row(row):
                                # Apply style to all visible columns if test period
                                style = [''] * len(row)
                                if row['is_test_period']:
                                    style = ['background-color: #90ee90; color: black'] * len(row)
                                return style

                            st.dataframe(
                                display_df.style
                                .apply(highlight_display_row, axis=1)
                                .format({
                                    'План. цена': fmt_price,
                                    'Факт. цена': fmt_price,
                                    'Себест.': fmt_price,
                                    'Выручка (нед.)': '{:,.2f}',
                                    'Прибыль (нед.)': '{:,.2f}'
                                })
                                .hide(subset=['is_test_period'], axis="columns"), 
                                use_container_width=True
                            )

                            st.markdown("---")
                            st.markdown("#### Детализация расчетов по неделям")
                            
                            # --- Iterate through ALL Test Weeks ---
                            # Sort by date
                            weekly_df_sorted = weekly_df.sort_values('week_start')
                            test_weeks_df = weekly_df_sorted[weekly_df_sorted['is_test_period'] == True]
                            
                            if not test_weeks_df.empty:
                                for _, row in test_weeks_df.iterrows():
                                    week_start = row['week_start']
                                    week_fmt = row['week_formatted']
                                    
                                    details = calc.get_weekly_details(selected_report_pid, week_start, activation_params)
                                    
                                    with st.expander(f"Детализация расчета: Неделя {week_fmt}"):
                                        # Show details even if no sales, to show structure with zeros
                                        # But keep check for display logic
                                        
                                        st.markdown("**1. Плановая цена**")
                                        st.markdown(details['plan_text'])
                                        
                                        # Helper function for highlighting Total row cells, greying unused, and bolding used
                                        # Use context to differentiate between 'Fact Price' (subset) and others (full sales)
                                        def highlight_cells(row, target_col, context='full'):
                                            styles = [''] * len(row)
                                            
                                            is_total = row['Дата'] == 'Итого'
                                            is_used_price = row.get('is_used', False) # Used for price calculation
                                            
                                            # Determine if row is "active" based on context
                                            # For 'price' context: active if is_used=True
                                            # For 'full' context (Rev/Cost/Profit): active if Quantity > 0 (real sale)
                                            
                                            # Check quantity safely
                                            qty = 0
                                            if 'Кол-во' in row:
                                                try:
                                                    qty = float(row['Кол-во'])
                                                except:
                                                    qty = 0
                                            
                                            is_active_row = False
                                            if context == 'price':
                                                is_active_row = is_used_price
                                            else:
                                                is_active_row = qty > 0

                                            # Grey out inactive rows (except Total)
                                            if not is_total and not is_active_row:
                                                return ['color: #aaaaaa'] * len(row)
                                            
                                            # Bold active rows
                                            if not is_total and is_active_row:
                                                styles = ['font-weight: bold'] * len(row)

                                            # Highlight target in Total row
                                            if is_total:
                                                if target_col in row.index:
                                                    idx = row.index.get_loc(target_col)
                                                    styles[idx] = 'color: green; font-weight: bold'
                                            
                                            return styles

                                        st.markdown("**2. Фактическая цена**")
                                        st.write("Транзакции, вошедшие в расчет цены:")
                                        # Use standard pandas formatting via styler, handling Total row
                                        # Convert mixed types to string for formatting where needed
                                        st.dataframe(details['fact_transactions'].style
                                            .format({
                                                'Цена': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Кол-во': lambda x: f"{x:.0f}" if isinstance(x, (int, float)) else str(x),
                                                'Дата': lambda x: f"{x:%d.%m.%Y}" if isinstance(x, (pd.Timestamp, datetime.date)) else str(x)
                                            })
                                            .apply(highlight_cells, target_col='Цена', context='price', axis=1)
                                            .hide(subset=['is_used'], axis="columns"), 
                                            use_container_width=True
                                        )
                                        st.markdown(details['fact_text'])
                                        
                                        # Full Week Data for subsequent blocks
                                        full_df = details['full_transactions']
                                        
                                        st.markdown("**3. Выручка**")
                                        st.write("Исходные данные с расчетом выручки (Цена * Кол-во):")
                                        st.dataframe(full_df[['Дата', 'День недели', 'Цена', 'Кол-во', 'Выручка', 'is_used']].style
                                            .format({
                                                'Цена': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Кол-во': lambda x: f"{x:.0f}" if isinstance(x, (int, float)) else str(x),
                                                'Выручка': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Дата': lambda x: f"{x:%d.%m.%Y}" if isinstance(x, (pd.Timestamp, datetime.date)) else str(x)
                                            })
                                            .apply(highlight_cells, target_col='Выручка', context='full', axis=1)
                                            .hide(subset=['is_used'], axis="columns"),
                                            use_container_width=True
                                        )
                                        st.markdown(details['revenue_text'])
                                        
                                        st.markdown("**4. Себестоимость**")
                                        st.write("История изменения себестоимости за период продаж:")
                                        if 'cost_source_data' in details and not details['cost_source_data'].empty:
                                            st.dataframe(details['cost_source_data'].style.format({
                                                'Период с': lambda x: f"{x:%d.%m.%Y}" if isinstance(x, (pd.Timestamp, datetime.date)) else str(x),
                                                'Период по': lambda x: f"{x:%d.%m.%Y}" if isinstance(x, (pd.Timestamp, datetime.date)) else str(x),
                                                'Себестоимость': '{:,.2f}'
                                            }), use_container_width=True)
                                        else:
                                            st.write("Нет данных о себестоимости вблизи периода продаж.")
                                            
                                        st.write("Расчет средней себестоимости по транзакциям:")
                                        st.dataframe(full_df[['Дата', 'День недели', 'Цена', 'Кол-во', 'Выручка', 'Себестоимость ед.', 'is_used']].style
                                            .format({
                                                'Цена': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Кол-во': lambda x: f"{x:.0f}" if isinstance(x, (int, float)) else str(x),
                                                'Выручка': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Себестоимость ед.': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Дата': lambda x: f"{x:%d.%m.%Y}" if isinstance(x, (pd.Timestamp, datetime.date)) else str(x)
                                            })
                                            .apply(highlight_cells, target_col='Себестоимость ед.', context='full', axis=1)
                                            .hide(subset=['is_used'], axis="columns"),
                                            use_container_width=True
                                        )
                                        st.markdown(details['cost_text'])
                                        
                                        st.markdown("**5. Прибыль**")
                                        st.write("Исходные данные с расчетом прибыли (Выручка - (Себ.ед * Кол-во)):")
                                        st.dataframe(full_df[['Дата', 'День недели', 'Цена', 'Кол-во', 'Выручка', 'Себестоимость ед.', 'Прибыль', 'is_used']].style
                                            .format({
                                                'Цена': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Кол-во': lambda x: f"{x:.0f}" if isinstance(x, (int, float)) else str(x),
                                                'Выручка': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Себестоимость ед.': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Прибыль': lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else str(x),
                                                'Дата': lambda x: f"{x:%d.%m.%Y}" if isinstance(x, (pd.Timestamp, datetime.date)) else str(x)
                                            })
                                            .apply(highlight_cells, target_col='Прибыль', context='full', axis=1)
                                            .hide(subset=['is_used'], axis="columns"),
                                            use_container_width=True
                                        )
                                        st.markdown(details['profit_text'])
                                        
                                        st.markdown("**6. Проверка недели на активацию цен**")
                                        st.markdown(details['activation_text'])
                            else:
                                st.info("Нет тестовых недель для отображения детализации.")

                            # --- 4. Pre-Test Period Selection ---
                            st.markdown("### 4. Выбор дотестового периода")
                            
                            pre_test_params = {
                                'pre_test_weeks_count': pre_test_weeks,
                                'pre_test_stock_threshold': pre_test_threshold,
                                'contiguous_pre_test': contiguous_pre_test
                            }
                            
                            contiguous_str = "подряд идущие" if contiguous_pre_test else "не обязательно подряд идущие"
                            
                            st.write(f"""
                            Наша задача выбрать ближайшие **{pre_test_weeks}** дотестовых полных недель ({contiguous_str}).
                            При этом важно, чтобы это были недели, когда товар был на остатке более **{pre_test_threshold}%** времени.
                            """)
                            
                            report_min = weekly_df['week_start'].min() if not weekly_df.empty else None
                            report_max = weekly_df['week_start'].max() if not weekly_df.empty else None
                            
                            pre_test_details = calc.get_pre_test_selection_details(
                                selected_report_pid,
                                pre_test_params,
                                report_min=report_min,
                                report_max=report_max
                            )
                            
                            if pre_test_details is not None:
                                pre_test_df = pre_test_details.get('df')
                                if pre_test_df is not None and not pre_test_df.empty:
                                    def highlight_selected(row):
                                        if row['Выбрана']:
                                            return ['background-color: #90ee90; color: black; font-weight: bold'] * len(row)
                                        return ['color: #888888'] * len(row)

                                    st.dataframe(
                                        pre_test_df.style
                                        .apply(highlight_selected, axis=1)
                                        .format({
                                            'Дней на остатке': '{:.0f}',
                                            'Доступность %': '{:.1f}%',
                                            'Средний остаток': '{:,.2f}'
                                        }),
                                        use_container_width=True
                                    )
                                    
                                    selected_stocks = pre_test_details.get('selected_stocks', [])
                                    base_stock_val = pre_test_details.get('base_stock', 0)
                                    
                                    if selected_stocks:
                                        stocks_formula = " + ".join([f"{val:,.2f}" for val in selected_stocks])
                                        st.markdown(f"Базовый остаток за дотестовый период: **{base_stock_val:,.2f}**")
                                        st.markdown(f"Формула: ({stocks_formula}) / {len(selected_stocks)} = {base_stock_val:,.2f}")
                                    else:
                                        st.info("Базовый остаток не рассчитан: не найдено подходящих дотестовых недель.")
                                else:
                                    st.warning("Не удалось определить дотестовый период (возможно, нет данных по стокам).")
                            else:
                                st.warning("Не удалось определить дотестовый период (возможно, нет данных по стокам).")

                            # --- 5. Final Selection of Test Periods ---
                            st.markdown("### 5. Финальный выбор тестовых периодов")
                            
                            st.write(f"""
                            На прошлых этапах мы выбрали периоды с корректной активацией цены. 
                            Однако, для финального замера эффекта нам нужно исключить:
                            1. **Транзитные недели** — недели, в которых произошла смена цены (так как они содержат продажи и по старой, и по новой цене, и эффект размывается).
                            2. **Недели с низким стоком** (если включено) — если остаток товара был ниже порогового значения от базового ({threshold_pct}% от {base_stock_val:.2f}), неделя исключается для чистоты эксперимента.
                            """)
                            
                            # Calculate timelines
                            timeline = calc.get_product_timeline(selected_report_pid)
                            
                            if not timeline.empty:
                                # Enrich status similar to Product Analysis tab
                                timeline['activation_status'] = timeline['week_start'].apply(
                                    lambda w: activation_status_map.get((selected_report_pid, w), "")
                                )
                                not_our_mask = timeline['activation_status'].str.startswith('не та цена', na=False)
                                not_our_test_mask = (timeline['period_label'] == 'Test') & not_our_mask
                                timeline.loc[not_our_test_mask, 'period_label'] = 'NotOurPrice'
                                
                                # Create display DF
                                final_selection_df = timeline.copy()
                                
                                # Define Status Translation
                                status_trans = {
                                    'Pre-Test': 'Дотестовый (База)',
                                    'Test': 'Тестовый (Включен)',
                                    'LowStock_Test': 'Исключен (Мало стока)',
                                    'LowStock_Before': 'Исключен (Мало стока до)',
                                    'Transit': 'Исключен (Транзитная)',
                                    'NotOurPrice': 'Исключен (Не та цена)',
                                    'Other': 'Другое'
                                }
                                final_selection_df['Статус'] = final_selection_df['period_label'].map(status_trans).fillna('Другое')
                                final_selection_df['Включена в тест'] = final_selection_df['period_label'] == 'Test'
                                final_selection_df['Включена в базу'] = final_selection_df['period_label'] == 'Pre-Test'
                                
                                cols_final = ['week_formatted', 'avg_stock', 'Статус', 'Включена в тест', 'Включена в базу']
                                display_final = final_selection_df[cols_final].rename(columns={
                                    'week_formatted': 'Неделя',
                                    'avg_stock': 'Средний остаток'
                                })
                                
                                def highlight_final_rows(row):
                                    style = [''] * len(row)
                                    status = row['Статус']
                                    
                                    if 'Тестовый (Включен)' in status:
                                        return ['background-color: #90ee90; color: black; font-weight: bold'] * len(row)
                                    elif 'Дотестовый (База)' in status:
                                        return ['background-color: #ffff99; color: black; font-weight: bold'] * len(row)
                                    elif 'Исключен' in status:
                                        return ['color: #888888'] * len(row)
                                    return style

                                st.dataframe(
                                    display_final.style
                                    .apply(highlight_final_rows, axis=1)
                                    .format({'Средний остаток': '{:,.2f}'}),
                                    use_container_width=True
                                )
                            else:
                                st.warning("Нет данных для формирования таймлайна.")
                                
                            # --- 6. Calculation of Effect ---
                            st.markdown("### 6. Расчет эффекта")
                            
                            calc_mode_desc = "по каждой неделе" if (test_calc_mode == "Значение текущей недели") else "по среднему"
                            
                            st.write(f"""
                            Мы выбрали тестовый и контрольный период (дотестовый) и посчитали показатели выручки и прибыли для тестовых периодов.
                            
                            Режим расчета: **{calc_mode_desc}**.
                            """)
                            
                            use_week_values = (test_calc_mode == "Значение текущей недели")
                            effect_details_df = calc.get_simple_effect_details(selected_report_pid, use_week_values=use_week_values)
                            
                            if not effect_details_df.empty:
                                # Summary Table
                                summary_cols = [
                                    'week_formatted', 
                                    'Fact_Revenue_Real', 'PreTest_Avg_Revenue', 
                                    'Uplift_Revenue_Pct',
                                    'Fact_Profit_Real', 'PreTest_Avg_Profit',
                                    'Uplift_Profit_Pct'
                                ]
                                
                                display_effect = effect_details_df[summary_cols].rename(columns={
                                    'week_formatted': 'Неделя',
                                    'Fact_Revenue_Real': 'Выручка (Test)',
                                    'PreTest_Avg_Revenue': 'Выручка (Pre-Test Avg)',
                                    'Uplift_Revenue_Pct': 'Прирост % (Rev)',
                                    'Fact_Profit_Real': 'Прибыль (Test)',
                                    'PreTest_Avg_Profit': 'Прибыль (Pre-Test Avg)',
                                    'Uplift_Profit_Pct': 'Прирост % (Prof)'
                                })
                                
                                st.dataframe(
                                    display_effect.style.format({
                                        'Выручка (Test)': '{:,.2f}',
                                        'Выручка (Pre-Test Avg)': '{:,.2f}',
                                        'Прирост % (Rev)': '{:+.2f}%',
                                        'Прибыль (Test)': '{:,.2f}',
                                        'Прибыль (Pre-Test Avg)': '{:,.2f}',
                                        'Прирост % (Prof)': '{:+.2f}%'
                                    }),
                                    use_container_width=True
                                )
                                
                                st.markdown("#### Детализация расчета по неделям (Тестовая группа)")
                                
                                for _, row in effect_details_df.iterrows():
                                    week_fmt = row['week_formatted']
                                    mode_cmt = row['Mode_Comment']
                                    
                                    # Revenue Values
                                    calc_rev = row['Calc_Revenue']
                                    base_rev = row['PreTest_Avg_Revenue']
                                    uplift_rev = row['Uplift_Revenue_Pct']
                                    
                                    # Profit Values
                                    calc_prof = row['Calc_Profit']
                                    base_prof = row['PreTest_Avg_Profit']
                                    uplift_prof = row['Uplift_Profit_Pct']
                                    
                                    with st.expander(f"Расчет прироста (Test): Неделя {week_fmt}"):
                                        c1, c2 = st.columns(2)
                                        
                                        with c1:
                                            st.markdown("**1. Эффект по Выручке**")
                                            st.markdown(f"""
                                            **Формула:** `(Fact_Revenue / PreTest_Avg) - 1`
                                            
                                            *   **Fact Revenue** {mode_cmt} = `{calc_rev:,.2f} ₽`
                                            *   **Pre-Test Avg** (база) = `{base_rev:,.2f} ₽`
                                            
                                            **Прирост (Uplift)** = `({calc_rev:,.2f} / {base_rev:,.2f}) - 1` = :green[**{uplift_rev*100:+.2f}%**]
                                            """)
                                            
                                        with c2:
                                            st.markdown("**2. Эффект по Прибыли**")
                                            st.markdown(f"""
                                            **Формула:** `(Fact_Profit / PreTest_Avg) - 1`
                                            
                                            *   **Fact Profit** {mode_cmt} = `{calc_prof:,.2f} ₽`
                                            *   **Pre-Test Avg** (база) = `{base_prof:,.2f} ₽`
                                            
                                            **Прирост (Uplift)** = `({calc_prof:,.2f} / {base_prof:,.2f}) - 1` = :green[**{uplift_prof*100:+.2f}%**]
                                            """)

                                # --- Control Group & Net Effect ---
                                st.markdown("### 6.2 Контрольная группа и Чистый эффект")
                                st.write("""
                                Все хорошо, но возможно такой рост произошел не по причине изменения цен, а, например, из-за сезонности. 
                                Чтобы исключить этот фактор, мы сравниваем прирост тестового товара с **Контрольной Группой** (ассортимент, который не участвовал в переоценке).
                                
                                **Логика:**
                                1. Считаем прирост Контрольной Группы за те же периоды.
                                2. Вычитаем прирост КГ из прироста Теста, чтобы получить **Чистый Эффект**.
                                """)
                                
                                # Control Group Table
                                summary_cols_control = [
                                    'week_formatted',
                                    'Control_Revenue_Real', 'Control_Avg_Revenue',
                                    'Control_Uplift_Revenue_Pct',
                                    'Control_Profit_Real', 'Control_Avg_Profit',
                                    'Control_Uplift_Profit_Pct'
                                ]
                                
                                display_control = effect_details_df[summary_cols_control].rename(columns={
                                    'week_formatted': 'Неделя',
                                    'Control_Revenue_Real': 'Выручка (КГ)',
                                    'Control_Avg_Revenue': 'Выручка (КГ База)',
                                    'Control_Uplift_Revenue_Pct': 'Прирост % (КГ Rev)',
                                    'Control_Profit_Real': 'Прибыль (КГ)',
                                    'Control_Avg_Profit': 'Прибыль (КГ База)',
                                    'Control_Uplift_Profit_Pct': 'Прирост % (КГ Prof)'
                                })
                                
                                st.dataframe(
                                    display_control.style.format({
                                        'Выручка (КГ)': '{:,.2f}',
                                        'Выручка (КГ База)': '{:,.2f}',
                                        'Прирост % (КГ Rev)': '{:+.2f}%',
                                        'Прибыль (КГ)': '{:,.2f}',
                                        'Прибыль (КГ База)': '{:,.2f}',
                                        'Прирост % (КГ Prof)': '{:+.2f}%'
                                    }),
                                    use_container_width=True
                                )
                                
                                st.markdown("#### Детализация расчета по неделям (Контрольная группа)")
                                
                                for _, row in effect_details_df.iterrows():
                                    week_fmt = row['week_formatted']
                                    mode_cmt = row['Mode_Comment']
                                    
                                    # Control Revenue Values
                                    calc_rev_c = row['Calc_Control_Revenue']
                                    base_rev_c = row['Control_Avg_Revenue']
                                    uplift_rev_c = row['Control_Uplift_Revenue_Pct']
                                    
                                    # Control Profit Values
                                    calc_prof_c = row['Calc_Control_Profit']
                                    base_prof_c = row['Control_Avg_Profit']
                                    uplift_prof_c = row['Control_Uplift_Profit_Pct']
                                    
                                    with st.expander(f"Расчет прироста (КГ): Неделя {week_fmt}"):
                                        c1, c2 = st.columns(2)
                                        
                                        with c1:
                                            st.markdown("**1. Прирост КГ по Выручке**")
                                            st.markdown(f"""
                                            **Формула:** `(Control_Fact / Control_Base) - 1`
                                            
                                            *   **Control Fact** {mode_cmt} = `{calc_rev_c:,.2f} ₽`
                                            *   **Control Base** (база) = `{base_rev_c:,.2f} ₽`
                                            
                                            **Прирост (Uplift)** = `({calc_rev_c:,.2f} / {base_rev_c:,.2f}) - 1` = :blue[**{uplift_rev_c*100:+.2f}%**]
                                            """)
                                            
                                        with c2:
                                            st.markdown("**2. Прирост КГ по Прибыли**")
                                            st.markdown(f"""
                                            **Формула:** `(Control_Fact / Control_Base) - 1`
                                            
                                            *   **Control Fact** {mode_cmt} = `{calc_prof_c:,.2f} ₽`
                                            *   **Control Base** (база) = `{base_prof_c:,.2f} ₽`
                                            
                                            **Прирост (Uplift)** = `({calc_prof_c:,.2f} / {base_prof_c:,.2f}) - 1` = :blue[**{uplift_prof_c*100:+.2f}%**]
                                            """)
                                
                                # Net Effect Table
                                st.markdown("#### Итоговый чистый прирост")
                                st.write("Вычитаем прирост Контрольной Группы из прироста Теста:")
                                
                                summary_cols_net = [
                                    'week_formatted',
                                    'Uplift_Revenue_Pct', 'Control_Uplift_Revenue_Pct', 'Net_Effect_Revenue_Pct',
                                    'Uplift_Profit_Pct', 'Control_Uplift_Profit_Pct', 'Net_Effect_Profit_Pct'
                                ]
                                
                                display_net = effect_details_df[summary_cols_net].rename(columns={
                                    'week_formatted': 'Неделя',
                                    'Uplift_Revenue_Pct': 'Test % (Rev)',
                                    'Control_Uplift_Revenue_Pct': 'Control % (Rev)',
                                    'Net_Effect_Revenue_Pct': 'Чистый Эффект % (Rev)',
                                    'Uplift_Profit_Pct': 'Test % (Prof)',
                                    'Control_Uplift_Profit_Pct': 'Control % (Prof)',
                                    'Net_Effect_Profit_Pct': 'Чистый Эффект % (Prof)'
                                })
                                
                                def highlight_net_effect(row):
                                    # Highlight Net Effect columns
                                    styles = [''] * len(row)
                                    if 'Чистый Эффект % (Rev)' in row.index:
                                        idx_r = row.index.get_loc('Чистый Эффект % (Rev)')
                                        val_r = row['Чистый Эффект % (Rev)']
                                        color_r = '#90ee90' if val_r > 0 else '#ffcccc'
                                        styles[idx_r] = f'background-color: {color_r}; color: black; font-weight: bold'
                                        
                                    if 'Чистый Эффект % (Prof)' in row.index:
                                        idx_p = row.index.get_loc('Чистый Эффект % (Prof)')
                                        val_p = row['Чистый Эффект % (Prof)']
                                        color_p = '#90ee90' if val_p > 0 else '#ffcccc'
                                        styles[idx_p] = f'background-color: {color_p}; color: black; font-weight: bold'
                                    
                                    return styles

                                st.dataframe(
                                    display_net.style
                                    .apply(highlight_net_effect, axis=1)
                                    .format({
                                        'Test % (Rev)': '{:+.2f}%',
                                        'Control % (Rev)': '{:+.2f}%',
                                        'Чистый Эффект % (Rev)': '{:+.2f}%',
                                        'Test % (Prof)': '{:+.2f}%',
                                        'Control % (Prof)': '{:+.2f}%',
                                        'Чистый Эффект % (Prof)': '{:+.2f}%'
                                    }),
                                    use_container_width=True
                                )
                                # --- Absolute Net Effect ---
                                st.markdown("### 6.3 Абсолютный чистый эффект")
                                st.write("""
                                Переводим относительный чистый прирост в деньги. 
                                
                                **Методология:** 
                                Абсолютный эффект рассчитывается путем умножения **фактического показателя** тестовой группы (за конкретную неделю) на **процент чистого прироста**.
                                
                                *Формула:* `Net_Abs_Effect = Net_Effect_% * Fact_Test_Metric`
                                """)
                                
                                summary_cols_abs = [
                                    'week_formatted',
                                    'Net_Effect_Revenue_Pct', 'Fact_Revenue_Real', 'Net_Abs_Effect_Revenue',
                                    'Net_Effect_Profit_Pct', 'Fact_Profit_Real', 'Net_Abs_Effect_Profit'
                                ]
                                
                                display_abs = effect_details_df[summary_cols_abs].rename(columns={
                                    'week_formatted': 'Неделя',
                                    'Net_Effect_Revenue_Pct': 'Чистый % (Rev)',
                                    'Fact_Revenue_Real': 'Факт (Test Rev)',
                                    'Net_Abs_Effect_Revenue': 'Абс. Эффект (Rev)',
                                    'Net_Effect_Profit_Pct': 'Чистый % (Prof)',
                                    'Fact_Profit_Real': 'Факт (Test Prof)',
                                    'Net_Abs_Effect_Profit': 'Абс. Эффект (Prof)'
                                })
                                
                                def highlight_abs_effect(row):
                                    styles = [''] * len(row)
                                    # Highlight Absolute Effect columns
                                    for col in ['Абс. Эффект (Rev)', 'Абс. Эффект (Prof)']:
                                        if col in row.index:
                                            idx = row.index.get_loc(col)
                                            val = row[col]
                                            # Use simple green/red logic
                                            color = '#90ee90' if val > 0 else '#ffcccc'
                                            # If zero, maybe neutral?
                                            if val == 0: color = ''
                                            else:
                                                styles[idx] = f'background-color: {color}; color: black; font-weight: bold'
                                    return styles

                                # Add Total row
                                total_row_abs = pd.DataFrame([{
                                    'Неделя': 'ИТОГО',
                                    'Чистый % (Rev)': effect_details_df['Net_Effect_Revenue_Pct'].mean(),
                                    'Факт (Test Rev)': effect_details_df['Fact_Revenue_Real'].sum(),
                                    'Абс. Эффект (Rev)': effect_details_df['Net_Abs_Effect_Revenue'].sum(),
                                    'Чистый % (Prof)': effect_details_df['Net_Effect_Profit_Pct'].mean(),
                                    'Факт (Test Prof)': effect_details_df['Fact_Profit_Real'].sum(),
                                    'Абс. Эффект (Prof)': effect_details_df['Net_Abs_Effect_Profit'].sum()
                                }])
                                display_abs_with_total = pd.concat([display_abs, total_row_abs], ignore_index=True)
                                
                                def highlight_abs_effect_with_total(row):
                                    styles = [''] * len(row)
                                    is_total = row['Неделя'] == 'ИТОГО'
                                    
                                    if is_total:
                                        # Bold all cells in total row
                                        styles = ['font-weight: bold'] * len(row)
                                    
                                    # Highlight Absolute Effect columns
                                    for col in ['Абс. Эффект (Rev)', 'Абс. Эффект (Prof)']:
                                        if col in row.index:
                                            idx = row.index.get_loc(col)
                                            val = row[col]
                                            if val != 0:
                                                color = '#90ee90' if val > 0 else '#ffcccc'
                                                styles[idx] = f'background-color: {color}; color: black; font-weight: bold'
                                    return styles

                                st.dataframe(
                                    display_abs_with_total.style
                                    .apply(highlight_abs_effect_with_total, axis=1)
                                    .format({
                                        'Чистый % (Rev)': '{:+.2f}%',
                                        'Факт (Test Rev)': '{:,.2f}',
                                        'Абс. Эффект (Rev)': '{:,.2f}',
                                        'Чистый % (Prof)': '{:+.2f}%',
                                        'Факт (Test Prof)': '{:,.2f}',
                                        'Абс. Эффект (Prof)': '{:,.2f}'
                                    }),
                                    use_container_width=True
                                )
                                
                                st.markdown("#### Детализация абсолютного эффекта по неделям")
                                for _, row in effect_details_df.iterrows():
                                    week_fmt = row['week_formatted']
                                    
                                    net_pct_rev = row['Net_Effect_Revenue_Pct']
                                    fact_rev = row['Fact_Revenue_Real']
                                    abs_rev = row['Net_Abs_Effect_Revenue']
                                    
                                    net_pct_prof = row['Net_Effect_Profit_Pct']
                                    fact_prof = row['Fact_Profit_Real']
                                    abs_prof = row['Net_Abs_Effect_Profit']
                                    
                                    with st.expander(f"Расчет абс. эффекта: Неделя {week_fmt}"):
                                        c1, c2 = st.columns(2)
                                        with c1:
                                            st.markdown("**Выручка**")
                                            st.markdown(f"`{net_pct_rev*100:+.2f}%` * `{fact_rev:,.2f} ₽` = :green[**{abs_rev:,.2f} ₽**]")
                                        with c2:
                                            st.markdown("**Прибыль**")
                                            st.markdown(f"`{net_pct_prof*100:+.2f}%` * `{fact_prof:,.2f} ₽` = :green[**{abs_prof:,.2f} ₽**]")
                                
                                # --- Summary / Итоги ---
                                st.markdown("### 7. Итоги")
                                
                                total_abs_revenue = effect_details_df['Net_Abs_Effect_Revenue'].sum()
                                total_abs_profit = effect_details_df['Net_Abs_Effect_Profit'].sum()
                                avg_pct_revenue = effect_details_df['Net_Effect_Revenue_Pct'].mean() * 100
                                avg_pct_profit = effect_details_df['Net_Effect_Profit_Pct'].mean() * 100
                                
                                st.success(f"""
                                **Итого данная позиция принесла:**
                                - Дополнительную выручку: **{total_abs_revenue:,.2f} ₽** (в среднем **{avg_pct_revenue:+.2f}%**)
                                - Дополнительную прибыль: **{total_abs_profit:,.2f} ₽** (в среднем **{avg_pct_profit:+.2f}%**)
                                """)
                                
                            else:
                                st.info("Нет валидных тестовых недель для расчета эффекта.")
                                
                        else:
                            st.warning("Нет данных для формирования понедельного отчета.")


                        # --- 8. Pilot Summary ---
                        st.markdown("---")
                        st.markdown("### 8. Итого по всему пилоту")
                        st.write("""
                        Таким способом мы считаем показатели по каждой позиции индивидуально. 
                        Если просуммировать эффект по всем протестированным позициям, мы получим итоговые показатели всего пилота.
                        """)
                        
                        st.subheader("Ключевые показатели эффективности")
                        
                        # --- Row 1: Revenue Metrics ---
                        st.markdown("**Выручка (Revenue)**")
                        r1, r2, r3, r4, r5 = st.columns(5)
                        
                        r1.metric("Эффект", f"{summary.get('total_abs_effect_revenue', 0):,.0f} ₽", 
                                  delta=f"{summary.get('effect_revenue_pct', 0):.2f}%")
                        
                        r2.metric("Без эффекта (Test)", f"{summary.get('revenue_without_effect', 0):,.0f} ₽")
                        r3.metric("С эффектом (Test)", f"{summary.get('total_fact_revenue', 0):,.0f} ₽")
                        r4.metric("Полный оборот (Global)", f"{summary.get('global_revenue', 0):,.0f} ₽")
                        r5.metric("Доля Test (%)", f"{summary.get('test_share_revenue', 0):.1f}%")
                        
                        st.divider()
                        
                        # --- Row 2: Profit Metrics ---
                        st.markdown("**Прибыль (Profit)**")
                        p1, p2, p3, p4, p5 = st.columns(5)
                        
                        p1.metric("Эффект", f"{summary.get('total_abs_effect_profit', 0):,.0f} ₽",
                                  delta=f"{summary.get('effect_profit_pct', 0):.2f}%")
                        
                        p2.metric("Без эффекта (Test)", f"{summary.get('profit_without_effect', 0):,.0f} ₽")
                        p3.metric("С эффектом (Test)", f"{summary.get('total_fact_profit', 0):,.0f} ₽")
                        p4.metric("Полная прибыль (Global)", f"{summary.get('global_profit', 0):,.0f} ₽")
                        p5.metric("Доля Test (%)", f"{summary.get('test_share_profit', 0):.1f}%")
                        
                        st.markdown("---")
                        
                        # --- Statistics Block ---
                        st.subheader("Статистика теста")
                        
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Протестировано позиций", summary.get('tested_count', 0))
                        c1.metric("Исключено позиций", summary.get('excluded_products_count', 0))
                        
                        c2.metric("Длительность теста", f"{summary.get('test_duration_weeks', 0):.0f} нед.")
                        c2.metric("Исключено периодов (недель)", summary.get('excluded_weeks_count', 0))
                        
                        c3.metric("Количество переоценок", summary.get('price_changes_count', 0))
                        
                        # Expanders for details
                        with st.expander("Детализация по переоценкам"):
                            changes_data = pd.DataFrame(
                                list(summary.get('products_per_change', {}).items()),
                                columns=['Дата старта цены', 'Кол-во товаров']
                            ).sort_values('Дата старта цены')
                            st.dataframe(changes_data, use_container_width=True)
                        
                        st.markdown("### Результаты по направлениям")
                        g_stats = summary.get('growth_stats', {})
                        d_stats = summary.get('decline_stats', {})
                        
                        col_g, col_d = st.columns(2)
                        
                        with col_g:
                            st.success(f"📈 РОСТ: {g_stats.get('count', 0)} позиций")
                            st.write(f"Эффект (Выручка): **{g_stats.get('revenue_effect', 0):,.0f} ₽**")
                            st.write(f"Эффект (Прибыль): **{g_stats.get('profit_effect', 0):,.0f} ₽**")
                            
                        with col_d:
                            st.error(f"📉 ПАДЕНИЕ: {d_stats.get('count', 0)} позиций")
                            st.write(f"Эффект (Выручка): **{d_stats.get('revenue_effect', 0):,.0f} ₽**")
                            st.write(f"Эффект (Прибыль): **{d_stats.get('profit_effect', 0):,.0f} ₽**")

                        # --- WORD EXPORT BUTTON ---
                        st.markdown("---")
                        col_export, _ = st.columns([1, 3])
                        with col_export:
                            # Collect params for generator
                            report_params = {
                                'pre_test_weeks_count': pre_test_weeks,
                                'pre_test_stock_threshold': pre_test_threshold,
                                'contiguous_pre_test': contiguous_pre_test,
                                'test_use_week_values': (test_calc_mode == "Значение текущей недели")
                            }
                            
                            generator = WordReportGenerator(
                                calc, 
                                selected_report_pid, 
                                report_params, 
                                results_summary=summary,
                                activation_params=activation_params
                            )
                            docx_file = generator.generate()
                            
                            st.download_button(
                                label="📄 Скачать отчет в Word",
                                data=docx_file,
                                file_name=f"report_{selected_report_pid}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        
                    else:
                        st.warning("Не удалось найти подходящий товар для истории (возможно, нет данных или результатов).")

            st.markdown("---")
            st.subheader("Экспорт результатов")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                pd.DataFrame([summary]).to_excel(writer, sheet_name='Summary', index=False)
                if not results.empty:
                    results.to_excel(writer, sheet_name='Detailed Results', index=False)
                
                if 'activation_df' in locals() and not activation_df.empty:
                    activation_df.to_excel(writer, sheet_name='Activation Analysis', index=False)
                
                calc.get_control_group_info().to_excel(writer, sheet_name='Control Group', index=False)
                
                all_timelines = []
                for pid in all_test_pids: 
                    tl = calc.get_product_timeline(pid)
                    if not tl.empty:
                        tl['product_id'] = pid
                        tl['product_name'] = calc.product_names.get(pid, "Unknown")
                        all_timelines.append(tl)
                
                if all_timelines:
                    pd.concat(all_timelines).to_excel(writer, sheet_name='Product Timelines', index=False)
                
            buffer.seek(0)
            st.download_button(
                label="📥 Скачать полный отчет (Excel)",
                data=buffer,
                file_name="pricing_effect_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            pres_fn = st.session_state.get("presentation_filename")
            if pres_fn is None:
                try:
                    _act_round = st.session_state.get("activation_round_direction_label", "Вверх до значения")
                    _act_dir = "up" if _act_round.startswith("Вверх") else ("down" if _act_round.startswith("Вниз") else "nearest")
                    _act_params = {
                        "threshold_pct": st.session_state.get("activation_threshold", 1),
                        "use_rounding": st.session_state.get("activation_use_rounding", True),
                        "round_value": st.session_state.get("activation_round_value", 90),
                        "round_direction": _act_dir,
                        "wap_from_change_date": st.session_state.get("activation_wap_from_change_date", True),
                        "min_days_threshold": st.session_state.get("activation_min_days_threshold", 2),
                    }
                    _calc_params = {
                        "pre_test_weeks": pre_test_weeks,
                        "pre_test_threshold": pre_test_threshold,
                        "contiguous_pre_test": contiguous_pre_test,
                        "use_stock_filter": use_stock_filter,
                        "stock_threshold_pct": threshold_pct,
                        "test_use_week_values": (test_calc_mode == "Значение текущей недели"),
                    }
                    stats_data = build_stats_data(calc, results, summary, _act_params, st.session_state.get("uploaded_file_name", "файл.xlsx"), _calc_params)
                    valid_pids = results[results["Is_Excluded"] == False]["product_id"].unique()
                    pid = int(valid_pids[0]) if len(valid_pids) > 0 else int(list(calc.test_product_ids)[0])
                    pres_data = build_presentation_data(calc, results, pid, _act_params, use_week_values=(test_calc_mode == "Значение текущей недели"))
                    html = generate_html(stats_data, pres_data)
                    pres_fn = save_presentation_and_manage_history(html, Path(__file__).parent, max_history=3)
                    st.session_state["presentation_filename"] = pres_fn
                except Exception:
                    pres_fn = None
            if pres_fn:
                static_path = Path(__file__).parent / "static" / pres_fn
                if static_path.exists():
                    html_bytes = static_path.read_bytes()
                    st.download_button(
                        label="📊 Скачать и открыть презентацию по расчету",
                        data=html_bytes,
                        file_name=pres_fn,
                        mime="text/html",
                        help="Скачайте файл и откройте в браузере — все данные уже встроены",
                    )
                    pdf_fn = static_path.with_suffix(".pdf").name
                    try:
                        pdf_path = export_html_to_pdf(static_path, static_path.with_suffix(".pdf"))
                        pdf_bytes = pdf_path.read_bytes()
                        st.download_button(
                            label="📄 Скачать презентацию в PDF",
                            data=pdf_bytes,
                            file_name=pdf_fn,
                            mime="application/pdf",
                            help="PDF через WebKit (как Safari), 1 слайд = 1 страница",
                            key="pdf_dl",
                        )
                    except Exception as pdf_err:
                        st.caption(f"PDF недоступен: {pdf_err}")
            
            # --- WORD EXPORT ---
            # Button is inside Tab 5 now.
            pass
            
        else:
            st.info("Нажмите кнопку 'Применить настройки' в сайдбаре для запуска расчета.")
            
    except Exception as e:
        st.error(f"Произошла ошибка при обработке файла: {e}")
        st.exception(e)
