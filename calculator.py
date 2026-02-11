import pandas as pd
import numpy as np
import datetime

class EffectCalculator:
    def __init__(self, file_path_or_buffer):
        self.xl = pd.ExcelFile(file_path_or_buffer)
        self.test_prices = pd.read_excel(self.xl, 'Тестовые цены')
        self.sales = pd.read_excel(self.xl, 'Продажи')
        self.costs = pd.read_excel(self.xl, 'Себестоимость') # Contains stock data
        
    def preprocess(self):
        # Convert dates
        self.test_prices['New_Price_Start'] = pd.to_datetime(self.test_prices['New_Price_Start'])
        self.sales['recorded_on'] = pd.to_datetime(self.sales['recorded_on'])
        self.costs['date'] = pd.to_datetime(self.costs['date'])

        # --- Calculate Cost at Sale (Logic: Backward -> Forward -> 0) ---
        # Sort for merge_asof
        self.sales = self.sales.sort_values('recorded_on')
        self.costs = self.costs.sort_values('date')

        # 1. Backward search (on or before)
        sales_w_cost = pd.merge_asof(
            self.sales,
            self.costs[['product_id', 'date', 'cost']],
            left_on='recorded_on',
            right_on='date',
            by='product_id',
            direction='backward'
        )

        # 2. Forward search (on or after) - fallback
        sales_w_cost_fwd = pd.merge_asof(
            self.sales,
            self.costs[['product_id', 'date', 'cost']],
            left_on='recorded_on',
            right_on='date',
            by='product_id',
            direction='forward',
            suffixes=('', '_fwd')
        )

        # Combine: take backward, fill NaN with forward, fill remaining NaN with 0
        sales_w_cost['cost_at_sale'] = sales_w_cost['cost'].fillna(sales_w_cost_fwd['cost']).fillna(0)
        
        # Cleanup extra columns from merge
        cols_to_drop = ['cost']
        if 'date' in sales_w_cost.columns:
            cols_to_drop.append('date')
            
        self.sales = sales_w_cost.drop(columns=cols_to_drop)
        
        # Calculate revenue for each transaction
        self.sales['revenue'] = self.sales['price'] * self.sales['quantity']
        self.sales['cost_volume'] = self.sales['cost_at_sale'] * self.sales['quantity']
        
        # Align dates to Week Start (Monday)
        self.sales['week_start'] = self.sales['recorded_on'].apply(lambda x: x - pd.Timedelta(days=x.weekday())).dt.normalize()
        self.costs['week_start'] = self.costs['date'].apply(lambda x: x - pd.Timedelta(days=x.weekday())).dt.normalize()
        
        # Determine Product Names mapping
        self.product_names = self.sales.groupby('product_id')['name_full'].first()
        
        # Identify Control and Test products
        self.all_products = set(self.sales['product_id'].unique())
        self.test_product_ids = set(self.test_prices['product_id'].unique())
        self.control_product_ids = self.all_products - self.test_product_ids
        
        # Aggregate Sales by Product and Week
        self.weekly_sales = self.sales.groupby(['product_id', 'week_start']).agg({
            'revenue': 'sum',
            'quantity': 'sum',
            'cost_volume': 'sum'
        }).reset_index()
        self.sales_lookup = set(zip(self.weekly_sales['product_id'], self.weekly_sales['week_start']))
        
        # Aggregate Control Group Sales by Week
        control_sales_df = self.sales[self.sales['product_id'].isin(self.control_product_ids)]
        self.control_weekly_sales = control_sales_df.groupby('week_start').agg({
            'revenue': 'sum',
            'cost_volume': 'sum'
        }).reset_index()
        self.control_weekly_sales.rename(columns={
            'revenue': 'control_revenue',
            'cost_volume': 'control_cost_volume'
        }, inplace=True)
        
        # STOCK AGGREGATION: 
        # 1. Average Daily Stock per Week (for filtering test weeks)
        self.weekly_stock = self.costs.groupby(['product_id', 'week_start'])['stock'].mean().reset_index()
        self.stock_lookup = dict(zip(zip(self.weekly_stock['product_id'], self.weekly_stock['week_start']), self.weekly_stock['stock']))
        
        # 2. Detailed Stock Info for Pre-test Search: Count days with stock > 0 per week
        # Filter days with stock > 0
        self.costs['has_stock'] = self.costs['stock'] > 0
        self.weekly_stock_days = self.costs.groupby(['product_id', 'week_start'])['has_stock'].sum().reset_index()
        # Lookup: (pid, week) -> days_with_stock
        self.stock_days_lookup = dict(zip(zip(self.weekly_stock_days['product_id'], self.weekly_stock_days['week_start']), self.weekly_stock_days['has_stock']))
        
        # Determine Global Data Range (Full weeks only)
        max_date = self.sales['recorded_on'].max()
        if max_date.weekday() != 6: # If not Sunday
            last_week_start = max_date - pd.Timedelta(days=max_date.weekday())
            last_week_start = pd.to_datetime(last_week_start).normalize()
            self.weekly_sales = self.weekly_sales[self.weekly_sales['week_start'] < last_week_start]
            self.control_weekly_sales = self.control_weekly_sales[self.control_weekly_sales['week_start'] < last_week_start]
            self.weekly_stock = self.weekly_stock[self.weekly_stock['week_start'] < last_week_start]
            
            # Rebuild lookups after filtering
            self.sales_lookup = set(zip(self.weekly_sales['product_id'], self.weekly_sales['week_start']))
            self.stock_lookup = dict(zip(zip(self.weekly_stock['product_id'], self.weekly_stock['week_start']), self.weekly_stock['stock']))

    def format_week(self, week_start):
        """Helper to format week for display: 'WeekNum (Start - End)'"""
        week_num = week_start.isocalendar()[1]
        week_end = week_start + pd.Timedelta(days=6)
        return f"{week_num} ({week_start.strftime('%d.%m.%Y')}-{week_end.strftime('%d.%m.%Y')})"

    def find_valid_pre_test_period(self, pid, start_search_week, n_weeks, threshold_pct, contiguous=True):
        max_lookback = 26 
        min_days_stock_per_week = 7 * (threshold_pct / 100.0)
        
        if contiguous:
            current_end_week = start_search_week
            for i in range(max_lookback):
                current_start_week = current_end_week - pd.Timedelta(weeks=n_weeks-1)
                window_weeks = pd.date_range(start=current_start_week, end=current_end_week, freq='W-MON')
                
                is_valid_window = True
                for w in window_weeks:
                    stock_days = self.stock_days_lookup.get((pid, w), 0)
                    if stock_days < min_days_stock_per_week:
                        is_valid_window = False
                        break
                
                if is_valid_window:
                    return list(window_weeks) # Return list of weeks
                
                current_end_week = current_end_week - pd.Timedelta(weeks=1)
        else:
            # Non-contiguous search
            valid_weeks = []
            current_week = start_search_week
            for i in range(max_lookback * 2): # Look further back if needed
                stock_days = self.stock_days_lookup.get((pid, current_week), 0)
                if stock_days >= min_days_stock_per_week:
                    valid_weeks.append(current_week)
                
                if len(valid_weeks) == n_weeks:
                    return sorted(valid_weeks) # Return sorted list of weeks
                
                current_week = current_week - pd.Timedelta(weeks=1)
                
        return None

    def get_pre_test_selection_details(self, pid, params, report_min=None, report_max=None):
        pre_test_weeks = params.get('pre_test_weeks_count', 2)
        threshold_pct = params.get('pre_test_stock_threshold', 10)
        contiguous = params.get('contiguous_pre_test', True)
        
        # Get Product Info
        product_test_info = self.test_prices[self.test_prices['product_id'] == pid]
        if product_test_info.empty:
            return None
            
        start_date = product_test_info['New_Price_Start'].min()
        start_week_monday = pd.to_datetime(start_date - pd.Timedelta(days=start_date.weekday())).normalize()
        
        if start_date.weekday() == 0:
            pre_test_search_end = start_week_monday - pd.Timedelta(weeks=1)
        else:
            pre_test_search_end = start_week_monday - pd.Timedelta(weeks=1)
            
        min_days_stock_per_week = 7 * (threshold_pct / 100.0)
        
        # Find selected weeks using existing logic
        selected_weeks = self.find_valid_pre_test_period(
            pid, pre_test_search_end, pre_test_weeks, threshold_pct, contiguous=contiguous
        )
        if selected_weeks is None:
            selected_weeks = []
        
        # Determine table range: all weeks in Weekly Indicators + selected pre-test weeks
        range_candidates = []
        if report_min is not None:
            range_candidates.append(pd.to_datetime(report_min).normalize())
        if report_max is not None:
            range_candidates.append(pd.to_datetime(report_max).normalize())
        if selected_weeks:
            range_candidates.append(min(selected_weeks))
            range_candidates.append(max(selected_weeks))
        
        if range_candidates:
            min_report_week = min(range_candidates)
            max_report_week = max(range_candidates)
        else:
            min_report_week = pre_test_search_end - pd.Timedelta(weeks=8)
            max_report_week = pre_test_search_end
        
        if min_report_week > max_report_week:
            min_report_week, max_report_week = max_report_week, min_report_week
        
        full_range_weeks = pd.date_range(start=min_report_week, end=max_report_week, freq='W-MON')
        
        rows = []
        selected_stocks = []
        
        for w in full_range_weeks:
            stock_days = self.stock_days_lookup.get((pid, w), 0)
            stock_pct = stock_days / 7 * 100
            
            # Avg stock for this week (from weekly_stock)
            avg_stock_val = self.stock_lookup.get((pid, w), 0)
            
            is_valid_stock_days = stock_days >= min_days_stock_per_week
            is_selected = w in selected_weeks
            
            if is_selected:
                selected_stocks.append(avg_stock_val)
            
            status = ""
            if is_selected:
                status = "Выбрана (Pre-Test)"
            elif w > pre_test_search_end:
                status = "Тестовый период"
            elif is_valid_stock_days:
                status = "Доступна (ОК)"
            else:
                status = "Мало стока"
            
            rows.append({
                'Неделя': self.format_week(w),
                'Дней на остатке': stock_days,
                'Доступность %': stock_pct,
                'Средний остаток': avg_stock_val,
                'Статус': status,
                'Выбрана': is_selected
            })
            
        # Calculate Base Stock
        base_stock_val = np.mean(selected_stocks) if selected_stocks else 0
            
        return {
            'df': pd.DataFrame(rows),
            'base_stock': base_stock_val,
            'selected_stocks': selected_stocks,
            'selected_weeks': selected_weeks
        }

    def get_activation_details(self, threshold_pct=10, use_rounding=False, round_value=90, round_direction="up", wap_from_change_date=True, min_days_threshold=3, specific_pid=None):
        activation_results = []
        
        target_pids = [specific_pid] if specific_pid else self.test_product_ids
        
        for pid in target_pids:
            product_prices = self.test_prices[self.test_prices['product_id'] == pid].sort_values('New_Price_Start')
            if product_prices.empty:
                continue

            # Ensure Current_Price is available (assuming it's in the Test Prices sheet)
            if 'Current_Price' not in product_prices.columns:
                 # Fallback if column missing, though user said it exists
                 product_prices['Current_Price'] = 0 
                
            first_test_start = product_prices['New_Price_Start'].iloc[0]
            
            # Align first test start to week - for Activation we check from the actual week of price change
            start_week_monday = first_test_start - pd.Timedelta(days=first_test_start.weekday())
            start_week_monday = pd.to_datetime(start_week_monday).normalize()
            
            # For activation analysis, use the week when price should have changed (not next week)
            test_start_week = start_week_monday
                
            last_week = self.weekly_sales['week_start'].max()
            if pd.isnull(last_week):
                 continue

            test_weeks = pd.date_range(start=test_start_week, end=last_week, freq='W-MON')
            
            last_known_is_our_price = False
            is_first_period = True
            blocked_same_price = False
            last_applied_plan_price = None
            if round_value is None:
                round_value = 0
            round_value = max(0, min(99, int(round_value)))

            def round_price(price):
                if price is None or price <= 0:
                    return None
                whole = int(price)
                frac = price - whole
                if frac < 1e-6:
                    return float(whole)
                target = round_value / 100.0
                
                # If already rounded to target value, don't change
                if abs(frac - target) < 0.01:
                    return price
                
                if round_direction == "nearest":
                    # Round to nearest target value
                    lower = float(whole) + target
                    upper = float(whole + 1) + target
                    # Choose closer one
                    if abs(price - lower) <= abs(price - upper):
                        return lower
                    return upper
                elif round_direction == "down":
                    if frac >= target:
                        return float(whole) + target
                    return float(whole - 1) + target
                else:  # "up"
                    if frac <= target:
                        return float(whole) + target
                    return float(whole + 1) + target
            
            # Initialize prev_fact_price with last known price BEFORE test period
            prev_fact_price = None
            pre_test_sales = self.weekly_sales[
                (self.weekly_sales['product_id'] == pid) & 
                (self.weekly_sales['week_start'] < test_start_week) &
                (self.weekly_sales['quantity'] > 0)
            ].sort_values('week_start', ascending=False)
            
            if not pre_test_sales.empty:
                last_pre_test = pre_test_sales.iloc[0]
                prev_fact_price = last_pre_test['revenue'] / last_pre_test['quantity']

            for w in test_weeks:
                week_end = w + pd.Timedelta(days=6)
                active_price_rows = product_prices[product_prices['New_Price_Start'] <= week_end]
                if active_price_rows.empty:
                    continue 
                
                current_record = active_price_rows.iloc[-1]
                plan_price_current_raw = current_record['New_Price']
                
                # Collect ALL plan prices for this product
                all_plan_prices_raw = active_price_rows['New_Price'].tolist()
                
                if len(active_price_rows) >= 2:
                    plan_price_previous_raw = active_price_rows.iloc[-2]['New_Price']
                else:
                    plan_price_previous_raw = None

                if use_rounding:
                    plan_price_current = round_price(plan_price_current_raw)
                    plan_price_previous = round_price(plan_price_previous_raw) if plan_price_previous_raw is not None else None
                    plan_price_unused = plan_price_current_raw
                    # Apply rounding to all plan prices
                    all_plan_prices_processed = [round_price(p) for p in all_plan_prices_raw]
                else:
                    plan_price_current = plan_price_current_raw
                    plan_price_previous = plan_price_previous_raw
                    plan_price_unused = round_price(plan_price_current_raw)
                    all_plan_prices_processed = all_plan_prices_raw
                
                # Check if this week is a Price Change Week FIRST
                is_price_change_week = not product_prices[
                    product_prices['New_Price_Start'].apply(
                        lambda x: (x - pd.Timedelta(days=x.weekday())).normalize()
                    ) == w
                ].empty
                
                # Get the price change date if exists
                price_change_date = None
                if is_price_change_week:
                    change_rows = product_prices[
                        product_prices['New_Price_Start'].apply(
                            lambda x: (x - pd.Timedelta(days=x.weekday())).normalize()
                        ) == w
                    ]
                    if not change_rows.empty:
                        price_change_date = change_rows.iloc[0]['New_Price_Start']
                
                # Get actual sales data for CURRENT week
                first_occurrence_date = None  # Track for status annotation
                if wap_from_change_date and is_price_change_week and price_change_date is not None:
                    # Get sales from plan date onwards
                    week_sales_from_plan = self.sales[
                        (self.sales['product_id'] == pid) & 
                        (self.sales['week_start'] == w) &
                        (self.sales['recorded_on'] >= price_change_date)
                    ].sort_values('recorded_on')
                    
                    # Try to find first occurrence of new price in tolerance
                    if not week_sales_from_plan.empty and plan_price_current > 0:
                        for idx, row in week_sales_from_plan.iterrows():
                            if abs(row['price'] - plan_price_current) / plan_price_current * 100 <= threshold_pct:
                                first_occurrence_date = row['recorded_on']
                                break
                    
                    # If found first occurrence, use sales from that date
                    if first_occurrence_date is not None:
                        # Calculate days from selected date to end of week
                        week_end = w + pd.Timedelta(days=6)  # Sunday
                        days_working = (week_end - first_occurrence_date).days + 1  # +1 to include the start day
                        
                        # If less than 3 days, fallback to whole week
                        if days_working < min_days_threshold:
                            first_occurrence_date = None  # Mark as if not found to trigger fallback
                            sales_row = self.weekly_sales[
                                (self.weekly_sales['product_id'] == pid) & 
                                (self.weekly_sales['week_start'] == w)
                            ]
                            if sales_row.empty or sales_row.iloc[0]['quantity'] == 0:
                                fact_price = 0
                                has_sales = False
                            else:
                                revenue = sales_row.iloc[0]['revenue']
                                quantity = sales_row.iloc[0]['quantity']
                                fact_price = revenue / quantity
                                has_sales = True
                        else:
                            # Use partial week (3+ days)
                            week_sales = self.sales[
                                (self.sales['product_id'] == pid) & 
                                (self.sales['week_start'] == w) &
                                (self.sales['recorded_on'] >= first_occurrence_date)
                            ]
                            if not week_sales.empty:
                                revenue = (week_sales['price'] * week_sales['quantity']).sum()
                                quantity = week_sales['quantity'].sum()
                                if quantity > 0:
                                    fact_price = revenue / quantity
                                    has_sales = True
                                else:
                                    fact_price = 0
                                    has_sales = False
                            else:
                                fact_price = 0
                                has_sales = False
                    else:
                        # If not found, fallback to whole week
                        sales_row = self.weekly_sales[
                            (self.weekly_sales['product_id'] == pid) & 
                            (self.weekly_sales['week_start'] == w)
                        ]
                        if sales_row.empty or sales_row.iloc[0]['quantity'] == 0:
                            fact_price = 0
                            has_sales = False
                        else:
                            revenue = sales_row.iloc[0]['revenue']
                            quantity = sales_row.iloc[0]['quantity']
                            fact_price = revenue / quantity
                            has_sales = True
                else:
                    # Use aggregated weekly sales as before
                    sales_row = self.weekly_sales[
                        (self.weekly_sales['product_id'] == pid) & 
                        (self.weekly_sales['week_start'] == w)
                    ]
                    if sales_row.empty or sales_row.iloc[0]['quantity'] == 0:
                        fact_price = 0
                        has_sales = False
                    else:
                        revenue = sales_row.iloc[0]['revenue']
                        quantity = sales_row.iloc[0]['quantity']
                        fact_price = revenue / quantity
                        has_sales = True
                
                # Logic Determination
                status = "не та цена"
                not_our_reason = None
                can_use_in_analysis = False
                
                # Check if factual price matches plan (current or previous)
                is_our_price = False
                is_match_current = False
                is_match_any_previous = False
                is_match_last_applied = False
                price_changed_from_prev = None
                
                if has_sales and plan_price_current > 0:
                    is_match_current = abs(fact_price - plan_price_current) / plan_price_current * 100 <= threshold_pct

                
                # Check if factual price matches ANY previous plan price
                if not is_first_period and has_sales and len(all_plan_prices_processed) > 1:
                    # Check all prices except the last one (current)
                    for prev_price in all_plan_prices_processed[:-1]:
                        if prev_price is not None and prev_price > 0:
                            if abs(fact_price - prev_price) / prev_price * 100 <= threshold_pct:
                                is_match_any_previous = True
                                break

                if has_sales and last_applied_plan_price is not None and last_applied_plan_price > 0:
                    is_match_last_applied = abs(fact_price - last_applied_plan_price) / last_applied_plan_price * 100 <= threshold_pct

                
                change_from_prev_pct = None
                if has_sales and prev_fact_price is not None and prev_fact_price > 0:
                    change_from_prev_pct = abs(fact_price - prev_fact_price) / prev_fact_price * 100
                    price_changed_from_prev = change_from_prev_pct > 0.01
                
                is_our_price = is_match_current or is_match_any_previous or is_match_last_applied
                
                # Special case: first period coincidence with previous known fact price
                if is_first_period and has_sales and prev_fact_price is not None and prev_fact_price > 0:
                    # Treat as "same price" with a small tolerance
                    same_price_first = change_from_prev_pct is not None and change_from_prev_pct <= 1.0
                    if same_price_first:
                        blocked_same_price = True
                        is_our_price = False
                        not_our_reason = "совпадение"
                
                # If we detected a coincidence, treat same price as not ours until it changes
                if blocked_same_price:
                    if has_sales and change_from_prev_pct is not None:
                        if change_from_prev_pct > 1.0:
                            blocked_same_price = False
                        else:
                            is_our_price = False
                            not_our_reason = "совпадение"
                
                if has_sales:
                    if is_our_price:
                        # Determine WAP method annotation for OK statuses
                        wap_method = ""
                        if wap_from_change_date and is_price_change_week:
                            if first_occurrence_date is not None:
                                wap_method = ", план. даты"
                            else:
                                wap_method = ", полная нед."
                        
                        if is_match_current:
                            status = f"ОК (текущая{wap_method})"
                        elif is_match_any_previous or is_match_last_applied:
                            status = f"ОК (предыдущая{wap_method})"
                        else:
                            status = "ОК"
                        can_use_in_analysis = True
                    else:
                        # Prioritize "совпадение" on first week
                        if not_our_reason == "совпадение":
                            status = "не та цена (совпадение)"
                        else:
                            status = "не та цена (мимо плана)"
                        can_use_in_analysis = False
                    
                    last_known_is_our_price = is_our_price
                    if is_our_price and not blocked_same_price:
                        if is_match_current:
                            last_applied_plan_price = plan_price_current
                        elif is_match_any_previous:
                            # Find and save the specific matched previous price
                            for prev_price in all_plan_prices_processed[:-1]:
                                if prev_price is not None and prev_price > 0:
                                    if abs(fact_price - prev_price) / prev_price * 100 <= threshold_pct:
                                        last_applied_plan_price = prev_price
                                        break
                else:
                    if last_known_is_our_price:
                        status = "наша цена не было продаж"
                        can_use_in_analysis = True
                    else:
                        status = "не та цена (нет продаж)"
                        can_use_in_analysis = False


                fact_price_changed = False
                if has_sales and prev_fact_price is not None and prev_fact_price > 0:
                    fact_price_changed = abs(fact_price - prev_fact_price) / prev_fact_price * 100 > 1.0
                
                fact_matches_plan = False
                if has_sales and plan_price_current > 0:
                    fact_matches_plan = abs(fact_price - plan_price_current) / plan_price_current * 100 <= threshold_pct
                
                # Calculate deviation percentages
                deviation_from_current_pct = None
                if has_sales and plan_price_current > 0:
                    deviation_from_current_pct = abs(fact_price - plan_price_current) / plan_price_current * 100
                
                deviation_from_unused_pct = None
                if has_sales and plan_price_unused is not None and plan_price_unused > 0:
                    deviation_from_unused_pct = abs(fact_price - plan_price_unused) / plan_price_unused * 100
                
                # For first period, don't calculate deviation from previous (no planned previous price)
                deviation_from_previous_pct = None
                if not is_first_period and has_sales and plan_price_previous is not None and plan_price_previous > 0:
                    deviation_from_previous_pct = abs(fact_price - plan_price_previous) / plan_price_previous * 100
                
                # Store previous fact price for display - always show if exists
                stored_prev_fact_price = prev_fact_price if prev_fact_price is not None else None
                
                # Calculate fact price change percentage - always calculate if we have previous price
                # Exception: only when there were NO sales before this week
                fact_price_change_pct = None
                is_fact_change = False
                
                # Calculate change if we have both current sales and previous fact price
                if has_sales and stored_prev_fact_price is not None and stored_prev_fact_price > 0:
                    fact_price_change_pct = abs(fact_price - stored_prev_fact_price) / stored_prev_fact_price * 100
                    # Is_Change is True if there's any change (>0.01% to avoid rounding issues)
                    is_fact_change = fact_price_change_pct > 0.01
                
                # Update prev_fact_price for next iteration
                if has_sales:
                    prev_fact_price = fact_price

                # Calculate Weekly Fact Cost (Weighted by Sales, always full week)
                weekly_sales_row = self.weekly_sales[
                    (self.weekly_sales['product_id'] == pid) & 
                    (self.weekly_sales['week_start'] == w)
                ]
                
                fact_cost = 0
                if not weekly_sales_row.empty:
                    row_data = weekly_sales_row.iloc[0]
                    if row_data['quantity'] > 0:
                        fact_cost = row_data['cost_volume'] / row_data['quantity']

                activation_results.append({
                    'product_id': pid,
                    'product_name': self.product_names.get(pid, f"ID {pid}"),
                    'week_start': w,
                    'week_formatted': self.format_week(w),
                    'Plan_Price_Current': plan_price_current,
                    'Plan_Price_Previous': plan_price_previous if not is_first_period else None,
                    'Plan_Price_Unused': plan_price_unused,
                    'Fact_Price': fact_price,
                    'Fact_Cost': fact_cost,
                    'Fact_Price_Prev': stored_prev_fact_price,
                    'Fact_Price_Change_Pct': fact_price_change_pct,
                    'Is_Fact_Change': is_fact_change,
                    'Deviation_From_Current_Pct': deviation_from_current_pct,
                    'Deviation_Unused_Pct': deviation_from_unused_pct,
                    'Deviation_From_Previous_Pct': deviation_from_previous_pct,
                    'Status': status,
                    'Can_Use_In_Analysis': can_use_in_analysis,
                    'Fact_Price_Changed': fact_price_changed,
                    'Fact_Matches_Plan': fact_matches_plan,
                    'Is_Price_Change_Week': is_price_change_week,
                    'Is_First_Period': is_first_period
                })
                
                # After first price change week, we're no longer in first period
                # Move AFTER append to ensure first week has empty Plan_Price_Previous
                if is_price_change_week:
                    is_first_period = False
                
        return pd.DataFrame(activation_results)

    def analyze_revaluation_activation(self, activation_threshold=10, use_rounding=False, round_value=90, round_direction="up", wap_from_change_date=True, min_days_threshold=3):
        """
        Анализирует активацию цен для каждой переоценки.
        
        Возвращает DataFrame с результатами анализа по каждой переоценке:
        - revaluation_date: дата переоценки
        - planned_week: плановая неделя активации (понедельник недели)
        - total_proposed: общее количество предложенных цен
        - activated_on_time: количество активированных вовремя (на плановой неделе)
        - activated_later: количество активированных позже плановой недели
        - rejected: количество отклоненных цен
        
        Также возвращает детальный DataFrame по каждому товару в каждой переоценке.
        """
        if self.test_prices.empty:
            return pd.DataFrame(), pd.DataFrame()
        
        # Получаем все данные активации
        activation_df = self.get_activation_details(
            threshold_pct=activation_threshold,
            use_rounding=use_rounding,
            round_value=round_value,
            round_direction=round_direction,
            wap_from_change_date=wap_from_change_date,
            min_days_threshold=min_days_threshold
        )
        
        if activation_df.empty:
            return pd.DataFrame(), pd.DataFrame()
        
        # Группируем переоценки по дате
        revaluations = self.test_prices.groupby('New_Price_Start')
        
        summary_results = []
        detail_results = []
        
        for reval_date, reval_group in revaluations:
            # Определяем плановую неделю (понедельник недели, в которую попадает дата)
            planned_week = reval_date - pd.Timedelta(days=reval_date.weekday())
            planned_week = pd.to_datetime(planned_week).normalize()
            
            # Получаем список товаров в этой переоценке
            product_ids = reval_group['product_id'].unique()
            total_proposed = len(product_ids)
            
            activated_on_time = 0
            activated_later = 0
            rejected = 0
            
            # Анализируем каждый товар
            for pid in product_ids:
                # Получаем данные активации для этого товара
                product_activation = activation_df[activation_df['product_id'] == pid].copy()
                
                if product_activation.empty:
                    # Если нет данных активации, считаем отклоненным
                    rejected += 1
                    detail_results.append({
                        'revaluation_date': reval_date,
                        'planned_week': planned_week,
                        'product_id': pid,
                        'product_name': self.product_names.get(pid, f"ID {pid}"),
                        'status': 'rejected',
                        'status_reason': 'Нет данных активации',
                        'activation_week': None,
                        'plan_price': None,
                        'fact_price': None,
                        'deviation_pct': None,
                        'activation_status': 'Нет данных активации',
                        'fact_cost': None
                    })
                    continue
                
                # Проверяем, активировалась ли цена на плановой неделе
                planned_week_data = product_activation[product_activation['week_start'] == planned_week]
                if not planned_week_data.empty:
                    planned_row = planned_week_data.iloc[0]
                    if planned_row['Can_Use_In_Analysis']:
                        activated_on_time += 1
                        detail_results.append({
                            'revaluation_date': reval_date,
                            'planned_week': planned_week,
                            'product_id': pid,
                            'product_name': self.product_names.get(pid, f"ID {pid}"),
                            'status': 'activated_on_time',
                            'status_reason': f"Активирована на плановой неделе ({self.format_week(planned_week)})",
                            'activation_week': planned_week,
                            'plan_price': planned_row.get('Plan_Price_Current'),
                            'fact_price': planned_row.get('Fact_Price'),
                            'deviation_pct': planned_row.get('Deviation_From_Current_Pct'),
                            'activation_status': planned_row.get('Status'),
                            'fact_cost': planned_row.get('Fact_Cost')
                        })
                        continue
                
                # Проверяем, активировалась ли цена позже плановой недели
                later_weeks_data = product_activation[
                    (product_activation['week_start'] > planned_week) & 
                    (product_activation['Can_Use_In_Analysis'] == True)
                ]
                
                if not later_weeks_data.empty:
                    # Берем первую неделю, когда активировалась
                    first_activation_week_row = later_weeks_data.iloc[0]
                    first_activation_week = first_activation_week_row['week_start']
                    activated_later += 1
                    
                    # Также получаем данные с плановой недели для сравнения
                    planned_row_data = planned_week_data.iloc[0] if not planned_week_data.empty else None
                    
                    # Для активированных позже показываем данные с недели активации (основные)
                    # и данные с плановой недели (для сравнения)
                    detail_results.append({
                        'revaluation_date': reval_date,
                        'planned_week': planned_week,
                        'product_id': pid,
                        'product_name': self.product_names.get(pid, f"ID {pid}"),
                        'status': 'activated_later',
                        'status_reason': f"Активирована позже на неделе {self.format_week(first_activation_week)}",
                        'activation_week': first_activation_week,
                        'plan_price': first_activation_week_row.get('Plan_Price_Current'),
                        'fact_price': first_activation_week_row.get('Fact_Price'),
                        'deviation_pct': first_activation_week_row.get('Deviation_From_Current_Pct'),
                        'activation_status': first_activation_week_row.get('Status'),
                        'fact_cost': first_activation_week_row.get('Fact_Cost'),
                        'plan_price_planned_week': planned_row_data.get('Plan_Price_Current') if planned_row_data is not None else None,
                        'fact_price_planned_week': planned_row_data.get('Fact_Price') if planned_row_data is not None else None,
                        'deviation_pct_planned_week': planned_row_data.get('Deviation_From_Current_Pct') if planned_row_data is not None else None,
                        'activation_status_planned_week': planned_row_data.get('Status') if planned_row_data is not None else None
                    })
                    continue
                
                # Если не активировалась ни на плановой, ни позже - проверяем статус
                # Логика должна совпадать с get_activation_details:
                # Отклонено = никогда не было Can_Use_In_Analysis = True
                has_ok_status = (product_activation['Can_Use_In_Analysis'] == True).any()
                
                # Берем данные с плановой недели для отклоненных (если есть)
                # Иначе берем первую доступную неделю
                if not planned_week_data.empty:
                    planned_row_data = planned_week_data.iloc[0]
                else:
                    # Если нет данных на плановой неделе, берем первую неделю после плановой
                    after_planned = product_activation[product_activation['week_start'] >= planned_week]
                    if not after_planned.empty:
                        planned_row_data = after_planned.iloc[0]
                    else:
                        # Если нет данных после плановой недели, берем любую доступную
                        planned_row_data = product_activation.iloc[0] if not product_activation.empty else None
                
                if planned_row_data is None:
                    rejected += 1
                    detail_results.append({
                        'revaluation_date': reval_date,
                        'planned_week': planned_week,
                        'product_id': pid,
                        'product_name': self.product_names.get(pid, f"ID {pid}"),
                        'status': 'rejected',
                        'status_reason': 'Нет данных активации',
                        'activation_week': None,
                        'plan_price': None,
                        'fact_price': None,
                        'deviation_pct': None,
                        'activation_status': 'Нет данных активации',
                        'fact_cost': None
                    })
                    continue
                
                # Если никогда не было статуса "ОК", значит отклонено
                if not has_ok_status:
                    rejected += 1
                    # Берем статус с плановой недели (или первой доступной)
                    detail_results.append({
                        'revaluation_date': reval_date,
                        'planned_week': planned_week,
                        'product_id': pid,
                        'product_name': self.product_names.get(pid, f"ID {pid}"),
                        'status': 'rejected',
                        'status_reason': planned_row_data.get('Status', 'Неизвестный статус'),
                        'activation_week': None,
                        'plan_price': planned_row_data.get('Plan_Price_Current'),
                        'fact_price': planned_row_data.get('Fact_Price'),
                        'deviation_pct': planned_row_data.get('Deviation_From_Current_Pct'),
                        'activation_status': planned_row_data.get('Status'),
                        'fact_cost': planned_row_data.get('Fact_Cost')
                    })
                else:
                    # Это не должно произойти, так как мы уже проверили has_ok_status выше
                    # Но на всякий случай обработаем
                    rejected += 1
                    detail_results.append({
                        'revaluation_date': reval_date,
                        'planned_week': planned_week,
                        'product_id': pid,
                        'product_name': self.product_names.get(pid, f"ID {pid}"),
                        'status': 'rejected',
                        'status_reason': 'Неожиданный статус',
                        'activation_week': None,
                        'plan_price': planned_row_data.get('Plan_Price_Current'),
                        'fact_price': planned_row_data.get('Fact_Price'),
                        'deviation_pct': planned_row_data.get('Deviation_From_Current_Pct'),
                        'activation_status': planned_row_data.get('Status'),
                        'fact_cost': planned_row_data.get('Fact_Cost')
                    })
            
            # Добавляем сводку по переоценке
            summary_results.append({
                'revaluation_date': reval_date,
                'planned_week': planned_week,
                'planned_week_formatted': self.format_week(planned_week),
                'total_proposed': total_proposed,
                'activated_on_time': activated_on_time,
                'activated_later': activated_later,
                'rejected': rejected,
                'activation_rate_on_time': (activated_on_time / total_proposed * 100) if total_proposed > 0 else 0,
                'activation_rate_total': ((activated_on_time + activated_later) / total_proposed * 100) if total_proposed > 0 else 0
            })
        
        summary_df = pd.DataFrame(summary_results)
        detail_df = pd.DataFrame(detail_results)
        
        # Сортируем по дате переоценки
        if not summary_df.empty:
            summary_df = summary_df.sort_values('revaluation_date').reset_index(drop=True)
        if not detail_df.empty:
            detail_df = detail_df.sort_values(['revaluation_date', 'product_id']).reset_index(drop=True)
        
        return summary_df, detail_df

    def get_product_weekly_report_data(self, pid, activation_params):
        # 1. Get Activation Details for this PID (to get Plan/Fact/Cost prices with logic)
        # We call get_activation_details with all params
        activation_df = self.get_activation_details(specific_pid=pid, **activation_params)
        
        if not activation_df.empty:
            activation_df = activation_df.set_index('week_start')
        
        # 2. Get Basic Weekly Sales (Revenue, Profit, Simple Fact/Cost) for the whole period
        sales_df = self.weekly_sales[self.weekly_sales['product_id'] == pid].copy()
        if sales_df.empty:
            return pd.DataFrame()
            
        sales_df['profit'] = sales_df['revenue'] - sales_df['cost_volume']
        
        # 3. Merge
        # We want all weeks from sales_df (full history)
        # We join activation details
        
        result_rows = []
        sales_df = sales_df.sort_values('week_start')
        
        for _, row in sales_df.iterrows():
            week = row['week_start']
            
            # Defaults
            plan_price = None
            fact_price = row['revenue'] / row['quantity'] if row['quantity'] > 0 else 0
            cost_price = row['cost_volume'] / row['quantity'] if row['quantity'] > 0 else 0
            is_test_period = False
            
            # If we have activation data for this week, override prices
            if not activation_df.empty and week in activation_df.index:
                act_row = activation_df.loc[week]
                plan_price = act_row['Plan_Price_Current']
                fact_price = act_row['Fact_Price'] # This might be smart WAP
                cost_price = act_row['Fact_Cost']
                # Mark as test period based on existence in activation details (which are filtered by test start)
                is_test_period = True
            
            result_rows.append({
                'product_id': pid,
                'product_name': self.product_names.get(pid, f"ID {pid}"),
                'week_start': week, # Keep for sorting/highlighting
                'week_formatted': self.format_week(week),
                'Plan_price': plan_price,
                'Fact_price': fact_price,
                'Cost_price': cost_price,
                'Суммарная выручка': row['revenue'],
                'Суммарная прибыль': row['profit'],
                'is_test_period': is_test_period
            })
            
        return pd.DataFrame(result_rows)

    def get_weekly_details(self, pid, week, params):
        """
        Returns a dict with details for calculation of Price, Cost, Revenue, Profit.
        Used for detailed report breakdown.
        """
        
        # Unpack params
        threshold_pct = params.get('threshold_pct', 10)
        use_rounding = params.get('use_rounding', False)
        round_value = params.get('round_value', 90)
        round_direction = params.get('round_direction', "up")
        wap_from_change_date = params.get('wap_from_change_date', True)
        min_days_threshold = params.get('min_days_threshold', 3)
        
        # --- 1. Plan Price Logic ---
        product_prices = self.test_prices[self.test_prices['product_id'] == pid].sort_values('New_Price_Start')
        week_end = week + pd.Timedelta(days=6)
        active_price_rows = product_prices[product_prices['New_Price_Start'] <= week_end]
        
        plan_price_current = None
        is_price_change_week = False
        price_change_date = None
        raw_price = None
        
        if not active_price_rows.empty:
            current_record = active_price_rows.iloc[-1]
            raw_price = current_record['New_Price']
            
            # Simplified rounding reproduction
            if use_rounding:
                if raw_price is not None and raw_price > 0:
                    whole = int(raw_price)
                    frac = raw_price - whole
                    target = round_value / 100.0
                    if abs(frac - target) < 0.01:
                        plan_price_current = raw_price
                    else:
                        if round_direction == "nearest":
                            lower = float(whole) + target
                            upper = float(whole + 1) + target
                            plan_price_current = lower if abs(raw_price - lower) <= abs(raw_price - upper) else upper
                        elif round_direction == "down":
                            plan_price_current = float(whole) + target if frac >= target else float(whole - 1) + target
                        else:
                            plan_price_current = float(whole) + target if frac <= target else float(whole + 1) + target
                else:
                    plan_price_current = raw_price
            else:
                plan_price_current = raw_price

            # Check change week
            week_start_norm = week
            is_price_change_week = not product_prices[
                product_prices['New_Price_Start'].apply(
                    lambda x: (x - pd.Timedelta(days=x.weekday())).normalize()
                ) == week_start_norm
            ].empty
            
            if is_price_change_week:
                change_rows = product_prices[
                    product_prices['New_Price_Start'].apply(
                        lambda x: (x - pd.Timedelta(days=x.weekday())).normalize()
                    ) == week_start_norm
                ]
                if not change_rows.empty:
                    price_change_date = change_rows.iloc[0]['New_Price_Start']

        # Plan Price Description Text
        plan_calc_text = ""
        if plan_price_current is not None:
            date_str = ""
            if is_price_change_week and price_change_date is not None:
                day_name_map = {
                    0: 'Понедельник', 1: 'Вторник', 2: 'Среда', 
                    3: 'Четверг', 4: 'Пятница', 5: 'Суббота', 6: 'Воскресенье'
                }
                day_name = day_name_map.get(price_change_date.weekday(), "")
                date_str = f"Плановая цена была отправлена **{price_change_date.strftime('%d.%m.%Y')}** ({day_name})."
            else:
                date_str = "Плановая цена действует с предыдущих недель."

            if use_rounding:
                plan_calc_text = f"{date_str} Цена в файле: {raw_price:,.2f} ₽.\n"
                plan_calc_text += f"После округления итоговая плановая цена: :green[**{plan_price_current:,.2f} ₽**]"
            else:
                plan_calc_text = f"{date_str}\nИтоговая плановая цена (без округления): :green[**{plan_price_current:,.2f} ₽**]"
        else:
            plan_calc_text = "**Плановая цена:** Не найдена для этой недели."

        # --- 2. Fact Price Logic ---
        method_description = ""
        calc_logic_text = ""
        
        if wap_from_change_date:
            method_description = f"Умный WAP (с даты переоценки, если дней >= {min_days_threshold})"
        else:
             method_description = "Стандартный (средняя за неделю)"

        used_transactions = pd.DataFrame()
        week_sales = self.sales[
            (self.sales['product_id'] == pid) & 
            (self.sales['week_start'] == week)
        ].sort_values('recorded_on')
        
        if week_sales.empty:
             return {
                'week_formatted': self.format_week(week),
                'has_sales': False,
                'plan_text': plan_calc_text,
                'fact_text': "Нет продаж.",
                'cost_text': "Нет продаж.",
                'revenue_text': "0 ₽",
                'profit_text': "0 ₽",
                'transactions': pd.DataFrame()
            }
            
        # Smart WAP Logic
        if wap_from_change_date and is_price_change_week and price_change_date is not None and plan_price_current:
            first_occ_date = None
            sales_from_plan = week_sales[week_sales['recorded_on'] >= price_change_date]
            
            for _, row in sales_from_plan.iterrows():
                if abs(row['price'] - plan_price_current) / plan_price_current * 100 <= threshold_pct:
                    first_occ_date = row['recorded_on']
                    break
            
            if first_occ_date:
                days_working = (week + pd.Timedelta(days=6) - first_occ_date).days + 1
                if days_working >= min_days_threshold:
                    used_transactions = week_sales[week_sales['recorded_on'] >= first_occ_date]
                    calc_logic_text = f"Это неделя переоценки. Мы нашли первую продажу по новой цене (~{plan_price_current:.2f} ₽) от **{first_occ_date.strftime('%d.%m')}**. Так как цена работала достаточно дней ({days_working} дн.), для расчета фактической цены берем только продажи с этой даты (чтобы не смешивать со старой ценой)."
                else:
                    used_transactions = week_sales
                    calc_logic_text = f"Это неделя переоценки. Цена найдена, но работала слишком мало дней ({days_working} < {min_days_threshold}). Поэтому считаем цену продажи по всем продажам за неделю."
            else:
                 used_transactions = week_sales
                 calc_logic_text = f"Это неделя переоценки, но продаж по новой плановой цене (в допуске {threshold_pct}%) не найдено. Считаем цену продажи по всем продажам за неделю."
        else:
            used_transactions = week_sales
            if is_price_change_week:
                calc_logic_text = "Это неделя переоценки, но условия для «Умного расчета» не выполнены (или он выключен). Считаем цену продажи по всем продажам за неделю."
            else:
                calc_logic_text = "Обычная неделя (без переоценки). Считаем цену продажи по всем продажам за неделю."

        total_rev = used_transactions['revenue'].sum()
        total_qty = used_transactions['quantity'].sum()
        final_price = total_rev / total_qty if total_qty > 0 else 0
        
        # Build formula string for Fact Price
        formula_parts = []
        denominator_parts = []
        for _, row in used_transactions.iterrows():
            formula_parts.append(f"{row['price']:.2f} * {row['quantity']:.0f}")
            denominator_parts.append(f"{row['quantity']:.0f}")
        
        formula_str = " + ".join(formula_parts)
        denom_str = " + ".join(denominator_parts)
        
        fact_calc_text = f"{calc_logic_text}\n\n**Формула:** ({formula_str}) / ({denom_str}) = :green[**{final_price:,.2f} ₽**]"

        # Prepare transactions for Fact Price display (subset based on logic)
        # We want to show ALL transactions, but mark those used in calculation
        # used_transactions is already filtered. week_sales has everything.
        
        # Prepare transactions for Fact Price display (subset based on logic)
        # We want to show ALL transactions, but mark those used in calculation
        # used_transactions is already filtered. week_sales has everything.
        
        day_map = {0: 'Пн', 1: 'Вт', 2: 'Ср', 3: 'Чт', 4: 'Пт', 5: 'Сб', 6: 'Вс'}
        
        # Ensure all 7 days are present
        full_week_dates = pd.date_range(start=week, periods=7, freq='D')
        
        # Merge with existing sales to fill gaps
        # week_sales must be indexed by date or we merge on column
        week_sales_full = pd.merge(
            pd.DataFrame({'recorded_on': full_week_dates}),
            week_sales,
            on='recorded_on',
            how='left'
        )
        
        # Fill NaNs for missing days
        week_sales_full['quantity'] = week_sales_full['quantity'].fillna(0)
        week_sales_full['revenue'] = week_sales_full['revenue'].fillna(0)
        week_sales_full['price'] = week_sales_full['price'].fillna(0) # Or keep 0 to show no price
        week_sales_full['cost_volume'] = week_sales_full['cost_volume'].fillna(0)
        week_sales_full['cost_at_sale'] = week_sales_full['cost_at_sale'].fillna(0)
        
        # Re-derive derived columns
        week_sales_full['day_name'] = week_sales_full['recorded_on'].dt.dayofweek.map(day_map)
        week_sales_full['profit'] = week_sales_full['revenue'] - week_sales_full['cost_volume']
        week_sales_full['unit_cost'] = week_sales_full['cost_at_sale']
        
        # Determine used rows (based on original indices or dates)
        # Since we rebuilt the dataframe, let's use date matching with used_transactions
        if not used_transactions.empty:
            used_dates = set(used_transactions['recorded_on'])
            # Logic: row is used if date is in used_dates AND quantity > 0 (it was a real sale)
            # Actually, the requirement is to mark unused sales as grey.
            # Sales that are 0 are also "unused" effectively, or just empty.
            # Let's mark as used only if it corresponds to a record in used_transactions
            # Note: used_transactions comes from week_sales which might have multiple rows per date?
            # week_sales was aggregated? No, self.sales is raw transactions.
            # If multiple transactions per day, merge might explode?
            # Wait, week_sales comes from self.sales filtering. It can have multiple rows per day.
            # Our merge above with full_week_dates (1 row per day) against week_sales (N rows per day)
            # is a 1:N merge, so it preserves all transactions and adds days with no transactions.
            # This is correct.
            
            # Re-verify is_used logic
            # We can simply check if the 'recorded_on' and other props match, or simpler:
            # We know used_transactions is a subset of week_sales.
            # Any row in week_sales_full that came from week_sales AND was in used_transactions is True.
            # Any row added by merge (quantity=0) is False.
            
            # Let's map by index? Indices are lost in merge.
            # Let's use date range logic from Fact Price section:
            # used_transactions was filtered by: week_sales['recorded_on'] >= first_occ_date
            
            is_used_list = []
            # We need to reconstruct the filter logic or pass it down
            # Logic was: if days_working >= min ... used = sales >= first_occ ...
            
            # It's cleaner to re-apply the condition on the full set
            # Conditions for exclusion:
            # 1. Row has quantity 0 (filled day) -> False
            # 2. Row date < first_occurrence_date (if Smart WAP active) -> False
            
            # Re-determine start date for usage
            usage_start_date = week # Default start of week
            if wap_from_change_date and is_price_change_week and price_change_date is not None and plan_price_current:
                 # ... (re-use logic from above) ...
                 # We already did this to get used_transactions.
                 # Let's find the min date in used_transactions if it was filtered
                 if not used_transactions.empty and len(used_transactions) < len(week_sales):
                     # It was filtered. The filter was date-based.
                     usage_start_date = used_transactions['recorded_on'].min()
                 elif used_transactions.empty and not week_sales.empty:
                     # Filtered to nothing?
                     usage_start_date = week + pd.Timedelta(days=8) # Future
            
            for _, row in week_sales_full.iterrows():
                if row['quantity'] == 0:
                    is_used_list.append(False)
                elif row['recorded_on'] >= usage_start_date:
                    is_used_list.append(True)
                else:
                    is_used_list.append(False)
            
            week_sales_full['is_used'] = is_used_list
            
        else:
            # No used transactions (maybe no sales at all)
            week_sales_full['is_used'] = False

        
        # Fact Price Table (Fact_Price, Quantity, Day)
        fact_display_transactions = week_sales_full[['recorded_on', 'day_name', 'price', 'quantity', 'is_used']].copy()
        fact_display_transactions.columns = ['Дата', 'День недели', 'Цена', 'Кол-во', 'is_used']
        
        # Add Total row
        total_row_fact = pd.DataFrame([{
            'Дата': 'Итого', 
            'День недели': '',
            'Цена': final_price, 
            'Кол-во': total_qty,
            'is_used': True # Keep total visible
        }])
        fact_display_transactions = pd.concat([fact_display_transactions, total_row_fact], ignore_index=True)

        # Prepare transactions df for display (FULL week sales - Revenue, Cost, Profit)
        display_transactions = week_sales_full[['recorded_on', 'day_name', 'price', 'quantity', 'revenue', 'unit_cost', 'profit', 'is_used']].copy()
        display_transactions.columns = ['Дата', 'День недели', 'Цена', 'Кол-во', 'Выручка', 'Себестоимость ед.', 'Прибыль', 'is_used']
        
        # Calculate totals for full week
        # Need to handle NaN/0 correctly in week_sales
        full_week_rev = week_sales['revenue'].sum()
        full_week_qty = week_sales['quantity'].sum()
        
        # Calculate full_week_profit if 'profit' column exists, else calculate it
        if 'profit' in week_sales.columns:
            full_week_profit = week_sales['profit'].sum()
        else:
            # Fallback if week_sales structure changed unexpectedly
            full_week_profit = (week_sales['revenue'] - week_sales['cost_volume']).sum()
            
        total_cost_volume = week_sales['cost_volume'].sum() # Use ALL sales for cost, consistent with weekly profit
        
        avg_cost = total_cost_volume / full_week_qty if full_week_qty > 0 else 0
        
        # Add Total row to full transactions
        # We can't easily sum 'Price' or 'Unit Cost' as they are weighted. We show averages.
        total_row_full = pd.DataFrame([{
            'Дата': 'Итого',
            'День недели': '',
            'Цена': final_price, # Weighted avg price
            'Кол-во': full_week_qty,
            'Выручка': full_week_rev,
            'Себестоимость ед.': avg_cost, # Weighted avg cost
            'Прибыль': full_week_profit,
            'is_used': True
        }])
        display_transactions = pd.concat([display_transactions, total_row_full], ignore_index=True)

        # --- 3. Cost Calculation Logic ---
        # Cost is taken from 'cost_at_sale' column which was calculated via merge_asof in preprocess
        
        # Get raw cost data for display (from self.costs)
        # We need to find costs relevant to this week's sales.
        
        min_sale_date = week_sales['recorded_on'].min()
        max_sale_date = week_sales['recorded_on'].max()
        
        # Fetch relevant cost records (e.g. 2 weeks lookback to show where cost came from)
        search_start = min_sale_date - pd.Timedelta(days=14)
        relevant_costs = self.costs[
            (self.costs['product_id'] == pid) & 
            (self.costs['date'] >= search_start) & 
            (self.costs['date'] <= max_sale_date)
        ].sort_values('date')
        
        # Group costs by period to simplify display
        grouped_cost_rows = []
        if not relevant_costs.empty:
            # Create periods where cost is constant
            relevant_costs['group'] = (relevant_costs['cost'] != relevant_costs['cost'].shift()).cumsum()
            for _, group_df in relevant_costs.groupby('group'):
                start_d = group_df['date'].min()
                end_d = group_df['date'].max()
                cost_val = group_df['cost'].iloc[0]
                grouped_cost_rows.append({
                    'Период с': start_d,
                    'Период по': end_d, # Ideally this should be until next change, but we show available records
                    'Себестоимость': cost_val
                })
        
        cost_source_df = pd.DataFrame(grouped_cost_rows)

        cost_formula_parts = []
        cost_denominator_parts = []
        for _, row in week_sales.iterrows():
            cost_formula_parts.append(f"{row['cost_at_sale']:.2f} * {row['quantity']:.0f}")
            cost_denominator_parts.append(f"{row['quantity']:.0f}")
        
        cost_formula_str = " + ".join(cost_formula_parts)
        cost_denom_str = " + ".join(cost_denominator_parts)

        cost_text = f"""
        Себестоимость определяется для каждой транзакции продажи на дату продажи (из файла «Себестоимость»).
        
        **Формула средней:** ({cost_formula_str}) / ({cost_denom_str}) = :green[**{avg_cost:,.2f} ₽**]
        """

        # --- 4. Revenue Calculation ---
        # Formula string for revenue
        rev_formula_parts = []
        for _, row in week_sales.iterrows():
            rev_formula_parts.append(f"{row['price']:.2f} * {row['quantity']:.0f}")
        rev_formula_str = " + ".join(rev_formula_parts)
        
        rev_text = f"""
        Сумма выручки по всем чекам за неделю.
        **Формула:** {rev_formula_str} = :green[**{full_week_rev:,.2f} ₽**]
        """

        # --- 5. Profit Calculation ---
        # Formula string for profit
        # profit_text = f"""
        # Выручка ({full_week_rev:,.2f}) - Себестоимость объема продаж ({total_cost_volume:,.2f}).
        # **Итого:** :green[**{full_week_profit:,.2f} ₽**]
        # """
        profit_text = f"""
        Выручка ({full_week_rev:,.2f}) - Себестоимость объема продаж ({total_cost_volume:,.2f}).
        **Итого:** :green[**{full_week_profit:,.2f} ₽**]
        """

        # --- 6. Activation Check Logic ---
        # Re-calculate status logic specifically for this week to generate detailed text
        # Reuse logic from early part of method:
        # plan_price_current, is_price_change_week, etc.
        
        # We need deviation %
        deviation_val = 0
        if plan_price_current and plan_price_current > 0 and final_price > 0:
            deviation_val = abs(final_price - plan_price_current) / plan_price_current * 100
        
        # Determine Status string similar to get_activation_details
        status_str = "не определен"
        is_valid = False
        
        # Check against current
        matches_current = False
        if plan_price_current and plan_price_current > 0 and final_price > 0:
            if deviation_val <= threshold_pct:
                matches_current = True
        
        # We don't have full history of previous prices easily here without re-fetching all
        # But for the report we focus on current/main plan price
        
        activation_text = ""
        
        has_sales_bool = not week_sales.empty and week_sales['quantity'].sum() > 0
        
        if not has_sales_bool:
             activation_text = """
             **Статус:** Не определен (нет продаж).
             
             В эту неделю продаж не зафиксировано. Неделя не может быть использована для анализа ценовой эластичности, так как нет фактической цены сделки.
             """
        else:
            if plan_price_current is None:
                activation_text = """
                **Статус:** Плановая цена не найдена.
                
                Для этой недели не удалось определить действующую плановую цену. Проверьте файл «Тестовые цены».
                """
            else:
                threshold_msg = f"Порог отклонения: **{threshold_pct}%**"
                
                if matches_current:
                    status_str = "ОК"
                    is_valid = True
                    activation_text = f"""
                    **Статус:** :green[**{status_str}**]
                    
                    1. **Фактическая цена:** {final_price:,.2f} ₽
                    2. **Плановая цена:** {plan_price_current:,.2f} ₽
                    3. **Отклонение:** {deviation_val:.2f}%
                    
                    {threshold_msg}.
                    
                    **Вывод:** Отклонение фактической цены от плановой ({deviation_val:.2f}%) укладывается в допустимый порог ({threshold_pct}%). 
                    Неделя считается **активированной** и корректной для анализа.
                    """
                else:
                    status_str = "не та цена"
                    is_valid = False
                    activation_text = f"""
                    **Статус:** :red[**{status_str}**]
                    
                    1. **Фактическая цена:** {final_price:,.2f} ₽
                    2. **Плановая цена:** {plan_price_current:,.2f} ₽
                    3. **Отклонение:** {deviation_val:.2f}%
                    
                    {threshold_msg}.
                    
                    **Вывод:** Отклонение фактической цены ({deviation_val:.2f}%) превышает допустимый порог ({threshold_pct}%).
                    Неделя **не активирована** (Mismatched Price) и должна быть исключена из финального расчета эффекта, так как цена на полке не соответствовала плановой.
                    """

        return {
            'week_formatted': self.format_week(week),
            'has_sales': True,
            'plan_text': plan_calc_text,
            'fact_text': fact_calc_text,
            'cost_text': cost_text,
            'revenue_text': rev_text,
            'profit_text': profit_text,
            'activation_text': activation_text,
            'fact_transactions': fact_display_transactions, # WAP calc specific
            'full_transactions': display_transactions, # All week transactions with details
            'cost_source_data': cost_source_df
        }

    # Alias for backward compatibility if needed, but we should update app.py
    get_wap_calculation_example = get_weekly_details

    def calculate(self, use_stock_filter=False, stock_threshold_pct=50, 
                  pre_test_weeks_count=3, pre_test_stock_threshold=70,
                  contiguous_pre_test=True, activation_threshold=None,
                  activation_use_rounding=False, activation_round_value=90,
                  activation_round_direction="up",
                  activation_wap_from_change_date=True,
                  activation_min_days_threshold=3,
                  test_use_week_values=True):
        results = []
        
        threshold_factor = 1 - (stock_threshold_pct / 100.0)
        
        activation_status_map = {}
        if activation_threshold is not None:
            activation_df = self.get_activation_details(
                activation_threshold,
                use_rounding=activation_use_rounding,
                round_value=activation_round_value,
                round_direction=activation_round_direction,
                wap_from_change_date=activation_wap_from_change_date,
                min_days_threshold=activation_min_days_threshold
            )
            if not activation_df.empty:
                activation_status_map = {
                    (row['product_id'], row['week_start']): row['Status']
                    for _, row in activation_df.iterrows()
                }
        
        for pid in self.test_product_ids:
            product_test_info = self.test_prices[self.test_prices['product_id'] == pid]
            if product_test_info.empty:
                continue
            start_date = product_test_info['New_Price_Start'].min()
            
            start_week_monday = start_date - pd.Timedelta(days=start_date.weekday())
            start_week_monday = pd.to_datetime(start_week_monday).normalize()
            
            if start_date.weekday() == 0:
                test_period_start_week = start_week_monday
                pre_test_search_end = start_week_monday - pd.Timedelta(weeks=1)
            else:
                test_period_start_week = start_week_monday + pd.Timedelta(weeks=1)
                pre_test_search_end = start_week_monday - pd.Timedelta(weeks=1)
            
            # --- 1. FIND DYNAMIC PRE-TEST PERIOD ---
            pre_test_weeks_list = self.find_valid_pre_test_period(
                pid, pre_test_search_end, pre_test_weeks_count, pre_test_stock_threshold,
                contiguous=contiguous_pre_test
            )
            
            if pre_test_weeks_list is None:
                continue
            
            # Use list of specific weeks
            pre_test_start_week = min(pre_test_weeks_list)
            pre_test_end_week = max(pre_test_weeks_list)
            
            # --- 2. Determine Baseline Stock ---
            pre_test_stocks = [self.stock_lookup.get((pid, w), 0) for w in pre_test_weeks_list]
            baseline_stock = np.mean(pre_test_stocks) if pre_test_stocks else 0
            
            # --- 3. Calculate R_tb and P_tb ---
            r_tb_sum = 0
            p_tb_sum = 0  # Profit Test Base Sum
            valid_pre_weeks_count = 0
            for w in pre_test_weeks_list:
                if (pid, w) in self.sales_lookup:
                    row = self.weekly_sales[(self.weekly_sales['product_id'] == pid) & (self.weekly_sales['week_start'] == w)]
                    rev = row['revenue'].sum()
                    cost_vol = row['cost_volume'].sum() if 'cost_volume' in row.columns else 0
                    
                    r_tb_sum += rev
                    p_tb_sum += (rev - cost_vol)
                    valid_pre_weeks_count += 1
            
            if valid_pre_weeks_count == 0:
                continue 
                
            r_tb = r_tb_sum / valid_pre_weeks_count
            p_tb = p_tb_sum / valid_pre_weeks_count
            
            r_cb_sum = 0
            p_cb_sum = 0 # Profit Control Base Sum
            for w in pre_test_weeks_list:
                if (pid, w) in self.sales_lookup:
                    row_c = self.control_weekly_sales[self.control_weekly_sales['week_start'] == w]
                    rev_c = row_c['control_revenue'].sum()
                    cost_vol_c = row_c['control_cost_volume'].sum() if 'control_cost_volume' in row_c.columns else 0
                    
                    r_cb_sum += rev_c
                    p_cb_sum += (rev_c - cost_vol_c)
                    
            r_cb = r_cb_sum / valid_pre_weeks_count
            p_cb = p_cb_sum / valid_pre_weeks_count
            
            # Allow P_tb or P_cb to be calculated even if 0, but handle division later
            if r_tb == 0 or r_cb == 0:
                continue
                
            # Iterate through Test Weeks
            available_weeks = self.control_weekly_sales[
                self.control_weekly_sales['week_start'] >= test_period_start_week
            ]['week_start'].sort_values().unique()
            
            week_status_map = {} 
            for w in available_weeks:
                week_stock = self.stock_lookup.get((pid, w), 0)
                is_excluded = False
                if use_stock_filter:
                    if week_stock < (baseline_stock * threshold_factor):
                        is_excluded = True
                activation_status = activation_status_map.get((pid, w), "")
                is_not_our_price = activation_status.startswith('не та цена')
                
                week_status_map[w] = {
                    'stock': week_stock,
                    'is_excluded': is_excluded,
                    'is_not_our_price': is_not_our_price
                }

            for week in available_weeks:
                current_status = week_status_map[week]
                week_formatted = self.format_week(week)
                
                if current_status['is_excluded'] or current_status['is_not_our_price']:
                    results.append({
                        'product_id': pid,
                        'product_name': self.product_names.get(pid, f"ID {pid}"),
                        'week_start': week,
                        'week_formatted': week_formatted,
                        'R_tt': 0, 'R_tb': r_tb, 'R_ct': 0, 'R_cb': r_cb,
                        'Effect_Revenue_pct': 0,
                        'Effect_Profit_pct': 0,
                        'Fact_Revenue': 0,
                        'Product_Profit': 0, 
                        'Control_Fact_Revenue': 0,
                        'Control_Profit': 0,
                        'Abs_Effect_Revenue': 0,
                        'Abs_Effect_Profit': 0,
                        'Test_Start_Date': start_date,
                        'PreTest_Start': pre_test_start_week,
                        'PreTest_End': pre_test_end_week,
                        'PreTest_Weeks': pre_test_weeks_list, # Store list for timeline
                        'Avg_Stock': current_status['stock'],
                        'Baseline_Stock': baseline_stock,
                        'Is_Excluded': True
                    })
                    continue

                expected_test_weeks = pd.date_range(start=test_period_start_week, end=week, freq='W-MON')
                
                valid_test_weeks = []
                for w_hist in expected_test_weeks:
                    if week_status_map.get(w_hist, {}).get('is_excluded', False):
                        continue
                    if week_status_map.get(w_hist, {}).get('is_not_our_price', False):
                        continue
                    valid_test_weeks.append(w_hist)
                
                if not valid_test_weeks:
                    r_tt = 0
                    r_ct = 0
                    p_tt = 0 # Profit Test Test
                    p_ct = 0 # Profit Control Test
                else:
                    if test_use_week_values:
                        # Revenue
                        r_tt = self.weekly_sales[
                            (self.weekly_sales['product_id'] == pid) & 
                            (self.weekly_sales['week_start'] == week)
                        ]['revenue'].sum()
                        r_ct = self.control_weekly_sales[
                            self.control_weekly_sales['week_start'] == week
                        ]['control_revenue'].sum()
                        
                        # Profit
                        ws_row = self.weekly_sales[
                            (self.weekly_sales['product_id'] == pid) & 
                            (self.weekly_sales['week_start'] == week)
                        ]
                        p_tt = 0
                        if not ws_row.empty:
                            p_tt = ws_row['revenue'].sum() - ws_row['cost_volume'].sum()

                        cs_row = self.control_weekly_sales[
                            self.control_weekly_sales['week_start'] == week
                        ]
                        p_ct = 0
                        if not cs_row.empty:
                            p_ct = cs_row['control_revenue'].sum() - cs_row['control_cost_volume'].sum()
                            
                    else:
                        # Aggregated logic for Revenue
                        r_tt_sum = self.weekly_sales[
                            (self.weekly_sales['product_id'] == pid) & 
                            (self.weekly_sales['week_start'].isin(valid_test_weeks))
                        ]['revenue'].sum()
                        r_tt = r_tt_sum / len(valid_test_weeks)
                        
                        r_ct_sum = self.control_weekly_sales[
                            self.control_weekly_sales['week_start'].isin(valid_test_weeks)
                        ]['control_revenue'].sum()
                        r_ct = r_ct_sum / len(valid_test_weeks)
                        
                        # Aggregated logic for Profit
                        # Not strictly requested to support 'test_use_week_values=False' for profit perfectly, 
                        # but adding similar logic for consistency.
                        
                        ws_rows = self.weekly_sales[
                            (self.weekly_sales['product_id'] == pid) & 
                            (self.weekly_sales['week_start'].isin(valid_test_weeks))
                        ]
                        p_tt_sum = (ws_rows['revenue'].sum() - ws_rows['cost_volume'].sum()) if not ws_rows.empty else 0
                        p_tt = p_tt_sum / len(valid_test_weeks)
                        
                        cs_rows = self.control_weekly_sales[
                            self.control_weekly_sales['week_start'].isin(valid_test_weeks)
                        ]
                        p_ct_sum = (cs_rows['control_revenue'].sum() - cs_rows['control_cost_volume'].sum()) if not cs_rows.empty else 0
                        p_ct = p_ct_sum / len(valid_test_weeks)

                # --- Revenue Effect Calculation ---
                if r_tt == 0:
                    term1 = -1
                    if r_ct == 0 and r_cb == 0:
                         term2 = 0
                    elif r_cb == 0:
                         term2 = 0
                    else:
                         term2 = (r_ct / r_cb) - 1
                    effect_revenue = term1 - term2
                else:
                    term1 = (r_tt / r_tb) - 1
                    term2 = (r_ct / r_cb) - 1
                    effect_revenue = term1 - term2
                
                # --- Profit Effect Calculation ---
                # P_tb, P_cb, P_tt, P_ct can be negative or zero.
                # Standard formula: Effect = (P_tt / P_tb) - (P_ct / P_cb) is risky if signs flip.
                # Keeping simple logic similar to revenue for now, as requested "same logic as revenue".
                # Handling zero division carefully.
                
                # Warning: Profit can be negative, percentages might be misleading.
                # Assuming standard uplift formula is desired.
                
                if p_tb == 0:
                     # Base profit is 0 -> Infinite effect or undefined. Set to 0 to avoid explosion.
                     effect_profit = 0
                else:
                    term1_p = (p_tt / p_tb) - 1
                    
                    if p_cb == 0:
                        term2_p = 0
                    else:
                        term2_p = (p_ct / p_cb) - 1
                        
                    effect_profit = term1_p - term2_p
                
                current_week_sales = self.weekly_sales[
                    (self.weekly_sales['product_id'] == pid) & 
                    (self.weekly_sales['week_start'] == week)
                ]
                fact_revenue = current_week_sales['revenue'].sum() if not current_week_sales.empty else 0
                test_cost_volume = current_week_sales['cost_volume'].sum() if not current_week_sales.empty else 0
                product_profit = fact_revenue - test_cost_volume
                
                current_week_control_sales = self.control_weekly_sales[
                    (self.control_weekly_sales['week_start'] == week)
                ]
                control_fact_revenue = current_week_control_sales['control_revenue'].sum() if not current_week_control_sales.empty else 0
                control_cost_volume = current_week_control_sales['control_cost_volume'].sum() if not current_week_control_sales.empty else 0
                control_profit = control_fact_revenue - control_cost_volume
                
                abs_effect_revenue = effect_revenue * fact_revenue
                abs_effect_profit = effect_profit * product_profit
                
                results.append({
                    'product_id': pid,
                    'product_name': self.product_names.get(pid, f"ID {pid}"),
                    'week_start': week,
                    'week_formatted': week_formatted,
                    'R_tt': r_tt,
                    'R_tb': r_tb,
                    'R_ct': r_ct,
                    'R_cb': r_cb,
                    'Effect_Revenue_pct': effect_revenue,
                    'Effect_Profit_pct': effect_profit,
                    'Fact_Revenue': fact_revenue,
                    'Product_Profit': product_profit,
                    'Control_Fact_Revenue': control_fact_revenue,
                    'Control_Profit': control_profit,
                    'Abs_Effect_Revenue': abs_effect_revenue,
                    'Abs_Effect_Profit': abs_effect_profit,
                    'Test_Start_Date': start_date,
                    'PreTest_Start': pre_test_start_week,
                    'PreTest_End': pre_test_end_week,
                    'PreTest_Weeks': pre_test_weeks_list,
                    'Avg_Stock': current_status['stock'],
                    'Baseline_Stock': baseline_stock,
                    'Is_Excluded': False
                })
        
        # Calculate Total_Product_Effect based on mode
        self.results_df = pd.DataFrame(results)
        if not self.results_df.empty:
            self.results_df['Total_Effect_Revenue'] = 0.0
            self.results_df['Total_Effect_Profit'] = 0.0
            
            for pid in self.results_df['product_id'].unique():
                pid_mask = self.results_df['product_id'] == pid
                prod_data = self.results_df[pid_mask]
                
                # Filter valid weeks for effect calculation
                valid_prod_data = prod_data[prod_data['Is_Excluded'] == False]
                
                if valid_prod_data.empty:
                    continue
                
                if test_use_week_values:
                    # Sum of weekly effects
                    total_effect_rev = valid_prod_data['Abs_Effect_Revenue'].sum()
                    total_effect_prof = valid_prod_data['Abs_Effect_Profit'].sum()
                else:
                    # Effect of last week * Total Revenue
                    last_week_row = valid_prod_data.sort_values('week_start').iloc[-1]
                    
                    final_effect_pct_rev = last_week_row['Effect_Revenue_pct']
                    total_revenue = valid_prod_data['Fact_Revenue'].sum()
                    total_effect_rev = final_effect_pct_rev * total_revenue
                    
                    final_effect_pct_prof = last_week_row['Effect_Profit_pct']
                    total_profit = valid_prod_data['Product_Profit'].sum()
                    total_effect_prof = final_effect_pct_prof * total_profit
                
                self.results_df.loc[pid_mask, 'Total_Effect_Revenue'] = total_effect_rev
                self.results_df.loc[pid_mask, 'Total_Effect_Profit'] = total_effect_prof
                
        return self.results_df

    def get_summary(self):
        if not hasattr(self, 'results_df') or self.results_df.empty:
            return None
        
        valid_df = self.results_df[self.results_df['Is_Excluded'] == False]
        
        # Sum Total_Product_Effect for each unique product
        unique_products_df = valid_df.drop_duplicates(subset=['product_id'])
        
        if 'Total_Effect_Revenue' in unique_products_df.columns:
            total_abs_effect_revenue = unique_products_df['Total_Effect_Revenue'].sum()
            total_abs_effect_profit = unique_products_df['Total_Effect_Profit'].sum()
        else:
            total_abs_effect_revenue = valid_df['Abs_Effect_Revenue'].sum()
            total_abs_effect_profit = valid_df['Abs_Effect_Profit'].sum()
            
        total_fact_revenue = valid_df['Fact_Revenue'].sum()
        total_fact_profit = valid_df['Product_Profit'].sum()
        
        # Revenue %
        revenue_without_effect = total_fact_revenue - total_abs_effect_revenue
        if revenue_without_effect == 0:
            effect_revenue_pct = 0
        else:
            ratio = total_fact_revenue / revenue_without_effect
            effect_revenue_pct = (ratio - 1) * 100
            
        # Profit %
        profit_without_effect = total_fact_profit - total_abs_effect_profit
        if profit_without_effect == 0:
            effect_profit_pct = 0
        else:
            ratio_p = total_fact_profit / profit_without_effect
            effect_profit_pct = (ratio_p - 1) * 100
            
        # --- Extended Stats for Dashboard ---
        # 1. Total Turnover (All products in the file for the test period)
        # Find min/max test dates
        if not valid_df.empty:
            test_start_global = valid_df['week_start'].min()
            test_end_global = valid_df['week_start'].max()
            test_duration_weeks = (test_end_global - test_start_global).days / 7 + 1
            
            # Calculate total revenue for ALL products in the sales file during the test period
            all_sales_in_period = self.weekly_sales[
                (self.weekly_sales['week_start'] >= test_start_global) &
                (self.weekly_sales['week_start'] <= test_end_global)
            ]
            global_revenue = all_sales_in_period['revenue'].sum()
            
            # Global Profit calculation (revenue - cost_volume)
            global_profit = (all_sales_in_period['revenue'] - all_sales_in_period['cost_volume']).sum()
            
        else:
            test_duration_weeks = 0
            global_revenue = 0
            global_profit = 0
            
        # Test Share
        test_share_revenue = (total_fact_revenue / global_revenue * 100) if global_revenue > 0 else 0
        test_share_profit = (total_fact_profit / global_profit * 100) if global_profit > 0 else 0
        
        # 2. Counts
        tested_count = len(self.test_product_ids)
        excluded_products_count = len(self.results_df['product_id'].unique()) - len(unique_products_df)
        excluded_weeks_count = len(self.results_df) - len(valid_df)
        
        # 3. Price Changes Analysis
        # Count unique start dates in test_prices
        price_changes = self.test_prices.groupby('New_Price_Start')['product_id'].count()
        price_changes_count = len(price_changes)
        products_per_change = price_changes.to_dict()
        
        # 4. Growth vs Decline
        # We need total effect per product
        if 'Total_Effect_Revenue' in unique_products_df.columns:
            prod_effects = unique_products_df[['product_id', 'Total_Effect_Revenue', 'Total_Effect_Profit']]
        else:
            # Fallback if not calculated in calculate()
            prod_effects = valid_df.groupby('product_id').agg({
                'Abs_Effect_Revenue': 'sum',
                'Abs_Effect_Profit': 'sum'
            }).reset_index()
            prod_effects.rename(columns={
                'Abs_Effect_Revenue': 'Total_Effect_Revenue',
                'Abs_Effect_Profit': 'Total_Effect_Profit'
            }, inplace=True)
            
        growth_mask = prod_effects['Total_Effect_Revenue'] > 0
        growth_df = prod_effects[growth_mask]
        decline_df = prod_effects[~growth_mask]
        
        growth_stats = {
            'count': len(growth_df),
            'revenue_effect': growth_df['Total_Effect_Revenue'].sum(),
            'profit_effect': growth_df['Total_Effect_Profit'].sum()
        }
        
        decline_stats = {
            'count': len(decline_df),
            'revenue_effect': decline_df['Total_Effect_Revenue'].sum(),
            'profit_effect': decline_df['Total_Effect_Profit'].sum()
        }

        return {
            'total_abs_effect_revenue': total_abs_effect_revenue,
            'total_abs_effect_profit': total_abs_effect_profit,
            'total_fact_revenue': total_fact_revenue,
            'total_fact_profit': total_fact_profit,
            'effect_revenue_pct': effect_revenue_pct,
            'effect_profit_pct': effect_profit_pct,
            
            'revenue_without_effect': revenue_without_effect,
            'profit_without_effect': profit_without_effect,
            
            'global_revenue': global_revenue,
            'global_profit': global_profit,
            'test_share_revenue': test_share_revenue,
            'test_share_profit': test_share_profit,
            
            'tested_count': tested_count,
            'excluded_products_count': excluded_products_count,
            'excluded_weeks_count': excluded_weeks_count,
            
            'test_duration_weeks': test_duration_weeks,
            'price_changes_count': price_changes_count,
            'products_per_change': products_per_change,
            
            'growth_stats': growth_stats,
            'decline_stats': decline_stats
        }

    def get_control_group_info(self):
        if not hasattr(self, 'control_product_ids'):
            return pd.DataFrame()
        control_list = []
        for pid in self.control_product_ids:
            control_list.append({
                'product_id': pid,
                'name': self.product_names.get(pid, "Unknown")
            })
        return pd.DataFrame(control_list)

    def get_simple_effect_details(self, pid, use_week_values=True):
        """
        Calculates simple effect (Test vs Pre-Test) metrics for each valid test week.
        Returns a DataFrame with details.
        """
        if not hasattr(self, 'results_df') or self.results_df.empty:
            return pd.DataFrame()
            
        prod_res = self.results_df[self.results_df['product_id'] == pid]
        if prod_res.empty:
            return pd.DataFrame()
            
        # Get baseline values (R_tb, P_tb are constant for the product)
        r_tb = prod_res['R_tb'].iloc[0]
        # Get Control baseline values (R_cb, P_cb) - Use R_cb from results if available, else calculate
        r_cb = prod_res['R_cb'].iloc[0] if 'R_cb' in prod_res.columns else 0
        
        # Calculate P_tb and P_cb manually if needed (as they might not be stored explicitly per row correctly or needed average)
        pre_test_weeks = prod_res['PreTest_Weeks'].iloc[0]
        if not isinstance(pre_test_weeks, list):
            pre_test_weeks = []
            
        p_tb = 0
        p_cb = 0
        valid_count = 0
        if pre_test_weeks:
            for w in pre_test_weeks:
                # Test Product Base Profit
                sales_row = self.weekly_sales[
                    (self.weekly_sales['product_id'] == pid) & 
                    (self.weekly_sales['week_start'] == w)
                ]
                if not sales_row.empty:
                    rev = sales_row['revenue'].sum()
                    cost = sales_row['cost_volume'].sum() if 'cost_volume' in sales_row.columns else 0
                    p_tb += (rev - cost)
                
                # Control Group Base Profit
                control_row = self.control_weekly_sales[
                    self.control_weekly_sales['week_start'] == w
                ]
                if not control_row.empty:
                    rev_c = control_row['control_revenue'].sum()
                    cost_c = control_row['control_cost_volume'].sum() if 'control_cost_volume' in control_row.columns else 0
                    p_cb += (rev_c - cost_c)
                    
                valid_count += 1
                
            if valid_count > 0:
                p_tb /= valid_count
                p_cb /= valid_count
        
        # Filter for valid test weeks only
        valid_test_weeks_df = prod_res[prod_res['Is_Excluded'] == False]
        
        # If using averages, calculate them once for both Test and Control
        avg_test_rev = 0
        avg_test_prof = 0
        avg_control_rev = 0
        avg_control_prof = 0
        
        if not use_week_values and not valid_test_weeks_df.empty:
            # Re-calculate actual sums per week
            avg_test_rev = valid_test_weeks_df['Fact_Revenue'].mean()
            avg_test_prof = valid_test_weeks_df['Product_Profit'].mean()
            avg_control_rev = valid_test_weeks_df['Control_Fact_Revenue'].mean()
            avg_control_prof = valid_test_weeks_df['Control_Profit'].mean()
        
        effect_rows = []
        
        for _, row in valid_test_weeks_df.iterrows():
            week = row['week_start']
            
            # --- 1. Test Group Data ---
            sales_row = self.weekly_sales[
                (self.weekly_sales['product_id'] == pid) & 
                (self.weekly_sales['week_start'] == week)
            ]
            fact_rev_real = 0
            fact_profit_real = 0
            if not sales_row.empty:
                fact_rev_real = sales_row['revenue'].sum()
                cost_vol = sales_row['cost_volume'].sum() if 'cost_volume' in sales_row.columns else 0
                fact_profit_real = fact_rev_real - cost_vol
            
            # --- 2. Control Group Data ---
            control_row = self.control_weekly_sales[
                self.control_weekly_sales['week_start'] == week
            ]
            control_rev_real = 0
            control_profit_real = 0
            if not control_row.empty:
                control_rev_real = control_row['control_revenue'].sum()
                cost_c = control_row['control_cost_volume'].sum() if 'control_cost_volume' in control_row.columns else 0
                control_profit_real = control_rev_real - cost_c

            # Determine values used for calculation based on mode
            if use_week_values:
                calc_rev = fact_rev_real
                calc_prof = fact_profit_real
                calc_control_rev = control_rev_real
                calc_control_prof = control_profit_real
                mode_comment = "(по этой неделе)"
            else:
                calc_rev = avg_test_rev
                calc_prof = avg_test_prof
                calc_control_rev = avg_control_rev
                calc_control_prof = avg_control_prof
                mode_comment = "(среднее за тест)"
            
            # --- 3. Uplift Calculations ---
            
            # Test Uplifts
            uplift_rev_pct = (calc_rev / r_tb - 1) if r_tb != 0 else 0
            uplift_prof_pct = (calc_prof / p_tb - 1) if p_tb != 0 else 0
            
            # Control Uplifts
            uplift_control_rev_pct = (calc_control_rev / r_cb - 1) if r_cb != 0 else 0
            uplift_control_prof_pct = (calc_control_prof / p_cb - 1) if p_cb != 0 else 0
            
            # Net Effects (Simple Difference)
            net_effect_rev = uplift_rev_pct - uplift_control_rev_pct
            net_effect_prof = uplift_prof_pct - uplift_control_prof_pct
            
            # Net Absolute Effects (Formula: Net_Effect_% * Fact_Real)
            # As per request: "like in Analysis of Goods" -> Effect % * Fact
            net_abs_rev = net_effect_rev * fact_rev_real
            net_abs_prof = net_effect_prof * fact_profit_real
            
            effect_rows.append({
                'week_formatted': row['week_formatted'],
                'week_start': week,
                'Mode_Comment': mode_comment,
                
                # Test Values
                'Fact_Revenue_Real': fact_rev_real,
                'Fact_Profit_Real': fact_profit_real,
                'Calc_Revenue': calc_rev,
                'Calc_Profit': calc_prof,
                'PreTest_Avg_Revenue': r_tb,
                'PreTest_Avg_Profit': p_tb,
                'Uplift_Revenue_Pct': uplift_rev_pct,
                'Uplift_Profit_Pct': uplift_prof_pct,
                
                # Control Values
                'Control_Revenue_Real': control_rev_real,
                'Control_Profit_Real': control_profit_real,
                'Calc_Control_Revenue': calc_control_rev,
                'Calc_Control_Profit': calc_control_prof,
                'Control_Avg_Revenue': r_cb,
                'Control_Avg_Profit': p_cb,
                'Control_Uplift_Revenue_Pct': uplift_control_rev_pct,
                'Control_Uplift_Profit_Pct': uplift_control_prof_pct,
                
                # Net Effects
                'Net_Effect_Revenue_Pct': net_effect_rev,
                'Net_Effect_Profit_Pct': net_effect_prof,
                'Net_Abs_Effect_Revenue': net_abs_rev,
                'Net_Abs_Effect_Profit': net_abs_prof
            })
            
        return pd.DataFrame(effect_rows)

    def get_product_timeline(self, pid):
        if not hasattr(self, 'results_df') or pid not in self.test_product_ids:
            return pd.DataFrame()
            
        prod_res = self.results_df[self.results_df['product_id'] == pid]
        
        # Product Info
        product_test_info = self.test_prices[self.test_prices['product_id'] == pid]
        if product_test_info.empty:
            return pd.DataFrame()
        
        # Determine theoretical Test Start (price change date)
        start_date = product_test_info['New_Price_Start'].min()
        start_week_monday = pd.to_datetime(start_date - pd.Timedelta(days=start_date.weekday())).normalize()
        
        if start_date.weekday() == 0:
            test_start_week = start_week_monday
            pre_test_search_end = start_week_monday - pd.Timedelta(weeks=1)
        else:
            test_start_week = start_week_monday + pd.Timedelta(weeks=1)
            pre_test_search_end = start_week_monday - pd.Timedelta(weeks=1)
            
        # Determine Period Boundaries
        pre_test_weeks_set = set()
        
        if not prod_res.empty:
            if 'PreTest_Weeks' in prod_res.columns:
                val = prod_res['PreTest_Weeks'].iloc[0]
                if isinstance(val, list) or isinstance(val, np.ndarray):
                    pre_test_weeks_set = set(val)
        
        test_end_week = self.weekly_sales['week_start'].max()
        
        if pre_test_weeks_set:
            start_display = min(pre_test_weeks_set) - pd.Timedelta(weeks=4)
        else:
            start_display = pre_test_search_end - pd.Timedelta(weeks=8)
            
        end_display = test_end_week
        
        # --- Prepare Sales & Profit Data ---
        prod_data = self.weekly_sales[self.weekly_sales['product_id'] == pid].set_index('week_start')
        prod_sales = prod_data['revenue']
        # Calculate Product Profit
        prod_cost = prod_data['cost_volume'] if 'cost_volume' in prod_data.columns else pd.Series(0, index=prod_data.index)
        prod_profit = prod_sales - prod_cost

        control_data = self.control_weekly_sales.set_index('week_start')
        control_sales = control_data['control_revenue']
        # Calculate Control Profit
        control_cost = control_data['control_cost_volume'] if 'control_cost_volume' in control_data.columns else pd.Series(0, index=control_data.index)
        control_profit = control_sales - control_cost

        prod_stock = self.weekly_stock[self.weekly_stock['product_id'] == pid].set_index('week_start')['stock']
        
        prod_res_idx = pd.DataFrame()
        if not prod_res.empty:
            prod_res_idx = prod_res.set_index('week_start')
        
        weeks = pd.date_range(start=start_display, end=end_display, freq='W-MON')
        
        timeline = []
        for w in weeks:
            w_norm = w.normalize()
            
            period_label = 'Other'
            
            if not prod_res.empty:
                if w_norm in pre_test_weeks_set:
                    period_label = 'Pre-Test'
                elif w_norm >= test_start_week:
                    if w_norm in prod_res_idx.index and prod_res_idx.loc[w_norm]['Is_Excluded']:
                        period_label = 'LowStock_Test'
                    else:
                        period_label = 'Test'
                elif w_norm < test_start_week:
                    if pre_test_weeks_set and w_norm > max(pre_test_weeks_set):
                        period_label = 'Transit'
                    else:
                        period_label = 'LowStock_Before'
            else:
                if w_norm < test_start_week:
                    period_label = 'LowStock_Before'
                elif w_norm >= test_start_week:
                    period_label = 'Test'
            
            has_sales = (pid, w_norm) in self.sales_lookup
            
            abs_effect_revenue = 0
            if w_norm in prod_res_idx.index:
                abs_effect_revenue = prod_res_idx.loc[w_norm]['Abs_Effect_Revenue']
                
            abs_effect_profit = 0
            if w_norm in prod_res_idx.index:
                abs_effect_profit = prod_res_idx.loc[w_norm]['Abs_Effect_Profit']
            
            timeline.append({
                'week_start': w_norm,
                'week_formatted': self.format_week(w_norm),
                'period_label': period_label,
                'product_revenue': prod_sales.get(w_norm, 0),
                'Product_Profit': prod_profit.get(w_norm, 0),
                'control_revenue': control_sales.get(w_norm, 0),
                'Control_Profit': control_profit.get(w_norm, 0),
                'avg_stock': prod_stock.get(w_norm, 0),
                'abs_effect_revenue': abs_effect_revenue,
                'abs_effect_profit': abs_effect_profit,
                'has_sales': has_sales,
                'is_excluded': period_label in ['LowStock_Test', 'LowStock_Before']
            })
            
        return pd.DataFrame(timeline)
