import pandas as pd
import os
import re
import warnings
from datetime import datetime, timedelta

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- CONFIGURATION ---
OUTPUT_FILENAME = "Monthly_Global_Analysis.xlsx"

# LIST OF EMPLOYEES TO EXCLUDE (Case insensitive)
EXCLUDED_EMPLOYEES = [
    # "ABOU HASNAA", 
    "HMOURI ALI"
]

# CODES THAT SIGNIFY AN "OUVRIER" (Worker)
OUVRIER_CODES = ['130', '140', '141', '131']

def clean_name_string(name):
    """Normalizes names to ensure matching works despite spaces/hidden chars."""
    if not name:
        return ""
    name = str(name).upper()
    name = name.replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ')
    name = re.sub(r'\s+', ' ', name)
    return name.strip()

def parse_scan_times_from_string(times_str):
    """Parses scan string to extract all time entries from pipe-separated format."""
    if not times_str or pd.isna(times_str):
        return [], 0
    times = times_str.split('|') if times_str else []
    times = [t for t in times if t and ':' in t]  # Filter out empty strings and ensure time format
    return times, len(times)

def calculate_hours_from_times_list(times):
    """Calculates total worked hours from a list of 'HH:MM' strings."""
    if not times:
        return 0.0
    
    total_seconds = 0
    for i in range(0, len(times) - 1, 2):
        try:
            t_in = datetime.strptime(times[i], '%H:%M')
            t_out = datetime.strptime(times[i+1], '%H:%M')
            if t_out < t_in: t_out += timedelta(days=1)
            total_seconds += (t_out - t_in).total_seconds()
        except:
            continue
            
    return round(total_seconds / 3600, 2)

def calculate_lunch_minutes(times):
    """Calculates the duration of the lunch break (gap between scan 2 and 3)."""
    if not times or len(times) < 4:
        return 0
    
    try:
        t_out_lunch = datetime.strptime(times[1], '%H:%M')
        t_in_lunch = datetime.strptime(times[2], '%H:%M')
        
        if t_in_lunch < t_out_lunch: 
            t_in_lunch += timedelta(days=1)
            
        diff_seconds = (t_in_lunch - t_out_lunch).total_seconds()
        return diff_seconds / 60 
    except:
        return 0

def analyze_record_from_csv(row):
    """Applies business rules to a single daily record from CSV data."""
    # Initialize Defaults
    is_late_930 = 0
    is_late_1000 = 0
    is_late_1400 = 0
    no_lunch = 0
    is_under = 0
    is_half_day = 0

    if row.get('is_leave', 0) == 1 or row.get('is_holiday', 0) == 1:
        return 0, 0, 0, 0, 0, 0

    times, scan_count = parse_scan_times_from_string(row.get('times_list', ''))
    if not times: 
        return 0, 0, 0, 0, 0, 0

    try:
        first_scan = datetime.strptime(times[0], '%H:%M')
        
        # --- TIME LIMITS ---
        limit_930 = first_scan.replace(hour=9, minute=30, second=0)
        limit_1000 = first_scan.replace(hour=10, minute=0, second=0)
        limit_1300 = first_scan.replace(hour=13, minute=0, second=0)
        limit_1400 = first_scan.replace(hour=14, minute=0, second=0)

        # --- LATENESS LOGIC (ANTI-DUPLICATION) ---
        if first_scan > limit_1400:
            is_late_1400 = 1
            is_late_1000 = 0 
            is_late_930 = 0
        elif first_scan > limit_1000:
            is_late_1400 = 0
            is_late_1000 = 1
            is_late_930 = 0
        elif first_scan > limit_930:
            is_late_1400 = 0
            is_late_1000 = 0
            is_late_930 = 1

        is_saturday = str(row.get('day_str', '')).startswith('Sa')
        
        if is_late_1400:
            no_lunch = 0
        else:
            no_lunch = 1 if (len(times) < 4 and not is_saturday) else 0

        target = 4.0 if is_saturday else 8.0
        hours_worked = row.get('hours_worked', 0.0)
        is_under = 1 if hours_worked > 0 and hours_worked < target else 0

        # --- HALF DAY LOGIC ---
        if row.get('is_day_worked', 0) == 1 and not is_saturday and len(times) >= 2:
            try:
                t_entry = datetime.strptime(times[0], '%H:%M')
                t_exit = datetime.strptime(times[-1], '%H:%M')
                if t_exit < t_entry: t_exit += timedelta(days=1)
                
                # --- Condition A: Afternoon Only (Entered after 13:00) ---
                cond_afternoon = (t_entry >= limit_1300)
                
                # --- Condition B: Morning Only (Left before 14:00) ---
                cond_morning = (t_exit <= limit_1400) and (hours_worked < 7.0)
                
                if cond_afternoon or cond_morning:
                    is_half_day = 1
                    
            except:
                pass 
    except:
        pass

    return is_late_930, is_late_1000, is_late_1400, no_lunch, is_under, is_half_day

def calculate_business_days_in_range(start_date, end_date):
    current = start_date
    business_days = 0
    while current <= end_date:
        wd = current.weekday()
        if wd != 6:  # Exclude Sunday
            business_days += 1
        current += timedelta(days=1)
    return business_days

def minutes_to_hhmm(mins):
    if pd.isna(mins) or mins == 0:
        return ""
    hours = int(mins // 60)
    minutes = int(round(mins % 60))
    if minutes == 60:
        hours += 1
        minutes = 0
    return f"{hours:02}:{minutes:02}"

def decimal_hours_to_hhmm(decimal_hours):
    if pd.isna(decimal_hours):
        return ""
    
    if decimal_hours == 0:
        return "00:00"
    
    is_negative = decimal_hours < 0
    minutes_total = abs(decimal_hours) * 60
    hours = int(minutes_total // 60)
    minutes = int(round(minutes_total % 60))
    
    if minutes == 60:
        hours += 1
        minutes = 0
        
    time_str = f"{hours:02}:{minutes:02}"
    return f"-{time_str}" if is_negative else time_str

def process_monthly_analysis_from_csv(csv_dir, output_dir):
    """
    Processes CSV files in csv_dir and generates monthly analysis in output_dir.
    Returns the path of the generated file or None.
    """
    if not os.path.exists(csv_dir):
        print(f"CSV directory not found: {csv_dir}")
        return None

    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    all_data = []
    print("Reading CSV files...")
    csv_files = [f for f in os.listdir(csv_dir) if f.lower().endswith('.csv')]
    
    if not csv_files:
        print("No CSV files found.")
        return None
    
    for csv_file in csv_files:
        print(f"Processing: {csv_file}")
        full_path = os.path.join(csv_dir, csv_file)
        try:
            df = pd.read_csv(full_path)
            all_data.append(df)
        except Exception as e:
            print(f"Error reading {csv_file}: {e}")
            continue

    if not all_data:
        print("No data found.")
        return None

    df = pd.concat(all_data, ignore_index=True)

    # --- CHRONOLOGICAL DETECTION ---
    if 'day_numeric' in df.columns and not df.empty:
        # 1. Get basic info
        month_num = df['month_num'].iloc[0] if 'month_num' in df.columns else '01'
        year_num = df['year_num'].iloc[0] if 'year_num' in df.columns else '2026'
        
        # 2. Identify real chronological sequence
        unique_days_in_order = []
        seen = set()
        for d in df['day_numeric']:
            if d not in seen:
                unique_days_in_order.append(d)
                seen.add(d)

        real_start_day = unique_days_in_order[0]
        real_end_day = unique_days_in_order[-1]
        
        # Detect month transition
        has_transition = False
        pivot_index = -1
        for i in range(len(unique_days_in_order) - 1):
            if unique_days_in_order[i] > unique_days_in_order[i+1]:
                has_transition = True
                pivot_index = i
                break
        
        print(f"\n--- PERIOD ANALYSIS ---")
        print(f"Detected sequence: {unique_days_in_order}")
        
        # 3. Define target day (last chronological day)
        target_report_day = real_end_day
        
        # 4. Check if last day is complete
        last_day_records = df[df['day_numeric'] == target_report_day]
        total_last_day = len(last_day_records)
        incomplete_count = len(last_day_records[last_day_records['scan_count'] <= 1])
        
        if total_last_day > 0 and (incomplete_count / total_last_day) > 0.5:
            print(f"DECISION: Day {target_report_day} is incomplete (in progress).")
            df = df[df['day_numeric'] != target_report_day].copy()
            if len(unique_days_in_order) > 1:
                target_report_day = unique_days_in_order[-2]
                real_end_day = target_report_day
            print(f"New target day: {target_report_day}")
        else:
            print(f"DECISION: Day {target_report_day} is complete.")

        # 5. Calculate month name for header
        month_names = {
            '01': 'Janvier', '02': 'Février', '03': 'Mars', '04': 'Avril',
            '05': 'Mai', '06': 'Juin', '07': 'Juillet', '08': 'Août',
            '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Décembre'
        }
        month_name = month_names.get(month_num, f'Mois {month_num}')
        
        # 6. Create complete dates for business days calculation
        if 'full_date' in df.columns and not df['full_date'].isnull().all():
            # Use real dates if available
            df['full_date'] = pd.to_datetime(df['full_date'], errors='coerce')
            final_min_date = df['full_date'].min()
            final_max_date = df['full_date'].max()
        else:
            # Recreate dates from extracted information
            final_min_date = datetime(int(year_num), int(month_num), real_start_day)
            final_max_date = datetime(int(year_num), int(month_num), real_end_day)
            
            # Handle multi-month periods
            if has_transition:
                if month_num == '12':
                    next_month_num = '01'
                    next_year_num = str(int(year_num) + 1)
                else:
                    next_month_num = f"{int(month_num) + 1:02d}"
                    next_year_num = year_num
                final_max_date = datetime(int(next_year_num), int(next_month_num), real_end_day)
        
        print(f"\n--- DETECTED DAY RANGE ---")
        print(f"First day found: {real_start_day}")
        print(f"Last day found: {real_end_day}")
        
        # Calculate total days correctly for multi-month periods
        if has_transition:
            first_month_days = unique_days_in_order[:pivot_index + 1]
            second_month_days = unique_days_in_order[pivot_index + 1:]
            total_days = len(first_month_days) + len(second_month_days)
            print(f"Multi-month period detected: {len(first_month_days)} days + {len(second_month_days)} days")
        else:
            total_days = len(unique_days_in_order)
        
        print(f"Total days analyzed: {total_days}")
        print(f"Final Analysis Period: {final_min_date.strftime('%d/%m/%Y')} to {final_max_date.strftime('%d/%m/%Y')}")
        global_expected_days = calculate_business_days_in_range(final_min_date, final_max_date)
        print(f"Theoretical Business Days (Mon-Sat) in period: {global_expected_days}")
        
        # Create dynamic filename based on analyzed period
        dynamic_filename = f"Monthly_Global_Analysis_{real_start_day:02d}-{month_num}-{year_num}_A_{real_end_day:02d}-{month_num}-{year_num}.xlsx"
        output_path = os.path.join(output_dir, dynamic_filename)
        header_text = f"Analyse Mensuelle - Période : {real_start_day} au {real_end_day} {month_name} {year_num}"

    else:
        print("Could not detect valid dates. Exiting.")
        return None

    if EXCLUDED_EMPLOYEES:
        print(f"\nFiltering out: {EXCLUDED_EMPLOYEES}")
        excluded_clean = [clean_name_string(name) for name in EXCLUDED_EMPLOYEES]
        df = df[~df['name'].isin(excluded_clean)]

    if df.empty:
        print("All data filtered out.")
        return None

    print("Analyzing metrics...")
    metrics = df.apply(analyze_record_from_csv, axis=1)
    
    df['ENTRY > 9H30'] = [x[0] for x in metrics]
    df['ENTRY > 10H'] = [x[1] for x in metrics]
    df['ENTRY > 14H'] = [x[2] for x in metrics] 
    df['NO LUNCH'] = [x[3] for x in metrics]
    df['UNDER 8H'] = [x[4] for x in metrics]
    df['IS HALF DAY'] = [x[5] for x in metrics]

    # Calculate additional metrics that were in the original CSV
    # Calculate daily_target_for_worked_day for ALL days that needed to be worked (excluding Sundays and holidays)
    df['daily_target_for_worked_day'] = df.apply(
        lambda row: 4.0 if (str(row['day_str']).startswith('Sa') and row.get('is_holiday', 0) == 0) 
        else (8.0 if (not str(row['day_str']).startswith('Sa') and not str(row['day_str']).startswith('Di') and row.get('is_holiday', 0) == 0) else 0.0), axis=1
    )
    
    # Calculate lunch break metrics - only for worked weekdays
    def calculate_lunch_metrics(row):
        if row.get('is_day_worked', 0) == 1 and not str(row['day_str']).startswith('Sa'):
            times, _ = parse_scan_times_from_string(row.get('times_list', ''))
            lunch_mins = calculate_lunch_minutes(times)
            has_break = 1 if len(times) >= 4 else 0
            return lunch_mins, has_break
        return 0, 0
    
    df[['daily_lunch_minutes', 'has_lunch_break']] = df.apply(calculate_lunch_metrics, axis=1, result_type='expand')

    report = df.groupby('name').agg({
        'is_day_worked': 'sum',
        'is_leave': 'sum',
        'is_holiday': 'sum',
        'daily_target_for_worked_day': 'sum', 
        'ENTRY > 10H': 'sum',
        'ENTRY > 14H': 'sum', 
        'ENTRY > 9H30': 'sum',
        'NO LUNCH': 'sum',
        'UNDER 8H': 'sum',
        'IS HALF DAY': 'sum',
        'hours_worked': 'sum',
        'daily_lunch_minutes': 'sum',
        'has_lunch_break': 'sum'
    }).reset_index()

    report.rename(columns={
        'name': 'Employee name',
        'is_day_worked': 'days worked',
        'daily_target_for_worked_day': 'TOTAL HOURS NEEDED', 
        'hours_worked': 'TOTAL HOURS WORKED',
        'IS HALF DAY': 'HALF DAYS'
    }, inplace=True)

    # Calculate total expected working days (including Saturdays as 0.5 days each)
    # Exclude both Saturdays AND Sundays from weekday records
    weekday_records = df[~df['day_str'].str.startswith('Sa', na=False) & ~df['day_str'].str.startswith('Di', na=False)]
    expected_weekdays = len(weekday_records['day_numeric'].unique())
    
    # Get Saturday records for expected days calculation
    saturday_records = df[df['day_str'].str.startswith('Sa', na=False)]
    expected_saturdays = len(saturday_records['day_numeric'].unique())
    
    # Calculate real working days with Saturday 0.5 adjustment (Sundays excluded completely)
    total_expected_days_adjusted = expected_weekdays + (expected_saturdays * 0.5)
    
    report['real working days'] = total_expected_days_adjusted
    
    # Calculate absence with Saturday adjustment
    saturday_work_by_employee = saturday_records[saturday_records['is_day_worked'] == 1].groupby('name').size().reset_index(name='saturdays_worked')
    saturday_work_by_employee['saturdays_worked'] = saturday_work_by_employee['saturdays_worked'].apply(lambda x: 1 if x > 0 else 0)
    
    # Merge Saturday work data back to report
    report = report.merge(saturday_work_by_employee, left_on='Employee name', right_on='name', how='left')
    report['saturdays_worked'] = report['saturdays_worked'].fillna(0)
    report.drop('name', axis=1, inplace=True)
    
    # Calculate adjusted absence - simplified formula with half day adjustment
    # ABSENCE = real working days - days worked + (HALF DAYS * 0.5)
    report['ABSENCE'] = report['real working days'] - report['days worked'] + (report['HALF DAYS'] * 0.5)
    
    # Clean up temporary columns (only drop saturdays_worked since it's the only one that exists)
    report.drop(['saturdays_worked'], axis=1, inplace=True)
    
    report['avg_lunch_raw'] = report.apply(
        lambda x: x['daily_lunch_minutes'] / x['has_lunch_break'] if x['has_lunch_break'] > 0 else (x['daily_lunch_minutes'] if x['daily_lunch_minutes'] > 0 else 0), axis=1
    )
    report['AVG LUNCH TIME'] = report['avg_lunch_raw'].apply(minutes_to_hhmm)

    report['balance_raw'] = report['TOTAL HOURS WORKED'] - report['TOTAL HOURS NEEDED']
    report['Balance of hours worked'] = report['balance_raw'].apply(decimal_hours_to_hhmm)

    # --- EXPORT ---
    final_cols = [
        'Employee name', 
        'real working days', 
        'days worked',
        'ABSENCE', 
        'HALF DAYS', 
        'UNDER 8H', 
        'NO LUNCH', 
        'AVG LUNCH TIME',
        'ENTRY > 14H', 
        'ENTRY > 10H', 
        'ENTRY > 9H30', 
        'TOTAL HOURS NEEDED', 
        'TOTAL HOURS WORKED', 
        'Balance of hours worked'
    ]
    
    final_df = report[final_cols]
    
    # Sort by ABSENCE column from highest to lowest
    final_df = final_df.sort_values('ABSENCE', ascending=False)

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Monthly Summary', index=False, startrow=2, header=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Monthly Summary']
            
            header_title = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 14, 'font_color': '#2F5597', 'border': 1
            })
            
            if len(final_df.columns) > 1:
                worksheet.merge_range(0, 0, 0, len(final_df.columns) - 1, header_text, header_title)
            else:
                worksheet.write(0, 0, header_text, header_title)
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })
            
            header_red = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#C00000', 'font_color': 'white', 'border': 1
            })

            header_orange = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#ED7D31', 'font_color': 'white', 'border': 1
            })

            body_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
            text_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
            
            absence_red_format = workbook.add_format({
                'border': 1, 'align': 'center', 'valign': 'vcenter', 
                'bg_color': '#FFC7CE', 'font_color': '#9C0006'
            })
            
            # Write Headers
            for col_num, value in enumerate(final_df.columns.values):
                if "14H" in value:
                     worksheet.write(1, col_num, value, header_red)
                elif "HALF DAYS" in value:
                    worksheet.write(1, col_num, value, header_orange)
                else:
                    worksheet.write(1, col_num, value, header_format)
            
            # Write Data
            for i, col in enumerate(final_df.columns):
                if col == 'Employee name':
                    cell_fmt = text_format
                    width = 20
                elif col in ['real working days', 'days worked', 'ABSENCE', 'HALF DAYS', 'UNDER 8H', 'NO LUNCH', 'ENTRY > 14H', 'ENTRY > 10H', 'ENTRY > 9H30']:
                    cell_fmt = body_format
                    width = 10
                elif col in ['AVG LUNCH TIME']:
                    cell_fmt = body_format
                    width = 12
                elif col in ['TOTAL HOURS NEEDED', 'TOTAL HOURS WORKED', 'Balance of hours worked']:
                    cell_fmt = body_format
                    width = 14
                else:
                    cell_fmt = body_format
                    width = 12
                
                worksheet.set_column(i, i, width)
                
                for row_idx, value in enumerate(final_df[col]):
                    if pd.isna(value): value = ""
                    
                    if col in ['real working days', 'days worked', 'ABSENCE', 'HALF DAYS', 'UNDER 8H', 'NO LUNCH', 'ENTRY > 14H', 'ENTRY > 10H', 'ENTRY > 9H30']:
                        if value == 0: value = 0
                    elif col in ['AVG LUNCH TIME', 'TOTAL HOURS WORKED']:
                        if value == 0 or value == "00:00": value = ""
                    # Keep balance as is - don't empty it for 00:00
                    elif col == 'Balance of hours worked':
                        pass  # Keep the formatted value as is
                    
                    if col == 'ABSENCE' and value > 0:
                        worksheet.write(row_idx + 2, i, value, absence_red_format)
                    else:
                        worksheet.write(row_idx + 2, i, value, cell_fmt)

        print(f"\nSUCCESS! Monthly report generated: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error saving file: {e}")
        return None

def process_monthly_analysis(input_dir, output_dir):
    """
    Wrapper function for compatibility with existing app.py
    This function first extracts CSV data, then processes it.
    """
    # First create temp_csv directory
    temp_csv_dir = os.path.join(os.path.dirname(input_dir), "temp_csv")
    
    # Extract CSV data from Excel files
    import csv_extractor
    if not csv_extractor.process_all_excel_to_csv(input_dir, temp_csv_dir):
        print("Failed to extract CSV data")
        return None
    
    # Process the CSV data
    return process_monthly_analysis_from_csv(temp_csv_dir, output_dir)

if __name__ == "__main__":
    # Test with CSV directory
    csv_directory = "temp_csv"
    output_directory = "temp_output"
    
    output = process_monthly_analysis_from_csv(csv_directory, output_directory)
    if output:
        print(f"Report generated: {output}")
