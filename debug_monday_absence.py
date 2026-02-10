import pandas as pd
import os
import analysis_per_month_csv

def debug_monday_absence():
    """Debug why Monday absence shows 0.5 instead of 1"""
    print("🔍 Debugging Monday Absence Issue...")
    
    # Test case: Employee absent on Monday, works rest
    test_data = [
        {'day_str': 'Lu', 'is_saturday': 0, 'is_sunday': 0, 'is_day_worked': 0, 'hours_worked': 0.0},  # Absent Monday
        {'day_str': 'Ma', 'is_saturday': 0, 'is_sunday': 0, 'is_day_worked': 1, 'hours_worked': 8.0},  # Work Tuesday
        {'day_str': 'Me', 'is_saturday': 0, 'is_sunday': 0, 'is_day_worked': 1, 'hours_worked': 8.0},  # Work Wednesday
        {'day_str': 'Je', 'is_saturday': 0, 'is_sunday': 0, 'is_day_worked': 1, 'hours_worked': 8.0},  # Work Thursday
        {'day_str': 'Ve', 'is_saturday': 0, 'is_sunday': 0, 'is_day_worked': 1, 'hours_worked': 8.0},  # Work Friday
        {'day_str': 'Sa', 'is_saturday': 1, 'is_sunday': 0, 'is_day_worked': 0, 'hours_worked': 0.0},  # Absent Saturday
    ]
    
    # Create test data
    df = pd.DataFrame([{
        'source_file': 'test.xlsx',
        'service': 'IT',
        'name': 'MONDAY ABSENT TEST',
        'matricule': '001',
        'full_date': '2025-01-15',
        'day_numeric': i + 15,
        'day_str': d['day_str'],
        'is_saturday': d['is_saturday'],
        'is_sunday': d['is_sunday'],
        'hj_code': '120',
        'scan_count': 0 if d['hours_worked'] == 0 else 4,
        'raw_pointages': '' if d['hours_worked'] == 0 else '08:00 12:00 13:00 17:00',
        'times_list': '' if d['hours_worked'] == 0 else '08:00|12:00|13:00|17:00',
        'hours_worked': d['hours_worked'],
        'is_day_worked': d['is_day_worked'],
        'is_leave': 0,
        'is_holiday': 0,
        'month_num': '01',
        'year_num': '2025'
    } for i, d in enumerate(test_data)])
    
    # Clear and create test files
    os.makedirs('temp_csv', exist_ok=True)
    os.makedirs('temp_output', exist_ok=True)
    
    csv_path = os.path.join('temp_csv', 'monday_absence_test.csv')
    df.to_csv(csv_path, index=False)
    
    print(f"✅ Created test CSV: {csv_path}")
    print("📋 Test Data:")
    print(df[['day_str', 'is_day_worked', 'hours_worked']])
    
    # Test monthly analysis
    print("\n🔄 Testing monthly analysis...")
    monthly_output = analysis_per_month_csv.process_monthly_analysis_from_csv('temp_csv', 'temp_output')
    if monthly_output:
        print(f"✅ Analysis successful: {os.path.basename(monthly_output)}")
    else:
        print("❌ Analysis failed")
        return
    
    # Verify results
    print("\n📊 Verifying Monday Absence:")
    report_df = pd.read_excel(monthly_output, sheet_name='Monthly Summary', header=1)
    employee_data = report_df[report_df['Employee name'] == 'MONDAY ABSENT TEST'].iloc[0]
    
    print(f"Employee: {employee_data['Employee name']}")
    print(f"Real working days: {employee_data['real working days']}")
    print(f"Days worked: {employee_data['days worked']}")
    print(f"HALF DAYS: {employee_data['HALF DAYS']}")
    print(f"ABSENCE: {employee_data['ABSENCE']}")
    
    # Expected calculation
    # 4 weekdays + 0.5 Saturday = 4.5 real working days
    # 4 days worked (Tue-Fri)
    # Expected absence: 4.5 - 4 + (0 * 0.5) = 0.5
    expected_real_days = 4.5
    expected_absence = 4.5 - 4 + (0 * 0.5)
    
    print(f"Expected real working days: {expected_real_days}")
    print(f"Expected absence: {expected_absence}")
    
    print(f"\n🔍 Analysis:")
    print(f"  Monday absent: Should be 1 day absence")
    print(f"  Saturday absent: Should be 0.5 day absence")
    print(f"  Total expected: 1.5 days absence")
    print(f"  System shows: {employee_data['ABSENCE']} days absence")
    
    if abs(employee_data['ABSENCE'] - expected_absence) < 0.01:
        print("✅ Calculation matches formula")
    else:
        print(f"❌ Calculation mismatch (expected {expected_absence}, got {employee_data['ABSENCE']})")

if __name__ == "__main__":
    debug_monday_absence()
