import pandas as pd
import os
import sys

# Add the project directory to path
sys.path.insert(0, r'C:\Users\PC\Desktop\rh-analysis-tool-main')

from analysis_bureau_daily import get_sheet_rows, clean_name_string

file_path = r'C:\Users\PC\Desktop\rh-analysis-tool-main\temp_input\POINTAGE 01-03- A 24-03-26 BUREAU.xls.xlsx'

print("Looking for NOM lines...")
count = 0
for row in get_sheet_rows(file_path):
    if not row:
        continue
    cell_0 = row[0]
    val_0 = str(cell_0.value).strip() if cell_0.value else ''
    
    if 'NOM :' in val_0:
        raw_name = val_0.replace('NOM :', '').strip()
        cleaned = clean_name_string(raw_name)
        print(f'Row {count}: Raw: "{raw_name}" -> Cleaned: "{cleaned}"')
        count += 1
        if count > 5:
            break

print(f"\nFound {count} names")
