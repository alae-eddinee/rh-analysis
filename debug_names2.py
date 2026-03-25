import pandas as pd
import os
import sys

# Add the project directory to path
sys.path.insert(0, r'C:\Users\PC\Desktop\rh-analysis-tool-main')

from analysis_bureau_daily import get_sheet_rows

file_path = r'C:\Users\PC\Desktop\rh-analysis-tool-main\temp_input\POINTAGE 01-03- A 24-03-26 BUREAU.xls.xlsx'

print("Scanning first 50 rows for NOM patterns...")
for i, row in enumerate(get_sheet_rows(file_path)):
    if i > 50:
        break
    if not row:
        continue
    cell_0 = row[0]
    val_0 = str(cell_0.value).strip() if cell_0.value else ''
    
    # Look for any row containing NOM
    if val_0 and 'nom' in val_0.lower():
        print(f'Row {i}: "{val_0}"')
        # Show next row too
        if i < 50:
            try:
                next_vals = []
                for r in get_sheet_rows(file_path):
                    if r and r[0].value and 'NOM' in str(r[0].value).upper():
                        continue
                    next_vals.append(str(r[0].value) if r[0].value else '')
                    if len(next_vals) >= 5:
                        break
                print(f"  Next rows: {next_vals[:3]}")
            except:
                pass
        break
else:
    print("No 'NOM' pattern found in first 50 rows")
    # Just show first 20 rows
    print("\nFirst 20 rows:")
    for i, row in enumerate(get_sheet_rows(file_path)):
        if i > 20:
            break
        cell_0 = row[0] if row else None
        val_0 = str(cell_0.value).strip()[:50] if cell_0 and cell_0.value else ''
        if val_0:
            print(f'Row {i}: "{val_0}"')
