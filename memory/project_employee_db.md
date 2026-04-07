---
name: Employee DB feature
description: employees_db.py persists employee-service mapping; used across all analysis scripts
type: project
---

Added persistent employee database so that "Service" column appears in all Excel reports.

**Why:** Service info from "liste personnels bureau.xlsx" needed in generated reports; users also need to add/edit workers inside the app.
**How to apply:** Any new analysis script should import employees_db and call `lookup_service(name)` to get the service for a worker.

## Key files
- `employees_db.py` — persistence module: load/save JSON, lookup_service(), get_service_map()
- `employees_db.json` — auto-created JSON (gitignored-worthy); auto-populated from Excel on first run
- `liste personnels bureau.xlsx` — source Excel (sheet="liste", skiprows=1 to skip title row)

## How it works
- `load_employees()` → returns list of dicts; auto-calls `_init_from_excel()` if JSON absent
- `lookup_service(name)` → cleans name (uppercase, normalize whitespace), tries exact match then prefix match in service_map
- `get_service_map()` → builds dict with keys: "NOM PRENOM", "PRENOM NOM", "NOM" (fallback)

## Integration in analysis scripts
Each analysis script: adds `import employees_db` (with sys.path guard for dynamic-load context)
- Daily scripts: `df['service'] = df['name'].apply(employees_db.lookup_service)` → passed through `create_category_dataframe` which now returns 4 cols [Name, Service, Count, %]
- Monthly scripts: `report['Service'] = report['Employee name'].apply(employees_db.lookup_service)` → added to final_cols after 'Employee name'

## app.py management tab
"👥 Gestion Employés" tab (4th tab) uses st.data_editor(num_rows="dynamic") for inline add/edit/delete. Save button calls `employees_db.save_employees()`.
