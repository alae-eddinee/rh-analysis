---
name: RH Analysis Tool — Project Overview
description: Full architecture of the Streamlit HR attendance analysis app
type: project
---

Streamlit app automating HR pointage analysis. Reads MEDIDIS "ETAT DES HEURES TRAVAILLEES" Excel badge-scan exports.

**Why:** Automates manual attendance report generation for an Moroccan company.
**How to apply:** All code changes must account for the 3+1 tab structure and two worker types.

## Tabs
1. Analyse Bureau (8h/day Mon-Fri, 4h Sat) → analysis_bureau_daily.py + analysis_bureau_monthly.py + late_arrivals_graph.py
2. Analyse Production (9h/day, codes 130/131/140/141) → analysis_production_daily.py + analysis_production_monthly.py
3. Pivot Annuel → pointage_pivot_V2.py
4. Gestion Employés → employees_db.py (CRUD for employee-service mapping)

## Input format (MEDIDIS Excel rows)
- `SERVICE / SECTION : <name>` → section header
- `NOM : <name>` → employee name
- `MATRICULE : <id>`
- Day rows: `"Lu 01/01/2025"` | HJ code | scan timestamps | Tps Dû (col3) | Tps Eff (col4)
- HJ codes 130/131/140/141 = production worker → excluded from bureau analysis (>50% ratio)

## Business rules
- Ramadan 2026: Feb 19 – Mar 20, 2026 → 7h target, no lunch check
- Moroccan public holidays 2025: hardcoded in pointage_pivot_V2.py
- Excluded employee: "HMOURI ALI" hardcoded in all scripts
- Hours: prefer Tps Eff column; fallback to scan-pair arithmetic
- Supabase Storage bucket "RH-Data": outputs uploaded after generation
- Credentials via st.secrets
