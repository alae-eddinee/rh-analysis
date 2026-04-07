"""employees_db.py - Persistent employee service database.

Stores the employee → service mapping as a JSON file (employees_db.json).
Auto-initializes from 'liste personnels bureau.xlsx' on first run.

Each employee record:
  { matricule, nom, prenom, responsable, service, poste, last_seen }
  - last_seen: ISO date string 'YYYY-MM-DD' of the most recent badge scan found in
    a bureau analysis run, or None if the employee was added manually and never seen.
"""
import json
import os
import re
from datetime import date, datetime, timedelta

_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DB_PATH    = os.path.join(_BASE_DIR, "employees_db.json")
EXCEL_PATH = os.path.join(_BASE_DIR, "liste personnels bureau.xlsx")

INACTIVE_DAYS = 30   # threshold for "stopped scanning" warning / auto-remove


def _clean(name: str) -> str:
    """Normalize a name for matching: uppercase, collapse all whitespace."""
    if not name:
        return ""
    s = str(name).upper().strip()
    s = s.replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ')
    return re.sub(r'\s+', ' ', s)


# ─── Persistence ─────────────────────────────────────────────────────────────

def load_employees() -> list:
    """Return the full employee list. Auto-initializes from Excel if JSON absent."""
    if not os.path.exists(DB_PATH):
        _init_from_excel()
    if not os.path.exists(DB_PATH):
        return []
    with open(DB_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_employees(employees: list) -> None:
    """Persist the employee list to JSON."""
    with open(DB_PATH, 'w', encoding='utf-8') as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)


def _init_from_excel() -> None:
    """Bootstrap the DB from the Excel file (first row is a merged title – skip it)."""
    if not os.path.exists(EXCEL_PATH):
        return
    try:
        import pandas as pd
        df = pd.read_excel(EXCEL_PATH, sheet_name='liste', skiprows=1)
        df.columns = ['matricule', 'nom', 'prenom', 'responsable', 'service', 'poste']
        employees = []
        for _, row in df.iterrows():
            nom = str(row.get('nom', '') or '').strip()
            if not nom or nom.lower() == 'nan':
                continue
            employees.append({
                'matricule':   str(row.get('matricule',   '') or '').strip(),
                'nom':         nom,
                'prenom':      str(row.get('prenom',      '') or '').strip(),
                'responsable': str(row.get('responsable', '') or '').strip(),
                'service':     str(row.get('service',     '') or '').strip(),
                'poste':       str(row.get('poste',       '') or '').strip(),
                'last_seen':   None,
            })
        save_employees(employees)
        print(f"Initialized employee DB with {len(employees)} employees from Excel.")
    except Exception as e:
        print(f"Warning: Could not init employee DB from Excel: {e}")


# ─── last_seen tracking ───────────────────────────────────────────────────────

def update_last_seen(name_date_map: dict) -> None:
    """
    Update the last_seen date for employees found in a bureau analysis run.

    name_date_map: { pointage_name (str) : last_scan_date (date|datetime|str) }
    Only updates if the new date is more recent than the stored one.
    """
    if not name_date_map:
        return

    employees = load_employees()
    changed = False

    for emp in employees:
        nom    = _clean(emp.get('nom', ''))
        prenom = _clean(emp.get('prenom', ''))
        # Build the set of name variants for this employee
        variants = set()
        if nom and prenom:
            variants.add(f"{nom} {prenom}")
            variants.add(f"{prenom} {nom}")
        if nom:
            variants.add(nom)

        # Find a matching entry in name_date_map
        best_date = None
        for pointage_name, scan_date in name_date_map.items():
            if _clean(pointage_name) in variants:
                if isinstance(scan_date, (date, datetime)):
                    d = scan_date.date() if isinstance(scan_date, datetime) else scan_date
                else:
                    try:
                        d = datetime.strptime(str(scan_date)[:10], '%Y-%m-%d').date()
                    except Exception:
                        continue
                if best_date is None or d > best_date:
                    best_date = d

        if best_date is None:
            continue

        current = emp.get('last_seen')
        if current:
            try:
                current_date = datetime.strptime(current[:10], '%Y-%m-%d').date()
            except Exception:
                current_date = None
        else:
            current_date = None

        if current_date is None or best_date > current_date:
            emp['last_seen'] = best_date.isoformat()
            changed = True

    if changed:
        save_employees(employees)


def get_inactive(reference_date=None, threshold_days: int = INACTIVE_DAYS) -> list:
    """
    Return employees who have been seen at least once but whose last_seen date is
    more than `threshold_days` before `reference_date` (default: today).

    Returns list of employee dicts with an extra 'days_inactive' key.
    Employees with last_seen=None (never seen in any scan) are NOT included.
    """
    if reference_date is None:
        reference_date = date.today()
    elif isinstance(reference_date, datetime):
        reference_date = reference_date.date()

    result = []
    for emp in load_employees():
        ls = emp.get('last_seen')
        if not ls:
            continue
        try:
            last = datetime.strptime(ls[:10], '%Y-%m-%d').date()
        except Exception:
            continue
        delta = (reference_date - last).days
        if delta > threshold_days:
            result.append({**emp, 'days_inactive': delta})
    return result


def remove_inactive(reference_date=None, threshold_days: int = INACTIVE_DAYS) -> int:
    """
    Remove employees inactive for more than `threshold_days`.
    Returns the number of employees removed.
    """
    inactive_names = {
        (_clean(e.get('nom', '')), _clean(e.get('prenom', '')))
        for e in get_inactive(reference_date, threshold_days)
    }
    if not inactive_names:
        return 0

    employees = load_employees()
    before = len(employees)
    employees = [
        e for e in employees
        if (_clean(e.get('nom', '')), _clean(e.get('prenom', ''))) not in inactive_names
    ]
    save_employees(employees)
    return before - len(employees)


# ─── Service lookup ──────────────────────────────────────────────────────────

def get_service_map() -> dict:
    """
    Build a lookup dict: cleaned name variant → capitalized service name.

    Keys added per employee:
      - 'NOM PRENOM'  (full name)
      - 'PRENOM NOM'  (reversed, for flexible matching)
      - 'NOM'         (last-name-only fallback – added only if not already present)
    """
    service_map: dict = {}
    for emp in load_employees():
        nom     = _clean(emp.get('nom', ''))
        prenom  = _clean(emp.get('prenom', ''))
        service = emp.get('service', '').strip().capitalize()
        if not nom:
            continue
        if prenom:
            service_map[f"{nom} {prenom}"] = service
            service_map[f"{prenom} {nom}"] = service
        if nom not in service_map:   # last-name fallback – don't override full-name entry
            service_map[nom] = service
    return service_map


def lookup_service(pointage_name: str) -> str:
    """
    Find the service for a name as it appears in a pointage file.

    Strategy:
      1. Exact match (full name or reversed variant).
      2. Prefix match: pointage_name starts with a known key, or vice-versa.

    Returns '' if nothing matches.
    """
    if not pointage_name:
        return ''
    cleaned = _clean(pointage_name)
    smap = get_service_map()

    if cleaned in smap:
        return smap[cleaned]

    for key, svc in smap.items():
        if cleaned.startswith(key) or key.startswith(cleaned):
            return svc

    return ''