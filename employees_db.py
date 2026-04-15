"""employees_db.py - Persistent employee service database.

Stores the employee → service mapping in Supabase.
Auto-initializes from 'liste personnels bureau.xlsx' on first run if table is empty.

Each employee record:
  { id, matricule, nom, prenom, responsable, service, poste, last_seen }
  - id: Supabase-generated UUID
  - last_seen: ISO date string 'YYYY-MM-DD' of the most recent badge scan found in
    a bureau analysis run, or None if the employee was added manually and never seen.
"""
import json
import os
import re
from datetime import date, datetime, timedelta

# Supabase imports
from supabase import create_client, Client

_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(_BASE_DIR, "liste personnels bureau.xlsx")

INACTIVE_DAYS = 30   # threshold for "stopped scanning" warning / auto-remove

# Supabase configuration (from environment variables or will be set later)
_supabase_url = os.environ.get("SUPABASE_URL", "")
_supabase_key = os.environ.get("SUPABASE_KEY", "")
_supabase_client: Client = None

# Table name in Supabase
TABLE_NAME = "employees"

# In-memory caches — invalidated whenever employees are saved
_employees_cache: list | None = None
_field_maps_cache: dict = {}


def _invalidate_cache():
    global _employees_cache, _field_maps_cache
    _employees_cache = None
    _field_maps_cache = {}


def _get_supabase_client() -> Client:
    """Get or create Supabase client."""
    global _supabase_client, _supabase_url, _supabase_key
    
    if _supabase_client is not None:
        return _supabase_client
    
    # Try to get from environment or use defaults for local testing
    url = _supabase_url or os.environ.get("SUPABASE_URL", "")
    key = _supabase_key or os.environ.get("SUPABASE_KEY", "")
    
    # Also try to get from streamlit secrets if available (for app.py context)
    if not url or not key:
        try:
            import streamlit as st
            url = st.secrets.get("SUPABASE_URL", "")
            key = st.secrets.get("SUPABASE_KEY", "")
        except Exception:
            pass
    
    if url and key:
        try:
            _supabase_client = create_client(url, key)
            return _supabase_client
        except Exception as e:
            print(f"Warning: Could not connect to Supabase: {e}")
    
    return None


def _clean(name: str) -> str:
    """Normalize a name for matching: uppercase, collapse all whitespace."""
    if not name:
        return ""
    s = str(name).upper().strip()
    s = s.replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ')
    return re.sub(r'\s+', ' ', s)


# ─── Persistence ─────────────────────────────────────────────────────────────

def load_employees() -> list:
    """Return the full employee list from Supabase. Auto-initializes from Excel if table is empty."""
    global _employees_cache
    if _employees_cache is not None:
        return _employees_cache

    client = _get_supabase_client()

    if client is None:
        _employees_cache = _load_from_local_fallback()
        return _employees_cache

    try:
        response = client.table(TABLE_NAME).select("*").execute()
        employees = response.data if response.data else []

        # Auto-initialize from Excel if empty
        if not employees and os.path.exists(EXCEL_PATH):
            _init_from_excel_to_supabase()
            response = client.table(TABLE_NAME).select("*").execute()
            employees = response.data if response.data else []

        _employees_cache = employees
        return _employees_cache
    except Exception as e:
        print(f"Warning: Could not load employees from Supabase: {e}")
        _employees_cache = _load_from_local_fallback()
        return _employees_cache


def _load_from_local_fallback() -> list:
    """Fallback to local JSON if Supabase is unavailable."""
    db_path = os.path.join(_BASE_DIR, "employees_db.json")
    if os.path.exists(db_path):
        try:
            with open(db_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_employees(employees: list) -> None:
    """Persist the employee list to Supabase."""
    _invalidate_cache()
    client = _get_supabase_client()
    
    if client is None:
        # Fallback: save to local JSON
        _save_to_local_fallback(employees)
        return
    
    try:
        # Get current employees to determine what to add/update/delete
        current = load_employees()
        current_ids = {e.get('id') for e in current if e.get('id')}
        new_ids = {e.get('id') for e in employees if e.get('id')}
        
        # Delete removed employees
        to_delete = current_ids - new_ids
        for emp_id in to_delete:
            client.table(TABLE_NAME).delete().eq('id', emp_id).execute()
        
        # Upsert (insert or update) employees
        for emp in employees:
            # Clean the employee data
            record = {
                'matricule': str(emp.get('matricule', '') or '').strip(),
                'nom': str(emp.get('nom', '') or '').strip().upper(),
                'prenom': str(emp.get('prenom', '') or '').strip().upper(),
                'responsable': str(emp.get('responsable', '') or '').strip(),
                'service': str(emp.get('service', '') or '').strip().lower(),
                'poste': str(emp.get('poste', '') or '').strip(),
                'last_seen': emp.get('last_seen'),
            }
            
            emp_id = emp.get('id')
            if emp_id:
                # Update existing
                client.table(TABLE_NAME).update(record).eq('id', emp_id).execute()
            else:
                # Insert new
                client.table(TABLE_NAME).insert(record).execute()
    except Exception as e:
        raise RuntimeError(f"Supabase save failed: {e}") from e


def _save_to_local_fallback(employees: list) -> None:
    """Fallback to save to local JSON if Supabase is unavailable."""
    db_path = os.path.join(_BASE_DIR, "employees_db.json")
    with open(db_path, 'w', encoding='utf-8') as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)


def _init_from_excel_to_supabase() -> None:
    """Bootstrap the DB from the Excel file (first row is a merged title – skip it)."""
    if not os.path.exists(EXCEL_PATH):
        return
    
    client = _get_supabase_client()
    if client is None:
        # Fallback: use local JSON initialization
        _init_from_excel_to_local()
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
                'nom':         nom.upper(),
                'prenom':      str(row.get('prenom',      '') or '').strip().upper(),
                'responsable': str(row.get('responsable', '') or '').strip(),
                'service':     str(row.get('service',     '') or '').strip().lower(),
                'poste':       str(row.get('poste',       '') or '').strip(),
                'last_seen':   None,
            })
        
        # Insert all employees to Supabase
        if employees:
            client.table(TABLE_NAME).insert(employees).execute()
            print(f"Initialized employee DB with {len(employees)} employees from Excel in Supabase.")
    except Exception as e:
        print(f"Warning: Could not init employee DB from Excel to Supabase: {e}")


def _init_from_excel_to_local() -> None:
    """Bootstrap local JSON from Excel as fallback."""
    db_path = os.path.join(_BASE_DIR, "employees_db.json")
    if os.path.exists(db_path):
        return
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
        _save_to_local_fallback(employees)
        print(f"Initialized employee DB with {len(employees)} employees from Excel (local fallback).")
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

def _build_field_map(field: str) -> dict:
    """
    Build a lookup dict: cleaned name variant → field value.

    Keys added per employee:
      - 'NOM PRENOM'  (full name)
      - 'PRENOM NOM'  (reversed, for flexible matching)
      - 'NOM'         (last-name-only fallback – added only if not already present)
    """
    if field in _field_maps_cache:
        return _field_maps_cache[field]

    field_map: dict = {}
    for emp in load_employees():
        nom    = _clean(emp.get('nom', ''))
        prenom = _clean(emp.get('prenom', ''))
        value  = emp.get(field, '').strip() if emp.get(field) else ''
        if field == 'service':
            value = value.capitalize()
        if not nom:
            continue
        if prenom:
            field_map[f"{nom} {prenom}"] = value
            field_map[f"{prenom} {nom}"] = value
        if nom not in field_map:
            field_map[nom] = value

    _field_maps_cache[field] = field_map
    return field_map


def _lookup_field(pointage_name: str, field: str) -> str:
    """
    Find the value of a given field for a name as it appears in a pointage file.

    Strategy:
      1. Exact match (full name or reversed variant).
      2. Prefix match: pointage_name starts with a known key, or vice-versa.

    Returns '' if nothing matches.
    """
    if not pointage_name:
        return ''
    cleaned = _clean(pointage_name)
    fmap = _build_field_map(field)

    if cleaned in fmap:
        return fmap[cleaned]

    for key, val in fmap.items():
        if cleaned.startswith(key) or key.startswith(cleaned):
            return val

    return ''


# Keep backward-compatible aliases
def get_service_map() -> dict:
    return _build_field_map('service')


def lookup_service(pointage_name: str) -> str:
    return _lookup_field(pointage_name, 'service')


def lookup_responsable(pointage_name: str) -> str:
    return _lookup_field(pointage_name, 'responsable')


def lookup_poste(pointage_name: str) -> str:
    return _lookup_field(pointage_name, 'poste')