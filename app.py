import streamlit as st
import os
import shutil
import importlib.util
import sys
import pandas as pd
from datetime import datetime
from supabase import create_client, Client

# --- CONFIGURATION ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_INPUT_DIR = os.path.join(BASE_DIR, "temp_input")
TEMP_OUTPUT_DIR = os.path.join(BASE_DIR, "temp_output")

# Supabase configuration from Streamlit secrets
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
SUPABASE_BUCKET = st.secrets.get("SUPABASE_BUCKET", "RH-Data")

# Initialize Supabase client if credentials are available
supabase_client = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.warning(f"Erreur connexion Supabase: {e}")

# --- SUPABASE UPLOAD FUNCTION ---
def upload_to_supabase(file_path, bucket_name=SUPABASE_BUCKET):
    """Upload a file to Supabase Storage."""
    if not supabase_client:
        return False, "Supabase client not initialized"
    
    if not os.path.exists(file_path):
        return False, f"File not found: {file_path}"
    
    try:
        # Add timestamp to filename to avoid duplicates
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_name = os.path.basename(file_path)
        name, ext = os.path.splitext(original_name)
        new_filename = f"{name}_{timestamp}{ext}"
        
        with open(file_path, "rb") as f:
            file_content = f.read()
        
        # Upload file to Supabase Storage
        response = supabase_client.storage.from_(bucket_name).upload(
            new_filename,
            file_content,
            file_options={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        )
        
        return True, new_filename
    except Exception as e:
        return False, str(e)

# --- IMPORT FUNCTIONS DYNAMICALLY ---
def load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

# Load Bureau analysis scripts (8h/day - standard office workers)
daily_script = load_module_from_path("bureau_daily_analysis", os.path.join(BASE_DIR, "analysis_bureau_daily.py"))
monthly_script = load_module_from_path("bureau_monthly_analysis", os.path.join(BASE_DIR, "analysis_bureau_monthly.py"))
graph_script = load_module_from_path("late_arrivals_graph", os.path.join(BASE_DIR, "late_arrivals_graph.py"))

# Load Production analysis scripts (9h/day - workers with codes 130, 131, 140, 141)
prod_daily_script = load_module_from_path("production_daily_analysis", os.path.join(BASE_DIR, "analysis_production_daily.py"))
prod_monthly_script = load_module_from_path("production_monthly_analysis", os.path.join(BASE_DIR, "analysis_production_monthly.py"))

# Load Annual Pivot analysis script (V2 with correct calculations)
annual_pivot_script = load_module_from_path("pointage_pivot_v2", os.path.join(BASE_DIR, "pointage_pivot_V2.py"))

# Load employee database module and ensure it is initialized
employees_db = load_module_from_path("employees_db", os.path.join(BASE_DIR, "employees_db.py"))
employees_db.load_employees()  # triggers auto-init from Excel if DB is absent

# --- UTILS ---
def reset_dirs():
    """Réinitialise les dossiers temporaires."""
    for folder in [TEMP_INPUT_DIR, TEMP_OUTPUT_DIR]:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
            except Exception as e:
                st.error(f"Erreur lors du nettoyage du dossier {folder}: {e}")
        os.makedirs(folder)

# --- STREAMLIT APP ---
st.set_page_config(page_title="RH Analysis Tool", page_icon="📊", layout="wide")

st.title("📊 RH Data Analysis Automation")
st.markdown("""
Cette application permet d'automatiser l'analyse des pointages.
Sélectionnez le type d'analyse, téléversez vos fichiers Excel et générez les rapports.
""")

# Create tabs for different analysis types
tab_bureau, tab_production, tab_annual, tab_employees = st.tabs([
    "📋 Analyse Bureau", "🔧 Analyse Production (9h)", "📊 Pivot Annuel", "👥 Gestion Employés"
])

# --- TAB 1: BUREAU ANALYSIS ---
with tab_bureau:
    st.header("Analyse Bureau (8h/jour)")
    st.markdown("""
    **Pour les employés de bureau :**
    - 8 heures de travail (Lundi-Vendredi)
    - 4 heures le Samedi
    - Règles standard de pause déjeuner
    """)
    
    uploaded_files_regular = st.file_uploader(
        "Téléversez les fichiers Excel pour l'analyse bureau (.xlsx, .xls)",
        type=['xlsx', '.xls'],
        accept_multiple_files=True,
        key="regular_uploader"
    )
    
    if st.button("🚀 Lancer l'Analyse Bureau", type="primary", key="regular_button"):
        if not uploaded_files_regular:
            st.warning("Veuillez d'abord téléverser des fichiers.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            status_text.text("Préparation de l'environnement...")
            reset_dirs()
            progress_bar.progress(10)

            status_text.text(f"Sauvegarde de {len(uploaded_files_regular)} fichiers...")
            for uploaded_file in uploaded_files_regular:
                file_path = os.path.join(TEMP_INPUT_DIR, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            progress_bar.progress(30)

            status_text.text("Exécution de l'analyse quotidienne...")
            try:
                daily_output = daily_script.process_daily_analysis(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if daily_output:
                    st.success(f"✅ Analyse Quotidienne générée : {os.path.basename(daily_output)}")
                else:
                    st.warning("⚠️ L'analyse quotidienne n'a rien généré (vérifiez les données).")
            except Exception as e:
                st.error(f"Erreur Analyse Quotidienne: {e}")
            progress_bar.progress(50)

            status_text.text("Exécution de l'analyse mensuelle...")
            try:
                monthly_output = monthly_script.process_monthly_analysis(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if monthly_output:
                    st.success(f"✅ Analyse Mensuelle générée : {os.path.basename(monthly_output)}")
                else:
                    st.warning("⚠️ L'analyse mensuelle n'a rien généré.")
            except Exception as e:
                st.error(f"Erreur Analyse Mensuelle: {e}")
            progress_bar.progress(70)

            status_text.text("Génération du graphique des retards...")
            graph_output = None
            try:
                graph_output = graph_script.generate_lateness_graph(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if graph_output:
                    st.success(f"✅ Graphique généré : {os.path.basename(graph_output)}")
                else:
                    st.warning("⚠️ Impossible de générer le graphique.")
            except Exception as e:
                st.error(f"Erreur Graphique: {e}")
            progress_bar.progress(90)

            status_text.text("Sauvegarde sur Supabase...")
            if daily_output and os.path.exists(daily_output):
                success, msg = upload_to_supabase(daily_output)
                if not success:
                    st.error(f"Erreur sauvegarde Supabase (daily): {msg}")
            if monthly_output and os.path.exists(monthly_output):
                success, msg = upload_to_supabase(monthly_output)
                if not success:
                    st.error(f"Erreur sauvegarde Supabase (monthly): {msg}")
            if graph_output and os.path.exists(graph_output):
                success, msg = upload_to_supabase(graph_output)
                if not success:
                    st.error(f"Erreur sauvegarde Supabase (graph): {msg}")

            status_text.text("Finalisation...")
            progress_bar.progress(100)
            
            st.divider()
            st.header("📂 Résultats Analyse Bureau")

            if graph_output and os.path.exists(graph_output):
                st.image(graph_output, caption="Graphique des Retards (>10h)", width='stretch')
                with open(graph_output, "rb") as file:
                    st.download_button(
                        label="⬇️ Télécharger le Graphique (PNG)",
                        data=file,
                        file_name=os.path.basename(graph_output),
                        mime="image/png"
                    )

            st.subheader("Rapports Excel")
            files_found = False
            if os.path.exists(TEMP_OUTPUT_DIR):
                for f in os.listdir(TEMP_OUTPUT_DIR):
                    if f.endswith(".xlsx") and not f.startswith("~$") and "Production" not in f:
                        files_found = True
                        file_path = os.path.join(TEMP_OUTPUT_DIR, f)
                        with open(file_path, "rb") as file:
                            st.download_button(
                                label=f"⬇️ Télécharger {f}",
                                data=file,
                                file_name=f,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
            
            if not files_found:
                st.info("Aucun rapport Excel trouvé dans le dossier de sortie.")

# --- TAB 2: PRODUCTION ANALYSIS ---
with tab_production:
    st.header("Analyse Production (9h/jour)")
    st.markdown("""
    **Pour les ouvriers Production (codes 130, 131, 140, 141) :**
    - 9 heures de travail (8h-18h, Lundi-Vendredi)
    - 5 heures le Samedi
    - Pause déjeuner Vendredi : 13h-14h30 (90 minutes)
    - Heure d'entrée : 8h00
    - Pénalité déjeuner : -1h si pas de scan
    """)
    
    uploaded_files_production = st.file_uploader(
        "Téléversez les fichiers Excel pour l'analyse Production (.xlsx, .xls)",
        type=['xlsx', '.xls'],
        accept_multiple_files=True,
        key="production_uploader"
    )
    
    if st.button("🚀 Lancer l'Analyse Production", type="primary", key="production_button"):
        if not uploaded_files_production:
            st.warning("Veuillez d'abord téléverser des fichiers.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            status_text.text("Préparation de l'environnement...")
            reset_dirs()
            progress_bar.progress(10)

            status_text.text(f"Sauvegarde de {len(uploaded_files_production)} fichiers...")
            for uploaded_file in uploaded_files_production:
                file_path = os.path.join(TEMP_INPUT_DIR, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            progress_bar.progress(30)

            status_text.text("Exécution de l'analyse quotidienne Production...")
            try:
                prod_daily_output = prod_daily_script.process_production_daily_analysis(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if prod_daily_output:
                    st.success(f"✅ Analyse Quotidienne Production générée : {os.path.basename(prod_daily_output)}")
                else:
                    st.warning("⚠️ L'analyse quotidienne Production n'a rien généré (vérifiez les données).")
            except Exception as e:
                st.error(f"Erreur Analyse Quotidienne Production: {e}")
                import traceback
                st.error(f"Détails: {traceback.format_exc()}")
            progress_bar.progress(60)

            status_text.text("Exécution de l'analyse mensuelle Production...")
            try:
                prod_monthly_output = prod_monthly_script.process_production_monthly_analysis(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if prod_monthly_output:
                    st.success(f"✅ Analyse Mensuelle Production générée : {os.path.basename(prod_monthly_output)}")
                else:
                    st.warning("⚠️ L'analyse mensuelle Production n'a rien généré.")
            except Exception as e:
                st.error(f"Erreur Analyse Mensuelle Production: {e}")
                import traceback
                st.error(f"Détails: {traceback.format_exc()}")
            progress_bar.progress(90)

            status_text.text("Sauvegarde sur Supabase...")
            if prod_daily_output and os.path.exists(prod_daily_output):
                success, msg = upload_to_supabase(prod_daily_output)
                if not success:
                    st.error(f"Erreur sauvegarde Supabase (prod daily): {msg}")
            if prod_monthly_output and os.path.exists(prod_monthly_output):
                success, msg = upload_to_supabase(prod_monthly_output)
                if not success:
                    st.error(f"Erreur sauvegarde Supabase (prod monthly): {msg}")

            status_text.text("Finalisation...")
            progress_bar.progress(100)
            
            st.divider()
            st.header("📂 Résultats Analyse Production")

            st.subheader("Rapports Excel Production")
            files_found = False
            if os.path.exists(TEMP_OUTPUT_DIR):
                for f in os.listdir(TEMP_OUTPUT_DIR):
                    if f.endswith(".xlsx") and not f.startswith("~$") and "production" in f.lower():
                        files_found = True
                        file_path = os.path.join(TEMP_OUTPUT_DIR, f)
                        with open(file_path, "rb") as file:
                            st.download_button(
                                label=f"⬇️ Télécharger {f}",
                                data=file,
                                file_name=f,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
            if not files_found:
                st.info("Aucun rapport Excel Production trouvé dans le dossier de sortie.")

# --- TAB 3: ANNUAL PIVOT ANALYSIS ---
with tab_annual:
    st.header("📊 Analyse Pivot Annuel")
    st.markdown("""
    **Pour l'analyse annuelle des heures travaillées :**
    - Génère un tableau pivot par employé
    - Heures travaillées par mois
    - Jours d'absence par mois
    - Totaux annuels
    - Fichiers Production uniquement
    """)
    
    uploaded_files_annual = st.file_uploader(
        "Téléversez les fichiers Excel pour l'analyse Pivot Annuelle (.xlsx, .xls)",
        type=['xlsx', '.xls'],
        accept_multiple_files=True,
        key="annual_uploader"
    )
    
    if st.button("🚀 Lancer l'Analyse Pivot Annuel", type="primary", key="annual_button"):
        if not uploaded_files_annual:
            st.warning("Veuillez d'abord téléverser des fichiers.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            status_text.text("Préparation de l'environnement...")
            reset_dirs()
            progress_bar.progress(10)

            status_text.text(f"Sauvegarde de {len(uploaded_files_annual)} fichiers...")
            for uploaded_file in uploaded_files_annual:
                file_path = os.path.join(TEMP_INPUT_DIR, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            progress_bar.progress(30)

            status_text.text("Exécution de l'analyse Pivot Annuel...")
            annual_output = None
            try:
                annual_output = annual_pivot_script.process_annual_pivot(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if annual_output:
                    st.success(f"✅ Analyse Pivot Annuel générée : {os.path.basename(annual_output)}")
                else:
                    st.warning("⚠️ L'analyse Pivot Annuel n'a rien généré (vérifiez les données Production).")
            except Exception as e:
                st.error(f"Erreur Analyse Pivot Annuel: {e}")
                import traceback
                st.error(f"Détails: {traceback.format_exc()}")
            progress_bar.progress(90)

            status_text.text("Finalisation...")
            progress_bar.progress(100)
            
            st.divider()
            st.header("📂 Résultats Analyse Pivot Annuel")

            st.subheader("Rapport Excel Pivot Annuel")
            files_found = False
            if os.path.exists(TEMP_OUTPUT_DIR):
                for f in os.listdir(TEMP_OUTPUT_DIR):
                    if f.endswith(".xlsx") and not f.startswith("~$") and ("ANNUAL_PIVOT" in f or "Annual_Pivot" in f):
                        files_found = True
                        file_path = os.path.join(TEMP_OUTPUT_DIR, f)
                        with open(file_path, "rb") as file:
                            st.download_button(
                                label=f"⬇️ Télécharger {f}",
                                data=file,
                                file_name=f,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
            if not files_found:
                st.info("Aucun rapport Pivot Annuel trouvé.")

# --- TAB 4: EMPLOYEE MANAGEMENT (BUREAU ONLY) ---
with tab_employees:
    st.header("👥 Gestion des Employés Bureau")
    st.markdown(
        "Gérez la liste des employés bureau et leurs services. "
        "Ces informations enrichissent les rapports **Bureau** avec les colonnes **Service**, **Responsable** et **Poste**."
    )

    employees = employees_db.load_employees()

    # ── Section 1: current list (editable + per-row delete) ──────────────────
    st.subheader("📋 Liste actuelle")
    st.info(f"**{len(employees)}** employé(s) dans la base de données.")

    if employees:
        df_emp = pd.DataFrame(employees)
    else:
        df_emp = pd.DataFrame(columns=['matricule', 'nom', 'prenom', 'service', 'poste', 'responsable', 'last_seen'])

    display_cols = ['matricule', 'nom', 'prenom', 'service', 'poste', 'responsable', 'last_seen']
    for col in display_cols:
        if col not in df_emp.columns:
            df_emp[col] = ''
    df_emp = df_emp[display_cols]

    st.caption("Double-cliquez une cellule pour modifier, puis cliquez **Sauvegarder**. Utilisez le bouton 🗑️ pour supprimer un employé individuellement.")

    edited_df = st.data_editor(
        df_emp,
        num_rows="fixed",
        width='stretch',
        column_config={
            "matricule":   st.column_config.TextColumn("Matricule",      width="small"),
            "nom":         st.column_config.TextColumn("Nom",            width="medium"),
            "prenom":      st.column_config.TextColumn("Prénom",         width="medium"),
            "service":     st.column_config.TextColumn("Service",        width="medium"),
            "poste":       st.column_config.TextColumn("Poste",          width="medium"),
            "responsable": st.column_config.TextColumn("Responsable",    width="medium"),
            "last_seen":   st.column_config.TextColumn("Dernier scan",   width="small", disabled=True),
        },
        key="employee_editor"
    )

    if st.button("💾 Sauvegarder les modifications", type="primary", key="save_employees"):
        records = edited_df.fillna('').to_dict('records')
        records = [r for r in records if str(r.get('nom', '')).strip()]
        employees_db.save_employees(records)
        st.success(f"✅ {len(records)} employés sauvegardés.")
        st.rerun()

    st.divider()

    # ── Section 2: individual remove ─────────────────────────────────────────
    st.subheader("🗑️ Supprimer un employé")
    if employees:
        emp_options = {
            f"{e.get('nom', '')} {e.get('prenom', '')} ({e.get('matricule', '') or 'sans mat.'})": i
            for i, e in enumerate(employees)
        }
        selected_label = st.selectbox("Sélectionner l'employé à supprimer", list(emp_options.keys()), key="del_select")
        if st.button("🗑️ Supprimer cet employé", type="secondary", key="del_single"):
            idx = emp_options[selected_label]
            removed_name = selected_label
            current = employees_db.load_employees()
            del current[idx]
            employees_db.save_employees(current)
            st.success(f"✅ **{removed_name}** supprimé.")
            st.rerun()
    else:
        st.info("Aucun employé dans la base de données.")

    st.divider()

    # ── Section 3: add a new employee ────────────────────────────────────────
    st.subheader("➕ Ajouter un employé")
    with st.form("add_employee_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        new_matricule  = c1.text_input("Matricule")
        new_nom        = c2.text_input("Nom *")
        new_prenom     = c3.text_input("Prénom")
        c4, c5, c6 = st.columns(3)
        new_service    = c4.text_input("Service *")
        new_poste      = c5.text_input("Poste")
        new_responsable = c6.text_input("Responsable")
        submitted = st.form_submit_button("Ajouter", type="primary")

    if submitted:
        if not new_nom.strip():
            st.error("Le champ **Nom** est obligatoire.")
        elif not new_service.strip():
            st.error("Le champ **Service** est obligatoire.")
        else:
            current = employees_db.load_employees()
            current.append({
                'matricule':    new_matricule.strip(),
                'nom':          new_nom.strip().upper(),
                'prenom':       new_prenom.strip().upper(),
                'responsable':  new_responsable.strip(),
                'service':      new_service.strip().lower(),
                'poste':        new_poste.strip(),
                'last_seen':    None,
            })
            employees_db.save_employees(current)
            st.success(f"✅ **{new_nom.upper()}** ajouté avec succès.")
            st.rerun()

    st.divider()

    # ── Section 4: inactive employees (>30 days without scan) ────────────────
    st.subheader("⚠️ Employés inactifs (> 30 jours sans scan)")
    inactive = employees_db.get_inactive()

    if not inactive:
        st.success("Aucun employé inactif détecté.")
    else:
        st.warning(
            f"**{len(inactive)}** employé(s) n'ont pas de scan depuis plus de 30 jours. "
            "Ils sont peut-être partis ou ont arrêté de pointer."
        )
        df_inactive = pd.DataFrame(inactive)[['matricule', 'nom', 'prenom', 'service', 'last_seen', 'days_inactive']]
        df_inactive.columns = ['Matricule', 'Nom', 'Prénom', 'Service', 'Dernier scan', 'Jours inactif']
        st.dataframe(df_inactive, width='stretch', hide_index=True)

        if st.button("🗑️ Supprimer tous les inactifs", type="secondary", key="remove_inactive"):
            removed = employees_db.remove_inactive()
            st.success(f"✅ {removed} employé(s) supprimé(s) de la base de données.")
            st.rerun()

st.sidebar.info("Application RH - Analyse Bureau & Production")
