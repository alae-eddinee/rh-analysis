import streamlit as st
import os
import shutil
import importlib.util
import sys

# --- CONFIGURATION ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_INPUT_DIR = os.path.join(BASE_DIR, "temp_input")
TEMP_OUTPUT_DIR = os.path.join(BASE_DIR, "temp_output")

# --- IMPORT FUNCTIONS DYNAMICALLY ---
def load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

# Load analysis scripts
daily_script = load_module_from_path("daily_analysis", os.path.join(BASE_DIR, "analysis_per_day+count.py"))
monthly_script = load_module_from_path("monthly_analysis", os.path.join(BASE_DIR, "analysis_per_month.py"))
graph_script = load_module_from_path("lateness_graph", os.path.join(BASE_DIR, "late_arrivals_graph.py"))

# Load TAP analysis scripts
tap_daily_script = load_module_from_path("tap_daily_analysis", os.path.join(BASE_DIR, "analysis_tap_daily.py"))
tap_monthly_script = load_module_from_path("tap_monthly_analysis", os.path.join(BASE_DIR, "analysis_tap_monthly.py"))

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
tab_regular, tab_tap = st.tabs(["📋 Analyse Standard", "🔧 Analyse TAP (9h)"])

# --- TAB 1: REGULAR ANALYSIS ---
with tab_regular:
    st.header("Analyse Standard (8h/jour)")
    st.markdown("""
    **Pour les employés réguliers :**
    - 8 heures de travail (Lundi-Vendredi)
    - 4 heures le Samedi
    - Règles standard de pause déjeuner
    """)
    
    uploaded_files_regular = st.file_uploader(
        "Téléversez les fichiers Excel pour l'analyse standard (.xlsx, .xls)",
        type=['xlsx', '.xls'],
        accept_multiple_files=True,
        key="regular_uploader"
    )
    
    if st.button("🚀 Lancer l'Analyse Standard", type="primary", key="regular_button"):
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

            status_text.text("Finalisation...")
            progress_bar.progress(100)
            
            st.divider()
            st.header("📂 Résultats Analyse Standard")

            if graph_output and os.path.exists(graph_output):
                st.image(graph_output, caption="Graphique des Retards (>10h)", use_container_width=True)
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
                    if f.endswith(".xlsx") and not f.startswith("~$") and "TAP" not in f:
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

# --- TAB 2: TAP ANALYSIS ---
with tab_tap:
    st.header("Analyse TAP (9h/jour)")
    st.markdown("""
    **Pour les ouvriers TAP (codes 130, 131, 140, 141) :**
    - 9 heures de travail (8h-18h, Lundi-Vendredi)
    - 5 heures le Samedi
    - Pause déjeuner Vendredi : 13h-14h30 (90 minutes)
    - Heure d'entrée : 8h00
    """)
    
    uploaded_files_tap = st.file_uploader(
        "Téléversez les fichiers Excel pour l'analyse TAP (.xlsx, .xls)",
        type=['xlsx', '.xls'],
        accept_multiple_files=True,
        key="tap_uploader"
    )
    
    if st.button("🚀 Lancer l'Analyse TAP", type="primary", key="tap_button"):
        if not uploaded_files_tap:
            st.warning("Veuillez d'abord téléverser des fichiers.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            status_text.text("Préparation de l'environnement...")
            reset_dirs()
            progress_bar.progress(10)

            status_text.text(f"Sauvegarde de {len(uploaded_files_tap)} fichiers...")
            for uploaded_file in uploaded_files_tap:
                file_path = os.path.join(TEMP_INPUT_DIR, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            progress_bar.progress(30)

            status_text.text("Exécution de l'analyse quotidienne TAP...")
            try:
                tap_daily_output = tap_daily_script.process_tap_daily_analysis(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if tap_daily_output:
                    st.success(f"✅ Analyse Quotidienne TAP générée : {os.path.basename(tap_daily_output)}")
                else:
                    st.warning("⚠️ L'analyse quotidienne TAP n'a rien généré (vérifiez les données).")
            except Exception as e:
                st.error(f"Erreur Analyse Quotidienne TAP: {e}")
                import traceback
                st.error(f"Détails: {traceback.format_exc()}")
            progress_bar.progress(60)

            status_text.text("Exécution de l'analyse mensuelle TAP...")
            try:
                tap_monthly_output = tap_monthly_script.process_tap_monthly_analysis(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
                if tap_monthly_output:
                    st.success(f"✅ Analyse Mensuelle TAP générée : {os.path.basename(tap_monthly_output)}")
                else:
                    st.warning("⚠️ L'analyse mensuelle TAP n'a rien généré.")
            except Exception as e:
                st.error(f"Erreur Analyse Mensuelle TAP: {e}")
                import traceback
                st.error(f"Détails: {traceback.format_exc()}")
            progress_bar.progress(90)

            status_text.text("Finalisation...")
            progress_bar.progress(100)
            
            st.divider()
            st.header("📂 Résultats Analyse TAP")

            st.subheader("Rapports Excel TAP")
            files_found = False
            if os.path.exists(TEMP_OUTPUT_DIR):
                for f in os.listdir(TEMP_OUTPUT_DIR):
                    if f.endswith(".xlsx") and not f.startswith("~$") and "TAP" in f:
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
                st.info("Aucun rapport Excel TAP trouvé dans le dossier de sortie.")

st.sidebar.info("Application RH - Analyse Standard & TAP")
