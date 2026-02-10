import streamlit as st
import os
import shutil
import importlib.util
import sys

# --- CONFIGURATION ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_INPUT_DIR = os.path.join(BASE_DIR, "temp_input")
TEMP_OUTPUT_DIR = os.path.join(BASE_DIR, "temp_output")
TEMP_CSV_DIR = os.path.join(BASE_DIR, "temp_csv")

# --- IMPORT FUNCTIONS DYNAMICALLY ---
def load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

# Charger les scripts d'analyse (versions CSV)
csv_extractor = load_module_from_path("csv_extractor", os.path.join(BASE_DIR, "csv_extractor.py"))
daily_script = load_module_from_path("daily_analysis_csv", os.path.join(BASE_DIR, "analysis_per_day_csv.py"))
monthly_script = load_module_from_path("monthly_analysis_csv", os.path.join(BASE_DIR, "analysis_per_month_csv.py"))
graph_script = load_module_from_path("lateness_graph", os.path.join(BASE_DIR, "late_arrivals_graph.py"))

# --- UTILS ---
def reset_dirs():
    """Réinitialise les dossiers temporaires."""
    for folder in [TEMP_INPUT_DIR, TEMP_OUTPUT_DIR, TEMP_CSV_DIR]:
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
Cette application permet d'automatiser l'analyse des pointages avec une extraction CSV pour éviter les erreurs d'interprétation.
1. **Téléversez** les fichiers Excel bruts dans la zone ci-dessous.
2. Cliquez sur **Lancer l'Analyse** pour extraire les données en CSV puis traiter.
3. **Téléchargez** les rapports Excel et le graphique générés.
""")

# 1. File Upload
uploaded_files = st.file_uploader("Déposez vos fichiers Excel ici (.xlsx, .xls)", type=['xlsx', 'xls'], accept_multiple_files=True)

if st.button("🚀 Lancer l'Analyse", type="primary"):
    if not uploaded_files:
        st.warning("Veuillez d'abord téléverser des fichiers.")
    else:
        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Step 1: Prep Environment
        status_text.text("Préparation de l'environnement...")
        reset_dirs()
        progress_bar.progress(10)

        # Step 2: Save Files
        status_text.text(f"Sauvegarde de {len(uploaded_files)} fichiers...")
        for uploaded_file in uploaded_files:
            file_path = os.path.join(TEMP_INPUT_DIR, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        progress_bar.progress(30)

        # Step 3: Extract CSV Data
        status_text.text("Extraction des données en format CSV...")
        try:
            csv_success = csv_extractor.process_all_excel_to_csv(TEMP_INPUT_DIR, TEMP_CSV_DIR)
            if csv_success:
                st.success("✅ Extraction CSV réussie")
            else:
                st.warning("⚠️ L'extraction CSV a échoué ou n'a pas généré de données")
                progress_bar.progress(100)
                st.stop()
        except Exception as e:
            st.error(f"Erreur Extraction CSV: {e}")
            progress_bar.progress(100)
            st.stop()
        progress_bar.progress(50)

        # Step 4: Run Daily Analysis (from CSV)
        status_text.text("Exécution de l'analyse quotidienne...")
        try:
            daily_output = daily_script.process_daily_analysis_from_csv(TEMP_CSV_DIR, TEMP_OUTPUT_DIR)
            if daily_output:
                st.success(f"✅ Analyse Quotidienne générée : {os.path.basename(daily_output)}")
            else:
                st.warning("⚠️ L'analyse quotidienne n'a rien généré (vérifiez les données).")
        except Exception as e:
            st.error(f"Erreur Analyse Quotidienne: {e}")
        progress_bar.progress(70)

        # Step 5: Run Monthly Analysis (from CSV)
        status_text.text("Exécution de l'analyse mensuelle...")
        try:
            monthly_output = monthly_script.process_monthly_analysis_from_csv(TEMP_CSV_DIR, TEMP_OUTPUT_DIR)
            if monthly_output:
                st.success(f"✅ Analyse Mensuelle générée : {os.path.basename(monthly_output)}")
            else:
                st.warning("⚠️ L'analyse mensuelle n'a rien généré.")
        except Exception as e:
            st.error(f"Erreur Analyse Mensuelle: {e}")
        progress_bar.progress(85)

        # Step 6: Generate Graph (from CSV)
        status_text.text("Génération du graphique des retards...")
        graph_output = None
        try:
            # Note: The graph script might need to be updated to work with CSV too
            # For now, we'll try the original approach
            graph_output = graph_script.generate_lateness_graph(TEMP_INPUT_DIR, TEMP_OUTPUT_DIR)
            if graph_output:
                st.success(f"✅ Graphique généré : {os.path.basename(graph_output)}")
            else:
                st.warning("⚠️ Impossible de générer le graphique.")
        except Exception as e:
            st.error(f"Erreur Graphique: {e}")
        progress_bar.progress(95)

        # Step 6: Finalize
        status_text.text("Finalisation...")
        progress_bar.progress(100)
        
        st.divider()
        st.header("📂 Résultats")

        # Display Graph
        if graph_output and os.path.exists(graph_output):
            st.image(graph_output, caption="Graphique des Retards (>10h)", use_container_width=True)
            with open(graph_output, "rb") as file:
                st.download_button(
                    label="⬇️ Télécharger le Graphique (PNG)",
                    data=file,
                    file_name=os.path.basename(graph_output),
                    mime="image/png"
                )

        # List Excel Files
        st.subheader("Rapports Excel")
        files_found = False
        if os.path.exists(TEMP_OUTPUT_DIR):
            for f in os.listdir(TEMP_OUTPUT_DIR):
                if f.endswith(".xlsx") and not f.startswith("~$"):
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
            st.info("Aucun rapport Excel n'a été trouvé dans le dossier de sortie.")

st.sidebar.info("Application créée pour l'automatisation RH.")
