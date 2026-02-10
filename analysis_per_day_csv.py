import pandas as pd
import os
import re
import warnings
from datetime import datetime, timedelta

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- CONFIGURATION ---
NOM_FICHIER_SORTIE = "Analyse_Quotidienne_Rapport_Avec_Comptages.xlsx"

# LISTE DES EMPLOYÉS À EXCLURE PAR NOM (Insensible à la casse)
EMPLOYES_EXCLUS = [
    # "ABOU HASNAA", 
    "HMOURI ALI"
]

# CODES QUI SIGNIFIENT UN "OUVRIER"
CODES_OUVRIER = ['130', '140', '141', '131']

def clean_name_string(name):
    """Normalise les noms pour assurer la correspondance malgré les espaces/caractères cachés."""
    if not name:
        return ""
    name = str(name).upper()
    name = name.replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ')
    name = re.sub(r'\s+', ' ', name)
    return name.strip()

def parse_scan_times_from_string(times_str):
    """Analyse la chaîne de temps séparée par des pipes pour trouver toutes les entrées de temps."""
    if not times_str or pd.isna(times_str):
        return [], 0
    times = times_str.split('|') if times_str else []
    times = [t for t in times if t and ':' in t]  # Filter out empty strings and ensure time format
    return times, len(times)

def analyze_row_from_csv(row):
    """Calcule les indicateurs pour retard, pas de déjeuner, heures et demi-journée depuis les données CSV."""
    # Parse times from the times_list column
    times, scan_count = parse_scan_times_from_string(row.get('times_list', ''))
    
    # Initialiser les valeurs par défaut
    late_930 = False
    late_1000 = False
    late_1400 = False
    no_lunch = False
    is_half_day = False
    hours_worked = row.get('hours_worked', 0.0)
    is_absent = False

    # --- DÉTECTION D'ABSENCE ---
    if not times or scan_count == 0:
        is_absent = True
        return late_930, late_1000, late_1400, no_lunch, hours_worked, is_half_day, is_absent
    
    # --- LOGIQUE DE RETARD (Hiérarchie Stricte) ---
    try:
        first_scan_dt = datetime.strptime(times[0], '%H:%M')
        
        limit_930 = first_scan_dt.replace(hour=9, minute=30, second=0)
        limit_1000 = first_scan_dt.replace(hour=10, minute=0, second=0)
        limit_1400 = first_scan_dt.replace(hour=14, minute=0, second=0)

        # Priorité 1 : Retard après 14:00 (Prend le dessus)
        if first_scan_dt > limit_1400:
            late_1400 = True
            late_1000 = False
            late_930 = False
        # Priorité 2 : Retard après 10:00 (mais avant 14:00)
        elif first_scan_dt > limit_1000:
            late_1400 = False
            late_1000 = True
            late_930 = False
        # Priorité 3 : Retard après 09:30 (mais avant 10:00)
        elif first_scan_dt > limit_930:
            late_1400 = False
            late_1000 = False
            late_930 = True
    except:
        pass

    # --- VÉRIFICATION PAS DE DÉJEUNER ---
    # Si début d'après-midi, Pas de Déjeuner n'est pas applicable/déjà signalé par Retard 14h
    if late_1400:
        no_lunch = False
    else:
        no_lunch = len(times) < 4 and len(times) > 0
    
    # --- LOGIQUE DEMI-JOURNÉE ---
    day_str = str(row.get('day_str', '')).lower()
    is_saturday = day_str.startswith('sa')

    if not is_saturday and len(times) >= 2 and hours_worked > 0:
        try:
            t_first = datetime.strptime(times[0], '%H:%M')
            t_last = datetime.strptime(times[-1], '%H:%M')
            
            # Gérer la logique de quart de nuit au cas où
            if t_last < t_first: 
                t_last += timedelta(days=1)
                
            limit_1300 = t_first.replace(hour=13, minute=0, second=0)
            limit_1400_exit = t_first.replace(hour=14, minute=0, second=0)
            
            # Condition A : Entré >= 13:00 (Quart d'après-midi / Retard)
            cond_afternoon = (t_first >= limit_1300)
            
            # Condition B : Parti <= 14:00 (Quart de matin / Départ anticipé) ET Heures < 7
            cond_morning = (t_last <= limit_1400_exit) and (hours_worked < 7.0)
            
            if cond_afternoon or cond_morning:
                is_half_day = True
        except:
            pass
    
    return late_930, late_1000, late_1400, no_lunch, hours_worked, is_half_day, is_absent

def create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, flag_column, output_header):
    """Crée un DataFrame à 3 colonnes : [Nom, Compte, %]"""
    subset = daily_df[daily_df[flag_column]].copy()
    result = subset[['name']].reset_index(drop=True)
    
    # Utiliser les bonnes statistiques selon le type de jour cible
    if flag_column == 'is_under_hours' and daily_df['day_str'].iloc[0].startswith('Sa'):
        stats_to_use = monthly_stats_saturday
    else:
        stats_to_use = monthly_stats
    
    # Mapper les comptes depuis les statistiques mensuelles appropriées
    result['Count'] = result['name'].map(stats_to_use[flag_column]).fillna(0).astype(int)
    
    # Total des jours basé sur HEURES > 0 (Logique stricte "Jour Travailé")
    total_days = result['name'].map(stats_to_use['total_attendance']).fillna(1) 
    
    result['%'] = (result['Count'] / total_days)
    
    result.columns = [output_header, 'Count', '%']
    return result

def process_daily_analysis_from_csv(csv_dir, output_dir):
    """
    Traite les fichiers CSV dans csv_dir et sauvegarde l'analyse dans output_dir.
    Retourne le chemin du fichier généré ou None.
    """
    if not os.path.exists(csv_dir):
        print(f"Dossier CSV non trouvé : {csv_dir}")
        return None
    
    # S'assurer que le dossier de sortie existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    all_data = []

    # --- LIRE LES DONNÉES CSV ---
    print("Analyse des fichiers CSV...")
    csv_files = [f for f in os.listdir(csv_dir) if f.lower().endswith('.csv')]
    
    if not csv_files:
        print("Aucun fichier CSV trouvé.")
        return None
    
    for csv_file in csv_files:
        print(f"Lecture : {csv_file}...")
        full_path = os.path.join(csv_dir, csv_file)
        try:
            df = pd.read_csv(full_path)
            all_data.append(df)
        except Exception as e:
            print(f"Erreur lors de la lecture de {csv_file}: {e}")
            continue

    if not all_data:
        print("Aucune donnée valide trouvée dans les fichiers CSV.")
        return None

    # Combiner toutes les données
    df = pd.concat(all_data, ignore_index=True)

    # --- EXCLURE LES EMPLOYÉS PAR NOM ---
    if EMPLOYES_EXCLUS:
        print(f"\nFiltrage des noms exclus : {EMPLOYES_EXCLUS}")
        excluded_clean = [clean_name_string(name) for name in EMPLOYES_EXCLUS]
        initial_len = len(df)
        df = df[~df['name'].isin(excluded_clean)]
        print(f"Supprimé {initial_len - len(df)} enregistrements basés sur la liste d'exclusion de noms.")
    
    if df.empty:
        print("Toutes les données filtrées.")
        return None

    # --- DÉTECTION CHRONOLOGIQUE AMÉLIORÉE ---
    if 'day_numeric' in df.columns and not df.empty:
        # 1. Récupérer les infos de base
        month_num = df['month_num'].iloc[0] if 'month_num' in df.columns else '01'
        year_num = df['year_num'].iloc[0] if 'year_num' in df.columns else '2026'
        
        # 2. Identifier la séquence chronologique réelle
        unique_days_in_order = []
        seen = set()
        for d in df['day_numeric']:
            if d not in seen:
                unique_days_in_order.append(d)
                seen.add(d)

        real_start_day = unique_days_in_order[0]
        real_end_day = unique_days_in_order[-1]
        
        # Détecter s'il y a une transition de mois
        has_transition = False
        pivot_index = -1
        for i in range(len(unique_days_in_order) - 1):
            if unique_days_in_order[i] > unique_days_in_order[i+1]:
                has_transition = True
                pivot_index = i
                break
        
        print(f"\n--- ANALYSE DE LA PÉRIODE ---")
        print(f"Séquence détectée : {unique_days_in_order}")
        
        # 3. Définir le jour cible (le dernier jour chronologique)
        target_report_day = real_end_day
        
        # 4. Vérifier si le dernier jour est complet
        last_day_records = df[df['day_numeric'] == target_report_day]
        total_last_day = len(last_day_records)
        incomplete_count = len(last_day_records[last_day_records['scan_count'] <= 1])
        
        if total_last_day > 0 and (incomplete_count / total_last_day) > 0.5:
            print(f"DÉCISION : Le jour {target_report_day} est incomplet (en cours).")
            df = df[df['day_numeric'] != target_report_day].copy()
            if len(unique_days_in_order) > 1:
                target_report_day = unique_days_in_order[-2]
                real_end_day = target_report_day
            print(f"Nouveau jour cible : {target_report_day}")
        else:
            print(f"DÉCISION : Le jour {target_report_day} est complet.")

        # 5. Calcul du nom du mois pour le header
        month_names = {
            '01': 'Janvier', '02': 'Février', '03': 'Mars', '04': 'Avril',
            '05': 'Mai', '06': 'Juin', '07': 'Juillet', '08': 'Août',
            '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Décembre'
        }
        month_name = month_names.get(month_num, f'Mois {month_num}')
        
    else:
        print("Erreur: Aucune donnée numérique de jour trouvée.")
        return None

    # --- CALCUL DES MÉTRIQUES ---
    print("\nCalcul des métriques...")
    results = df.apply(analyze_row_from_csv, axis=1)
    
    df['is_late_930'] = [x[0] for x in results]
    df['is_late_1000'] = [x[1] for x in results]
    df['is_late_1400'] = [x[2] for x in results]
    df['no_lunch'] = [x[3] for x in results]
    df['hours_worked'] = [x[4] for x in results]
    df['is_half_day'] = [x[5] for x in results] 
    df['is_absent'] = [x[6] for x in results] 
    
    if 'day_str' in df.columns:
        mask_saturday = df['day_str'].astype(str).str.startswith('Sa')
        df.loc[mask_saturday, 'no_lunch'] = False

    # Calculate target hours for ALL days that needed to be worked (excluding Sundays and holidays)
    df['target_hours'] = df.apply(
        lambda row: 4.0 if (str(row['day_str']).startswith('Sa') and row.get('is_holiday', 0) == 0) 
        else (8.0 if (not str(row['day_str']).startswith('Sa') and not str(row['day_str']).startswith('Di') and row.get('is_holiday', 0) == 0) else 0.0), axis=1
    )
    df['is_under_hours'] = (df['scan_count'] > 0) & (df['hours_worked'] < df['target_hours'])

    # --- GÉNÉRATION DES STATISTIQUES ---
    cols_to_sum = ['is_late_930', 'is_late_1000', 'is_late_1400', 'no_lunch', 'is_half_day', 'is_absent']
    
    valid_days_df = df[df['hours_worked'] > 0]
    
    saturday_records = valid_days_df[valid_days_df['day_str'].str.startswith('Sa')]
    weekday_records = valid_days_df[~valid_days_df['day_str'].str.startswith('Sa')]
    
    monthly_stats_weekday = weekday_records.groupby('name')[cols_to_sum].sum()
    monthly_stats_weekday['total_attendance'] = weekday_records.groupby('name').size()
    monthly_stats_weekday['is_under_hours'] = weekday_records.groupby('name')['is_under_hours'].sum()
    
    monthly_stats_saturday = saturday_records.groupby('name')[cols_to_sum].sum()
    monthly_stats_saturday['total_attendance'] = saturday_records.groupby('name').size()
    monthly_stats_saturday['is_under_hours'] = saturday_records.groupby('name')['is_under_hours'].sum()
    
    monthly_stats = monthly_stats_weekday.combine_first(monthly_stats_saturday)
    
    # --- FILTRER POUR LE JOUR CIBLE DU RAPPORT ---
    if 'day_numeric' in df.columns:
        daily_df = df[df['day_numeric'] == target_report_day].copy()
    else:
        daily_df = pd.DataFrame()

    if daily_df.empty:
        print(f"\nATTENTION : Aucun enregistrement trouvé pour le Jour {target_report_day}.")
        return None

    sample_day_str = daily_df.iloc[0]['day_str'] if not daily_df.empty else ""
    is_target_saturday = str(sample_day_str).startswith('Sa')

    # --- PRÉPARER LES LISTES DE SORTIE ---
    under_header = "Moins de 4h" if is_target_saturday else "Moins de 8h"
    df_under = create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, 'is_under_hours', under_header)

    df_late_10 = create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, 'is_late_1000', "Entrée > 10:00")
    df_late_930 = create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, 'is_late_930', "Entrée > 09:30")
    df_late_1400 = create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, 'is_late_1400', "Entrée > 14:00")
    df_half_day = create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, 'is_half_day', "Demi-Journée")
    df_absent = create_category_dataframe(daily_df, monthly_stats, monthly_stats_saturday, 'is_absent', "Absence")

    if is_target_saturday:
        main_list = pd.concat([df_under, df_late_10, df_late_930, df_late_1400, df_absent], axis=1)
    else:
        df_no_lunch = create_category_dataframe(daily_df[~daily_df['is_half_day']], monthly_stats, monthly_stats_saturday, 'no_lunch', "Pas de Déjeuner")
        main_list = pd.concat([df_under, df_half_day, df_no_lunch, df_late_10, df_late_930, df_late_1400, df_absent], axis=1)

    # --- EXPORTER VERS EXCEL ---
    if not df.empty and 'day_numeric' in df.columns:
        if has_transition:
            first_month_days = unique_days_in_order[:pivot_index + 1]
            second_month_days = unique_days_in_order[pivot_index + 1:]
            total_days = len(first_month_days) + len(second_month_days)
        else:
            total_days = len(unique_days_in_order)
        
        dynamic_filename = f"POINTAGE ANALYSE DU {real_start_day:02d}-{month_num}-{year_num} A {real_end_day:02d}-{month_num}-{year_num}.xlsx"
        output_path = os.path.join(output_dir, dynamic_filename)
        header_text = f"Analyse Quotidienne - Période : {real_start_day} au {real_end_day} {month_name} {year_num}"
    else:
        header_text = "Analyse Quotidienne - Période non spécifiée"
        output_path = os.path.join(output_dir, NOM_FICHIER_SORTIE)
    
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            main_list.to_excel(writer, sheet_name='Analyse Quotidienne', index=False, header=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['Analyse Quotidienne']
            
            header_title = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 14, 'font_color': '#2F5597', 'border': 1
            })
            
            if len(main_list.columns) > 1:
                worksheet.merge_range(0, 0, 0, len(main_list.columns) - 1, header_text, header_title)
            else:
                worksheet.write(0, 0, header_text, header_title)
            
            header_blue = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })
            header_orange = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'fg_color': '#ED7D31', 'font_color': 'white', 'border': 1
            })
            header_red = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'fg_color': '#C00000', 'font_color': 'white', 'border': 1
            })
            
            body_left = workbook.add_format({'border': 1, 'align': 'left'})
            body_center = workbook.add_format({'border': 1, 'align': 'center'})
            body_pct = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0%'})

            max_rows = len(main_list)
            columns = main_list.columns.tolist()

            for i, col_name in enumerate(columns):
                col_name_str = str(col_name)
                
                col_format = body_left
                header_style = header_blue

                if "Count" in col_name_str:
                    header_style = header_orange
                    col_format = body_center
                elif "%" in col_name_str:
                    header_style = header_orange
                    col_format = body_pct
                elif "14:00" in col_name_str:
                    header_style = header_red
                elif "Demi-Journée" in col_name_str:
                    header_style = header_orange
                elif "Absence" in col_name_str:
                    header_style = header_red
                
                worksheet.write(1, i, col_name, header_style)

                col_data = main_list.iloc[:, i]
                max_data_len = 0
                if "%" in col_name_str:
                    max_data_len = 5
                else:
                    valid_data = col_data.dropna().astype(str)
                    if not valid_data.empty:
                        max_data_len = valid_data.map(len).max()
                
                final_width = max(max_data_len, len(col_name_str)) + 4
                worksheet.set_column(i, i, final_width)

                for row_idx in range(max_rows):
                    cell_val = main_list.iloc[row_idx, i]
                    if pd.isna(cell_val):
                        worksheet.write(row_idx + 2, i, "", col_format)
                    else:
                        worksheet.write(row_idx + 2, i, cell_val, col_format)

        print(f"\nSUCCÈS ! Rapport sauvegardé : {output_path}")
        return output_path

    except Exception as e:
        print(f"Erreur lors de la sauvegarde du fichier : {e}")
        return None

def process_daily_analysis(input_dir, output_dir):
    """
    Wrapper function for compatibility with existing app.py
    This function first extracts CSV data, then processes it.
    """
    # First create temp_csv directory
    temp_csv_dir = os.path.join(os.path.dirname(input_dir), "temp_csv")
    
    # Extract CSV data from Excel files
    import csv_extractor
    if not csv_extractor.process_all_excel_to_csv(input_dir, temp_csv_dir):
        print("Failed to extract CSV data")
        return None
    
    # Process the CSV data
    return process_daily_analysis_from_csv(temp_csv_dir, output_dir)

if __name__ == "__main__":
    # Test with CSV directory
    csv_directory = "temp_csv"
    output_directory = "temp_output"
    
    output = process_daily_analysis_from_csv(csv_directory, output_directory)
    if output:
        print(f"Fichier généré : {output}")
