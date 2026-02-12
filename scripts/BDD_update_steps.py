import os
import datetime
import shutil
import pandas as pd
import numpy as np

from scripts.utils import (
    date_to_str,
    _standardize_text,
    _standardize_list_ValueDateValue
)


# ==================================================================
# STEP 1 : Chargement des données
# ==================================================================

def load_data(path_dict: dict[str, str]):
    
    # Chargement des 4 onglets de BDD
    path = os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['bdd_file'])
    vision_personne = pd.read_excel(path, sheet_name="Vision personne", engine='calamine')
    vision_contrat = pd.read_excel(path, sheet_name="Vision contrat", engine='calamine')
    utils_personne = pd.read_excel(path, sheet_name="Utils personne", engine='calamine')
    utils_contrat = pd.read_excel(path, sheet_name="Utils contrat", engine='calamine')
    
    # Chargement des données agrégées.
    aggregated_data = pd.read_excel(os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['aggregated_data_file']), engine='calamine')

    # Chargement des informations de statut.
    status_information = pd.read_excel(os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['projects_info_file']), sheet_name="STATUS_INFORMATIONS", engine='calamine')

    # Chargement des informations de contrats.
    budgets_dates_contrats = pd.read_excel(os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['projects_info_file']), sheet_name="BUDGETS_DATES", engine='calamine')
    
    return {
        'bdd_vision_personne': vision_personne, # Feuille de BDD donnant la vision du suivi des heures à la maille personne
        'bdd_vision_contrat': vision_contrat, # Feuille de BDD donnant la vision du suivi des heures à la maille contrat (projet)
        'bdd_utils_personne': utils_personne, # Feuille de BDD donnant l'historique du statut des personnes (grade, taux horaire, appartenance labs)
        'bdd_utils_contrat': utils_contrat, # Feuille de BDD donnant les informations sur les contrats (dates, budget)
        'aggregated_data': aggregated_data, # Données agrégées issues des fichiers individuels de suivi des heures (attention, certaines lignes de 2024 sont dans BDD mais pas dans les fichiers individuels)
        'status_information': status_information, # Information sur le statut actuel des personnes tiré du fichier de staffing prévisionnel
        'budgets_dates_contrats': budgets_dates_contrats # Informations des contrats tirés du fichier de staffing prévisionnel
    }
    


# ==================================================================
# STEP 2 : Standardisation des données chargées
# ==================================================================
    
def standardize_bdd_vision_personne(df: pd.DataFrame):
    
    df['prenom_nom'] = df['prenom_nom'].apply(_standardize_text)
    
    return df


def standardize_bdd_utils_personne(df: pd.DataFrame):
    
    df['prenom_nom'] = df['prenom_nom'].apply(_standardize_text)
    
    example = df.iloc[0].values
    if type(example[1]) is str: 
        df["grade"] = df["grade"].apply(_standardize_list_ValueDateValue)
    if type(example[2]) is str: 
        df["taux_horaire"] = df["taux_horaire"].apply(_standardize_list_ValueDateValue)
    if type(example[3]) is str: 
        df["appartenance_lab"] = df["appartenance_lab"].apply(_standardize_list_ValueDateValue)
        
    return df


def standardize_bdd_vision_contrat(df: pd.DataFrame): 
    # Rien pour l'instant, mais à modifier si on décide d'ajouter une étape sur la vision contrat tirée de BDD (avant MAJ)
    return df


def standardize_bdd_utils_contrat(df: pd.DataFrame):
    # Rien pour l'instant, mais à modifier si on décide d'ajouter une étape sur les infos de contrats tirées de BDD (avant MAJ)
    return df


def standardize_aggragated_data(df: pd.DataFrame):
    """
    Nettoie les données agrégées en retirant les lignes avec noms erronés.
    
    Returns:
        Tuple: (DataFrame nettoyé, liste de messages de log)
    """
    try:    
        log = []    
           
        # Masques de validation
        valid_project = (~df["Projet ?"].isna()) & (df["Projet ?"] != "ERREUR")
        valid_person = (~df["Qui ?"].isna()) & (df["Qui ?"] != "prenom.nom")
        
        # Reporting des problèmes
        if not valid_project.all():
            n_invalid = (~valid_project).sum()
            log.append(f"ATTENTION: {n_invalid} lignes avec nom de projet erroné (retirées). {df["Projet ?"][~valid_project].unique()}")
            
        if not valid_person.all():
            n_invalid = (~valid_person).sum()
            log.append(f"ATTENTION: {n_invalid} lignes avec nom de personne erroné (retirées). {df["Qui ?"][~valid_person].unique()}")
        
        # Filtrage et standardisation
        df = df[valid_project & valid_person].copy()
        df['Qui ?'] = df['Qui ?'].apply(_standardize_text)
        
        df["Projet ?"] = df["Projet ?"].apply(lambda x: ''.join(x.split("_Terminé - ")))

    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise

    return df, log


def standardize_status_information(df: pd.DataFrame):
    
    df = df[~ df["Prenom nom"].isna()]
    df['Prenom nom'] = df['Prenom nom'].apply(_standardize_text)
    
    return df


def standardize_contrat_information(df: pd.DataFrame):
    """
    Formate les informations budgétaires des contrats (en input colonnes _DATA et _ESG séparées).
    A adapter si on ajoute un nouveau lab
    
    Returns:
        DataFrame avec colonnes: nom_contrat, Budget, Plan_de_charge, Reste_a_faire,
                                 Date_start, Date_end, root_name
    """
    
    columns = ["nom_contrat", "Budget", "Plan_de_charge", "Reste_a_faire", "Date_start", "Date_end"]
    
    # Extraction DATA et ESG
    data_cols = {f"{c}_DATA": c for c in columns}
    esg_cols = {f"{c}_ESG": c for c in columns}
    
    df_data = df[list(data_cols.keys())].rename(columns=data_cols)
    df_data = df_data[~df_data["nom_contrat"].isna()]
    
    df_esg = df[list(esg_cols.keys())].rename(columns=esg_cols)
    df_esg = df_esg[~df_esg["nom_contrat"].isna()]
    
    df = pd.concat([df_data, df_esg], axis=0, ignore_index=True)
    
    # Ajout root_name (retire "_Terminé - ")
    df["root_name"] = df["nom_contrat"].apply(lambda x: ''.join(x.split("_Terminé - ")))
    
    # Conversion types
    df['Budget'] = pd.to_numeric(df['Budget'], errors='coerce')
    df['Date_start'] = pd.to_datetime(df['Date_start'], errors='coerce')
    df['Date_end'] = pd.to_datetime(df['Date_end'], errors='coerce')
    
    return df


def complete_aggregated_data_with_old_rows(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    # Ajoute les anciennes lignes de vision personne (BDD) dans les données agrégées (hedbo)
    
    aggregated_data = data["aggregated_data"]
    vision_personne = data["bdd_vision_personne"]
    
    aggregated_data["key"] = aggregated_data["Qui ?"] + aggregated_data["Date"].astype(str)
    vision_personne["key"] = vision_personne["prenom_nom"] + vision_personne["date"].astype(str)

    aggregated_data_keys = aggregated_data["key"].unique()
    vision_personne_keys = vision_personne["key"].unique()

    keys_to_add = vision_personne_keys[~ np.isin(vision_personne_keys, aggregated_data_keys)]
    old_rows_to_add = vision_personne[np.isin(vision_personne["key"], keys_to_add)][["prenom_nom", "root_name", "date", "h"]].copy()
    old_rows_to_add.columns = ['Qui ?', 'Projet ?', 'Date', 'h']

    log = [f"{len(old_rows_to_add)} lignes (combinaison personne X date nouvelle) trouvées dans la version précédente de la BDD, mais pas dans les données agrégées, ont été ajoutées."]
    
    aggregated_data.drop(columns=["key"], inplace=True)
    vision_personne.drop(columns=["key"], inplace=True)

    aggregated_data = pd.concat([old_rows_to_add, aggregated_data]).reset_index(drop=True)
    
    data["aggregated_data"] = aggregated_data

    return data, log


def uniformize_identities(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """
    Normalize person identifiers across datasets by resolving
    'first.last' ↔ 'last.first' mismatches.

    The reference source is `status_information["Prenom nom"]`.
    If a name is found in other datasets but not in the reference,
    its inverted form is tested and replaced if found.
    """

    vision_person = data["bdd_vision_personne"]
    utils_person = data["bdd_utils_personne"]
    aggregated_data = data["aggregated_data"]
    status_info = data["status_information"]

    reference_names = set(status_info["Prenom nom"])

    names_in_datasets = (
        set(vision_person["prenom_nom"])
        | set(utils_person["prenom_nom"])
        | set(aggregated_data["Qui ?"])
    )

    missing_in_reference = names_in_datasets - reference_names

    replacement_map = {}
    log_inverted, log_unknown = [], []

    for name in missing_in_reference:
        inverted_name = ".".join(name.split(".")[::-1])

        if inverted_name in reference_names:
            replacement_map[name] = inverted_name
            log_inverted.append(f"{name} → {inverted_name}")
        else:
            log_unknown.append(name)

    if replacement_map:
        vision_person.replace({"prenom_nom": replacement_map}, inplace=True)
        utils_person.replace({"prenom_nom": replacement_map}, inplace=True)
        aggregated_data.replace({"Qui ?": replacement_map}, inplace=True)

        data["bdd_vision_personne"] = vision_person
        data["bdd_utils_personne"] = utils_person
        data["aggregated_data"] = aggregated_data

    log = [] 
    if log_inverted:
        log.append(f"prenom.nom inversé(s) trouvé(s) : {', '.join(log_inverted)}")
    if log_unknown:
        log.append(f"prenom.nom inconnu(s) trouvé(s) (laissés tels quels) : {', '.join(log_unknown)}")

    return data, log



def standardize_pipe(loaded_data):
    # Pipe de standardization complet
    try:
        log = []
        
        loaded_data['bdd_vision_personne'] = standardize_bdd_vision_personne(loaded_data['bdd_vision_personne'])
        loaded_data['bdd_vision_contrat'] = standardize_bdd_vision_contrat(loaded_data['bdd_vision_contrat'])
        loaded_data['bdd_utils_personne'] = standardize_bdd_utils_personne(loaded_data['bdd_utils_personne'])
        loaded_data['bdd_utils_contrat'] = standardize_bdd_utils_contrat(loaded_data['bdd_utils_contrat'])
        
        loaded_data['aggregated_data'], log_ = standardize_aggragated_data(loaded_data['aggregated_data'])
        log += log_
        loaded_data['status_information'] = standardize_status_information(loaded_data['status_information'])
        loaded_data['budgets_dates_contrats'] = standardize_contrat_information(loaded_data['budgets_dates_contrats'])
        
        loaded_data, log_ = complete_aggregated_data_with_old_rows(loaded_data)
        log += log_
        
        loaded_data, log_ = uniformize_identities(loaded_data)
        log += log_
    
    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise
    
    return loaded_data, log




# ==================================================================
# STEP 3 : Mise à jour des feuilles de BDD
# ==================================================================

def update_utils_personne(status_information: pd.DataFrame,
                          utils_personne_prev: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """
    Met à jour les informations de statut (grade, taux horaire, appartenance lab), BDD 'Utils personne'.
    
    Logique:
        - Détecte les ajouts/suppressions/modifications
        - Pour chaque changement, ajoute une entrée datée dans le format adapté
        - Format: [value_0, threshold_date, value_1, ..., threshold_date_n, value_n]
        
    Args:
        status_information: Nouvelles informations de statut (source)
        utils_personne_prev: Informations précédentes (BDD actuelle)
        
    Returns:
        Tuple: (utils_personne mis à jour, log des changements)
    """
    
    def get_last_saturday() -> pd.Timestamp:
        """
        Retourne le dernier samedi à 00:00.
        
        Utilisé comme date seuil pour les mises à jour de statut.
        """
        today = datetime.datetime.today()
        weekday = today.weekday()
        
        # Calcul du nombre de jours à soustraire
        days_diff = weekday + 2 if weekday < 5 else weekday - 5
        last_saturday = today - datetime.timedelta(days=days_diff)
        
        return pd.to_datetime(last_saturday).normalize()

    try:
        log = []
        threshold_date = get_last_saturday()
        
        # Validation de la source
        duplicates = status_information['Prenom nom'].duplicated()
        if duplicates.any():
            dup_names = status_information.loc[duplicates, 'Prenom nom'].unique()
            log.append(f"ATTENTION: Doublons dans la source de statut: {', '.join(dup_names)}")
        
        # Analyse des changements
        prev_names = set(utils_personne_prev['prenom_nom'])
        new_names = set(status_information['Prenom nom'])
        
        removed_names = prev_names - new_names
        added_names = new_names - prev_names
        
        if removed_names:
            log.append(f"Personnes retirées ({len(removed_names)}): {', '.join(sorted(removed_names))}")
            log.append("  → Les dernières informations seront conservées.")
        
        if added_names:
            log.append(f"Personnes ajoutées ({len(added_names)}): {', '.join(sorted(added_names))}")
        
        # Construction des nouvelles données
        all_names = sorted(prev_names | new_names)
        new_data = []
        
        for prenom_nom in all_names:
            if prenom_nom in removed_names:
                # Conserver les dernières valeurs connues
                prev_row = utils_personne_prev[utils_personne_prev['prenom_nom'] == prenom_nom].iloc[0]
                new_data.append([
                    prenom_nom,
                    prev_row['grade'],
                    prev_row['taux_horaire'],
                    prev_row['appartenance_lab']
                ])
                
            else:
                # Extraire les nouvelles valeurs
                source_row = status_information[status_information['Prenom nom'] == prenom_nom].iloc[0]
                
                grade_new = source_row['Grade']
                taux_new = float(source_row['Taux horaire (€)'])
                
                # Appartenance lab: trouver toutes les colonnes "lab*"
                lab_names = [col for col in source_row.index if 'lab' in col.lower()]
                lab_values = [float(source_row[col]) for col in lab_names]
                appartenance_new = [lab_names, lab_values]
                
                if prenom_nom in added_names:
                    # Nouvelle personne: initialiser
                    new_data.append([
                        prenom_nom,
                        [grade_new],
                        [taux_new],
                        [appartenance_new]
                    ])
                    
                else:
                    # Personne existante: vérifier les changements
                    prev_row = utils_personne_prev[utils_personne_prev['prenom_nom'] == prenom_nom].iloc[0]
                    
                    grade_prev = prev_row['grade'][-1]
                    taux_prev = prev_row['taux_horaire'][-1]
                    appartenance_prev = prev_row['appartenance_lab'][-1]
                    
                    # Copie profonde pour modification
                    row_data = [
                        prenom_nom,
                        list(prev_row['grade']),
                        list(prev_row['taux_horaire']),
                        list(prev_row['appartenance_lab'])
                    ]
                    
                    changes = []
                    
                    # Vérification grade
                    if grade_prev != grade_new:
                        row_data[1].extend([threshold_date, grade_new])
                        changes.append(f"GRADE : {grade_prev} → {grade_new}")
                    
                    # Vérification taux horaire
                    if taux_prev != taux_new:
                        row_data[2].extend([threshold_date, taux_new])
                        changes.append(f"TAUX HORAIRE : {taux_prev:.1f}€ → {taux_new:.1f}€")
                    
                    # Vérification appartenance lab
                    if appartenance_prev != appartenance_new:
                        row_data[3].extend([threshold_date, appartenance_new])
                        changes.append(f"LABS : {appartenance_new}")
                    
                    if changes:
                        log.append(f"  > {prenom_nom}: {' | '.join(changes)}")
                    
                    new_data.append(row_data)
        
        # Création du nouveau DataFrame
        utils_personne_new = pd.DataFrame(
            new_data,
            columns=['prenom_nom', 'grade', 'taux_horaire', 'appartenance_lab']
        )
    
    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise
    
    return utils_personne_new, log



def update_utils_contrat(budgets_dates_contrats: pd.DataFrame,
                         utils_contrat_prev: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Met à jour le feuille "Utils contrat" de BDD.
    La nouvelle version est simplement budgets_dates_contrats.
    Cette fonction détecte et rapporte les changements pour information.
    
    Returns:
        Tuple: (budgets_dates_contrats, log des changements)
    """
    try:
        log = []
        
        # Helper pour comparer et rapporter les changements
        def report_changes(prev_df: pd.DataFrame, 
                        new_df: pd.DataFrame,
                        column: str,
                        label: str,
                        is_monetary: bool = True) -> list[str]:
            """
            Compare deux DataFrames et rapporte les différences sur une colonne.
            """
            changes = []
            
            # Filtrage des valeurs non-null
            prev_data = prev_df[['root_name', column]].dropna()
            new_data = new_df[['root_name', column]].dropna()
            
            # Vérification si changements
            if prev_data.shape != new_data.shape or \
            not prev_data.sort_values('root_name').equals(new_data.sort_values('root_name')):
                
                changes.append(f"\n=== Changements: {label} ===")
                
                prev_names = set(prev_data['root_name'])
                new_names = set(new_data['root_name'])
                
                # Suppressions
                removed = prev_names - new_names
                if removed:
                    changes.append(f"  Retirés ({len(removed)}):")
                    for name in sorted(removed):
                        changes.append(f"    - {name}")
                
                # Ajouts
                added = new_names - prev_names
                if added:
                    changes.append(f"  Ajoutés ({len(added)}):")
                    for name in sorted(added):
                        value = new_data[new_data['root_name'] == name][column].iloc[0]
                        if is_monetary:
                            changes.append(f"    - {name}: {value:,.0f}€")
                        else:
                            changes.append(f"    - {name}: {date_to_str(value)}")
                
                # Modifications
                merged = prev_data.merge(
                    new_data,
                    on='root_name',
                    suffixes=('_prev', '_new')
                )
                modified = merged[merged[f'{column}_prev'] != merged[f'{column}_new']]
                
                if len(modified) > 0:
                    changes.append(f"  Modifiés ({len(modified)}):")
                    for _, row in modified.iterrows():
                        if is_monetary:
                            changes.append(
                                f"    - {row['root_name']}: "
                                f"{row[f'{column}_prev']:,.0f}€ → {row[f'{column}_new']:,.0f}€"
                            )
                        else:
                            changes.append(
                                f"    - {row['root_name']}: "
                                f"{date_to_str(row[f'{column}_prev'])} → {date_to_str(row[f'{column}_new'])}"
                            )
            
            return changes
        
        # Vérification des changements pour chaque type d'information
        log.extend(report_changes(utils_contrat_prev, budgets_dates_contrats, 'Budget', 'Budgets', is_monetary=True))
        log.extend(report_changes(utils_contrat_prev, budgets_dates_contrats, 'Date_start', 'Dates de début', is_monetary=False))
        log.extend(report_changes(utils_contrat_prev, budgets_dates_contrats, 'Date_end', 'Dates de fin', is_monetary=False))
        
        if not log:
            log = ["Aucun changement détecté dans les contrats."]
    
    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise
    
    return budgets_dates_contrats, log


def _apply_time_dependent_values(
    vision_personne,
    utils_personne,
    column,
    value_processor,
):
    """
    Moteur générique pour appliquer des valeurs dépendantes du temps (basées sur grade, taux horaire).

    value_processor(mask, value, prenom_nom) doit modifier directement
    les structures externes (ex: tableau résultat).
    """
    try:
        log = []

        mapping = dict(zip(utils_personne["prenom_nom"], utils_personne[column]))
        personnes = vision_personne["prenom_nom"].values
        dates = vision_personne["date"].values

        far_past = dates.min() - pd.Timedelta(days=1)
        far_future = dates.max() + pd.Timedelta(days=1)

        for prenom_nom, infos in mapping.items():

            if not infos or len(infos) % 2 == 0:
                log.append(
                    f"Format invalide pour {column} ({prenom_nom}, {infos})"
                )
                continue

            mask_person = personnes == prenom_nom

            # Ajout des dates aux bornes pour obtenir un format :
            # [date_min, valeur, date_max, ..., date_max_final]
            infos_completed = [far_past] + infos + [far_future]

            for i in range(len(infos_completed) // 2):
                d_min = infos_completed[2 * i]
                value = infos_completed[2 * i + 1]
                d_max = infos_completed[2 * i + 2]

                mask = mask_person & (dates > d_min) & (dates <= d_max)

                value_processor(mask, value, prenom_nom)
                
    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise
    
    return log



def build_column_tool(vision_personne, column, default_value, utils_personne):
    
    result = np.full(len(vision_personne), default_value, dtype=object)
    
    def processor(mask, value, _):
        result[mask] = value

    log = _apply_time_dependent_values(
        vision_personne,
        utils_personne,
        column,
        processor,
    )
    
    return result, log



def build_labs_columns(vision_personne, utils_personne):

    try:
        log = []

        # Collecte sécurisée des labs
        
        all_labs = set([lab for labs in [
            infos[-1][0] for infos in utils_personne["appartenance_lab"].values
            ] for lab in labs])
        all_labs = sorted(all_labs)
        lab_index = {lab: i for i, lab in enumerate(all_labs)}
        result = np.zeros((len(vision_personne), len(all_labs)))

        def processor(mask, value, prenom_nom):

            labs, proportions = value
            total = np.sum(proportions)

            if total == 0:
                log.append(f"PROBLEME - Répartition sommant à 0 : {prenom_nom} ({value}).")

            if total != 1:
                log.append(f"PROBLEME - Répartition ne sommant pas à 1 (corrigée) : {prenom_nom} ({value}).")
                proportions = proportions / total

            for lab, prop in zip(labs, proportions):
                result[mask, lab_index[lab]] = prop

        log_engine = _apply_time_dependent_values(
            vision_personne,
            utils_personne,
            "appartenance_lab", 
            processor,
        )

        log.extend(log_engine)
        
    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise
    
    return all_labs, result, log


def update_vision_personne(
    aggregated_data,
    utils_personne,
    utils_contrat,
    path_dict
):
    """
    Met à jour le feuille "Vision personne" de BDD.
    """

    try:
        log = []

        vision_personne = aggregated_data.rename(columns={"Qui ?": "prenom_nom", "Projet ?": "root_name",  "Date": "date"}).copy()
        
        # prénom.nom non trouvés dans utils_personne (suivi de staffing prévisionnel)
        mask_unknown_prenom_nom = ~ vision_personne["prenom_nom"].isin(utils_personne["prenom_nom"])
        n_unknown_prenom_nom = np.sum(mask_unknown_prenom_nom)
        if n_unknown_prenom_nom:
            unique_unknown = vision_personne["prenom_nom"][mask_unknown_prenom_nom].unique()
            log.append(f"{n_unknown_prenom_nom} prénom.nom inconnus détectés (valeurs : {unique_unknown}).")
        vision_personne = vision_personne[~ mask_unknown_prenom_nom].copy() #!


        # Retire "_Terminé - " de la colonne root_name (cohérent avec le root_name des contrats)
        vision_personne["root_name"] = vision_personne["root_name"].apply(lambda x: ''.join(x.split("_Terminé - ")))
        
        
        # Ajout de la colonne Entreprise
        entreprises = pd.read_excel(os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['projects_info_file']), engine='calamine')["Entreprises"].dropna().values
        # dict(root_name: entreprise | None)
        root_to_entreprise = {root: next((e for e in entreprises if f'- {e} -' in root), None) for root in vision_personne["root_name"].unique()}
        vision_personne['Entreprise'] = vision_personne['root_name'].replace(root_to_entreprise)
        
        
        # Ajouter le Type (interne, projet ...) et le sujet de chaque projet (root_name)
        candidate_folders = [folder_name for folder_name in os.listdir(os.path.join(path_dict['root_folder'], path_dict['link_rpath'])) if 'Analyses_budgetaires_categorisation_V' in folder_name]
        if not candidate_folders: 
            log.append("ERREUR: Aucun fichier excel dont le nom est de la forme 'Analyses_budgetaires_categorisation_V'+date(%Y%m%d)+'.xlsx n'a été trouvé. " \
                    + "Ce fichier donne des informations supplémentaires sur les projets (type de projet, type de financement, etc.)")
        
        path_to_categorisation = os.path.join(path_dict['root_folder'], path_dict['link_rpath'], str(np.sort(candidate_folders)[-1])) # On prend la plus grande date
        categorisation_projets = pd.read_excel(path_to_categorisation, index_col=0, engine='calamine') 
        
        if not categorisation_projets.index.is_unique:
            log.append(f"ERREUR: Attention, des doublons on été trouvés dans le fichier '{path_to_categorisation}', cela peut fausser les résultats. Veuillez corriger ce problème.")
        
        categorisation_projets = categorisation_projets.groupby(categorisation_projets.index).first()
        vision_personne = vision_personne.merge(categorisation_projets[["Type", "Sujet"]], left_on="root_name", right_index=True, how='left')
        
        projets_sans_categorie = set(vision_personne[vision_personne[["Type", "Sujet"]].isna().any(axis=1)].root_name.values)
        projets_sans_categorie = projets_sans_categorie - set(["ILB - Congés & Arrêts (hors jours fériés)", "ILB - Jour Férié"])
        if projets_sans_categorie:
            log.append("Projets pour lesquels il manque l'information de Type / Sujet :")
            log.append(' - '+'\n - '.join(projets_sans_categorie))
            log.append(f"Où modifier ? {path_to_categorisation}")

        
        # Ajoute grade à date, taux horaire associé
        grade_column, log_temp = build_column_tool(vision_personne, "grade", "UNKNOWN", utils_personne)
        log.extend(log_temp)
        taux_horaire_column, log_temp = build_column_tool(vision_personne, "taux_horaire", np.nan, utils_personne)
        log.extend(log_temp)
        
        vision_personne["grade"] = grade_column
        vision_personne["taux_horaire"] = taux_horaire_column
        vision_personne["coût"] = vision_personne["taux_horaire"]*vision_personne["h"]
        
        # Ajout des données des contrats dans la vision personne
        vision_personne = vision_personne.merge(
            utils_contrat[["root_name", "Budget", "Plan_de_charge", "Reste_a_faire", "Date_start", "Date_end"]]
            .rename(columns={
                "Budget": "budget_contrat",
                "Plan_de_charge": "plan_de_charge",
                "Reste_a_faire": "reste_a_faire",
                "Date_start": "debut_contrat",
                "Date_end": "fin_contrat",
            }),
            on="root_name",
            how="left",
        )
        
        # Ajoute appartenance labs
        labs, new_labs_cols, log_temp = build_labs_columns(vision_personne, utils_personne)
        log.extend(log_temp)
        
        for i in range(len(labs)): 
            vision_personne[labs[i]] = new_labs_cols[:,i]
            vision_personne["h_"+labs[i]] = vision_personne["h"]*vision_personne[labs[i]]
            vision_personne["coût_"+labs[i]] = vision_personne["coût"]*vision_personne[labs[i]]
            
        for grade in vision_personne["grade"].unique():
            vision_personne["h_"+grade] = np.where(vision_personne["grade"]==grade, vision_personne["h"], 0)
            vision_personne["coût_"+grade] = np.where(vision_personne["grade"]==grade, vision_personne["coût"], 0)

        vision_personne["finished_tag"] = (vision_personne["date"] > vision_personne["fin_contrat"])

        vision_personne = vision_personne.drop_duplicates().reset_index(drop=True)
        
    except Exception: # Si ça plante, on print le log courant avant de raise
        print('\n'.join(log))
        raise
  
    return vision_personne, log


def update_vision_contrat(vision_personne):
    """
    Met à jour le feuille "Vision contrat" de BDD.
    """
    
    agg_dict = {
        "Entreprise": "first",
        "Type": "first",
        "Sujet": "first",
        "budget_contrat": "first",
        "plan_de_charge": "first",
        "reste_a_faire": "first",
        "debut_contrat": "first",
        "fin_contrat": "first",
        "finished_tag": "max", # True is True somewhere
        
        "h": "sum",
        "coût": "sum",
    }
        
    for col in vision_personne.columns:
        if col[:2] == 'h_' or col[:5] == 'coût_':
            agg_dict[col] = "sum"
    
    vision_contrat = (
        vision_personne
        .groupby(["root_name", "date"])
        .agg(agg_dict)
        .reset_index()
    )

    return vision_contrat



# ==================================================================
# STEP 4 : Enregistrement
# ==================================================================

def save_BDD(path_dict, vision_personne, vision_contrat, utils_personne, utils_contrat):
    
    """
    Objectif : Mettre à jour la base de données BDD et créer une sauvegarde (backup).
    
    Points d'attention : Les colonnes au format [valeur_0, date_1, valeur_1, ...., date_n, valeur_n] (grade, taux_horaire et appartenance_lab)\
        sont stockées sous format str.
    """
    # transformation de utils_personne avec de sauver (dates en str compréhensibles)
    # Il est nécessaire de faire cette transformation pour simplifier la lecture du fichier.
    def transform_utils_format_to_str(elt):
        for i in range(1, len(elt), 2):
            elt[i] = date_to_str(elt[i])
        return elt
    
    # Pour éviter de modifier utils_personne.
    utils_personne_copy = utils_personne.copy()
    
    utils_personne_copy["grade"] = utils_personne_copy["grade"].apply(transform_utils_format_to_str)
    utils_personne_copy["taux_horaire"] = utils_personne_copy["taux_horaire"].apply(transform_utils_format_to_str)
    utils_personne_copy["appartenance_lab"] = utils_personne_copy["appartenance_lab"].apply(transform_utils_format_to_str)
    
    path_to_bdd_file = os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['bdd_file'])
    
    # Mise à jour de la Base de Données BDD
    with pd.ExcelWriter(path_to_bdd_file) as writer:
        
        vision_personne.to_excel(writer, sheet_name="Vision personne", index=False)
        vision_contrat.to_excel(writer, sheet_name="Vision contrat", index=False)
        utils_personne_copy.to_excel(writer, sheet_name="Utils personne", index=False)
        utils_contrat.to_excel(writer, sheet_name="Utils contrat", index=False)
    
    # Création du backup
    suffix = date_to_str(datetime.datetime.today())
    shutil.copyfile(
        src = path_to_bdd_file,
        dst = os.path.join(path_dict['root_folder'], path_dict['link_rpath'], path_dict['backup_bdd_rpath'], f"BDD_{suffix}.xlsx")
    )