import os
import datetime
import pandas as pd
from scripts.utils import _standardize_text


def explode_weekly_df(df, id_cols=['id'], date_col='date'):
    """
    Transforme un DataFrame avec des colonnes 'lundi' à 'vendredi' en un format long,
    avec une ligne par jour et une date calculée.

    Args:
        df (pd.DataFrame): DataFrame d'entrée.
        id_cols (str): Noms des colonnes identifiantes.
        date_col (str): Colonne contenant la date du vendredi de la semaine.

    Returns:
        pd.DataFrame: DataFrame transformé avec colonnes: id, date, h (valeur), jour (nom du jour).
    """

    day_cols=['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi']
    
    # Mapping pour reculer à partir de vendredi
    day_offsets = {
        'Lundi': -4,
        'Mardi': -3,
        'Mercredi': -2,
        'Jeudi': -1,
        'Vendredi': 0
    }

    # Transformation en long format
    df_long = df.melt(id_vars=id_cols+[date_col], value_vars=day_cols,
                      var_name='jour', value_name='h')

    # Calcul de la vraie date pour chaque jour
    df_long[date_col] = pd.to_datetime(df_long[date_col]) + df_long['jour'].map(lambda d: datetime.timedelta(days=day_offsets[d]))
    
    df_long['h'] = pd.to_numeric(df_long['h'], errors='coerce')

    # Réorganiser les colonnes (facultatif)
    df_long = df_long[df_long['h']>0][id_cols+[date_col, 'jour', 'h']]
    return df_long.sort_values(by=date_col)


# Reading every excel and aggregating in a single df
def aggregate_individual_data(folder_path):
    """
    Agrège les données individuelles de temps passé par projet à partir de fichiers Excel (.xlsm)
    situés dans un dossier spécifié. Effectue des vérifications de cohérence sur les noms et les projets,
    et génère un journal des problèmes rencontrés.

    Cette fonction lit chaque fichier Excel dans le dossier, extrait les données de la feuille "Archive",
    vérifie la cohérence entre le nom du fichier et les noms trouvés dans les données, ainsi que
    la présence de projets manquants. Les données sont ensuite concaténées en un seul DataFrame.

    Paramètres
    ----------
    folder_path : str
        Chemin absolu ou relatif vers le dossier contenant les fichiers Excel (.xlsm) à agréger.

    Retourne
    --------
    tuple
        Un tuple contenant :
        - aggregated_public_archives : pandas.DataFrame
            DataFrame agrégé contenant toutes les données individuelles de temps passé par projet.
        - log : str
    """


    file_names = [f for f in os.listdir(folder_path) if f.endswith('.xlsm')]
    df_list = []
    
    problematic_files = []
    problematic_names = []
    missing_projects = []
    
    for file in file_names:
        try:
            df = pd.read_excel(os.path.join(folder_path, file), sheet_name="Archive", engine='calamine', date_format={'Date': '%Y-%m-%d'})
            df = explode_weekly_df(df, id_cols=['Qui ?', 'Projet ?'], date_col='Date')
            
            expected_prenom_nom = _standardize_text('.'.join(file.split('.')[0].split('_')[2:][::-1])) 
            found_prenom_nom = [_standardize_text(name) for name in df['Qui ?'].unique()]
            
            if len(found_prenom_nom) > 1: # Plusieurs prenom.nom trouvé...
                problematic_names.append(f"{len(found_prenom_nom)} 'prenom.nom' trouvé(s) pour {expected_prenom_nom} [{', '.join(found_prenom_nom)}]")
            elif found_prenom_nom[0] != expected_prenom_nom: # prenom.nom trouvé différent de celui tiré du nom du fichier
                problematic_names.append(f"'prenom.nom' {found_prenom_nom[0]} trouvé pour {expected_prenom_nom}")
            
            mask_missing_projects = df['Projet ?'].isna()
            if mask_missing_projects.sum():
                missing_projects.append(f"{mask_missing_projects.sum()} projet(s) manquant(s) pour {expected_prenom_nom}")
            
            df_list.append(df)
            
        except Exception:
            problematic_files.append(file)
    
    aggregated_public_archives = pd.concat(df_list, ignore_index=True)
    aggregated_public_archives["Date"] = pd.to_datetime(aggregated_public_archives["Date"])
    
    # LOG
    log = []
    if problematic_files:
        log.append(f"Fichiers n'ayant pas pu être lus (potentiellement ouverts pendant l'agrégation) : {', '.join(problematic_files)}")
    if problematic_names:
        log.append(f"Fichiers ayant un problème de 'prenom.nom' :\n - {'\n - '.join(problematic_names)}")
    if missing_projects:
        log.append(f"Fichiers ayant un problème de projets :\n - {'\n - '.join(missing_projects)}")
    
    log = '\n\n' + '\n\n'.join(log) if log else ''

    return aggregated_public_archives, log