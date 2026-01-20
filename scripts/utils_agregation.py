#utils 

import numpy as np
import os
import datetime
import pandas as pd
from datetime import timedelta


def get_last_friday():
    today = datetime.datetime.today()
    week_day = today.weekday()
    if week_day>=4: diff = week_day-4
    else: diff = week_day+3
    return today-datetime.timedelta(days=diff)


def date_to_str(date):
    return str(date.year)+'-'+str(date.month).zfill(2)+'-'+str(date.day).zfill(2)


def get_all_fridays(min_date, max_date):
    first_week_day = min_date.weekday() #4 is friday
    if first_week_day<=4: diff = 4 - first_week_day
    else: diff = 11 - first_week_day
    friday = min_date + datetime.timedelta(days=diff)
    all_fridays = []
    while friday<=max_date:
        all_fridays.append(friday)
        friday = friday + datetime.timedelta(days=7)
    return all_fridays


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
    df_long[date_col] = pd.to_datetime(df_long[date_col]) + df_long['jour'].map(lambda d: timedelta(days=day_offsets[d]))
    
    df_long['h'] = pd.to_numeric(df_long['h'], errors='coerce')

    # Réorganiser les colonnes (facultatif)
    df_long = df_long[df_long['h']>0][id_cols+[date_col, 'jour', 'h']]
    return df_long.sort_values(by=date_col)


# Reading every excel and aggregating in a single df
def aggregate_public_archives(folder_path):
    
    if folder_path[-1]!='/': folder_path += '/'
    
    file_names = []
    for (dirpath, dirnames, filenames) in os.walk(folder_path):
        file_names.extend(filenames)
        break
    
    df_list = []
    
    problematic_files = []
    for file in file_names:
        try:
            df = pd.read_excel(folder_path + file, sheet_name="Archive", engine='calamine', date_format={'Date': '%Y-%m-%d'})
            df = explode_weekly_df(df, id_cols=['Qui ?', 'Projet ?'], date_col='Date')
            df_list.append(df)
        except:
            problematic_files.append(file)
    
    aggregated_public_archives = pd.concat(df_list, ignore_index=True)
    
    aggregated_public_archives["Date"] = pd.to_datetime(aggregated_public_archives["Date"])
    
    return aggregated_public_archives, problematic_files


def aggregate_public_archives_with_old(folder_path):
    
    aggregated_public_archives, problematic_files = aggregate_public_archives(folder_path)
    
    path_to_old = "P:/Equipe/Suivi de projet staffing/Archive des anciens/Suivi_Projet_Agrégé - Anciens.xlsx"
    old_data = pd.read_excel(path_to_old, engine='calamine')
    
    aggregated_public_archives = pd.concat([aggregated_public_archives, old_data]).drop_duplicates(ignore_index=True)
    
    return aggregated_public_archives, problematic_files


def get_all_project_names(path):
    
    all_projects = pd.read_excel(path, engine='calamine', sheet_name="ALL_PROJECTS",nrows=10000)
    all_projects.fillna('no_name', inplace=True)

    all_project_names = []
    for col in all_projects.columns:
        if col!='Entreprises':
            names = all_projects[col].values
            if col=="Kanban":
                all_project_names += ["Kanban - "+name for name in names[names!='no_name']]
            else:
                all_project_names += list(names[names!='no_name'])


    all_project_names = np.array(all_project_names)

    for i in range(len(all_project_names)):    #remove "_Terminé - "
        if "_Terminé - " in all_project_names[i]: all_project_names[i] = ''.join(all_project_names[i].split("_Terminé - "))
        
    return all_project_names