import datetime
import pandas as pd
import ast
from unidecode import unidecode


def get_last_friday():
    today = datetime.datetime.today()
    week_day = today.weekday()
    if week_day>=4: 
        diff = week_day-4
    else: 
        diff = week_day+3
    return today-datetime.timedelta(days=diff)


def date_to_str(date):
    return str(date.year)+'-'+str(date.month).zfill(2)+'-'+str(date.day).zfill(2)


def get_all_fridays(min_date, max_date):
    first_week_day = min_date.weekday() #4 is friday
    if first_week_day<=4: 
        diff = 4 - first_week_day
    else: 
        diff = 11 - first_week_day
    friday = min_date + datetime.timedelta(days=diff)
    all_fridays = []
    while friday<=max_date:
        all_fridays.append(friday)
        friday = friday + datetime.timedelta(days=7)
    return all_fridays


def get_all_project_names(path):
    
    all_projects = pd.read_excel(path, engine='calamine', sheet_name="ALL_PROJECTS", nrows=10000)
    all_projects.fillna('no_name', inplace=True)

    all_project_names = []
    for col in all_projects.columns:
        if col!='Entreprises':
            names = all_projects[col].values
            if col=="Kanban":
                all_project_names += ["Kanban - "+name for name in names[names!='no_name']]
            else:
                all_project_names += list(names[names!='no_name'])

    #remove "_Terminé - "
    all_project_names = [''.join(x.split("_Terminé - ")) for x in all_project_names]        
        
    return all_project_names


def _standardize_text(s: str) -> str: 
    return unidecode(s).lower().strip()


def _standardize_list_ValueDateValue(s: str) -> list: 
    # Example: s = "['Junior', '2025-06-12', 'Senior']" -> ["Junior", 2025-06-12, "Senior"]
    l: list = ast.literal_eval(s)
    for i in range(1, len(l), 2):
        l[i] = pd.to_datetime(l[i])
    return l