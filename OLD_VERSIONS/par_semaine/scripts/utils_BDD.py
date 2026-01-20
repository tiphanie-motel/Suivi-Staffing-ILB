import numpy as np
import pandas as pd
import datetime
import ast
import copy
import shutil
import os
from unidecode import unidecode


path_to_folder = "P:/Equipe/Suivi de projet staffing/Données agrégées/Liens/"

def load_BDD(path_to_folder):
    
    """
    Objectif : Charger la base de données actuelle (que l'on souhaite mettre à jour avec les nouvelles données agrégées) \
        à partir du fichier Excel décrit par path_to_folder.
    
    Fonctionnement :
     - Lit les feuilles "Vision personne", "Vision contrat", "Utils personne" et "Utils contrat" du fichier BDD.xlsx.
     - Convertit les champs prenom_nom en minuscules pour uniformiser les données.
     - Les colonnes grade, taux_horaire et appartenance_lab ont un format spécifique du type : [value, date, value, date, ..., date, value] \
         mais sont stockée en excel en tant que str. Il est donc nécessaire formater ces données, sous forme de liste, et avec les dates \
         sous format pandas.DateTime.
     
    Interactions : Appelle read_utils_format_from_str pour convertir les chaînes de caractères en listes avec des dates.
    
    Points d'attention : Vérifie que les colonnes spécifiques sont bien au format attendu avant de les convertir.
    """
    
    vision_personne = pd.read_excel(path_to_folder + "BDD.xlsx", sheet_name="Vision personne", engine='calamine')
    vision_contrat = pd.read_excel(path_to_folder + "BDD.xlsx", sheet_name="Vision contrat", engine='calamine')
    utils_personne = pd.read_excel(path_to_folder + "BDD.xlsx", sheet_name="Utils personne", engine='calamine')
    utils_contrat = pd.read_excel(path_to_folder + "BDD.xlsx", sheet_name="Utils contrat", engine='calamine')
    
    vision_personne.replace({"prenom_nom": lambda x: x.lower()}, inplace=True)
    utils_personne.replace({"prenom_nom": lambda x: x.lower()}, inplace=True)
    
    def read_utils_format_from_str(elt):
        elt_to_list = ast.literal_eval(elt)
        for i in range(1,len(elt_to_list),2):
            elt_to_list[i] = pd.to_datetime(elt_to_list[i])
        return elt_to_list
    
    example = utils_personne.iloc[0].values
    if type(example[1])==str: utils_personne["grade"] = utils_personne["grade"].apply(read_utils_format_from_str)
    if type(example[2])==str: utils_personne["taux_horaire"] = utils_personne["taux_horaire"].apply(read_utils_format_from_str)
    if type(example[3])==str: utils_personne["appartenance_lab"] = utils_personne["appartenance_lab"].apply(read_utils_format_from_str)
    
    return vision_personne, vision_contrat, utils_personne, utils_contrat



def load_current_data(path_to_folder, add_old_aggregated_data=False):
    
    """
    Objectif : Charger les données agrégées actuelles et optionnellement les anciennes données agrégées \
        (si ces dernière ont été modifiées par exemple, ou si l'on reconstruit la BDD de zero).
    
    Fonctionnement :
     - Lit les données agrégées à partir du fichier Suivi_Projet_Agrégé.xlsx.
     - Si add_old_aggregated_data est vrai, lit également les anciennes données agrégées et les concatène avec les données actuelles.
     - Supprime les lignes avec des noms de projet ou de personne erronés (valeurs manquantes, ou spéciales : "ERREUR" et "prenom.nom").
     - Charge les informations de statut et de contrats à partir du fichier ILB_Projects.xlsx.
    
    Points d'attention : Standardise le nom des projets en supprimant le préfix '_Terminé - '.
    """
    log = ""
    
    ### Chargement des informations les plus récentes.
    
    # Chargement des données agrégées.

    aggregated_data = pd.read_excel(path_to_folder + "Suivi_Projet_Agrégé.xlsx", engine='calamine')
    
    if add_old_aggregated_data:
        log += "Données agrégées 'ancienne version' (2024) chargées.\n"
        aggregated_data_old = pd.read_excel(path_to_folder+'BACKUPS/Suivi_Projet_Agrégé_pre_MAJ_241121_corrected.xlsx', engine='calamine')
        aggregated_data = pd.concat([aggregated_data_old, aggregated_data]).reset_index(drop=True)
    
    # Suppression des lignes pour lesquelles le nom du projet ou de la personne erroné de manière évidente.
    mask_projet = (~aggregated_data["Projet ?"].isna()) & (aggregated_data["Projet ?"]!="ERREUR")
    mask_personne = (~aggregated_data["Qui ?"].isna()) & (aggregated_data["Qui ?"]!="prenom.nom")
    
    if not mask_projet.all(): log += "ATTENTION : Problème lors du chargement des données agrégées. "+str(np.sum(~mask_projet))+" lignes avec nom de projet erroné.\n"
    if not mask_personne.all(): log += "ATTENTION : Problème lors du chargement des données agrégées. "+str(np.sum(~mask_personne))+" lignes avec nom de personne erroné.\n"
    
    aggregated_data = aggregated_data[mask_projet & mask_personne]
    aggregated_data["Qui ?"] = aggregated_data["Qui ?"].apply(unidecode).apply(lambda x: x.lower().strip())

    # Chargement des informations de statut.
    
    status_information = pd.read_excel(path_to_folder + "ILB_Projects.xlsx", sheet_name="STATUS_INFORMATIONS", engine='calamine')
    status_information = status_information[~ status_information["Prenom nom"].isna()]
    status_information["Prenom nom"] = status_information["Prenom nom"].apply(unidecode).apply(lambda x: x.lower().strip())

    # Chargement des informations de contrats.
    
    budgets_dates_raw = pd.read_excel(path_to_folder + "ILB_Projects.xlsx", sheet_name="BUDGETS_DATES", engine='calamine')
    columns = ["nom_contrat", "Budget", "Plan_de_charge", "Reste_a_faire", "Date_start", "Date_end"]
    budgets_dates_raw_data = budgets_dates_raw[[c+"_DATA" for c in columns]].rename(columns=dict(zip([c+"_DATA" for c in columns], columns)))
    budgets_dates_raw_data = budgets_dates_raw_data[~ budgets_dates_raw_data["nom_contrat"].isna()]
    budgets_dates_raw_esg = budgets_dates_raw[[c+"_ESG" for c in columns]].rename(columns=dict(zip([c+"_ESG" for c in columns], columns)))
    budgets_dates_raw_esg = budgets_dates_raw_esg[~ budgets_dates_raw_esg["nom_contrat"].isna()]
    budgets_dates = pd.concat([budgets_dates_raw_data, budgets_dates_raw_esg], axis=0)

    root_names_budg_dates = budgets_dates["nom_contrat"].values.copy()
    for i, name in enumerate(root_names_budg_dates):
        if "_Terminé - " in name:
            root_names_budg_dates[i] = ''.join(name.split("_Terminé - "))
    budgets_dates["root_name"] = root_names_budg_dates
    
    
    # Correction des types des colonnes dates et budgets
    budgets_dates['Budget'] = pd.to_numeric(budgets_dates['Budget'], errors='coerce')
    budgets_dates['Date_start'] = pd.to_datetime(budgets_dates['Date_start'], errors='coerce', format='%y-%m-%d')
    budgets_dates['Date_end'] = pd.to_datetime(budgets_dates['Date_end'], errors='coerce', format='%y-%m-%d')
    
    return aggregated_data, status_information, budgets_dates, log



def check_loaded_coherence(vision_personne, utils_personne, aggregated_data, status_information):
    
    """
    Objectif : Vérifier la cohérence des données agrégées chargées (issues de load_current_data)\
        avec les données courrantes de la BDD (issues de load_BDD). Effectue des modifications si besoin.
    
    Fonctionnement :
     - Vérifie les prenom_nom (identifiant les personnes) uniques dans les différentes DataFrames et \
         identifie les inversions de prenom_nom (prénom.nom au lieu de nom.prénom). Remplace les prenom_nom inversés.
     - Identifie les lignes de vision_personne qui ne sont pas présentes dans aggregated_data, et les ajoute. Pour \
         identifier ces nouvelles lignes, on utilise une clé composée pour les personnes de prenom_nom + date \
         (respectivement nom du projet + date pour les projets)
    
    Points d'attention : Vérifie que tous les noms sont cohérents et que toutes les lignes nécessaires sont présentes dans aggregated_data.
    """
    
    # Vérification de la cohérence des données chargées
    
    names_vision_personne = vision_personne["prenom_nom"].unique()
    names_utils_personne = utils_personne["prenom_nom"].unique()
    names_aggregated_data = aggregated_data["Qui ?"].unique()
    names_status_information = status_information["Prenom nom"].unique()

    # Cohérence des noms : inversions (nom.prénom au lieu de prénom.nom).
    all_names = np.concatenate([names_status_information, names_vision_personne, names_utils_personne, names_aggregated_data])
    
    # renvoie un dictionaire pour remplacer les noms inversés (prenom.nom == nom.prenom)
    inverted_names = ['.'.join(name.split('.')[::-1]) for name in all_names]
    replacement_dict, black_list = {}, []
    for i, name in enumerate(all_names):
        if name not in black_list and name in inverted_names:
            inverted_version = inverted_names[i]
            black_list += [name, inverted_version]
            replacement_dict[inverted_version] = name
            
    vision_personne.replace({"prenom_nom": replacement_dict}, inplace=True)
    utils_personne.replace({"prenom_nom": replacement_dict}, inplace=True)
    aggregated_data.replace({"Qui ?": replacement_dict}, inplace=True)
    status_information.replace({"Prenom nom": replacement_dict}, inplace=True)
    
    # On ajoute dans aggregated_data les lignes de vision_personne qui n'y sont pas déjà 
    aggregated_data["key"] = aggregated_data["Qui ?"] + aggregated_data["Date"].astype(str)
    vision_personne["key"] = vision_personne["prenom_nom"] + vision_personne["date"].astype(str)

    aggregated_data_keys = aggregated_data["key"].unique()
    vision_personne_keys = vision_personne["key"].unique()

    keys_to_add = vision_personne_keys[~ np.isin(vision_personne_keys, aggregated_data_keys)]
    old_rows_to_add = vision_personne[np.isin(vision_personne["key"], keys_to_add)][["prenom_nom", "root_name", "date", "h"]].copy()
    old_rows_to_add.columns = ['Qui ?', 'Projet ?', 'Date', 'h']
    
    log = str(len(old_rows_to_add))+" lignes (combinaison personne X date nouvelle) trouvées dans la version précédente de la BDD, mais pas dans les données agrégées, ont été ajoutées.\n"

    aggregated_data.drop(columns=["key"], inplace=True)
    vision_personne.drop(columns=["key"], inplace=True)
    
    aggregated_data = pd.concat([old_rows_to_add, aggregated_data]).reset_index(drop=True)
    
    return vision_personne, utils_personne, aggregated_data, status_information, log



def get_last_end_of_last_week():
    # last saturday 00:00
    today = datetime.datetime.today()
    week_day = today.weekday()
    if week_day>=5: diff = week_day-5
    else: diff = week_day+2
    return pd.to_datetime(today-datetime.timedelta(days=diff)).normalize()


def date_to_str(date):
    return str(date.year)+'-'+str(date.month).zfill(2)+'-'+str(date.day).zfill(2)


def update_utils_personne(status_information, utils_personne_prev):
    
    """
    Objectif : Mettre à jour les informations de statut dans la base de données.
    
    Fonctionnement :
     - Vérifie les doublons et les nouveaux noms ou noms disparus dans status_information.
     - Met à jour les informations de grade, taux horaire et appartenance aux labs en fonction des nouvelles données.
     
    Interactions : Utilise status_information pour mettre à jour utils_personne_prev.
    
    Points d'attention : Vérifie les changements et conserve les dernières informations enregistrées pour les noms disparus.
    """
    
    log = ""
    
    # Mise à jour des informations de statuts dans la BDD.
    threshold_date = get_last_end_of_last_week()
    
    prenom_nom_new = status_information["Prenom nom"].values
    prenom_nom_prev = utils_personne_prev["prenom_nom"].values
    
    # Vérification de la présence de doublons dans la source
    u, c = np.unique(prenom_nom_new, return_counts=True)
    if len(prenom_nom_new)!=len(u):
        log += "ATTENTION : La source de utils personne contient des doublons :\n   - "+"\n   - ".join(u[c > 1])+'\n\n'
        
    # Vérification des nouveaux noms, ou des noms disparus :
    prenom_nom_removed = prenom_nom_prev[~np.isin(prenom_nom_prev, prenom_nom_new)]
    if len(prenom_nom_removed)>0:
        log += "Les informations de statut concernant les personnes suivantes ont été retirées :\n   - "+"\n   - ".join(prenom_nom_removed) + '\n'
        log += " * Les dernières informations enregistrées seront conservées.\n\n"
        
    prenom_nom_added = prenom_nom_new[~np.isin(prenom_nom_new, prenom_nom_prev)]
    if len(prenom_nom_added)>0:
        log += "Les informations de statut concernant les personnes suivantes ont été ajoutées :\n   - "+"\n   - ".join(prenom_nom_added)+'\n\n'
    
    all_prenom_nom = np.unique(np.concatenate([prenom_nom_new, prenom_nom_prev]))
    data = []
    for prenom_nom in all_prenom_nom:
        
        if prenom_nom in prenom_nom_removed: 
            data.append(utils_personne_prev[utils_personne_prev["prenom_nom"]==prenom_nom][["prenom_nom","grade","taux_horaire","appartenance_lab"]].values[0,:].tolist())
            
        else:
            source_info = status_information[status_information["Prenom nom"]==prenom_nom]
            labs_new = [col for col in source_info.columns if "lab" in col.lower()]
            grade_new, taux_horaire_new, appartenance_lab_new = source_info["Grade"].values[0], source_info["Taux horaire (€)"].values[0], list(source_info[labs_new].values[0])
            
            if prenom_nom in prenom_nom_added: 
                data.append([prenom_nom, [grade_new], [float(taux_horaire_new)], [[labs_new, [float(e) for e in appartenance_lab_new]]]])
            
            else:
                data_prev = utils_personne_prev[utils_personne_prev["prenom_nom"]==prenom_nom][["prenom_nom","grade","taux_horaire","appartenance_lab"]].values[0,:].tolist()
                grade_prev, taux_horaire_prev, appartenance_lab_prev = data_prev[1][-1], data_prev[2][-1], data_prev[3][-1][-1]
                
                data_to_append = copy.deepcopy(data_prev)
                changes = np.zeros(3, dtype=bool)
                if grade_prev!=grade_new:
                    changes[0] = True
                    data_to_append[1] += [threshold_date, grade_new]
                if taux_horaire_prev!=taux_horaire_new:
                    changes[1] = True
                    data_to_append[2] += [threshold_date, float(taux_horaire_new)]
                if len(appartenance_lab_prev)!=len(appartenance_lab_new) or appartenance_lab_prev!=appartenance_lab_new:
                    changes[2] = True
                    data_to_append[3] += [labs_new, [float(e) for e in appartenance_lab_new]]
                
                if changes.any():
                    log += " * Les informations de "+prenom_nom+" ont été mises à jour : \n"
                    if changes[0]: log += "   - GRADE : "+str(grade_prev)+" --> "+str(grade_new) + '\n'
                    if changes[1]: log += "   - TAUX HORAIRE : "+str(round(taux_horaire_prev,1))+" --> "+str(round(taux_horaire_new,1)) + '\n'
                    if changes[2]: log += "   - LABS : "+str(labs_new)+" : "+str(appartenance_lab_new) + '\n'
                
                data.append(data_to_append)    
    
    utils_personne_new = pd.DataFrame(data, columns=["prenom_nom","grade","taux_horaire","appartenance_lab"])
    
    if len(log)>0: log = log[:-1]

    return utils_personne_new, log



def update_utils_contrat(budgets_dates, utils_contrat_prev):
    
    """
    Objectif : Renseigner les changements repérés dans la nouvelle version des informations sur les contrats.\
        La nouvelle version renvoyée correspond tout simplement à budgets_dates.
         
    Points d'attention : Vérifie les ajouts, suppressions et modifications des informations de contrats. Print le tout.
    """
    
    log = ""
    
    # Mise à jour des informations de contrats dans la BDD.
    
    utils_contrat_new = budgets_dates.copy()
    
    # check and report changes
    budget_info_prev = utils_contrat_prev[["root_name", "Budget"]].dropna()
    budget_info_new = utils_contrat_new[["root_name", "Budget"]].dropna()
    start_date_info_prev = utils_contrat_prev[["root_name", "Date_start"]].dropna()
    start_date_info_new = utils_contrat_new[["root_name", "Date_start"]].dropna()
    end_date_info_prev = utils_contrat_prev[["root_name", "Date_end"]].dropna()
    end_date_info_new = utils_contrat_new[["root_name", "Date_end"]].dropna()
    
    def check_and_report_changes(info_prev, info_new, name, column_name, is_montant=True):
        
        log = ""
        
        if (info_prev.shape!=info_new.shape) or (info_prev.sort_values(by=list(info_prev.columns)).values != info_new.sort_values(by=list(info_new.columns)).values).any():
            log += "Les informations concernant ["+name+"] ont été mises à jour.\n\n"
            root_names_prev, root_names_new = info_prev.root_name.values, info_new.root_name.values
            in_common = root_names_prev[np.isin(root_names_prev, root_names_new)]
            just_prev, just_new = root_names_prev[~ np.isin(root_names_prev, in_common)], root_names_new[~ np.isin(root_names_new, in_common)]
            
            if len(just_prev)>0: log += " - Les informations de ["+name+"] suivantes ont été retirées :\n   - "+"\n   - ".join(just_prev) + '\n\n'
            if len(just_new)>0: 
                log += " - Les informations de ["+name+"] suivantes ont été ajoutées :\n"
                for root_name in just_new:
                    values = info_new[info_new["root_name"]==root_name][["root_name", column_name]].values[0,:]
                    if is_montant: log += "   - "+str(values[0])+": "+str(np.round(values[1],0))+"€\n"
                    else:  log += "   - "+str(values[0])+": "+date_to_str(values[1]) + '\n'
                log += '\n'
                    
            changes = pd.merge(info_prev.rename(columns={column_name: column_name+"_prev"}), 
                            info_new.rename(columns={column_name: column_name+"_new"}), 
                            left_on="root_name", right_on="root_name", how='inner')
            
            changes = changes[changes[column_name+"_prev"]!=changes[column_name+"_new"]]
            if len(changes)>0:
                log += " - Les informations de ["+name+"] suivantes ont été modifiée :\n"
                values = changes.values
                for i in range(len(changes)):
                    if is_montant: log += "   - contrat ["+str(values[i,0])+"], ["+name+"] : "+str(np.round(values[i,1],0))+"€ --> "+str(np.round(values[i,2],0))+"€\n"
                    else: log += "   - contrat ["+str(values[i,0])+"], ["+name+"] : "+date_to_str(values[i,1])+" --> "+date_to_str(values[i,2]) + '\n'
                log += '\n'
        
        return log
    
    log += check_and_report_changes(budget_info_prev, budget_info_new, "budget", "Budget", is_montant=True)
    log += check_and_report_changes(start_date_info_prev, start_date_info_new, "date de début", "Date_start", is_montant=False)
    log += check_and_report_changes(end_date_info_prev, end_date_info_new, "date de fin", "Date_end", is_montant=False)
    
    if len(log)>1: log = log[:-2]
    
    return utils_contrat_new, log



def build_column_tool(vision_personne, column, default_value, utils_personne):
    
    """
    Objectif : Construire une colonne pour vision_personne en fonction des informations de utils_personne.
    
    Fonctionnement :
     - Permet de construire une colonne (correspondant à grade ou taux horaire) pour 'vision_personne' (en fonction de ces colonnes prenom_nom et date).
     - Pour trouver les informations nécessaires, on utilise utils_personne, avec le format [valeur_0, date_1, valeur_1, ...., date_n, valeur_n].
     - la colonne prend la valeur_0 jusqu'à la date_1 (inclus), puis valeur_1 jusqu'à date_1, etc. jusqu'à valeur_n
         
    Points d'attention : Les valeurs changent au cours du temps, indépendamment pour chaque personne.
        Vérifie le format des données ([valeur_0, date_1, valeur_1, ...., date_n, valeur_n]). \
        Si le format n'est pas vérifié pour certaines personnes, la valeur par défaut est utilisée.
    """
    
    log = ""
    
    dictionary = dict(zip(utils_personne["prenom_nom"].to_list(), utils_personne[column].to_list()))
    personnes, dates = vision_personne["prenom_nom"].values, vision_personne["date"].values

    far_old_date = min(dates) - 1
    far_future_date = max(dates) + 1

    new_col = np.empty(vision_personne.shape[0], dtype=object)
    new_col[:] = default_value
    
    personnes_manquantes = np.unique(personnes)
    personnes_manquantes = list(personnes_manquantes[~ np.isin(personnes_manquantes, list(dictionary.keys()))])

    for prenom_nom, infos in dictionary.items(): # on parcourt les infos de tout le monde.
        
        if len(infos)<1 or len(infos)%2==0: # vérification du format [v, d, v, ... v] (non vide te taille impaire)
            log+="Mauvais format de "+column+" pour "+prenom_nom+": "+str(infos)+'\n'
            personnes_manquantes.append(prenom_nom)
        else:
            
            mask = personnes==prenom_nom
            # on ajoute des date au bord pour se ramener à un format [d, v, d, ..., d],
            # duquel on peut extraire des tronçons [d_début, valeur, date_fin]
            infos_completed = [far_old_date] + infos + [far_future_date] # on ajoute des date au bord pour se ramener à un format [d, v, d, ..., d],
            
            for i in range(len(infos_completed)//2):
                date_min, date_max = infos_completed[2*i], infos_completed[2*i + 2]
                new_col[mask & (dates>date_min) & (dates<=date_max)] = infos_completed[2*i + 1]
                
    if len(personnes_manquantes)>1:
        log+="ATTENTION : Pas d'information de "+column+" concernant les personnes suivantes:\n  - "+"\n  - ".join(np.unique(personnes_manquantes))+'\n'
        
    return new_col, log



def build_labs_columns(vision_personne, utils_personne):
    """
    Objectif : Construire les colonnes d'appartenance aux labs pour vision_personne. \
        L'appartenance est représentée sous forme de proportions (sommant à 1). \
        Exemple [1, 0] pour Data 100%, [0.5, 0.5] pour 50/50 Data ESG.\
        En prévision de futures évolutions, ce code peut gérer des proportions entre un nombre indéterminé de Labs. 
    
    Points d'attention : Les valeurs changent au cours du temps, indépendamment pour chaque personne.
        Vérifie le format des données ([valeur_0, date_1, valeur_1, ...., date_n, valeur_n]).
    """
    
    log = ""
    
    dictionary = dict(zip(utils_personne["prenom_nom"].to_list(), utils_personne["appartenance_lab"].to_list()))
    personnes, dates = vision_personne["prenom_nom"].values, vision_personne["date"].values
    
    far_old_date = min(dates) - 1
    far_future_date = max(dates) + 1

    all_labs = np.unique([elt for L in [v[-1][0] for v in dictionary.values()] for elt in L])
    new_cols = np.zeros((vision_personne.shape[0], len(all_labs)), dtype=float)
    position_dict = dict(zip(all_labs, range(len(all_labs))))
    
    personnes_manquantes = np.unique(personnes)
    personnes_manquantes = list(personnes_manquantes[~ np.isin(personnes_manquantes, list(dictionary.keys()))])

    for prenom_nom, infos in dictionary.items(): # on parcourt les infos de tout le monde.
        
        if len(infos)<1 or len(infos)%2==0: # vérification du format [v, d, v, ... v] (non vide te taille impaire)
            log += "Mauvais format d'appartenance aux labs pour "+prenom_nom+": "+str(infos)+'\n'
            personnes_manquantes.append(prenom_nom)
        else:
            
            mask = personnes==prenom_nom
            # on ajoute des date au bord pour se ramener à un format [d, v, d, ..., d],
            # duquel on peut extraire des tronçons [d_début, valeur, date_fin]
            infos_completed = [far_old_date] + infos + [far_future_date] # on ajoute des date au bord pour se ramener à un format [d, v, d, ..., d],
            
            for i in range(len(infos_completed)//2):
                date_min, date_max = infos_completed[2*i], infos_completed[2*i + 2]
                current_labs, proportions = infos_completed[2*i + 1]
                
                sum_proportions = np.sum(proportions)
                if sum_proportions==0: log += "PROBLEME - Répartition sommant à 0 : "+prenom_nom+" ("+str(infos)+").\n"
                elif sum_proportions!=1:
                    log += "PROBLEME - Répartition ne sommant pas à 1 (corrigée) : "+prenom_nom+" ("+str(infos)+").\n" 
                    proportions = proportions/sum_proportions
                
                for pos, lab in enumerate(current_labs): 
                    new_cols[mask & (dates>date_min) & (dates<=date_max), position_dict[lab]] = proportions[pos]
                        
    if len(personnes_manquantes)>1:
        log += "ATTENTION : Pas d'information d'appartenance aux Labs concernant les personnes suivantes:\n  - "+"\n  - ".join(np.unique(personnes_manquantes)) + '\n'
            
    return all_labs, new_cols, log



def create_vision_personne_and_contrat(aggregated_data, utils_personne, utils_contrat, path_to_folder=path_to_folder):
    
    """ !!! A revoir
    Objectif : Créer les DataFrames vision_personne et vision_contrat à jour à partir des données agrégées et des \
        informations utiles sur les personnes et les contrats.
    
    Fonctionnement :
     - Met à jour vision_personne avec les informations de grade, taux horaire, appartenance aux labs et informations de contrats.
     - Crée vision_contrat en agrégeant les données de vision_personne.
     
    Interactions : Appelle build_column_tool et build_labs_columns pour construire les colonnes nécessaires.
    
    Points d'attention : Vérifie les informations manquantes et les incohérences dans les données.
    """
    # Permet de créer les 2 dataframes pandas vision_personne et vision_contrat qui seront des feuilles de BDD.
    
    log = ""
    
    ### Mise à jour de la vision personne.
    i = 0
    vision_personne = aggregated_data.copy() # On part des données agrégées.
    vision_personne.rename(columns={"Qui ?": "prenom_nom", "Projet ?": "root_name", "Date": "date"}, inplace=True)

    # Retrait des lignes pour lesquelles les noms sont mauvais
    prenom_nom_unique = vision_personne["prenom_nom"].unique()
    prenom_nom_missing = prenom_nom_unique[~ np.isin(prenom_nom_unique, utils_personne["prenom_nom"].unique())] #les prenom.nom qui ne sont pas dans le suivi du statut

    if len(prenom_nom_missing)>0:
        log += "Problèmes sur les prenom_noms (tentative d'inversion prenom<->nom):\n"
        for prenom_nom in prenom_nom_missing:
            elts = prenom_nom.split('.')
            if len(elts)==2:
                inverted = elts[1]+'.'+elts[0]
                if inverted in utils_personne["prenom_nom"].unique(): #on avait nom.prénom au lieu de prénom.nom
                    vision_personne["prenom_nom"] = vision_personne["prenom_nom"].replace({prenom_nom: inverted})
                    log += " - "+prenom_nom+" remplacé par "+inverted+".\n"
                else:
                    log += " - problème avec "+prenom_nom+" : inconnu.\n"
            else:
                log += " - problème avec "+prenom_nom+" : trop d'éléments.\n"
        log+='\n'
    
    len_before = len(vision_personne)
    vision_personne = vision_personne[np.isin(vision_personne.prenom_nom, utils_personne["prenom_nom"].unique())]
    len_after = len(vision_personne)
    log += str(len_before-len_after)+" lignes ont été retirées pour cause de prénom.nom erroné.\n"
    
    # On enlève les tag "_Terminé - " des noms des projets.
    root_names = np.array([''.join(name.split("_Terminé - ")) for name in vision_personne["root_name"]])
    unique_root_names = np.unique(root_names)
    vision_personne["root_name"] = root_names
    
    entreprises_vect = pd.read_excel(path_to_folder + "ILB_Projects.xlsx", sheet_name="ALL_PROJECTS", engine='calamine')["Entreprises"].dropna().values

    root_to_entreprise = {} #dict, keys = root_names, values = entreprises/None
    for root_name in unique_root_names:
        root_to_entreprise[root_name] = next((e for e in entreprises_vect if f'- {e} -' in root_name), None)
    vision_personne['Entreprise'] = vision_personne['root_name'].replace(root_to_entreprise)
    
    # Ajouter le Type (interne, projet ...) et le sujet 
    candidate_folders = os.listdir("P:/Equipe/Suivi de projet staffing/Données agrégées/Liens/")
    candidate_folders = [folder_name for folder_name in candidate_folders if 'Analyses_budgetaires_categorisation_V' in folder_name]
    if len(candidate_folders)<1:
        log += "ERREUR: Aucun fichier excel dont le nom est de la forme 'Analyses_budgetaires_categorisation_V'+date(%Y%m%d)+'.xlsx n'a été trouvé. Ce fichier donne des inforamtions supplémentaires sur les projets (type de projet, type de financement, etc.)\n"

    folder_name_categorisation = str(np.sort(candidate_folders)[-1]) #on prend la plus grande date
    info_type_et_sujet =  pd.read_excel(f"P:/Equipe/Suivi de projet staffing/Données agrégées/Liens/{folder_name_categorisation}", index_col=0)
    
    if not info_type_et_sujet.index.is_unique:
        log += "ERREUR: Attention, des doublons on été trouvés dans le fichier '{folder_name_categorisation}', cela peut complétement fausser les résultats. Veuillez corriger ce problème.\n"
    
    info_type_et_sujet = info_type_et_sujet.groupby(info_type_et_sujet.index).first()
    vision_personne = vision_personne.merge(info_type_et_sujet[["Type", "Sujet"]], left_on="root_name", right_index=True, how='left')

    vision_personne["finished_tag"] = False
    # ajouter grade à date, appartenance labs, taux horaire associé
    grade_column, log_temp = build_column_tool(vision_personne, "grade", "UNKNOWN", utils_personne)
    log += log_temp
    taux_horaire_column, log_temp = build_column_tool(vision_personne, "taux_horaire", np.nan, utils_personne)
    log += log_temp
    labs, new_labs_cols, log_temp = build_labs_columns(vision_personne, utils_personne)
    log += log_temp
    
    vision_personne["grade"] = grade_column
    vision_personne["taux_horaire"] = taux_horaire_column
    vision_personne["coût"] = vision_personne["taux_horaire"]*vision_personne["h"]
    vision_personne = vision_personne.merge(utils_contrat[["root_name", "Budget", "Plan_de_charge", "Reste_a_faire", "Date_start", "Date_end"]].rename(
        columns={"Budget":"budget_contrat", "Plan_de_charge": "plan_de_charge", "Reste_a_faire": "reste_a_faire", "Date_start":"debut_contrat", "Date_end":"fin_contrat"}), 
                                            left_on="root_name", right_on="root_name", how='left')
    
    for i in range(len(labs)): vision_personne[labs[i]] = new_labs_cols[:,i]

    h_cout_par_labs_cols, h_cout_par_grade_cols = [], []
    for i in range(len(labs)): 
        h_cout_par_labs_cols.append("h_"+labs[i])
        h_cout_par_labs_cols.append("coût_"+labs[i])
        vision_personne["h_"+labs[i]] = vision_personne["h"]*vision_personne[labs[i]]
        vision_personne["coût_"+labs[i]] = vision_personne["coût"]*vision_personne[labs[i]]
    for grade in vision_personne["grade"].unique():
        h_cout_par_grade_cols.append("h_"+grade)
        h_cout_par_grade_cols.append("coût_"+grade)
        vision_personne["h_"+grade] = np.where(vision_personne["grade"]==grade, vision_personne["h"], 0)
        vision_personne["coût_"+grade] = np.where(vision_personne["grade"]==grade, vision_personne["coût"], 0)

    vision_personne["finished_tag"] = vision_personne["finished_tag"].values | (vision_personne["date"]>vision_personne["fin_contrat"])

    vision_personne.reset_index(inplace=True, drop=True)
    
    # Mise à jour de la vision personne terminée.
    
    ### Mise à jour de la vision contrat.
    
    agg_dict = dict(zip(["Entreprise", "Type", "Sujet", "budget_contrat", "plan_de_charge", "reste_a_faire", "debut_contrat", "fin_contrat", "finished_tag", "h", "coût"] + h_cout_par_labs_cols + h_cout_par_grade_cols, ["first"]*8+["min"]+["sum"]*(2+len(h_cout_par_labs_cols)+len(h_cout_par_grade_cols))))
    vision_contrat = vision_personne.groupby(["root_name", "date"]).agg(agg_dict)
    vision_contrat.reset_index(inplace=True)
    
    log += "\nProjets pour lesquels il manque l'information de Type / Sujet :\n"
    log += ' - '+'\n - '.join(vision_personne[vision_personne[["Type", "Sujet"]].isna().any(axis=1)].root_name.unique())+'\n'
    log += "Où modifier ? "+ path_to_folder + folder_name_categorisation+"\n"
    
    return vision_personne, vision_contrat, log



def update_BDD(path_to_folder, vision_personne, vision_contrat, utils_personne, utils_contrat):
    
    """
    Objectif : Mettre à jour la base de données BDD et créer une sauvegarde (backup).
    
    Points d'attention : Les colonnes au format [valeur_0, date_1, valeur_1, ...., date_n, valeur_n] (grade, taux_horaire et appartenance_lab)\
        sont stockées sous format str.
    """
    # transformation de utils_personne avec de sauver (dates en str compréhensibles)
    # Il est nécessaire de faire cette transformation pour simplifier la lecture du fichier.
    def transform_utils_format_to_str(elt):
        for i in range(1,len(elt),2):
            elt[i] = date_to_str(elt[i])
        return elt
    
    # Pour éviter de modifier utils_personne.
    utils_personne_copy = utils_personne.copy()
    
    utils_personne_copy["grade"] = utils_personne_copy["grade"].apply(transform_utils_format_to_str)
    utils_personne_copy["taux_horaire"] = utils_personne_copy["taux_horaire"].apply(transform_utils_format_to_str)
    utils_personne_copy["appartenance_lab"] = utils_personne_copy["appartenance_lab"].apply(transform_utils_format_to_str)
    
    # Mise à jour de la Base de Données BDD
    with pd.ExcelWriter(path_to_folder+"BDD.xlsx") as writer:
        
        vision_personne.to_excel(writer, sheet_name="Vision personne", index=False)
        vision_contrat.to_excel(writer, sheet_name="Vision contrat", index=False)
        utils_personne_copy.to_excel(writer, sheet_name="Utils personne", index=False)
        utils_contrat.to_excel(writer, sheet_name="Utils contrat", index=False)
    
    # Création du backup
    suffix = date_to_str(datetime.datetime.today())
    shutil.copyfile(path_to_folder+"BDD.xlsx", path_to_folder+"BACKUPS_BDD/BDD_"+suffix+".xlsx")