import os
import datetime
import tkinter as tk
from tkinter import filedialog # important pour messagebox
import numpy as np
import pandas as pd

from scripts.utils import (
    date_to_str,
    get_last_friday,
    get_all_project_names
)

from scripts.aggregate_individual_data import aggregate_individual_data

from scripts.BDD_update_steps import (
    load_data,
    standardize_pipe,
    update_utils_personne,
    update_utils_contrat,
    update_vision_personne,
    update_vision_contrat,
    save_BDD
)


# GUI

class suivi_projet_GUI(tk.Tk):
    """
    Interface graphique tkinter.
    """
    
    def __init__(self):
        super().__init__()
        
        self.title("OUTIL - SUIVI DES HEURES")
        self.geometry("1000x360")
        self.resizable(width=False, height=False)
        
        self.last_friday = get_last_friday()
        self.aggregated_public_archives = None
        self.available_fridays = ["---"]
        self.available_fridays_translation_dict = {}
        
        self.PATHS = {
            'root_folder': 'P:/Equipe/Suivi de projet staffing/', # Dossier où sont rangés les fichiers individuels de suivi des heures 
            'agg_backup_rpath': 'Données agrégées', # Dossier où sont rangés les backups hebdos des données agréges 
            'link_rpath': 'Données agrégées/Liens/', # Chemin relatif depuis root où sont rangés les fichiers agrégés et faisant le lien avec d'autres ressources
            'bdd_file': 'BDD.xlsx', # BDD, fichier excel de sortie alimentant le PBI
            'aggregated_data_file': "Suivi_Projet_Agrégé.xlsx", # Données agrégées des fichiers individuels de suivi des heures
            'projects_info_file': "ILB_Projects.xlsx", # Liens avec le fichier de staffing prévisionnel
            'backup_bdd_rpath': "BACKUPS_BDD/" # Sous dossier dans 'link_rpath' où sont rangés les backups hebdo de BDD
        }
        
        self.all_project_names = get_all_project_names(os.path.join(self.PATHS['root_folder'], self.PATHS['link_rpath'], self.PATHS['projects_info_file']))
        
        self.init_window() 

        self.analysis_window_is_open = False
        


    def init_window(self):
        
        LF1 = tk.LabelFrame(self)
        LF1.place(width=980, height=340, x=10, y=10)
        
        LF11 = tk.LabelFrame(LF1)
        LF11.place(width=400, height=185, x=10, y=60)
        
        LF12 = tk.LabelFrame(LF1)
        LF12.place(width=400, height=65, x=10, y=255)
        
        
        # ==================================================================
        # INFO : Dates des vendredis précédents (GAUCHE)
        # ==================================================================
        
        friday_info = f"Date du dernier vendredi :      {date_to_str(self.last_friday)}\nDate du vendredi précédent : {date_to_str(self.last_friday - datetime.timedelta(days=7))}"
        tk.Label(LF1, text=friday_info, anchor="w").place(width=900, height=30, x=20, y=15)
        
        
        # ==================================================================
        # UPDATE : Cadre d'agrégation (GAUCHE)
        # ==================================================================
        
        def agreger():
            """
            Agrège les données individuelles de suivi des heures et met à jour les widgets associés.
            On effectue également quelques tests sur les noms associés à des projets suspects
            """
            path = self.PATHS.get('root_folder', 'UNKNOWN')
            
            try:
                # Agrégation des données individuelles
                aggregated_public_archives, agg_log = aggregate_individual_data(path)
                
            except Exception as e:
                
                self.aggregated_public_archives = None
                tk.messagebox.showerror("!", f"Erreur lors de l'agrégation des données (vérifier le chemin fourni [{path}]).\n{e}")
                
                # Mise à jour des widgets
                light.config(bg="red")
                n_rows_L.config(text="")
                n_persons_L.config(text="")
                n_projects_L.config(text="")
                self.available_fridays = ["---"]
                friday_menu["menu"].delete(0, "end")
                friday_var.set("---")
                self.available_fridays_translation_dict = {}
                
                return
            
            self.aggregated_public_archives = aggregated_public_archives
            tk.messagebox.showinfo("", f"Agrégation des archives privées terminée [{path}].{agg_log}")
            
            # Mise à jour des widgets
            light.config(bg="lightgreen")
            n_rows_L.config(text=str(len(self.aggregated_public_archives)))
            n_persons_L.config(text=str(self.aggregated_public_archives["Qui ?"].nunique()))
            n_projects_L.config(text=str(self.aggregated_public_archives["Projet ?"].nunique()))
            
            # On prend uniquement les 12 derniers vendredi pour le menu déroulant
            n_friday_displayed = 12 #! MODIFIABLE
            self.available_fridays = self.aggregated_public_archives[self.aggregated_public_archives['jour'] == 'Vendredi']['Date'].drop_duplicates().sort_values().to_list()[:-n_friday_displayed -1:-1]
            self.available_fridays_translation_dict = dict(zip([date_to_str(date) for date in self.available_fridays], self.available_fridays))
            friday_menu["menu"].delete(0, "end")
            for option in self.available_fridays_translation_dict.keys():
                friday_menu["menu"].add_command(label=option, command=tk._setit(friday_var, option))
            
            def _highlight_suspect_projects():
                """
                On met en avant les projets trouvés qui n'apparaissent pas dans le fichier de staffing prévisionnel (et les personnes associées)
                """
                # On ne regarde qu'après threshold_date, car les modifications ont été apportées à ce moment
                threshold_date = '2026-01-01' #! MODIFIABLE
                filtered_agg_data = self.aggregated_public_archives[self.aggregated_public_archives['Date'] > pd.to_datetime(threshold_date)]
                used_project_names = filtered_agg_data["Projet ?"].unique()
                suspect_project_names = used_project_names[~ np.isin(used_project_names, self.all_project_names)]

                if len(suspect_project_names) > 0:
                    warning_message = f"Noms de projet suspects (depuis {threshold_date}):\n"
                    for suspect_project in suspect_project_names:
                        people = filtered_agg_data[filtered_agg_data["Projet ?"]==suspect_project]["Qui ?"].unique()
                        warning_message += f" - {suspect_project}: {', '.join(people)}\n"
                    tk.messagebox.showwarning("", warning_message)
            
            _highlight_suspect_projects()   
            
            
        tk.Button(LF11, text="AGREGER", bg="white", command=agreger).place(width=300, height=30, x=20, y=10)
        light = tk.Label(LF11, text="", bg="red", relief=tk.SOLID)
        light.place(width=30, height=30, x=340, y=10)
        
        tk.Label(LF11, text="Nombre de lignes :", anchor="w").place(width=130, height=20, x=20, y=50)
        tk.Label(LF11, text="Nombre de personnes :", anchor="w").place(width=130, height=20, x=20, y=75)
        tk.Label(LF11, text="Nombre de projets :", anchor="w").place(width=130, height=20, x=20, y=100)
        
        n_rows_L = tk.Label(LF11, text="", bg="white", anchor="w")
        n_rows_L.place(width=50, height=20, x=160, y=50)
        n_persons_L = tk.Label(LF11, text="", bg="white", anchor="w")
        n_persons_L.place(width=50, height=20, x=160, y=75)
        n_projects_L = tk.Label(LF11, text="", bg="white", anchor="w")
        n_projects_L.place(width=50, height=20, x=160, y=100)
        
        # ___________________
        
        def enregistrer():
            """
            Enregistrer les données agrégées de self.aggregated_public_archives
            """
            if type(self.aggregated_public_archives) is not type(None):
                
                main_file_path = os.path.join(self.PATHS['root_folder'], self.PATHS['link_rpath'], self.PATHS['aggregated_data_file'])
                backup_file_name = f"Suivi_projet_agrégé_{date_to_str(self.last_friday)}.xlsx"
                backup_file_path = os.path.join(self.PATHS['root_folder'], self.PATHS['agg_backup_rpath'], backup_file_name)
                
                # check if already exists, if so, ask for validation before replacing.
                if os.path.exists(backup_file_path):
                    if not tk.messagebox.askokcancel("", f"Les données semblent déjà avoir été agrégées cette semaine.\nSouhaitez-vous remplacer le fichier existant ?\nMAIN : {main_file_path}\nBACKUP : {backup_file_path}"): 
                        return
                    
                try:
                    self.aggregated_public_archives.to_excel(backup_file_path, index=False, sheet_name="Détails")
                    self.aggregated_public_archives.to_excel(main_file_path, index=False, sheet_name="Détails")
                    tk.messagebox.showinfo("", f"Le fichier a été enregistré avec succès !\nMAIN : {main_file_path}\nBACKUP : {backup_file_path}") 
                except Exception as e:
                    tk.messagebox.showerror("!", f"ERREUR: Erreur lors de l'enregistrement du fichier.\n{e}") 
            
            else:
                tk.messagebox.showerror("!", "ERREUR: Pas de fichier agrégé disponible.")
        
        tk.Button(LF11, text="ENREGISTRER", bg="white", command=enregistrer).place(width=300, height=30, x=20, y=135)
        
        
        # ==================================================================
        # OBSERVATION : FAIT VS à FAIRE (DROITE)
        # ==================================================================
        
        def update_done_vs_todo(*args):
            """
            Met à jour les 2 panneaux à droite FAIT / A FAIRE lors qu'on sélectionne un vendredi dans le menu déroulant
            """
            try:
                look_at_friday = self.available_fridays_translation_dict[friday_var.get()]
                df = self.aggregated_public_archives[self.aggregated_public_archives["Date"] == look_at_friday]
                
                # self.aggregated_public_archives ne contient pas les données des anciens
                all_names = set(self.aggregated_public_archives["Qui ?"])
                done = set(df["Qui ?"])
                todo = all_names - done
                
            except Exception as e:
                tk.messagebox.showerror("", e)
                done, todo = [], []
                    
            complete_list.config(state="normal")
            complete_list.delete(1.0, tk.END)
            complete_list.insert(tk.END, '\n'.join(done))
            complete_list.config(state="disable")
            
            incomplete_list.config(state="normal")
            incomplete_list.delete(1.0, tk.END)
            incomplete_list.insert(tk.END, '\n'.join(todo))
            incomplete_list.config(state="disable")
            
            
        tk.Label(LF1, text="Pour le vendredi :", anchor="w").place(width=100, height=20, x=460, y=10)
        friday_var = tk.StringVar(LF1) 
        friday_var.set("---")
        friday_var.trace_add("write", update_done_vs_todo)
        friday_menu = tk.OptionMenu(LF1, friday_var, *self.available_fridays)
        friday_menu.config(bg="white")
        friday_menu["menu"].config(bg="white")
        friday_menu.place(width=100, height=30, x=570, y=5)
        
        tk.Label(LF1, text="FAIT").place(width=220, height=20, x=460, y=40)
        complete_list = tk.Text(LF1, wrap=tk.WORD)
        complete_list.place(width=220, height=260, x=460, y=60)

        def on_scroll(*args):
            complete_list.yview(*args)
            
        scrollbar1 = tk.Scrollbar(LF1, command=on_scroll)
        scrollbar1.place(width=13, height=260, x=680, y=60)
        complete_list.config(yscrollcommand=scrollbar1.set)
        complete_list.config(state="disable")
        
        tk.Label(LF1, text="A FAIRE").place(width=220, height=20, x=710, y=40)
        incomplete_list = tk.Text(LF1, wrap=tk.WORD)
        incomplete_list.place(width=220, height=260, x=710, y=60)

        def on_scroll(*args):
            incomplete_list.yview(*args)
            
        scrollbar2 = tk.Scrollbar(LF1, command=on_scroll)
        scrollbar2.place(width=13, height=260, x=930, y=60)
        incomplete_list.config(yscrollcommand=scrollbar2.set)
        incomplete_list.config(state="disable")
        
        # ==================================================================
        # UPDATE : Cadre de mise à jour BDD (GAUCHE)
        # ==================================================================
        
        def load_update_save_BDD():
            """
            On applique successivement toutes les étapes de mise à jour de la BDD (utils_BDD)
            Avec try/except et pop-ups pour prévenir l'utilisateur des changements 
            """
            
            # LOADING DATA
            try:
                data = load_data(
                    path_dict=self.PATHS
                )
                
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors du chargement des données.\nAbandon.\n\n{e}")
                return
            
            # STANDARDIZE DATA
            try:
                data, log = standardize_pipe(
                    loaded_data=data
                )
                if log:
                    log_str = f"{'\n\n'.join(log)}\n\n"
                else:
                    log_str = ""
                if not tk.messagebox.askokcancel("Standardisation des données", f"{log_str}Souhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                    return
                
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la satandardization des données.\nAbandon.\n\n{e}")
                return
            
            # UPDATE UTILS PERSONNE
            try:
                utils_personne_new, log = update_utils_personne(
                    status_information = data['status_information'], 
                    utils_personne_prev = data['bdd_utils_personne']
                )
                if log:
                    log_str = f"{'\n\n'.join(log)}\n\n"
                else:
                    log_str = ""
                if not tk.messagebox.askokcancel("Mise à jour 'Utils personne' (BDD)", f"{log_str}Souhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                    return
                
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour de 'Utils personne' (BDD).\nAbandon.\n\n{e}")
                return
            
            # UPDATE UTILS CONTRAT
            try:
                utils_contrat_new, log = update_utils_contrat(
                    budgets_dates_contrats=data['budgets_dates_contrats'], 
                    utils_contrat_prev=data['bdd_utils_contrat']
                )
                if log:
                    log_str = f"{'\n'.join(log)}\n"
                else:
                    log_str = ""
                if not tk.messagebox.askokcancel("Mise à jour 'Utils contrat' (BDD)", f"{log_str}Souhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                    return
                
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour de 'Utils contrat' (BDD).\nAbandon.\n\n{e}")
                return
            
            # UPDATE VISION PERSONNE
            try:
                vision_personne_new, log = update_vision_personne(
                    aggregated_data=data['aggregated_data'], 
                    utils_personne=utils_personne_new, 
                    utils_contrat=utils_contrat_new, 
                    path_dict=self.PATHS
                )
                if log:
                    log_str = f"{'\n\n'.join(log)}\n\n"
                else:
                    log_str = ""
                if not tk.messagebox.askokcancel("Mise à jour 'Vision personne' (BDD)", f"{log_str}Souhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                    return
                
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour de 'Vision personne' (BDD).\nAbandon.\n\n{e}")
                return
            
            # UPDATE VISION CONTRAT
            try:
                vision_contrat_new = update_vision_contrat(
                    vision_personne=vision_personne_new
                )
                
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour de 'Vision contrat' (BDD).\nAbandon.\n\n{e}")
                return

            # SAVING UPDATED BDD
            if tk.messagebox.askokcancel("Enregistrement de BDD", "Souhaitez-vous enregistrer la version mise à jour de BDD ? Peut prendre plusieurs secondes, un message s'affichera à la fin de l'enregistrement.\nUne version archivée sera également crée."):
                try:
                    save_BDD(
                        path_dict=self.PATHS, 
                        vision_personne=vision_personne_new, 
                        vision_contrat=vision_contrat_new, 
                        utils_personne=utils_personne_new, 
                        utils_contrat=utils_contrat_new
                    )
                    tk.messagebox.showinfo("Enregistrement", "Fichier enregistré et archivé avec succès !")
                    
                except Exception as e:
                    tk.messagebox.showerror("!", f"ERREUR : Problème lors l'enregistrement de BDD.\nAbandon.\n\n{e}")
            
        
        tk.Button(LF12, text="Chargement et Mise à jour BDD", bg="white", command=load_update_save_BDD).place(width=350, height=30, x=20, y=16)
        
        
        
if __name__ == "__main__":
    app = suivi_projet_GUI() 
    app.mainloop()
    
    
    