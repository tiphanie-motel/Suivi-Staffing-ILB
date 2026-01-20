import tkinter as tk
from tkinter import filedialog

from scripts.utils_agregation import *
from scripts.utils_archivage_anciens_fichiers import *
from scripts.utils_BDD import *


# GUI

class suivi_projet_GUI(tk.Tk):
    
    def __init__(self):
        super().__init__()
        
        self.title("OUTIL - SUIVI DE PROJET")
        self.geometry("1000x600")
        self.resizable(width=False, height=False)
        
        self.last_friday = get_last_friday()
        self.aggregated_public_archives = None
        self.available_fridays = ["---"]
        self.available_fridays_translation_dict = {}
        
        self.all_project_names = get_all_project_names("P:/Equipe/Suivi de projet staffing/Données agrégées/Liens/ILB_Projects.xlsx")
        
        self.init_window_1() # Cadre dans lequel on précise les chemins.
        self.init_window_11() 
        self.init_window_2() 
        self.init_window_3() 

        self.analysis_window_is_open = False
        
        friday_info = "Date du dernier vendredi : "+date_to_str(self.last_friday)+"   |   Date du vendredi précédent : "+date_to_str(self.last_friday - datetime.timedelta(days=7))
        tk.Label(self, text=friday_info, anchor="w").place(width=900, height=20, x=20, y=20)
        
        
    def init_window_1(self): # Cadre dans lequel on précise les chemins.
        
        LF1 = tk.LabelFrame(self)
        LF1.place(width=980, height=90, x=10, y=50)
        
        tk.Label(LF1, text="Dossier contenant les fichiers de suivi individuels : ", anchor="w").place(width=270, height=30, x=20, y=5)
        tk.Label(LF1, text="Dossier contenant les fichiers de suivi agrégés : ", anchor="w").place(width=270, height=30, x=20, y=45)
        
        self.path_to_single_files = tk.Text(LF1, wrap=tk.WORD)
        self.path_to_single_files.insert(tk.END, "P:/Equipe/Suivi de projet staffing")
        self.path_to_single_files.place(width=600, height=20, x=290, y=10)
        self.path_to_single_files.config(state="disable")
        
        self.path_to_aggregated_files = tk.Text(LF1, wrap=tk.WORD)
        self.path_to_aggregated_files.insert(tk.END, "P:/Equipe/Suivi de projet staffing/Données agrégées")
        self.path_to_aggregated_files.place(width=600, height=20, x=290, y=50)
        self.path_to_aggregated_files.config(state="disable")
        
        def browse_1():
            folder_path = filedialog.askdirectory()
            if folder_path!="":
                self.path_to_single_files.config(state="normal")
                self.path_to_single_files.delete('1.0', tk.END)
                self.path_to_single_files.insert(tk.END, folder_path)
                self.path_to_single_files.config(state="disable")
            
        def browse_2():
            folder_path = filedialog.askdirectory()
            if folder_path!="":
                self.path_to_aggregated_files.config(state="normal")
                self.path_to_aggregated_files.delete('1.0', tk.END)
                self.path_to_aggregated_files.insert(tk.END, folder_path)
                self.path_to_aggregated_files.config(state="disable")
            
        tk.Button(LF1, text="...", bg="white", command=browse_1).place(width=30, height=20, x=910, y=10)
        tk.Button(LF1, text="...", bg="white", command=browse_2).place(width=30, height=20, x=910, y=50)
    
    
    def init_window_11(self): # Cadre par lequel on peut mettre à jour l'archive des anciens (là où sont stockés/archivés les fichiers des gens partis)
        # La mise à jour consiste à updater le fichier agrégé (et corrigé) qui rassemble et fige les données des anciens.
        
        LF11 = tk.LabelFrame(self)
        LF11.place(width=980, height=45, x=10, y=148)
        
        tk.Label(LF11, text="OPTIONNEL - M.A.J. archive des anciens (si fichiers modfiés ou ajoutés) :", anchor="w").place(width=400, height=30, x=20, y=5)
        
        def MAJ_archive_anciens():
            
            path = self.path_to_single_files.get("1.0", tk.END)
            if path[-1]=="\n": path=path[:-1]
            if path[-1]=="/": path=path[:-1]
            path_archive_anciens = path + '/Archive des anciens'
            
            try:
                aggregated_old_excels, problematic_files = aggregate_public_archives(path_archive_anciens + "/Fichiers des anciens - Archive")
            except:
                tk.messagebox.showerror("!", f"ERREUR: Mauvais chemin fourni pour l'archive des anciens [{path_archive_anciens}/Fichiers des anciens - Archive].")
                return
            
            if len(problematic_files)==0:
                tk.messagebox.showinfo("", f"Agrégation de l'archive des anciens.\nTous les fichiers ont été agrégés avec succès.\nShape : {aggregated_old_excels.shape}.")
            elif tk.messagebox.askyesno("!", f"Agrégation de l'archive des anciens.\nFichiers n'ayant pas pu être lus (potentiellement ouverts pendant l'agrégation) : {problematic_files}\nShape : {aggregated_old_excels.shape}.\nIl est vivement recommendé de régler le problème avant de poursuivre.\nSouhaitez-vous annuler ?"):
                return
                
            corrected_aggregated_old_excels = correct_aggregated_old_excels(aggregated_old_excels)
            
            try:
                corrected_aggregated_old_excels.to_excel(path_archive_anciens + "/Suivi_Projet_Agrégé - Anciens.xlsx", index=False, sheet_name="Détails")
                corrected_aggregated_old_excels.to_excel(path_archive_anciens + "/Données agrégés des anciens - Archive/Suivi_Projet_Agrégé_Anciens_"+date_to_str(datetime.datetime.today())+".xlsx", index=False, sheet_name="Détails")
                tk.messagebox.showinfo("", f"Fichier enregistré et archivé avec succès !")
            except:
                tk.messagebox.showerror("!", f"ERREUR: Problème lors de l'enregistrement des archive des anciens agrégées et corrigées. Abandon.")
        
        tk.Button(LF11, text="Mettre à jour", bg="white", command=MAJ_archive_anciens).place(width=200, height=25, x=420, y=7)
    

    def init_window_2(self):
        
        LF2 = tk.LabelFrame(self)
        LF2.place(width=980, height=340, x=10, y=200)
        
        LF21 = tk.LabelFrame(LF2)
        LF21.place(width=400, height=150, x=10, y=10)
        
        LF22 = tk.LabelFrame(LF2)
        LF22.place(width=400, height=150, x=10, y=175)
        
        # __________________
        
        def agreger():
            path = self.path_to_single_files.get("1.0", tk.END)
            if path[-1]=="\n": path=path[:-1]
            if path[-1]=="/": path=path[:-1]
            
            try:
                aggregated_public_archives, problematic_files = aggregate_public_archives_with_old(path)
                
            except:
                tk.messagebox.showerror("!", "ERREUR: Mauvais chemin fourni ["+path+"].")
                light.config(bg="red")
                n_rows_L.config(text="")
                n_persons_L.config(text="")
                n_projects_L.config(text="")
                self.aggregated_public_archives = None
                self.available_fridays = ["---"]
                friday_menu["menu"].delete(0, "end")
                friday_var.set("---")
                self.available_fridays_translation_dict = {}
                return
            
            self.aggregated_public_archives = aggregated_public_archives
            tk.messagebox.showinfo("", "Agrégation des archives privées terminée ["+path+"].\nFichiers n'ayant pas pu être lus (potentiellement ouverts pendant l'agrégation) : "+str(problematic_files))
            light.config(bg="lightgreen")
            n_rows_L.config(text=str(len(self.aggregated_public_archives)))
            n_persons_L.config(text=str(self.aggregated_public_archives["Qui ?"].nunique()))
            n_projects_L.config(text=str(self.aggregated_public_archives["Projet ?"].nunique()))
            self.available_fridays = list(pd.to_datetime(np.sort(self.aggregated_public_archives["Date"].unique())[::-1]))
            self.available_fridays_translation_dict = dict(zip([date_to_str(date) for date in self.available_fridays], self.available_fridays))
            friday_menu["menu"].delete(0, "end")
            for option in self.available_fridays_translation_dict.keys():
                friday_menu["menu"].add_command(label=option, command=tk._setit(friday_var, option))
            #friday_menu.config(command=update_done_vs_todo)
            
            # On ne regarde qu'après 2026-01-18, car on a fait des modifs depuis, et des projets de l'époques ne sont plus valides aujourd'hui
            used_project_names = self.aggregated_public_archives[self.aggregated_public_archives["Date"]>pd.to_datetime('2026-01-18')]["Projet ?"].unique()
            
            for i, projet in enumerate(used_project_names):
                if isinstance(projet, str) and "_Terminé - " in projet: 
                    used_project_names[i] = ''.join(projet.split("_Terminé - "))
            
            suspect_project_names = used_project_names[~ np.isin(used_project_names, self.all_project_names)]
            if len(suspect_project_names)>0:
                warning_message = "Suspect entries:\n"
                for suspect_project in suspect_project_names:
                    people = self.aggregated_public_archives[self.aggregated_public_archives["Projet ?"]==suspect_project]["Qui ?"].unique()
                    warning_message += f" - {suspect_project}: {', '.join(people)}\n"
                tk.messagebox.showinfo("", warning_message)
            
        tk.Button(LF21, text="AGREGER", bg="white", command=agreger).place(width=300, height=30, x=20, y=20)
        light = tk.Label(LF21, text="", bg="red", relief=tk.SOLID)
        light.place(width=30, height=30, x=340, y=20)
        
        tk.Label(LF21, text="Nombre de lignes :", anchor="w").place(width=130, height=20, x=20, y=60)
        tk.Label(LF21, text="Nombre de personnes :", anchor="w").place(width=130, height=20, x=20, y=85)
        tk.Label(LF21, text="Nombre de projets :", anchor="w").place(width=130, height=20, x=20, y=110)
        
        n_rows_L = tk.Label(LF21, text="", bg="white", anchor="w")
        n_rows_L.place(width=50, height=20, x=160, y=60)
        n_persons_L = tk.Label(LF21, text="", bg="white", anchor="w")
        n_persons_L.place(width=50, height=20, x=160, y=85)
        n_projects_L = tk.Label(LF21, text="", bg="white", anchor="w")
        n_projects_L.place(width=50, height=20, x=160, y=110)
        
        # ___________________
        
        tk.Label(LF22, text="Dossier où sauver le nouveau fichier : ", anchor="w").place(width=250, height=20, x=20, y=5)
        
        self.path_to_save_folder = tk.Text(LF22, wrap=tk.WORD)
        self.path_to_save_folder.insert(tk.END, self.path_to_aggregated_files.get("1.0", tk.END))
        self.path_to_save_folder.place(width=310, height=30, x=20, y=30)
        self.path_to_save_folder.config(state="disable")
        
        def browse():
            folder_path = filedialog.askdirectory()
            if folder_path!="":
                self.path_to_save_folder.config(state="normal")
                self.path_to_save_folder.delete('1.0', tk.END)
                self.path_to_save_folder.insert(tk.END, folder_path)
                self.path_to_save_folder.config(state="disable")
            
        tk.Button(LF22, text="...", bg="white", command=browse).place(width=30, height=20, x=340, y=30)
        
        tk.Label(LF22, text="Nom du nouveau fichier :", anchor="w").place(width=150, height=30, x=20, y=65)
        new_file_entry = tk.Entry(LF22)
        new_file_entry.insert(0,"Suivi_projet_agrégé_"+date_to_str(self.last_friday))
        new_file_entry.place(width=190, height=20, x=180, y=70)
        
        def enregistrer():
            if type(self.aggregated_public_archives)!=type(None):
                if new_file_entry.get()!="":
                    path_to_folder = self.path_to_save_folder.get("1.0", tk.END)
                    while path_to_folder[-1] in ["\n", "/"]: path_to_folder = path_to_folder[:-1]
                    new_file_name = new_file_entry.get()
                    while new_file_name[-1] in ["\n", "/"]: new_file_name = new_file_name[:-1]
                    if new_file_name[:-5]!=".xlsx": new_file_name+=".xlsx"
                    path_to_new_file = path_to_folder + "/" + new_file_name
                    # check if already exists, if so, ask for validation before replacing.
                    existing_files = os.listdir(path_to_folder)
                    if new_file_name in existing_files: 
                        if not tk.messagebox.askokcancel("", "Le nom de fichier '"+new_file_name+"' est déjà utilisé.\nSouhaitez-vous remplacer le fichier existant ?"): return
                        
                    try:
                        self.aggregated_public_archives.to_excel(path_to_new_file, index=False, sheet_name="Détails")
                        self.aggregated_public_archives.to_excel(path_to_folder + "/Liens/Suivi_Projet_Agrégé.xlsx", index=False, sheet_name="Détails")
                        tk.messagebox.showinfo("", "Le fichier a été enregistré avec succès !\n\n"+path_to_new_file) 
                    except:
                        tk.messagebox.showerror("!", f"ERREUR: Erreur lors de l'enregistrement du fichier. Vérifiez le chemin saisi. {path_to_folder}/Liens/Suivi_Projet_Agrégé.xlsx") 
                    
                else:
                   tk.messagebox.showerror("!", "ERREUR: Pas de nom de fichier renseigné.") 
            else:
                tk.messagebox.showerror("!", "ERREUR: Pas de fichier agrégé disponible.")
        
        tk.Button(LF22, text="ENREGISTRER", bg="white", command=enregistrer).place(width=350, height=30, x=20, y=105)
        
        # __________________
        
        def update_done_vs_todo(*args):
            
            try:
                df = self.aggregated_public_archives[self.aggregated_public_archives["Date"]==self.available_fridays_translation_dict[friday_var.get()]]
                all_names = self.aggregated_public_archives["Qui ?"].unique()
                done = df["Qui ?"].unique()
                todo = np.array([name for name in all_names if name not in done])
            except:
                tk.messagebox.showerror("","ERREUR : Problème lors de la lecture des données agrégées.")
                done, todo = [], []

            nan_pos = np.where(df["Qui ?"].isna())[0]
            if len(nan_pos)>0:
                message = ""
                for i in nan_pos:
                    next_to = []
                    if i>0: next_to.append(df["Qui ?"].iloc[i-1])
                    if i<len(df): next_to.append(df["Qui ?"].iloc[i+1])
                    message += f"\n - nan trouvé près de : {np.unique(next_to)}"
                        
                tk.messagebox.showwarning("",f"Warning : Identités manquantes ({len(nan_pos)}) lors de la lecture des données agrégées."+message)
                    
            complete_list.config(state="normal")
            complete_list.delete(1.0, tk.END)
            complete_list.insert(tk.END, '\n'.join(done))
            complete_list.config(state="disable")
            
            incomplete_list.config(state="normal")
            incomplete_list.delete(1.0, tk.END)
            incomplete_list.insert(tk.END, '\n'.join(todo))
            incomplete_list.config(state="disable")
            
            
        tk.Label(LF2, text="Pour le vendredi :", anchor="w").place(width=100, height=20, x=460, y=10)
        friday_var = tk.StringVar(LF2) 
        friday_var.set("---")
        friday_var.trace_add("write", update_done_vs_todo)
        friday_menu = tk.OptionMenu(LF2, friday_var, *self.available_fridays)
        friday_menu.config(bg="white")
        friday_menu["menu"].config(bg="white")
        friday_menu.place(width=100, height=30, x=570, y=5)
        
        tk.Label(LF2, text="FAIT").place(width=220, height=20, x=460, y=40)
        complete_list = tk.Text(LF2, wrap=tk.WORD)
        complete_list.place(width=220, height=260, x=460, y=60)

        def on_scroll(*args):
            complete_list.yview(*args)
            
        scrollbar1 = tk.Scrollbar(LF2, command=on_scroll)
        scrollbar1.place(width=13, height=260, x=680, y=60)
        complete_list.config(yscrollcommand=scrollbar1.set)
        complete_list.config(state="disable")
        
        tk.Label(LF2, text="A FAIRE").place(width=220, height=20, x=710, y=40)
        incomplete_list = tk.Text(LF2, wrap=tk.WORD)
        incomplete_list.place(width=220, height=260, x=710, y=60)

        def on_scroll(*args):
            incomplete_list.yview(*args)
            
        scrollbar2 = tk.Scrollbar(LF2, command=on_scroll)
        scrollbar2.place(width=13, height=260, x=930, y=60)
        incomplete_list.config(yscrollcommand=scrollbar2.set)
        incomplete_list.config(state="disable")
        
    def init_window_3(self):
        
        LF3 = tk.LabelFrame(self)
        LF3.place(width=980, height=45, x=10, y=548)
        
        tk.Label(LF3, text="M.A.J. BDD - Base de données pour le suivi budgétaire :", anchor="w").place(width=400, height=30, x=20, y=5)
        
        def load_update_save_BDD():
            
            path = self.path_to_aggregated_files.get("1.0", tk.END)
            if path[-1]=="\n": path=path[:-1]
            if path[-1]=="/": path=path[:-1]
            path_to_folder = path + '/Liens/'
            
            # LOADING DATA
            try:
                vision_personne_prev, vision_contrat_prev, utils_personne_prev, utils_contrat_prev = load_BDD(path_to_folder)
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la lecture de l'ancienne BDD.\nPath : {path_to_folder}\nAbandon.\n{e}")
                return
                
            try:
                aggregated_data, status_information, budgets_dates, log = load_current_data(path_to_folder, add_old_aggregated_data=False)
                if len(log)>0:
                    if not tk.messagebox.askokcancel("Lecture des données agrégées", f"{log}\nSouhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                        return
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la lecture des données agrégées.\nPath : {path_to_folder}\nAbandon.\n{e}")
                return
            
            # CHECKING DATA
            try:
                vision_personne_prev, utils_personne_prev, aggregated_data, status_information, log = check_loaded_coherence(vision_personne_prev, utils_personne_prev, aggregated_data, status_information)
                if len(log)>0:
                    if not tk.messagebox.askokcancel("Vérification des données chargées", f"{log}\nSouhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                        return
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la vérification des données chargées.\nAbandon.\n{e}")
                return
            
            # BUILDING NEW BDD
            try:
                utils_personne_new, log = update_utils_personne(status_information, utils_personne_prev)
                if len(log)>0:
                    if not tk.messagebox.askokcancel("Mise à jour de 'utils_personne'", f"{log}\nSouhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                        return
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour de 'utils_personne'.\nAbandon.\n{e}")
                return
            
            try:
                utils_contrat_new, log = update_utils_contrat(budgets_dates, utils_contrat_prev)
                if len(log)>0:
                    if not tk.messagebox.askokcancel("Mise à jour de 'utils_contrat'", f"{log}\nSouhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                        return
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour de 'utils_contrat'.\nAbandon.\n{e}")
                return
            
            try:
                vision_personne_new, vision_contrat_new, log = create_vision_personne_and_contrat(aggregated_data, utils_personne_new, utils_contrat_new)
                if len(log)>0:
                    print(log)
                    if not tk.messagebox.askokcancel("Mise à jour des visions personne et contrat", f"{log}\nSouhaitez vous poursuivre ? 'Annuler' pour revenir au menu principal."):
                        return
            except Exception as e:
                tk.messagebox.showerror("!", f"ERREUR : Problème lors de la mise à jour des visions personne et contrat.\nAbandon.\n{e}")
                return
            
            if tk.messagebox.askokcancel("Enregistrement", "Souhaitez-vous enregistrer la version mise à jour de BDD ? Peut prendre plusieurs secondes, un message s'affichera à la fin de l'enregistrement.\nUne version archivée sera également crée."):
                try:
                    update_BDD(path_to_folder, vision_personne_new, vision_contrat_new, utils_personne_new, utils_contrat_new)
                    tk.messagebox.showinfo("Enregistrement", "Fichier enregistré et archivé avec succès !")
                except Exception as e:
                    tk.messagebox.showerror("!", f"ERREUR : Problème lors l'enregistrement de BDD.\nPath : {path_to_folder}\nAbandon.\n{e}")
                    
            
        tk.Button(LF3, text="Chargement et Mise à jour BDD", bg="white", command=load_update_save_BDD).place(width=300, height=25, x=350, y=7)
        
if __name__ == "__main__":
    app = suivi_projet_GUI() 
    app.mainloop()