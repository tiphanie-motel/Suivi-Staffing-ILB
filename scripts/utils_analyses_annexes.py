import matplotlib.pyplot as plt
import numpy as np
import pandas as pd



def get_details_about_someone(vision_personne, prenom_nom, year, month):
    
    mask_date = (vision_personne["date"].dt.year == year) & (vision_personne["date"].dt.month == month)
    mask_nom = vision_personne["prenom_nom"]==prenom_nom

    mask = mask_date & mask_nom
    
    details = vision_personne[mask].groupby("root_name").h.sum().to_frame()
    details['%'] = (100*details.h/vision_personne[mask].h.sum()).round(1)

    return details



def checking_the_good_labsiens(vision_personne, k_months_ago=6):
    
    names_per_date = vision_personne.groupby("date")["prenom_nom"].unique().to_dict()
    sorted_dates = np.sort(list(names_per_date.keys()))
    sorted_names = np.sort(vision_personne["prenom_nom"].unique())

    rows = []
    for date in sorted_dates:
        names = names_per_date[date]
        rows.append(np.isin(sorted_names, names).astype(int))
        
    reported_weeks = pd.DataFrame(rows, columns = sorted_names, index = sorted_dates)
    
    last_date = reported_weeks.index.max()
    k_months_ago = last_date - pd.DateOffset(months=k_months_ago)
    reported_weeks = reported_weeks.loc[k_months_ago:]
    
    empty_cols = reported_weeks.sum()
    kept_individuals = empty_cols[empty_cols!=0].index
    removed_individuals = empty_cols[empty_cols==0].index
    print(f"Removed: {', '.join(list(removed_individuals))}")
    reported_weeks = reported_weeks[kept_individuals]
    
    semaines = reported_weeks.index
    individus = reported_weeks.columns
    condition = reported_weeks.values.T

    # Initialiser la figure
    fig, ax = plt.subplots(figsize=(10, 0.2*len(individus)))

    # Pour chaque individu, dessiner une ligne épaisse
    for i, individu in enumerate(individus):
        # Tracer les carrés quand la condition est remplie
        for j, semaine in enumerate(semaines):
            if condition[i, j]:
                ax.add_patch(plt.Rectangle((j, i - 0.4), 1, 0.8, color="green"))

        # Tracer la ligne continue pour l'individu
        ax.plot(np.where(condition[i, :])[0], [i] * np.sum(condition[i, :]), lw=1, color='black')

    # Configurer les axes
    ax.set_xticks(np.arange(len(semaines)))
    ax.set_xticklabels([sem.strftime('%d-%m') for sem in semaines], rotation=45)
    ax.set_yticks(np.arange(len(individus)))
    ax.set_yticklabels(individus)

    # Ajouter des labels et une légende
    ax.set_xlabel("Semaines")
    ax.set_ylabel("Individus")
    ax.set_title("Semaines renseignées sur les 6 derniers mois.")

    # Affichage
    plt.tight_layout()
    plt.show()