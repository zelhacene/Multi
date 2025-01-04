import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.title("Tableau des indicateurs : ")

# Liste des mois
months_available = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

# Convertire le mois de la forme en chiffre vers la forme en lettre
current_month = datetime.now().strftime('%B')

# Selectinner le mois courant
mois_selectionne = st.selectbox('Sélectionner le(s) mois', months_available, index=months_available.index(current_month))
aujourd_hui = datetime.now()

#=======================================================================================================================================

import pandas as pd

# Définir le chemin de base
chemin_fichier_1 = r"A:\\ACHAT\\ACHAT 2024\\Streamlit_app\\Planning_achat.xlsx"
# Vérifier si le fichier existe et le lire
ds_1 = pd.read_excel(chemin_fichier_1,sheet_name='Sheet1')

# Filtrer les colonnes et enlever les lignes avec des valeurs manquantes
ds_1 = ds_1[['Description', f'Prévision {mois_selectionne}']]
ds_1 = ds_1.dropna(subset=[f'Prévision {mois_selectionne}'])


# Définir le nom du fichier avec f-string pour insérer les variables
fichier_recherche_2 = f"Suivi situation stock {mois_selectionne} {aujourd_hui.year}.xlsx"
# Définir le chemin de base
chemin_fichier_2 = r"A:\ACHAT\ACHAT 2024\SITUATION STOCK"

chemin_fichier_2 = os.path.join(chemin_fichier_2, fichier_recherche_2)

# Initialiser ds pour éviter les erreurs si jamais il n'est pas assigné
ds_2 = pd.DataFrame()
ds = pd.DataFrame()
if mois_selectionne:
    try:
        # Chargement du fichier Excel et vérification des colonnes existantes
        ds_2 = pd.read_excel(chemin_fichier_2, sheet_name="situation")

        if 'Reception' in ds_2.columns and 'Ret_Fournisseur' in ds_2.columns:
            # Calculer la nouvelle colonne basée sur le mois sélectionné
            ds_2[f'Rea_{mois_selectionne}'] = ds_2['Reception'] - ds_2['Ret_Fournisseur']

            # Filtrer et garder seulement les colonnes nécessaires
            ds_2 = ds_2[['Description', f'Rea_{mois_selectionne}']]

            # Supprimer les lignes où 'Description' est NaN
            ds_2 = ds_2.dropna(subset=['Description'])

            # Supprimer les lignes où la valeur calculée est 0
            ds_2 = ds_2.loc[ds_2[f'Rea_{mois_selectionne}'] != 0]

            # Supprimer la ligne où la description est "TOTAL DU MOIS"
            ds_2 = ds_2.loc[ds_2['Description'] != "TOTAL DU MOIS"]

        else:
            st.error("Les colonnes 'Reception' ou 'Ret_Fournisseur' sont manquantes dans le fichier Excel.")
    
    except FileNotFoundError:
        st.error("Le fichier Excel n'a pas été trouvé. Vérifie le chemin.")
    
    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")

# Vérifier si ds_2 contient des données avant la fusion
if not ds_2.empty:
    ds = pd.merge(ds_1, ds_2, on='Description', how='left')
    st.write("Fusion réussie :", ds)

    # Calcul des ratios après la fusion
    if f'Rea_{mois_selectionne}' in ds.columns and f'Prévision {mois_selectionne}' in ds.columns:
        ds[f'Ratio_{mois_selectionne}'] = ds[f'Rea_{mois_selectionne}'] / ds[f'Prévision {mois_selectionne}']
        ds[f'Ratio_{mois_selectionne}'] = ds[f'Ratio_{mois_selectionne}'] * 100
        st.write(f"Ratio calculé pour {mois_selectionne} :", ds[[f'Ratio_{mois_selectionne}']])

        # Vérification des colonnes avant de les sélectionner
        required_columns = ['Description', f'Prévision {mois_selectionne}', f'Rea_{mois_selectionne}', f'Ratio_{mois_selectionne}']
        if all(col in ds.columns for col in required_columns):
            ds = ds[required_columns]
            ds = ds.fillna(0)
            st.write(f"Données finales pour {mois_selectionne} :", ds)
        else:
            st.error("Une ou plusieurs colonnes nécessaires sont manquantes.")
    else:
        st.error(f"Les colonnes 'Réa_{mois_selectionne}' ou 'Prévision {mois_selectionne}' sont manquantes.")
else:
    st.error("ds_2 est vide ou n'a pas été chargé correctement, la fusion ne peut pas être effectuée.")

# Vérification si ds contient les données avant d'utiliser 'Mom'
if not ds.empty:
    # Calcul des valeurs 'Mom'
    Mom = np.array(ds[f'Ratio_{mois_selectionne}'])

    for i in range(len(Mom)):
        if Mom[i] >= 100:
            Mom[i] = 100

    taux_respect_planning_approvisionnement = np.nanmean(Mom)
    st.write(f"Taux de respect du planning d'approvisionnement : {taux_respect_planning_approvisionnement:.2f}%")
else:
    st.error("Le DataFrame 'ds' est vide, le calcul du taux de respect ne peut pas être effectué.")

#==================================================================================================================
