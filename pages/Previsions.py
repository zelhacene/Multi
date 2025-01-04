import pandas as pd
import streamlit as st
from datetime import datetime

# Charger le DataFrame depuis le fichier Excel
df_prev = pd.read_excel('A:\\ACHAT\\ACHAT 2024\\Streamlit_app\\Prévision.xlsx')
df_prev = df_prev.query('Description != "Litre de produit fini 125"')
df_prev = df_prev.query('Description != "Litre de produit fini 200"')
df_prev = df_prev.query('Description != "Total Litre de produit fini"')

# Liste des mois
months_available = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

# Obtenir le mois en lettre
current_month = datetime.now().strftime('%B')

# Selectionner le mois courrant
mois_selectionne = st.selectbox('Sélectionner le(s) mois', months_available, index=months_available.index(current_month))

if mois_selectionne:

    # Filter le DataFrame on foction le mois selectionner
    filtered_data = df_prev[['Description', mois_selectionne]].copy()

    # Charger le DataFrame depuis le fichier Excel
    aujourd_hui = datetime.now()
    chemin_fichier_excel = f"A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock {mois_selectionne} {aujourd_hui.year}.xlsx"
    df_reas = pd.read_excel(chemin_fichier_excel, sheet_name="situation")
    df_reas['consommation'] = df_reas['Transfer'] - df_reas['Ret_Production']
    df_reas = df_reas[['Description', 'consommation']]
    df_reas = df_reas.dropna(subset=['Description'])
    df_reas = df_reas.query('Description != "TOTAL DU MOIS"')

    df_fin = pd.merge(filtered_data, df_reas, on='Description', how='left')
    df_fin['ratio'] = (df_fin['consommation'] / df_fin[mois_selectionne]*100).round(0)
    st.table(df_fin)

else:
    st.warning("Aucun mois sélectionné. Veuillez sélectionner au moins un mois.")    