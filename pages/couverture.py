import pandas as pd
import streamlit as st
import calendar
from datetime import datetime


# Charger le DataFrame depuis le fichier Excel
df_cn = pd.read_excel(r"A:\ACHAT\ACHAT 2024\Streamlit_app\Formule.xlsx", sheet_name="Consommation")

# Filtrer les lignes avec des descriptions spécifiques
df_cn = df_cn[~df_cn['Description'].isin([
    "Vitesse (Boites\\Heures)",
    "Nombre de Boites produit par jours",
    "Nombre de Boites par carton",
    "Nombre de Carton par palette",
    "Nombre d intercalaires par palette",
    "Nombre de palettes produits par jour",
    "Contenance (Litre)",
    "Litre de sirop fini produit en 24"
])]

# Liste des lignes disponibles
lignes_disponibles = ['Ligne 1/ 200 ml', 'Ligne 2/ 125 ml', 'Ligne 3/ 125 ml', 'Ligne 4/ 125 ml']

# Interface utilisateur avec Streamlit
st.title("Calcul couverture des Matières Premières")
selected_lignes = st.multiselect("Sélectionnez les lignes en marche", lignes_disponibles)

# Filtrer le DataFrame en fonction des lignes sélectionnées
filtered_df = df_cn[['Description'] + selected_lignes].copy()

# Créer une nouvelle colonne pour calculer le total pour les lignes sélectionnées
filtered_df['Consommation'] = filtered_df[selected_lignes].sum(axis=1)
filtered_df['Consommation Mensuelle'] = filtered_df['Consommation'] * 30

# Arrondir la colonne 'Consommation' à deux décimales
filtered_df['Consommation'] = filtered_df['Consommation'].round(2)
filtered_df['Consommation Mensuelle'] = filtered_df['Consommation Mensuelle'].round(2)

aujourd_hui = datetime.now()
Mois = aujourd_hui.month
Year = aujourd_hui.year
mois_en_lettres = calendar.month_name[Mois]

# Construisez le chemin du fichier Excel
chemin_fichier_excel = fr"A:\ACHAT\ACHAT 2024\SITUATION STOCK\Suivi situation stock {mois_en_lettres} {Year}.xlsx"

# Charger le DataFrame depuis le fichier Excel
df_st = pd.read_excel(chemin_fichier_excel, sheet_name="situation")
df_st = df_st[['Description', 'St_Final']].dropna(subset=['Description'])
df_st = df_st[df_st['Description'] != "TOTAL DU MOIS"]

# Merge les deux tableaux
df_final = pd.merge(filtered_df, df_st, on='Description', how='left')

# Ajouter une nouvelle colonne 'Couverture/Jours' à df_final
df_final['Couverture/Jours'] = df_final['St_Final'] / df_final['Consommation']
df_final.fillna(0, inplace=True)

# Appliquer le format 123 333.00
df_final['St_Final'] = df_final['St_Final'].astype(int)
df_final['Consommation'] = df_final['Consommation'].astype(int)
df_final['Couverture/Jours'] = df_final['Couverture/Jours'].round(2)

# Afficher le tableau avec les colonnes "Description" et le total pour les lignes sélectionnées
st.table(df_final[['Description', 'St_Final', 'Consommation', 'Consommation Mensuelle', 'Couverture/Jours']])

# Afficher le graphique à barres
st.bar_chart(df_final.set_index('Description')['Couverture/Jours'])
