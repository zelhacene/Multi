import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.title("Consommation")

# Liste des mois pour charger les fichiers correspondants
months = ["January", "February", "March", "April", "May", "June", 
          "July", "August", "September", "October", "November", "December"]

dataframes = []

# Chargement des données pour chaque mois
for month in months:
    file_path = f"A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock {month} 2024.xlsx"
    
    df = pd.read_excel(file_path, sheet_name="situation")
    df[month] = df['Transfer'] - df['Ret_Production']
    df = df[['Description', month]].dropna(subset=['Description'])
    df = df.query('Description != "TOTAL DU MOIS"')
    
    dataframes.append(df)

# Fusion de tous les DataFrames en un seul
df_final = dataframes[0]

for df in dataframes[1:]:
    if not df.empty:  # Vérifie que le DataFrame n'est pas vide
        df_final = pd.merge(df_final, df, on='Description', how='right')

# Affichage du DataFrame final
st.dataframe(df_final)

# Liste des désignations à sélectionner
designations = ['Sucre Cristallisé', 'Emulsion Tropical Cosmos', 'Emulsion Tropical Givaudan',
                'Acide Citrique', 'CMC', 'Carton 48x125ml Cocktail Kiddy',
                'Carton 30x200ml Cocktail Yumy', 'Cartouche encre I-TECH', 'Cartouche make-up I-TECH']

# Interface utilisateur Streamlit
st.title('Sélection d\'articles et affichage correspondant')

# Cases à cocher pour la sélection des articles
selected_designations = st.multiselect('Sélectionnez les articles', designations)

# Vérifier si des désignations sont sélectionnées
if selected_designations:
    # Filtrer le DataFrame en fonction des articles sélectionnés
    df_selected = df_final[df_final['Description'].isin(selected_designations)]

    # Afficher le graphique si des articles sont sélectionnés
    if not df_selected.empty:
        fig, ax = plt.subplots()
        for _, row in df_selected.iterrows():
            ax.plot(df_selected.columns[1:], row[1:], label=row['Description'], marker='o')
            # Ajouter les valeurs au-dessus des points
            for i, value in enumerate(row[1:]):
                plt.annotate(f'{value:.2f}', (df_selected.columns[1:][i], value),
                             textcoords="offset points", xytext=(0, 5), ha='center')

        ax.set_xlabel('Mois')
        ax.set_ylabel('Consommation')
        ax.legend()
        st.pyplot(fig)
    else:
        st.warning('Aucun article sélectionné. Veuillez choisir au moins un article.')
else:
    st.warning('Veuillez sélectionner au moins un article.')
