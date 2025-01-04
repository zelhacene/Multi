import pandas as pd
import streamlit as st
from datetime import datetime

# Liste des mois
months_available = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

# Obtenir le mois en lettre
current_month =datetime.now().strftime('%B')

# Selectionner le mois courant
mois_selectionne = st.selectbox('Sélectionner le(s) mois', months_available, index=months_available.index(current_month))
aujourd_hui = datetime.now()

if mois_selectionne:
    chemin_fichier_excel = f"A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock {mois_selectionne} {aujourd_hui.year}.xlsx"
    ds_3=pd.read_excel(chemin_fichier_excel,sheet_name="situation")

ds_3=ds_3[['Description','St_Deb_Mois','Ret_Production','Transfer','St_Final']]
ds_3['St_Moyen']=(ds_3['St_Deb_Mois']+ds_3['St_Final'])/2
ds_3['Consommation']=ds_3['Transfer']-ds_3['Ret_Production']
ds_3['Rotation_Stock']=30*ds_3['St_Moyen']/ds_3['Consommation']
ds_3=ds_3[['Description','Rotation_Stock']]
ds_3=ds_3.dropna(subset=['Description'])
ds_3= ds_3.query('Description != "TOTAL DU MOIS"')
ds_3= ds_3.query('Description != "Jus Coktail Yummy 200 ml"')
ds_3= ds_3.query('Description != "Jus coktail Kiddy 125 ml"')

st.table(ds_3)