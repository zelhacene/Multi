import pandas as pd
import streamlit as st
from datetime import datetime
import io



st.title("Planning des achats")

df_1=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\Streamlit_app\\Planning_achat.xlsx")

df_ja=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock January 2024.xlsx",sheet_name="situation")
df_ja['Rea_January']=df_ja['Reception']-df_ja['Ret_Fournisseur']
df_ja=df_ja[['Description', 'Rea_January']]
df_ja=df_ja.dropna(subset=['Description'])
df_ja= df_ja.query('Description != "TOTAL DU MOIS"')

df_fe=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock February 2024.xlsx",sheet_name="situation")
df_fe['Rea_February']=df_fe['Reception']-df_fe['Ret_Fournisseur']
df_fe=df_fe[['Description', 'Rea_February']]
df_fe=df_fe.dropna(subset=['Description'])
df_fe= df_fe.query('Description != "TOTAL DU MOIS"')

df_ma=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock March 2024.xlsx",sheet_name="situation")
df_ma['Rea_March']=df_ma['Reception']-df_ma['Ret_Fournisseur']
df_ma=df_ma[['Description', 'Rea_March']]
df_ma=df_ma.dropna(subset=['Description'])
df_ma= df_ma.query('Description != "TOTAL DU MOIS"')

df_ap=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock April 2024.xlsx",sheet_name="situation")
df_ap['Rea_April']=df_ap['Reception']-df_ap['Ret_Fournisseur']
df_ap=df_ap[['Description', 'Rea_April']]
df_ap=df_ap.dropna(subset=['Description'])
df_ap= df_ap.query('Description != "TOTAL DU MOIS"')

df_mi=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock May 2024.xlsx",sheet_name="situation")
df_mi['Rea_May']=df_mi['Reception']-df_mi['Ret_Fournisseur']
df_mi=df_mi[['Description', 'Rea_May']]
df_mi=df_mi.dropna(subset=['Description'])
df_mi= df_mi.query('Description != "TOTAL DU MOIS"')

df_ju=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock June 2024.xlsx",sheet_name="situation")
df_ju['Rea_June']=df_ju['Reception']-df_ju['Ret_Fournisseur']
df_ju=df_ju[['Description', 'Rea_June']]
df_ju=df_ju.dropna(subset=['Description'])
df_ju= df_ju.query('Description != "TOTAL DU MOIS"')

df_jl=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock July 2024.xlsx",sheet_name="situation")
df_jl['Rea_July']=df_jl['Reception']-df_jl['Ret_Fournisseur']
df_jl=df_jl[['Description', 'Rea_July']]
df_jl=df_jl.dropna(subset=['Description'])
df_jl= df_jl.query('Description != "TOTAL DU MOIS"')

df_ag=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock August 2024.xlsx",sheet_name="situation")
df_ag['Rea_August']=df_ag['Reception']-df_ag['Ret_Fournisseur']
df_ag=df_ag[['Description', 'Rea_August']]
df_ag=df_ag.dropna(subset=['Description'])
df_ag= df_ag.query('Description != "TOTAL DU MOIS"')

df_sp=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock September 2024.xlsx",sheet_name="situation")
df_sp['Rea_September']=df_sp['Reception']-df_sp['Ret_Fournisseur']
df_sp=df_sp[['Description', 'Rea_September']]
df_sp=df_sp.dropna(subset=['Description'])
df_sp= df_sp.query('Description != "TOTAL DU MOIS"')

df_oc=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock October 2024.xlsx",sheet_name="situation")
df_oc['Rea_October']=df_oc['Reception']-df_oc['Ret_Fournisseur']
df_oc=df_oc[['Description', 'Rea_October']]
df_oc=df_oc.dropna(subset=['Description'])
df_oc= df_oc.query('Description != "TOTAL DU MOIS"')

df_nv=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock November 2024.xlsx",sheet_name="situation")
df_nv['Rea_November']=df_nv['Reception']-df_nv['Ret_Fournisseur']
df_nv=df_nv[['Description', 'Rea_November']]
df_nv=df_nv.dropna(subset=['Description'])
df_nv= df_nv.query('Description != "TOTAL DU MOIS"')

df_dc=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock December 2024.xlsx",sheet_name="situation")
df_dc['Rea_December']=df_dc['Reception']-df_dc['Ret_Fournisseur']
df_dc=df_dc[['Description', 'Rea_December']]
df_dc=df_dc.dropna(subset=['Description'])
df_dc= df_dc.query('Description != "TOTAL DU MOIS"')

dataframes = [df_fe, df_ma, df_ap, df_mi, df_ju, df_jl, df_ag, df_sp, df_oc, df_nv, df_dc]
df_2 = df_ja

for df in dataframes:
    df_2 = pd.merge(df_2, df, on='Description', how='right')

df = pd.merge(df_1, df_2, on='Description', how='left')

# Calcul des ratios
df['Ratio_January'] = df['Rea_January'] / df['Prévision January']
df['Ratio_February'] = df['Rea_February'] / df['Prévision February']
df['Ratio_March'] = df['Rea_March'] / df['Prévision March']
df['Ratio_April'] = df['Rea_April'] / df['Prévision April']
df['Ratio_May'] = df['Rea_May'] / df['Prévision May']
df['Ratio_June'] = df['Rea_June'] / df['Prévision June']
df['Ratio_July'] = df['Rea_July'] / df['Prévision July']
df['Ratio_August'] = df['Rea_August'] / df['Prévision August']
df['Ratio_September'] = df['Rea_September'] / df['Prévision September']
df['Ratio_October'] = df['Rea_October'] / df['Prévision October']
df['Ratio_November'] = df['Rea_November'] / df['Prévision November']
df['Ratio_December'] = df['Rea_December'] / df['Prévision December']

# Sélection des colonnes
df = df[['Description', 'Prévision January', 'Rea_January', 'Ratio_January',
    'Prévision February', 'Rea_February', 'Ratio_February',
    'Prévision March', 'Rea_March', 'Ratio_March',
    'Prévision April', 'Rea_April', 'Ratio_April',
    'Prévision May', 'Rea_May', 'Ratio_May',
    'Prévision June', 'Rea_June', 'Ratio_June',
    'Prévision July', 'Rea_July', 'Ratio_July',
    'Prévision August', 'Rea_August', 'Ratio_August',
    'Prévision September', 'Rea_September', 'Ratio_September',
    'Prévision October', 'Rea_October', 'Ratio_October',
    'Prévision November', 'Rea_November', 'Ratio_November',
    'Prévision December', 'Rea_December', 'Ratio_December']]



# Creer boutton pour extraire le fichier Excel
if st.button('Extraire et sauvegarder les données'):

    buffer = io.BytesIO()
    # Enregistrer le DataFrame dans un fichier Excel
    df.to_excel(buffer, index=False, engine='xlsxwriter')
    buffer.seek(0)

    # Boutton pour telecharger le fichier
    st.download_button(
        label="Télécharger le fichier Excel",
        data=buffer,
        file_name="Actuals_and_forecasts.xlsx",
        key="download_button",
    )

def Actuals_and_forcats():
    # Liste des mois
    months_available = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

    # Selectionner le mois en lettre
    current_month = datetime.now().strftime('%B')

    # Selectionner le mois courrant
    mois_selectionne = st.selectbox('Sélectionner le(s) mois', months_available, index=months_available.index(current_month))
    aujourd_hui = datetime.now()

    ds_1=pd.read_excel("A:\\ACHAT\\ACHAT 2024\\Streamlit_app\\Planning_achat.xlsx")
    ds_1=ds_1[['Description',f'Prévision {mois_selectionne}']]
    ds_1 = ds_1.dropna(subset=[f'Prévision {mois_selectionne}'])

    if mois_selectionne:
        chemin_fichier_excel = f"A:\\ACHAT\\ACHAT 2024\\SITUATION STOCK\\Suivi situation stock {mois_selectionne} {aujourd_hui.year}.xlsx"
        ds_2=pd.read_excel(chemin_fichier_excel,sheet_name="situation")

        ds_2[f'Rea_{mois_selectionne}']=ds_2['Reception']-ds_2['Ret_Fournisseur']
        ds_2=ds_2[['Description',f'Rea_{mois_selectionne}']]
        ds_2=ds_2.dropna(subset=['Description'])
        ds_2 = ds_2.loc[ds_2[f'Rea_{mois_selectionne}'] != 0]
        ds_2= ds_2.query('Description != "TOTAL DU MOIS"')

    ds = pd.merge(ds_1, ds_2, on='Description', how='left')

    # Calcul des ratios
    ds[f'Ratio_{mois_selectionne}'] = ds[f'Rea_{mois_selectionne}'] / ds[f'Prévision {mois_selectionne}']
    ds[f'Ratio_{mois_selectionne}']=ds[f'Ratio_{mois_selectionne}']*100

    # Sélection des colonnes
    ds = ds[['Description', f'Prévision {mois_selectionne}', f'Rea_{mois_selectionne}', f'Ratio_{mois_selectionne}']]
    ds=ds.fillna(0)
    st.table(ds)

if __name__ == "__main__":
    Actuals_and_forcats()


