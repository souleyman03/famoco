import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import tempfile

st.title("📈 SUIVI CHALLENGE NFC")

VTO_URL = "https://docs.google.com/spreadsheets/d/165bFP7MjYjaIUHTdV1xo4E1PHa4EQRA8i_VJv3Z5rJ4/export?format=csv&gid=1269838156"
vto_df = pd.read_csv(VTO_URL, sep=",", on_bad_lines="skip")

uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (Données Challenge NFC)", type=["xlsx", "csv"])
    
if uploaded_file: 
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
        else:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("🗂️ Choisir la feuille à exploiter :", options=sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        # ✅ Charger logins depuis fichier VTO
        
        logins_concernes = vto_df["LOGIN"].astype(str).tolist()

        

        # ✅ Nettoyage des colonnes
        df = df.rename(columns={
            
            'ACCUEIL': 'PVT',
            
            'AGENCE': 'DR'
        })

        
        df['DR'] = df['DR'].astype(str).str.strip().str.upper()
        df['NOM'] = df['NOM'].astype(str).str.strip().str.upper()
        df['PRENOM'] = df['PRENOM'].astype(str).str.strip().str.upper()

        # 🔍 Filtrage
        df_filtre = df[df['LOGIN'].isin(logins_concernes) ]

        st.success("✅ Fichier filtré avec succès !")
        st.write("📊 Ventes via NFC :", df_filtre.shape[0], "lignes")
        st.dataframe(df_filtre)

        # 📊 Résumé par VTO
        # Regrouper par LOGIN, NUMERO, AGENCE, ACCUEIL et sommer les opérations
        df_summary = df_filtre.groupby(
            ["LOGIN", "DR", "PVT"], as_index=False
        )[["OPERATION NFC", "OPERATION MANUELLE", "TOTAL OPERATION"]].sum()

        # Ajouter la colonne TAUX DE NFC VTO
        df_summary["TAUX DE NFC VTO"] = (
            df_summary["OPERATION NFC"] / df_summary["TOTAL OPERATION"]
        ).fillna(0).apply(lambda x: "{:.0%}".format(x))
        

        
        # 📊 Résumé par PVT
        # Regrouper par PVT et sommer les opérations
        df_summary1 = df.groupby(
            ["PVT"], as_index=False
        )[["OPERATION NFC", "OPERATION MANUELLE", "TOTAL OPERATION"]].sum()

        # Ajouter la colonne TAUX DE NFC PVT
        df_summary1["TAUX DE NFC PVT"] = (
            df_summary1["OPERATION NFC"] / df_summary1["TOTAL OPERATION"]
        ).fillna(0).apply(lambda x: "{:.0%}".format(x))
        
        # 📊 Résumé par DR
        # Regrouper par DR et sommer les opérations
        df_summary2 = df.groupby(
            ["DR"], as_index=False
        )[["OPERATION NFC", "OPERATION MANUELLE", "TOTAL OPERATION"]].sum()

        # Ajouter la colonne TAUX DE NFC DR
        df_summary2["TAUX DE NFC DR"] = (
            df_summary2["OPERATION NFC"] / df_summary2["TOTAL OPERATION"]
        ).fillna(0).apply(lambda x: "{:.0%}".format(x))




        # 🧾 Export Excel
        temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='TAUX NFC VTO', index=False)
            df_summary1.to_excel(writer, sheet_name='TAUX NFC PVT', index=False)
            df_summary2.to_excel(writer, sheet_name='TAUX NFC DR', index=False)
        wb = load_workbook(temp_file.name)
        wb.save(temp_file.name)

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.success("✅ Fichier généré avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=final_buffer,
            file_name="CHALLENGE NFC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
