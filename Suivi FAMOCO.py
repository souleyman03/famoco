import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile

# Titre de l'application
st.title("📦 Générateur de Daily Reporting Distribution Famoco ")

# === Feuille 1 : base ===
BASE_URL = "https://docs.google.com/spreadsheets/d/165bFP7MjYjaIUHTdV1xo4E1PHa4EQRA8i_VJv3Z5rJ4/export?format=csv&gid=891360123"
base = pd.read_csv(BASE_URL, sep=",", on_bad_lines="skip")

# === Feuille 2 : correspondance ===
CORRESPONDANCE_URL = "https://docs.google.com/spreadsheets/d/165bFP7MjYjaIUHTdV1xo4E1PHa4EQRA8i_VJv3Z5rJ4/export?format=csv&gid=0"  
correspondance_df = pd.read_csv(CORRESPONDANCE_URL, sep=",", on_bad_lines="skip")


# Uploader du fichier Excel brut
uploaded_file = st.file_uploader("🗂️ Choisir la feuille des distributions Famoco)", type=["xlsx", "csv"])


if uploaded_file : 

    if uploaded_file.name.endswith('.csv'):
        df_dist = pd.read_csv(uploaded_file, encoding='utf-8', sep=';', header=None, names=["Code", "Type", "Total"])


    else:
        # Charger toutes les feuilles sans les lire entièrement
        xls = pd.ExcelFile(uploaded_file)
            

        # Afficher les noms de feuilles disponibles
        sheet_names = xls.sheet_names
        

        # Étape 1 : choix des 2 feuilles AVANT de lire
        selected_sheet1 = st.selectbox("🗂️ Choisir la feuille des distributions Famoco :", options=sheet_names, key="sheet1")
        selected_sheet2 = st.selectbox("🟢 Choisir la feuille des Famoco actifs :", options=sheet_names, key="sheet2")
        

        if selected_sheet1 and selected_sheet2 and selected_sheet1 != selected_sheet2:
            # Étape 2 : maintenant on lit les feuilles sélectionnées
            df_dist = pd.read_excel(uploaded_file, sheet_name=selected_sheet1, header=None, names=["Code", "Type", "Total"])
            df_dist1 = pd.read_excel(uploaded_file, sheet_name=selected_sheet2)
            df_dist1 = df_dist1.rename(columns={
            "Étiquettes de lignes": "FLOTTE",
            "Nombre de Custom Identifier": "NBRE DE FAMOCO ACTIF"
            })

            if uploaded_file.name.endswith('.csv') or selected_sheet1:
                # === Extraire la structure DR (DR1, DR2...) ===
                def extraire_structure(code):
                    code = str(code)
                    if code.startswith("DRV1_"):
                        return "DR1"
                    elif code.startswith("DRV2_"):
                        return "DR2"
                    elif code.startswith("DRVC_"):
                        return "DRC"
                    elif code.startswith("DRVE_"):
                        return "DRE"
                    elif code.startswith("DRVN_"):
                        return "DRN"
                    elif code.startswith("DRVS_"):
                        return "DRS"
                    elif code.startswith("DRVSE_"):
                        return "DRSE"
                    else:
                        return None
                
            
            df_dist["Structure"] = df_dist["Code"].apply(extraire_structure)
            distrib_par_structure = df_dist.dropna(subset=["Structure"])
            total_distribue = distrib_par_structure.groupby("Structure")["Total"].sum().reset_index()
            total_distribue.columns = ["Structure", "Total Distribue"]

            base["Structure"] = base["Structure"].replace({ 
            "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
            "DV-DRV1_DIRECTION REGIONALE DES VENTES DAKAR 1": "DR1",
            "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
            "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
            "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
            "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
            "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE"
                })
            
            # === Fusion et calcul du taux ===
            final_df = pd.merge(base, total_distribue, on="Structure", how="left")
            final_df["Total Distribue"] = final_df["Total Distribue"].fillna(0).astype(int)
            final_df["Taux de distribution (%)"] = (
                (final_df["Total Distribue"] / final_df["Nombre de Famoco"]) * 100).round().astype("Int64")
            
            # Supprimer les lignes où Structure = 'Total général' 
            base = base[base["Structure"] != "Total général"]
            
            # Garder uniquement les DR et Total général
            dr_liste = ["DR1", "DR2", "DRC", "DRE", "DRN", "DRS", "DRSE"]
            final_df = final_df[final_df["Structure"].isin(dr_liste)].copy()

            # === Total général ===
            total_row = pd.DataFrame({
                    "Structure": ["Total général"],
                    "Nombre de Famoco": [final_df["Nombre de Famoco"].sum()],
                    "Total Distribue": [final_df["Total Distribue"].sum()],
                    "Taux de distribution (%)": [f"{round((final_df["Total Distribue"].sum() / final_df["Nombre de Famoco"].sum()) * 100)}%"]
                })
            
            final_df = pd.concat([final_df, total_row], ignore_index=True)

            


            #-----------------------------------------------------------#
            df_ravt = df_dist[df_dist["Type"] == "RAVT"].copy()
            df_ravt = df_ravt.merge(correspondance_df[["Etiquettes", "Correspondance"]],
                                    left_on="Code", right_on="Etiquettes", how="left")
  
            ravt_distrib = df_ravt.groupby("Correspondance")["Total"].sum().reset_index()
            ravt_distrib.columns = ["Structure", "Total Distribué"]

            base_ravt = base[base["Structure"].str.startswith("RAVT")].copy()
            ravt_df = pd.merge(base_ravt, ravt_distrib, on="Structure", how="left")
            ravt_df["Total Distribué"] = ravt_df["Total Distribué"].fillna(0).astype(int)
            ravt_df["Taux de distribution (%)"] = ravt_df.apply(
                lambda row: f"{round((row['Total Distribué'] / row['Nombre de Famoco']) * 100)}%" if row["Nombre de Famoco"] > 0 else "0%",
                axis=1
            )

            # Ligne Total général RAVT
            total_line_ravt = pd.DataFrame({
                "Structure": ["Total général"],
                "Nombre de Famoco": [ravt_df["Nombre de Famoco"].sum()],
                "Total Distribué": [ravt_df["Total Distribué"].sum()],
                "Taux de distribution (%)": [f"{round((ravt_df['Total Distribué'].sum() / ravt_df['Nombre de Famoco'].sum()) * 100)}%"]
            })

            final_ravt_df = pd.concat([ravt_df, total_line_ravt], ignore_index=True)

            #---------------------------------------------------------------#
            df_rz = df_dist[df_dist["Type"] == "PDV - RZ"].copy()
            df_rz = df_rz.merge(correspondance_df[["Etiquettes", "Correspondance"]],
                                    left_on="Code", right_on="Etiquettes", how="left")

            rz_distrib = df_rz.groupby("Correspondance")["Total"].sum().reset_index()
            rz_distrib.columns = ["Structure", "Total Distribué"]

            base_rz = base[base["Structure"].str.startswith("RZ")].copy()
            rz_df = pd.merge(base_rz, rz_distrib, on="Structure", how="left")
            rz_df["Total Distribué"] = rz_df["Total Distribué"].fillna(0).astype(int)
            rz_df["Taux de distribution (%)"] = rz_df.apply(
                lambda row: f"{round((row['Total Distribué'] / row['Nombre de Famoco']) * 100)}%" if row["Nombre de Famoco"] > 0 else "0%",
                axis=1
            )

            # Ligne Total général RAVT
            total_line_rz = pd.DataFrame({
                "Structure": ["Total général"],
                "Nombre de Famoco": [rz_df["Nombre de Famoco"].sum()],
                "Total Distribué": [rz_df["Total Distribué"].sum()],
                "Taux de distribution (%)": [f"{round((rz_df['Total Distribué'].sum() / rz_df['Nombre de Famoco'].sum()) * 100)}%"]
            })

            final_rz_df = pd.concat([rz_df, total_line_rz], ignore_index=True)


            #-----------------TAUX FAMOCO ACTIF-----------------------------------#
            # Renommer les colonnes si besoin
            

            # Merge avec correspondance pour obtenir noms propres
            df = df_dist.merge(correspondance_df, left_on="Code", right_on="Etiquettes", how="left")
            df = df.rename(columns={
                "Code": "FLOTTE",
                "Total": "NBRE DE FAMOCO LIVRE",
                "Correspondance": "LIBELLE"
            })

            df["DR"] = df["FLOTTE"].apply(extraire_structure)

            # Joindre avec les actifs
            df = df.merge(df_dist1, on="FLOTTE", how="left")
            df["NBRE DE FAMOCO ACTIF"] = pd.to_numeric(df["NBRE DE FAMOCO ACTIF"], errors="coerce").fillna(0).astype(int)


            #Calcul du taux
            df["TAUX D'ACTIF"] = (df["NBRE DE FAMOCO ACTIF"] / df["NBRE DE FAMOCO LIVRE"]) * 100
            df["TAUX D'ACTIF"] = df["TAUX D'ACTIF"].fillna(0).round(0).astype(int).astype(str) + "%"


            # Final columns
            final_df1 = df[["DR", "LIBELLE", "FLOTTE", "NBRE DE FAMOCO LIVRE", "NBRE DE FAMOCO ACTIF", "TAUX D'ACTIF"]]
            final_df1 = final_df1.sort_values(by=["DR", "LIBELLE"])

            # Export Excel
            buffer_paiement = BytesIO()
            with pd.ExcelWriter(buffer_paiement, engine='openpyxl') as writer:
                    final_df1.to_excel(writer, sheet_name='TAUX FAMOCO ACTIFS', index=False)
                    final_df.to_excel(writer, sheet_name='SUIVI PAR DR', index=False)
                    final_ravt_df.to_excel(writer, sheet_name='SUIVI PAR RAVT ', index=False)
                    final_rz_df.to_excel(writer, sheet_name='SUIVI PAR RZ', index=False)
                    #df_filtre[cols_affichage].to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
                    #df_par_pvt.to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
            buffer_paiement.seek(0)

            st.download_button(
                        label="📥 Télécharger le fichier Daily reporting",
                        data=buffer_paiement,
                        file_name="Daily Reporting FAMOCO.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            

            