import streamlit as st
import pandas as pd
from io import BytesIO

# Fonction de traitement pour le fichier Excel
def process_excel(file):
    try:
        # Lire le fichier Excel sans header
        df = pd.read_excel(file, header=None)
        
        # Trouver l'index de la ligne d'en-tête
        header_index = next((idx for idx, row in df.iterrows() 
                             if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])), None)
        
        if header_index is None:
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables dans le fichier.")
            return None
        
        # Redéfinir les en-têtes et supprimer les lignes précédentes
        df.columns = df.iloc[header_index]
        df = df.iloc[header_index + 1:].reset_index(drop=True)
        
        # Vérification si le fichier est vide après nettoyage
        if df.empty:
            st.error("Aucune donnée valide après le traitement des lignes.")
            return None
        
        # Vérification des colonnes nécessaires (double vérification)
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            st.error(f"Colonnes manquantes après traitement : {', '.join(missing)}")
            return None
        
        # Aperçu des données
        st.write("Aperçu des données après traitement automatique :")
        st.write(df.head(10))
        
        # Nettoyage des données
        df = df.dropna(subset=['Nom', 'Prénom', 'Code'])
        df['Name'] = df['Code'].astype(str) + ' ' + df['Nom'] + ' ' + df['Prénom']
        df = df[['Code', 'Name']].drop_duplicates()
        return df
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None


# Fonction de traitement pour le fichier CSV
def process_csv(excel_file, csv_file):
    try:
        # Lire le fichier Excel sans header
        df_xls = pd.read_excel(excel_file, header=None)
        df_csv = pd.read_csv(csv_file, delimiter=';', encoding='utf-8')
        
        # Trouver l'index de la ligne d'en-tête
        header_index = next((idx for idx, row in df_xls.iterrows() 
                             if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])), None)
        
        if header_index is None:
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables dans le fichier.")
            return None, None
        
        # Redéfinir les en-têtes et supprimer les lignes précédentes
        df_xls.columns = df_xls.iloc[header_index]
        dm_xls = df_xls.iloc[:header_index]
        df_xls = df_xls.iloc[header_index + 1:].reset_index(drop=True)
        
        
        # Vérification si le fichier est vide après nettoyage
        if df_xls.empty:
            st.error("Aucune donnée valide après le traitement des lignes.")
            return None, None
        
        # Vérification des colonnes nécessaires (double vérification)
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in df_xls.columns]
        if missing:
            st.error(f"Colonnes manquantes après traitement : {', '.join(missing)}")
            return None, None
        
        # Aperçu des données
        #st.write("Aperçu du fichier administraif nettoyé :")
        #st.write(df_xls.head(10))

            
        # Vérification des colonnes nécessaires
        required_columns = ['A:Code', 'Code', 'Note']
        missing_columns = [col for col in required_columns if col not in df_csv.columns]
        if missing_columns:
            st.error(f"Colonnes manquantes : {', '.join(missing_columns)}")
            return None, None
        
        # Initialiser un DataFrame pour les anomalies
        anomalies = pd.DataFrame()
        
        # 1. Filtrer les lignes avec 'NONE' ou valeurs vides
        df_csv = df_csv[['A:Code', 'Code', 'Nom', 'Note']]
        none_mask = (df_csv['A:Code'] == 'NONE') 
        anomalies = pd.concat([anomalies, df_csv[none_mask]])
        df_clean = df_csv[~none_mask].copy()
        
        # 2. Vérifier la cohérence entre 'A:Code' et 'Code' converti
        df_clean['A:Code'] = pd.to_numeric(df_clean['A:Code'], errors='coerce')
        mismatch_mask = df_clean['A:Code'] != df_clean['Code']
        anomalies = pd.concat([anomalies, df_clean[mismatch_mask]])
        df_clean = df_clean[~mismatch_mask]
        
        # Sauvegarder les anomalies dans un fichier Excel
        if not anomalies.empty:
            anomalies_file = BytesIO()
            anomalies.to_excel(anomalies_file, index=False)
            anomalies_file.seek(0)
        else:
            anomalies_file = None
        
        # Vérifier si le fichier nettoyé est vide
        if df_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, anomalies
        
        # Afficher un aperçu du fichier CSV après nettoyage
        st.write("Aperçu du fichier des notes nettoyé:")
        st.write(df_clean.head(10))
    

        df_merged = pd.merge(df_xls, df_csv[['Code', 'Note']], on='Code', how='left')
        df_merged.rename(columns={'Note_y': 'Note'}, inplace=True)
        df_merged.drop(columns=['Note_x'], inplace=True)
        df_final = pd.concat([dm_xls, df_merged], axis=0)

        # Enregistrement du DataFrame fusionné dans le fichier Excel
        #output_excel_file = 'fusion.xlsx'
        #df_merged.to_excel(output_excel_file, index=False)

        
        # Aperçu des données
        st.write("Aperçu du fichier des notes après fusion :")
        st.write(df_merged.head(10))
        
        return df_final, anomalies
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None
    
# Interface utilisateur
st.title("Traitements de fichiers Excel et CSV pour AMC")

# Onglets pour séparer les sections
tab1, tab2 = st.tabs(["Liste des étudiants", "Traitement des notes"])

with tab1:
    st.header("Préparation de la liste des étudiants")
    st.info("""
    - Téléchargez un fichier Excel contenant les colonnes 'Nom', 'Prénom' et 'Code'.
    - La détection des en-têtes est automatique.
    - Les lignes avant les en-têtes seront automatiquement supprimées.
    """)
    
    uploaded_excel_file = st.file_uploader(
        "Chargez votre fichier Excel", 
        type="xlsx", 
        key="excel_uploader"
    )
    
    if uploaded_excel_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            processed_data = process_excel(uploaded_excel_file)
            
            if processed_data is not None:
                st.success(f"Lecture du fichier Excel réussie ! {len(processed_data)} étudiants trouvés.")
                
                st.write('Voici un aperçu de la liste des étudiants:')
                st.write(processed_data.head(10))
                st.success(f"La liste contient {len(processed_data)} étudiants.")

                # Convertir le DataFrame en CSV
                csv = processed_data.to_csv(index=False)
                
                # Permettre le téléchargement du fichier CSV
                st.download_button(
                    label="Télécharger la liste des étudiants au format CSV",
                    data=csv,
                    file_name="liste.csv",
                    mime="text/csv"
                )

with tab2:
    st.header("Traitement des notes")
    st.info("""
    - Téléchargez un fichier CSV contenant les colonnes 'A:Code', 'Nom' et 'Note'.
    - Les anomalies seront exportées dans un fichier Excel séparé.
    - Les notes seront automatiquement associées aux étudiants de l'onglet 1.
    """)

    uploaded_excel_file2 = st.file_uploader(
        "Chargez votre fichier Excel", 
        type="xlsx", 
        key="excel_uploader2"
    )

    if uploaded_excel_file2 is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            processed_data = process_excel(uploaded_excel_file2)
            
            if processed_data is not None:
                st.success(f"Le fichier Excel de l'administration contient {len(processed_data)} étudiants.")


    uploaded_csv_file = st.file_uploader(
        "Charger le fichier CSV des étudiants", 
        type="csv", 
        key="csv_uploader"
    )

    if uploaded_excel_file2 is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            processed_data = process_excel(uploaded_excel_file2)
            
            if processed_data is not None:
                st.success(f"{len(processed_data)} étudiants ont passé l'examen.")
    
    if uploaded_csv_file is not None and uploaded_excel_file2 is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            final_data, anomalies2 = process_csv(uploaded_excel_file2, uploaded_csv_file)
            
            if final_data is not None and anomalies2 is not None:
                st.success("Fusion réussie !")
                
                # Afficher les statistiques
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Étudiants trouvés", len(final_data))
                with col2:
                    st.metric("Anomalies trouvées", len(anomalies2))
                
                # Téléchargement des résultats
                
                #st.download_button(
                #    label="📥 Télécharger les données finales",
                #    data=final_data,#.to_csv(index=False, sep=';'),
                #    file_name="etudiants_avec_notes.xlsx",
                #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"#"text/csv"
                #)
                
                
                #st.download_button(
                #    label="🚨 Télécharger les anomalies",
                #    data=anomalies2,
                #    file_name="anomalies.xlsx",
                #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                #)
                
                
                # Aperçu interactif
                
                #with st.expander("Aperçu des données fusionnées"):
                #    st.dataframe(final_data.style.highlight_null(color='#FF6666'))
                
