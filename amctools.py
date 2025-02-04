import streamlit as st
import pandas as pd
from io import BytesIO

# Fonction de traitement pour le fichier Excel
def process_excel(file):
    try:
        # Lire le fichier Excel sans header
        xls = pd.read_excel(file, header=None)
                
        # Trouver l'index de la ligne d'en-tête
        header_index = next((idx for idx, row in xls.iterrows() 
                             if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])), None)
        
        if header_index is None:
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables dans le fichier.")
            return None, None
        
        # Redéfinir les en-têtes et supprimer les lignes précédentes
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)
        
        # Vérification si le fichier est vide après nettoyage
        if xls.empty:
            st.error("Aucune donnée valide après le traitement des lignes.")
            return None, None
        
        # Vérification des colonnes nécessaires (double vérification)
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in xls.columns]
        if missing:
            st.error(f"Colonnes manquantes après traitement : {', '.join(missing)}")
            return None, None
        
        # Nettoyage des données
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        return xls, liste
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None


# Fonction de traitement pour le fichier CSV
def process_csv(excel_file, csv_file):
    try:
        xls, liste = process_excel(excel_file) 
        
        csv = pd.read_csv(csv_file, delimiter=';', encoding='utf-8')

        # Initialiser un DataFrame pour les anomalies
        anomalies = pd.DataFrame()

        # 1. Filtrer les lignes avec 'NONE' ou valeurs vides
        csv = csv[['A:Code', 'Code', 'Nom', 'Note']]
        none_mask = (csv['A:Code'] == 'NONE') 
        anomalies = pd.concat([anomalies, csv[none_mask]])
        csv_clean = csv[~none_mask].copy()
        
        # 2. Vérifier la cohérence entre 'A:Code' et 'Code' converti
        #df_clean['A:Code'] = pd.to_numeric(df_clean['A:Code'], errors='coerce')
        #mismatch_mask = df_clean['A:Code'] != df_clean['Code']
        #anomalies = pd.concat([anomalies, df_clean[mismatch_mask]])
        #df_clean = df_clean[~mismatch_mask]
        
        # Sauvegarder les anomalies dans un fichier Excel
        if not anomalies.empty:
            anomalies_file = BytesIO()
            anomalies.to_excel(anomalies_file, index=False)
            anomalies_file.seek(0)
        else:
            anomalies_file = None
        
        # Vérifier si le fichier nettoyé est vide
        if csv_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None, None
    
        df_merged = pd.merge(xls, csv_clean[['Code', 'Note']], on='Code', how='left')
        df_merged.rename(columns={'Note_y': 'Note'}, inplace=True)
        df_merged = df_merged.drop(columns=['Note_x'])
        return csv_clean, df_merged, anomalies
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None, None
    
# Interface utilisateur
st.title("Traitements de fichiers Excel et CSV pour AMC")

# Onglets pour séparer les sections
tab1, tab2 = st.tabs(["Liste des étudiants", "Traitement des notes"])

with tab1:
    st.header("Préparation de la liste des étudiants")
    st.info("""
    - Télécharger le fichier Excel de l'administration. 
    - Les en-têtes 'Nom', 'Prénom' et 'Code' seront détectés.
    - Les lignes avant les en-têtes seront automatiquement supprimées.
    - La liste des étudiants à fournir à AMC sera préparée au format CSV.
    """)
    
    uploaded_excel_file = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader"
    )
    
    if uploaded_excel_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file)
            
            if xls is not None:
                st.success(f"Lecture du fichier Excel réussie ! {len(xls)} étudiants trouvés.")  
                st.write("Aperçu de la base de données des étudiants avant traitement automatique :")
                st.write(xls.head(10))              
                st.write('Aperçu de la liste des étudiants à fournir à AMC:')
                st.write(liste.head(10))
                st.success(f"La liste contient {len(xls)} étudiants.")

                # Convertir le DataFrame en CSV
                csv = liste.to_csv(index=False)
                
                # Permettre le téléchargement du fichier CSV
                st.download_button(
                    label="📥 Télécharger la liste des étudiants au format CSV",
                    data=csv,
                    file_name="liste.csv",
                    mime="text/csv"
                )

with tab2:
    st.header("Traitement des notes")
    st.info("""
    - Télécharger le fichier Excel de l'administration. 
    - Télécharger le fichier CSV des notes calculées par AMC.
    - Les notes seront automatiquement associées aux étudiants.
    - Le nom nombre d'étudiants mal identifiés sera indiqué.
    """)

    uploaded_excel_file2 = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader2"
    )

    uploaded_csv_file = st.file_uploader(
        "Télécharger le fichier CSV des notes calculées par AMC", 
        type="csv", 
        key="csv_uploader"
    )

    if uploaded_excel_file2 is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file2)

    
    if uploaded_csv_file is not None and uploaded_excel_file2 is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            csv_clean, df_merged, anomalies = process_csv(uploaded_excel_file2, uploaded_csv_file)

            st.write("Aperçu de la base de données des étudiants :")
            st.write(xls.head(10))   
            
            #if final_data is not None and anomalies2 is not None:
            #st.success("Fusion réussie !")

            # Aperçu des données
            st.write("Aperçu du fichier des notes :")
            st.write(csv_clean.head(10))

            # Aperçu des données
            st.write("Aperçu de la base de données des étudiants alimentée par les notes :")
            st.write(df_merged.head(10))
                
            # Afficher les statistiques
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Effectif total", len(df_merged))
            with col2:
                st.metric("Présents", len(csv_clean))
            with col3:
                st.metric("Absents", len(df_merged)-len(csv_clean)-len(anomalies))
            with col4:
                st.metric("Mal identifiés", len(anomalies))
                
            # Téléchargement des résultats
             
            st.download_button(
                label="📥 Télécharger le fichier final des notes au format Excel",
                data=df_merged.to_csv(index=False, sep=';'),
                file_name="etudiants_avec_notes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"#"text/csv"
                )
                
            if len(anomalies) > 0:   
                st.error(f"Attention! {len(anomalies)} étudiants ont été mal identifiés. Vérifier leurs copies.")        
            #st.download_button(
            #    label="🚨 Télécharger les anomalies",
            #    data=df_merged,
            #    file_name="notes.xlsx",
            #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            #)
               
                
            # Aperçu interactif
                
            #with st.expander("Aperçu des données fusionnées"):
            #    st.dataframe(final_data.style.highlight_null(color='#FF6666'))
                
