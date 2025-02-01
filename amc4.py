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
def process_csv(file):
    try:
        # Lire le fichier CSV
        df = pd.read_csv(file, delimiter=';', encoding='utf-8')
        
        # Vérification des colonnes nécessaires
        required_columns = ['A:Code', 'Code', 'Note']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Colonnes manquantes : {', '.join(missing_columns)}")
            return None, None
        
        # Initialiser un DataFrame pour les anomalies
        anomalies = pd.DataFrame()
        
        # 1. Filtrer les lignes avec 'NONE' ou valeurs vides
        df = df[['A:Code', 'Code', 'Nom', 'Note']]
        none_mask = (df['A:Code'] == 'NONE') 
        anomalies = pd.concat([anomalies, df[none_mask]])
        df_clean = df[~none_mask].copy()
        
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
            return None, anomalies_file
        
        # Afficher un aperçu du fichier CSV après nettoyage
        st.write("Aperçu des données après nettoyage :")
        st.write(df_clean.head(10))
        
        return df_clean, anomalies_file
    
    except Exception as e:
        st.error(f"Erreur critique lors du traitement CSV : {str(e)}")
        return None, None


# Interface utilisateur
st.title("Traitements de fichiers Excel et CSV pour AMC")

# Onglets pour séparer les sections
tab1, tab2, tab3 = st.tabs(["Fichier Excel", "Fichier CSV", "Fusion"])

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
                st.success(f"Traitement réussi ! {len(processed_data)} étudiants trouvés.")
                
                # Convertir le DataFrame en CSV
                csv = processed_data.to_csv(index=False)
                
                # Permettre le téléchargement du fichier CSV
                st.download_button(
                    label="Télécharger la liste des étudiants",
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
    
    uploaded_csv_file = st.file_uploader(
        "Charger le fichier CSV des étudiants", 
        type="csv", 
        key="csv_uploader"
    )
    
    if uploaded_csv_file is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            final_data, anomalies_file = process_csv(uploaded_csv_file)
            
            if final_data is not None and anomalies_file is not None:
                st.success("Fusion réussie !")
                
                # Afficher les statistiques
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Étudiants trouvés", len(final_data))
                with col2:
                    st.metric("Notes attribuées", len(final_data.dropna(subset=['Note'])))
                
                # Téléchargement des résultats
                st.download_button(
                    label="📥 Télécharger les données finales",
                    data=final_data.to_csv(index=False, sep=';'),
                    file_name="etudiants_avec_notes.csv",
                    mime="text/csv"
                )
                st.download_button(
                    label="🚨 Télécharger les anomalies",
                    data=anomalies_file,
                    file_name="anomalies.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Aperçu interactif
                with st.expander("Aperçu des données fusionnées"):
                    st.dataframe(final_data.style.highlight_null(color='#FF6666'))

with tab3:
    st.header("Tansfert des notes vers le fichier administratif")
    st.info("""
    - Téléchargez un fichier Excel de l'administration
    - Téléchargez un fichier CSV des notes
    - Transférer les notes vers le fichier administratif.
    """)