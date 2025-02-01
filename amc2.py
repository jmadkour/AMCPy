import streamlit as st
import pandas as pd
from io import BytesIO

# Fonction de traitement pour le fichier Excel MODIFIÉE
def process_excel(file):
    try:
        # Lire le fichier Excel sans header
        df = pd.read_excel(file, header=None)
        
        # Trouver l'index de la ligne d'en-tête
        header_index = None
        for idx, row in df.iterrows():
            if all(col in row.values for col in ['Code', 'Nom', 'Prénom']):
                header_index = idx
                break
        
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
        if missing := [col for col in required_columns if col not in df.columns]:
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
        st.error(f"Erreur de traitement : {str(e)}")
        st.info("Vérifiez que le fichier est bien formaté et contient les colonnes requises.")
        return None

# (Le reste du code original reste inchangé jusqu'à l'interface utilisateur)

def process_csv(file):
    try:
        # Lire le fichier CSV
        df = pd.read_csv(file, delimiter=';', encoding='utf-8')

        # Vérification des colonnes nécessaires
        required_columns = ['A:Code', 'Nom', 'Note']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Les colonnes suivantes sont manquantes : {', '.join(missing_columns)}")
            return None, None

        # Nettoyage approfondi
        anomalies = pd.DataFrame()
        
        # 1. Filtrer les lignes avec 'NONE' ou valeurs vides
        none_mask = (df['A:Code'] == 'NONE') | df[['A:Code', 'Note']].isnull().any(axis=1)
        none_rows = df[none_mask]
        anomalies = pd.concat([anomalies, none_rows])
        df_clean = df[~none_mask].copy()

        # 2. Vérifier la cohérence entre 'A:Code' et 'Code' converti
        df_clean['Code'] = pd.to_numeric(df_clean['A:Code'], errors='coerce')
        mismatch_mask = df_clean['A:Code'].astype(str) != df_clean['Code'].astype(str)
        mismatch_rows = df_clean[mismatch_mask]
        anomalies = pd.concat([anomalies, mismatch_rows])
        df_clean = df_clean[~mismatch_mask]

        # Vérifier si le fichier nettoyé est vide
        if df_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None

        # Gérer les notes manquantes
        missing_grades = df_clean['Note'].isnull().sum()
        if missing_grades > 0:
            st.warning(f"{missing_grades} étudiants n'ont pas de note correspondante")

        # Préparer le fichier d'anomalies
        anomalies_file = BytesIO()
        anomalies.to_excel(anomalies_file, index=False)
        anomalies_file.seek(0)

        return df_clean, anomalies_file

    except Exception as e:
        st.error(f"Erreur critique lors du traitement CSV : {str(e)}")
        return None, None

# Interface utilisateur 
st.title("Préparer et traiter les fichiers des étudiants")

# Onglets pour séparer les sections
tab1, tab2 = st.tabs(["Fichier Excel", "Fichier CSV"])

with tab1:
    st.header("1. Préparation de la liste des étudiants (fichier Excel)")
    st.info("""
    - Téléchargez un fichier Excel contenant les colonnes 'Nom', 'Prénom' et 'Code'
    - La détection des en-têtes est automatique
    - Les lignes avant les en-têtes seront automatiquement supprimées
    """)
    
    # Télécharger le fichier Excel
    uploaded_excel_file = st.file_uploader("Charger le fichier Excel de l'administration", 
                                         type="xlsx", 
                                         key="excel_uploader")
    
    if uploaded_excel_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            processed_data = process_excel(uploaded_excel_file)
            
            if processed_data is not None:
                st.success(f"Traitement réussi ! {len(processed_data)} étudiants valides trouvés.")
                
                # Convertir le DataFrame en CSV
                csv = processed_data.to_csv(index=False)
                
                # Permettre le téléchargement du fichier CSV
                st.download_button(
                    label="Télécharger le fichier CSV final",
                    data=csv,
                    file_name="liste_etudiants.csv",
                    mime="text/csv"
                )

with tab2:
    st.header("2. Traitement des fichiers CSV des étudiants")
    st.info("""
    - Téléchargez un fichier CSV contenant les colonnes 'A:Code', 'Nom' et 'Note'
    - Les anomalies seront exportées dans un fichier Excel séparé
    - Les notes seront automatiquement associées aux étudiants de l'onglet 1
    """)
    
    uploaded_csv_file = st.file_uploader("Charger le fichier CSV des étudiants", type="csv", key="csv_uploader")
    
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
                    file_name="anomalies.csv",
                    mime="application/vnd.ms-excel"
                )

                # Aperçu interactif
                with st.expander("Aperçu des données fusionnées"):
                    st.dataframe(final_data.style.highlight_null(color='#FF6666'))