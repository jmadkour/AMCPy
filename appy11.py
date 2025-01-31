import streamlit as st
import pandas as pd
from io import BytesIO

# Fonction de traitement pour le fichier Excel
def process_excel(file, rows_to_delete):
    try:
        # Lire l'Excel et ignorer les 'rows_to_delete' premières lignes
        df = pd.read_excel(file, skiprows=rows_to_delete)
        
        # Vérification si le fichier est vide
        if df.empty:
            st.error("Le fichier Excel est vide. Veuillez vérifier le fichier téléchargé.")
            return None

        # Vérification de la présence des colonnes nécessaires
        required_columns = ['Nom', 'Prénom', 'Code']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Les colonnes suivantes sont manquantes dans le fichier Excel : {', '.join(missing_columns)}")
            st.info("Assurez-vous que le fichier contient les colonnes 'Nom', 'Prénom' et 'Code'.")
            return None

        # Vérification des valeurs manquantes dans les colonnes essentielles
        if df[['Nom', 'Prénom', 'Code']].isnull().any().any():
            st.warning("Certaines lignes contiennent des valeurs manquantes dans les colonnes 'Nom', 'Prénom' ou 'Code'. Elles seront supprimées.")

        # Affichage d'un aperçu des données
        st.write("Aperçu des données :")
        st.write(df.head(10))

        # Nettoyage des données : suppression des lignes avec des valeurs manquantes dans les colonnes essentielles
        df = df.dropna(subset=['Nom', 'Prénom', 'Code'])

        # Suppression des premières lignes selon la valeur de 'rows_to_delete'
        if rows_to_delete > 0:
            df = df.iloc[rows_to_delete:]  # Supprimer les premières 'rows_to_delete' lignes

        # Création de la colonne 'Name' à partir des colonnes 'Code', 'Nom', et 'Prénom'
        df['Name'] = df['Code'].astype(str) + ' ' + df['Nom'] + ' ' + df['Prénom']
        
        # Ne garder que les colonnes 'Code' et 'Name'
        df = df[['Code', 'Name']]

        # Suppression des doublons
        df = df.drop_duplicates(subset=['Code', 'Name'])

        return df

    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement du fichier Excel : {e}")
        st.info("Veuillez vérifier que le fichier est bien au format Excel (.xlsx) et qu'il contient des données valides.")
        return None

# Fonction de traitement du fichier CSV
def process_csv(file):
    try:
        # Lire le fichier CSV avec le bon délimiteur et gestion de l'encodage
        df = pd.read_csv(file, delimiter=';', encoding='utf-8')

        # Vérification si le fichier est vide
        if df.empty:
            st.error("Le fichier CSV est vide. Veuillez vérifier le fichier téléchargé.")
            return None, None, None

        # Vérification des colonnes nécessaires
        required_columns = ['A:Code', 'Nom', 'Note']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Les colonnes suivantes sont manquantes dans le fichier CSV : {', '.join(missing_columns)}")
            st.info("Assurez-vous que le fichier contient les colonnes 'A:Code', 'Nom' et 'Note'.")
            return None, None, None

        # Conversion de la colonne 'Code' en numérique, avec gestion des erreurs
        df['Code'] = pd.to_numeric(df['Code'], errors='coerce')

        # Vérification si la conversion a échoué pour certaines lignes
        if df['Code'].isnull().any():
            st.warning("Certaines valeurs dans la colonne 'Code' n'ont pas pu être converties en nombres. Ces lignes seront supprimées.")

        # Supprimer les lignes où 'A:Code' est 'NONE'
        none_rows = df[df['A:Code'] == 'NONE']
        df_cleaned = df[df['A:Code'] != 'NONE']

        # Enregistrer les lignes avec 'NONE' dans un fichier Excel séparé
        none_rows_file = BytesIO()
        none_rows.to_excel(none_rows_file, index=False)
        none_rows_file.seek(0)  # Rewind le fichier pour le téléchargement

        # Vérifier les discordances entre 'A:Code' et 'Code'
        mismatch_rows = df_cleaned[df_cleaned['A:Code'] != df_cleaned['Code']]
        mismatch_rows_file = BytesIO()
        mismatch_rows.to_excel(mismatch_rows_file, index=False)
        mismatch_rows_file.seek(0)  # Rewind le fichier pour le téléchargement

        # Retourner les fichiers séparés et le DataFrame nettoyé
        return df_cleaned, none_rows_file, mismatch_rows_file

    except pd.errors.EmptyDataError:
        st.error("Le fichier CSV est vide ou mal formé. Veuillez vérifier le fichier.")
        return None, None, None
    except pd.errors.ParserError:
        st.error("Le fichier CSV est mal formé. Vérifiez le format et les délimiteurs.")
        st.info("Assurez-vous que le fichier utilise le point-virgule (;) comme délimiteur.")
        return None, None, None
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement du fichier CSV : {e}")
        st.info("Veuillez vérifier que le fichier est bien au format CSV et qu'il contient des données valides.")
        return None, None, None

# Interface utilisateur Streamlit
st.title("Préparer et traiter les fichiers des étudiants")

# Onglets pour séparer les sections
tab1, tab2 = st.tabs(["Fichier Excel", "Fichier CSV"])

with tab1:
    st.header("1. Préparation de la liste des étudiants (fichier Excel)")
    st.info("""
    - Téléchargez un fichier Excel contenant les colonnes 'Nom', 'Prénom' et 'Code'.
    - Indiquez le nombre de lignes à ignorer en haut du fichier.
    - Le programme générera un fichier CSV avec les colonnes 'Code' et 'Name'.
    """)
    
    # Télécharger le fichier Excel
    uploaded_excel_file = st.file_uploader("Charger le fichier Excel de l'administration", type="xlsx", key="excel_uploader")
    
    # Demander à l'utilisateur combien de lignes supprimer
    rows_to_delete = st.number_input("Combien de lignes souhaitez-vous supprimer ?", min_value=0, value=0, key="rows_to_delete")
    
    if uploaded_excel_file is not None:
        with st.spinner("Traitement du fichier Excel en cours..."):
            processed_data = process_excel(uploaded_excel_file, rows_to_delete)
            if processed_data is not None:
                st.success("Traitement terminé !")
                st.write("Données après traitement :")
                st.write(processed_data.head(10))
                
                # Convertir le DataFrame en CSV
                csv = processed_data.to_csv(index=False)
                
                # Permettre le téléchargement du fichier CSV
                if st.download_button(
                    label="Télécharger le fichier CSV",
                    data=csv,
                    file_name="liste.csv",
                    mime="text/csv"
                ):
                    st.success("Fichier CSV téléchargé avec succès !")

with tab2:
    st.header("2. Traitement des fichiers CSV des étudiants")
    st.info("""
    - Téléchargez un fichier CSV contenant les colonnes 'A:Code', 'Nom' et 'Note'.
    - Le programme nettoiera les données et générera des fichiers Excel pour les lignes problématiques.
    """)
    
    # Télécharger le fichier CSV
    uploaded_csv_file = st.file_uploader("Charger le fichier CSV des étudiants", type="csv", key="csv_uploader")
    
    if uploaded_csv_file is not None:
        with st.spinner("Traitement du fichier CSV en cours..."):
            df_cleaned, none_rows_file, mismatch_rows_file = process_csv(uploaded_csv_file)
            if df_cleaned is not None:
                st.success("Traitement terminé !")
                st.write("Données après nettoyage :")
                st.write(df_cleaned.head())
                
                # Permettre le téléchargement des fichiers Excel séparés
                if st.download_button(
                    label="Télécharger les lignes avec 'NONE' dans 'A:Code'",
                    data=none_rows_file,
                    file_name="lignes_NONE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ):
                    st.success("Fichier des lignes 'NONE' téléchargé avec succès !")
                
                if st.download_button(
                    label="Télécharger les lignes avec des discordances entre 'A:Code' et 'Code'",
                    data=mismatch_rows_file,
                    file_name="lignes_incorrectes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ):
                    st.success("Fichier des lignes incorrectes téléchargé avec succès !")

                # Permettre le téléchargement du fichier nettoyé
                if st.download_button(
                    label="Télécharger le fichier nettoyé",
                    data=df_cleaned.to_csv(index=False).encode('utf-8'),
                    file_name="fichier_nettoye.csv",
                    mime="text/csv"
                ):
                    st.success("Fichier nettoyé téléchargé avec succès !")