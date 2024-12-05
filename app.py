import streamlit as st
import pandas as pd
import requests

# Titel der App
st.title("Dokumenten- und Vertrags-ID-Manager mit SharePoint-Integration")

# Datei-Uploader
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

# Funktion zum Abrufen der Dateien von SharePoint
def search_files_in_sharepoint(base_url, directory, doc_id, username, password):
    # URL für SharePoint API
    search_url = f"{base_url}/_api/web/GetFolderByServerRelativeUrl('{directory}')/Files"
    try:
        response = requests.get(
            search_url,
            auth=(username, password),
            headers={"Accept": "application/json;odata=verbose"}
        )
        
        # Statuscode und Antwortinhalt debuggen
        st.write("Antwort Statuscode:", response.status_code)
        if response.status_code != 200:
            st.error(f"Fehler bei der Anfrage. Status Code: {response.status_code}")
            return []

        # Antwortinhalt als JSON verarbeiten
        files = response.json()['d']['results']
        matching_files = [f for f in files if doc_id.lower() in f['Name'].lower()]  # Unterscheidung der Groß-/Kleinschreibung ignorieren
        return matching_files

    except requests.exceptions.RequestException as e:
        st.error(f"Fehler bei der Anfrage: {e}")
        return []
    except ValueError:
        st.error("Fehler: Die Antwort ist kein gültiges JSON.")
        st.write("Antwortinhalt:", response.text)
        return []

# Wenn eine Datei hochgeladen wird
if uploaded_file is not None:
    try:
        # Excel-Datei einlesen
        df = pd.read_excel(uploaded_file)

        # Überprüfen, ob die erforderlichen Spalten vorhanden sind
        required_columns = ['Document ID', 'Contract ID']
        if all(col in df.columns for col in required_columns):
            st.write("Hochgeladene Daten:")
            st.dataframe(df)

            # Eingabe der SharePoint-Details
            sharepoint_url = st.text_input("Gib die SharePoint-URL ein (z. B. https://example.sharepoint.com/sites/your-site)")
            sharepoint_directory = st.text_input("Gib das SharePoint-Verzeichnis ein (z. B. /Shared Documents)")

            # Eingabe der Anmeldedaten
            sharepoint_username = st.text_input("Gib deinen SharePoint-Benutzernamen ein")
            sharepoint_password = st.text_input("Gib dein SharePoint-Passwort ein", type="password")

            if sharepoint_url and sharepoint_directory and sharepoint_username and sharepoint_password:
                # Dateien für jede Document ID suchen
                for doc_id in df['Document ID']:
                    st.write(f"Suche nach Dateien für Document ID: {doc_id}...")
                    matches = search_files_in_sharepoint(
                        sharepoint_url, sharepoint_directory, doc_id, sharepoint_username, sharepoint_password
                    )
                    if matches:
                        for file in matches:
                            file_url = file['ServerRelativeUrl']
                            st.write(f"Gefundene Datei: {file['Name']} ({file_url})")
                            file_data = requests.get(
                                f"{sharepoint_url}{file_url}",
                                auth=(sharepoint_username, sharepoint_password)
                            ).content
                            st.download_button(
                                label=f"Download {file['Name']}",
                                data=file_data,
                                file_name=file['Name'],
                                mime='application/octet-stream',
                            )
                    else:
                        st.write(f"Keine Dateien gefunden für Document ID: {doc_id}")

        else:
            st.error("Die Datei muss die Spalten 'Document ID' und 'Contract ID' enthalten!")
    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.write("Bitte lade eine Excel-Datei hoch, um fortzufahren.")
