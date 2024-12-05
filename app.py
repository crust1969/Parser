import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# Titel der App
st.title("Dokumenten- und Vertrags-ID-Manager mit SharePoint-Dateisuche")

# Datei-Uploader
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

# Überprüfen, ob eine Datei hochgeladen wurde
if uploaded_file is not None:
    try:
        # Excel-Datei lesen
        df = pd.read_excel(uploaded_file)

        # Sicherstellen, dass die benötigten Spalten vorhanden sind
        required_columns = ['Document ID', 'Contract ID']
        if all(col in df.columns for col in required_columns):
            # Dataframe anzeigen
            st.write("Hochgeladene Daten:")
            st.dataframe(df)

            # Datenzusammenfassung
            st.write("Zusammenfassung:")
            st.write(f"Gesamtanzahl der Datensätze: {len(df)}")
            st.write(f"Einzigartige Dokument-IDs: {df['Document ID'].nunique()}")
            st.write(f"Einzigartige Vertrags-IDs: {df['Contract ID'].nunique()}")

            # Eingabe der SharePoint-URL
            sharepoint_url = st.text_input("Gib die SharePoint-URL ein")
            sharepoint_token = st.text_input("Gib dein SharePoint-Token ein", type="password")

            if sharepoint_url and sharepoint_token:
                st.write("Suche Dateien auf SharePoint...")

                # Fortschrittsanzeige
                progress = st.progress(0)

                # Funktion, um Dateien auf SharePoint zu suchen
                def search_files_on_sharepoint(base_url, token, document_id):
                    search_query = f"?$filter=substringof('{document_id}', Name)"
                    headers = {
                        "Authorization": f"Bearer {token}",
                        "Accept": "application/json"
                    }
                    response = requests.get(f"{base_url}/_api/web/lists/getbytitle('Documents')/items{search_query}", headers=headers)
                    if response.status_code == 200:
                        results = response.json()
                        return results.get('value', [])
                    else:
                        st.error(f"Fehler bei der Suche für Document ID: {document_id}. {response.text}")
                        return []

                # Dateien für jede Dokument-ID suchen und Download-Buttons erstellen
                for idx, doc_id in enumerate(df['Document ID']):
                    results = search_files_on_sharepoint(sharepoint_url, sharepoint_token, doc_id)
                    if results:
                        for file in results:
                            file_name = file['Name']
                            file_url = file['ServerRelativeUrl']
                            st.write(f"Gefundene Datei für {doc_id}: {file_name}")
                            st.download_button(
                                label=f"Download {file_name}",
                                data=requests.get(f"{sharepoint_url}{file_url}").content,
                                file_name=file_name,
                                mime='application/octet-stream',
                            )
                    progress.progress((idx + 1) / len(df['Document ID']))

            # Option, die Daten als CSV herunterzuladen
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Daten als CSV herunterladen",
                data=csv,
                file_name='document_contract_data.csv',
                mime='text/csv',
            )
        else:
            st.error("Die Datei muss die Spalten 'Document ID' und 'Contract ID' enthalten!")
    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.write("Bitte lade eine Excel-Datei hoch, um fortzufahren.")
