import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# Title of the app
st.title("Document and Contract ID Uploader with SharePoint Integration")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Display the dataframe
    st.write("Uploaded Data:")
    st.dataframe(df)

    # Basic functionalities
    st.write("Summary:")
    st.write(f"Total number of records: {len(df)}")
    st.write(f"Unique Document IDs: {df['Document ID'].nunique()}")
    st.write(f"Unique Contract IDs: {df['Contract ID'].nunique()}")

    # SharePoint URL input
    sharepoint_url = st.text_input("Enter SharePoint URL")

    if sharepoint_url:
        # Function to fetch files from SharePoint
        def fetch_files_from_sharepoint(url, document_id):
            response = requests.get(f"{url}/{document_id}")
            if response.status_code == 200:
                return BytesIO(response.content)
            else:
                st.error(f"Failed to fetch file for Document ID: {document_id}")
                return None

        # Fetch and display files for each Document ID
        for doc_id in df['Document ID']:
            file_content = fetch_files_from_sharepoint(sharepoint_url, doc_id)
            if file_content:
                st.write(f"File for Document ID: {doc_id}")
                st.download_button(
                    label=f"Download {doc_id}",
                    data=file_content,
                    file_name=f"{doc_id}.pdf",
                    mime='application/pdf',
                )

    # Option to download the data as CSV
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download data as CSV",
        data=csv,
        file_name='document_contract_data.csv',
        mime='text/csv',
    )
else:
    st.write("Please upload an Excel file to proceed.")
