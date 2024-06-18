import streamlit as st
import pandas as pd
import os
import base64

# Function to extract data from Excel files
def extract_data(file):
    try:
        data = pd.read_excel(file.getvalue())
        columns = ['MeterNo', 'AccountNo.', 'CONSUMPTION', 'Previous Reading', 'Current Reading', 'READ STATUS', 'District']
        extracted_data = data[columns]
        extracted_data['File'] = os.path.basename(file.name)
        return extracted_data
    except Exception as e:
        st.error(f"Error processing file {file.name}: {str(e)}")
        return pd.DataFrame()

# Function to download the template
def download_template(df):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="template.csv">Download Template</a>'
    return href

# Streamlit app
def main():
    st.title("Extract Data from Excel Files")

    files = st.file_uploader("Upload Excel files", accept_multiple_files=True, type=['xlsx', 'xls'])

    if files:
        data = pd.DataFrame()
        any_file_processed = False

        for file in files:
            st.info(f"Processing file: {file.name}")
            extracted = extract_data(file)
            if not extracted.empty:
                data = pd.concat([data, extracted], ignore_index=True)
                any_file_processed = True
                st.success(f"File {file.name} processed successfully!")

        if any_file_processed:
            st.subheader("Extracted Data")
            st.write(data)
            st.markdown(download_template(data), unsafe_allow_html=True)
        else:
            st.warning("No files were processed successfully. Please check the errors and try again.")

if __name__ == "__main__":
    main()