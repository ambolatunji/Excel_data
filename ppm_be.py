import streamlit as st
import pandas as pd

st.title('PPM BAND EXTRACT')

# Function to normalize column names
def normalize_columns(columns):
    return [col.lower() for col in columns]

# Function to process a single file
def process_file(file, file_label):
    df = pd.read_excel(file)
    df.columns = normalize_columns(df.columns)
    
    # Select required columns and rename them for consistency
    df_selected = df[['meterno', 'custacc', 'district', 'tariff']].copy()
    
    # Add a column for the BAND derived from the TARIFF column
    df_selected['band'] = df_selected['tariff'].apply(lambda x: x[4] if len(x) > 4 else None)
    
    # Add a column indicating the source file
    df_selected['source_file'] = file_label
    
    return df_selected

# File uploader
uploaded_files = st.file_uploader("Upload Excel files", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    combined_df = pd.DataFrame()
    
    # Process each uploaded file
    for uploaded_file in uploaded_files:
        file_label = uploaded_file.name
        processed_df = process_file(uploaded_file, file_label)
        combined_df = pd.concat([combined_df, processed_df], ignore_index=True)
    
    # Sort by meterno and custacc to detect changes
    combined_df = combined_df.sort_values(by=['meterno', 'custacc'])
    
    # Detect changes in custacc for the same meterno
    combined_df['custacc_change'] = combined_df.groupby('meterno')['custacc'].apply(lambda x: x != x.shift())
    
    # Detect changes in meterno for the same custacc
    combined_df['meterno_change'] = combined_df.groupby('custacc')['meterno'].apply(lambda x: x != x.shift())
    
    # Display the combined dataframe with detected changes
    st.write("Processed Data:")
    st.dataframe(combined_df)
    
    # Optionally, save the combined DataFrame to a new Excel file
    output_file = 'processed_data.xlsx'
    combined_df.to_excel(output_file, index=False)
    st.success(f'Data processed successfully. Download the file: [processed_data.xlsx](./{output_file})')
