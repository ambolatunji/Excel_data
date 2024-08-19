import streamlit as st
import pandas as pd
from pathlib import Path

# Define functions to handle file uploads and merging
def read_and_concatenate(files):
    dataframes = []
    for file in files:
        df = pd.read_excel(file)
        df['SourceFile'] = Path(file.name).stem
        dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)

def align_and_concatenate(files):
    dataframes = []
    all_columns = set()

    # First pass to gather all unique columns
    for file in files:
        df = pd.read_excel(file)
        all_columns.update(df.columns)

    # Second pass to align all dataframes to the same columns
    for file in files:
        df = pd.read_excel(file)
        for col in all_columns:
            if col not in df.columns:
                df[col] = pd.NA  # Fill missing columns with NaN
        df = df[sorted(all_columns)]  # Ensure the same column order
        df['SourceFile'] = Path(file.name).stem
        dataframes.append(df)
    
    return pd.concat(dataframes, ignore_index=True)

# Main App
st.title("Excel File Merger")

# Define navigation options
page = st.sidebar.radio("Select a page:", ["Merge Matching Files", "Merge and Align Different Files"])

# Page 1: Merge Files with Matching Columns
if page == "Merge Matching Files":
    st.header("Merge Files with Matching Columns")

    uploaded_files = st.file_uploader("Upload Excel Files", type="xlsx", accept_multiple_files=True)
    
    if uploaded_files:
        # Read and concatenate files
        result_df = read_and_concatenate(uploaded_files)
        st.write("Merged Data:")
        st.write(result_df)
        
        # Provide download option
        csv = result_df.to_csv(index=False)
        st.download_button(label="Download Merged CSV", data=csv, file_name="merged_data.csv", mime="text/csv")
    
# Page 2: Merge Files with Different Columns
elif page == "Merge and Align Different Files":
    st.header("Merge Files with Different Columns")

    uploaded_files = st.file_uploader("Upload Excel Files", type="xlsx", accept_multiple_files=True)
    
    if uploaded_files:
        # Align and concatenate files
        result_df = align_and_concatenate(uploaded_files)
        st.write("Aligned and Merged Data:")
        st.write(result_df)
        
        # Provide download option
        csv = result_df.to_csv(index=False)
        st.download_button(label="Download Merged and Aligned CSV", data=csv, file_name="aligned_merged_data.csv", mime="text/csv")
