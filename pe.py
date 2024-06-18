import streamlit as st
import pandas as pd
import base64

# Function to extract data from Excel files
def extract_data(file, columns):
    try:
        data = pd.read_excel(file, dtype=str)  # Read all data as strings
        extracted_data = data.loc[:, columns]  # Use .loc to avoid SettingWithCopyWarning
        extracted_data['File'] = file.name
        return extracted_data
    except Exception as e:
        st.error(f"Error processing file {file.name}: {str(e)}")
        return pd.DataFrame()

# Function to merge data based on AccountNo.
def merge_data(base_data, other_files, columns_to_extract):
    merged_data = base_data.copy()

    for file in other_files:
        st.info(f"Processing file: {file.name}")
        other_data = extract_data(file, ['AccountNo.', 'MeterNo'] + columns_to_extract)
        if not other_data.empty:
            merged_data = pd.merge(merged_data, other_data, how='outer', on='AccountNo.', suffixes=('', f'_{file.name}'))
            st.success(f"File {file.name} processed successfully!")

    return merged_data

# Function to create the template
def create_template(data):
    template_data = data.melt(id_vars=['AccountNo.', 'District', 'MeterNo'], var_name='File_Column', value_name='Value')
    template_data[['Column', 'File']] = template_data['File_Column'].str.split('_', expand=True, n=1)
    template_data.drop(columns=['File_Column'], inplace=True)
    template_pivot = template_data.pivot_table(index=['AccountNo.', 'District', 'MeterNo', 'Column'], columns='File', values='Value', aggfunc='first').reset_index()
    return template_pivot

# Function to download the template
def download_template(df):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="template.csv">Download Template</a>'
    return href

# Function to filter and sort data based on selected columns
def filter_and_sort_data(data, filter_columns):
    sorted_data = data.sort_values(by=filter_columns)
    return sorted_data

# Streamlit app
def main():
    st.title("Excel Data Extractor and Merger")

    # First Page: Extract and Merge Data
    st.header("Upload and Merge Excel Files")

    files = st.file_uploader("Upload Excel files", accept_multiple_files=True, type=['xlsx', 'xls'])

    if files:
        st.subheader("Select the base file to extract MeterNo, District, and AccountNo. from")
        base_file = st.selectbox("Base file", files, format_func=lambda x: x.name)

        if base_file:
            # Define the columns to extract
            base_columns = ['MeterNo', 'AccountNo.', 'District']
            additional_columns = ['CONSUMPTION', 'Previous Reading', 'Current Reading', 'READ STATUS']
            base_data = extract_data(base_file, base_columns + additional_columns)
            other_files = [file for file in files if file != base_file]
            merged_data = merge_data(base_data, other_files, additional_columns)

            if not merged_data.empty:
                st.subheader("Merged Data")
                st.write(merged_data)
                st.markdown(download_template(merged_data), unsafe_allow_html=True)

                # Filter and sort options
                st.subheader("Filter and Sort Options")
                filter_columns = st.multiselect("Select columns to filter and sort by", options=merged_data.columns)

                if filter_columns:
                    filtered_sorted_data = filter_and_sort_data(merged_data, filter_columns)
                    st.subheader("Filtered and Sorted Data")
                    st.write(filtered_sorted_data)
                    st.markdown(download_template(filtered_sorted_data), unsafe_allow_html=True)

    # Second Page: Template Creation
    if st.button("Create and Download Template"):
        if 'merged_data' in locals() and not merged_data.empty:
            template_data = create_template(merged_data)
            st.subheader("Template")
            st.write(template_data)
            st.markdown(download_template(template_data), unsafe_allow_html=True)
        else:
            st.warning("Please upload and merge files first.")

if __name__ == "__main__":
    main()
