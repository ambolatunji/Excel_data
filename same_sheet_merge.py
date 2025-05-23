import streamlit as st
import pandas as pd
import zipfile
import io
import os
from pathlib import Path
import re

st.set_page_config(page_title="Merge Files into Excel", layout="centered")

st.title("📊 Merge Files into a Single Excel Workbook (Memory-Efficient + Live Progress)")

uploaded_files = st.file_uploader(
    "Upload .csv, .xls, .xlsx, or .zip files (ZIP can contain any of the allowed types):",
    type=["csv", "xls", "xlsx", "zip"],
    accept_multiple_files=True
)

def clean_sheet_name(name):
    name = Path(name).stem
    name = re.sub(r'[^A-Za-z0-9_]', '_', name)
    return name[:31]  # Excel sheet name max length

def read_file(file, filename, chunksize=None):
    suffix = Path(filename).suffix.lower()
    if suffix == '.csv':
        if chunksize:
            return pd.read_csv(file, chunksize=chunksize)
        return pd.read_csv(file)
    elif suffix in ['.xls', '.xlsx']:
        return pd.read_excel(file)
    else:
        raise ValueError("Unsupported file format")

def extract_zip(file):
    try:
        extracted = []
        with zipfile.ZipFile(file) as z:
            for member in z.namelist():
                if member.endswith(('.csv', '.xls', '.xlsx')):
                    with z.open(member) as extracted_file:
                        extracted.append((member, io.BytesIO(extracted_file.read())))
        return extracted
    except zipfile.BadZipFile:
        raise RuntimeError("Invalid ZIP file")

if uploaded_files:
    st.info("Preparing to process files...")

    sheet_data = {}
    sheet_names_set = set()
    files_to_process = []

    # Step 1: Unpack all files into a flat list with filename, file_obj
    try:
        for uploaded_file in uploaded_files:
            if uploaded_file.name.endswith('.zip'):
                extracted_files = extract_zip(uploaded_file)
                for fname, file_obj in extracted_files:
                    files_to_process.append((fname, file_obj))
            else:
                files_to_process.append((uploaded_file.name, uploaded_file))
    except Exception as e:
        st.error(f"❌ Failed to extract ZIP file: {e}")
        st.stop()

    total_files = len(files_to_process)
    progress_bar = st.progress(0)
    completed_files = 0

    error_occurred = False
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, (fname, file_obj) in enumerate(files_to_process):
            sheet_name = clean_sheet_name(fname)
            if sheet_name in sheet_names_set:
                st.warning(f"⚠️ Duplicate sheet name: `{sheet_name}` – Skipping `{fname}`.")
                continue

            st.write(f"🔄 Processing `{fname}` ...")
            try:
                # For CSVs, support memory-efficient read (if needed, here chunking is skipped for Excel output)
                df = read_file(file_obj, fname)
                if hasattr(df, '__iter__') and not isinstance(df, pd.DataFrame):
                    df = pd.concat(df)

                sheet_data[sheet_name] = df
                sheet_names_set.add(sheet_name)

                df.to_excel(writer, sheet_name=sheet_name, index=False)

                with st.expander(f"✅ Completed: {sheet_name}"):
                    st.dataframe(df.head(50))

                completed_files += 1
                progress_bar.progress(completed_files / total_files)

            except Exception as e:
                st.error(f"❌ Error while processing `{fname}`: {e}")
                error_occurred = True
                break

    if completed_files > 0:
        st.success(f"✅ Processed {completed_files} file(s) successfully.")
        st.download_button(
            label="📥 Download Partial Excel File" if error_occurred else "📥 Download Merged Excel File",
            data=output.getvalue(),
            file_name="merged_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if error_occurred:
        st.warning("⚠️ Further processing stopped due to an error.")
