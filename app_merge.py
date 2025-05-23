import streamlit as st
import pandas as pd
import zipfile
import io
import os
from pathlib import Path
import re
from collections import defaultdict
import tempfile
import shutil

#st.set_option('server.maxUploadSize', 500 * 1024)  # 500 GB in MB, but Streamlit will cap at 200 MB
st.set_page_config(page_title="Excel Merger + Summary", layout="centered")
st.title("📦 Merge ZIPs → Excel Sheets + Summary + ZIP Export")


uploaded_zips = st.file_uploader(
    "Upload multiple ZIP files (each ZIP should contain .csv, .xls, .xlsx):",
    type=["zip"],
    accept_multiple_files=True
)

def clean_sheet_name(name):
    name = Path(name).stem
    name = re.sub(r'[^A-Za-z0-9_]', '_', name)
    return name[:31]

def remove_blank_rows(df):
    df = df.dropna(how='all')
    df = df.loc[:, df.notna().any()]
    return df.reset_index(drop=True)

def detect_data_and_count_rows(df):
    df = remove_blank_rows(df)
    max_scan_rows = min(30, len(df))

    for i in range(max_scan_rows):
        row = df.iloc[i]
        lowercased = [str(cell).strip().lower() for cell in row]
        if any(re.match(r's[\s\\/_-]*n|serial[\s_-]*no', col) for col in lowercased):
            header_idx = i
            header_row = df.iloc[header_idx]
            header_cols = [str(x).strip() if pd.notna(x) else f"col_{idx}" for idx, x in enumerate(header_row.values)]
            df.columns = header_cols
            df_clean = df.iloc[header_idx + 1:]

            df_clean = remove_blank_rows(df_clean)
            df_clean = df_clean[df_clean[df.columns[0]].notna()]
            contiguous_rows = df_clean[df_clean[df.columns[0]].astype(str).str.strip() != ""]
            return contiguous_rows.reset_index(drop=True), len(contiguous_rows)
    return pd.DataFrame(), 0

def extract_zip_files(zip_file):
    try:
        with zipfile.ZipFile(zip_file) as z:
            return [
                (member, io.BytesIO(z.read(member)))
                for member in z.namelist()
                if member.endswith(('.csv', '.xls', '.xlsx'))
            ]
    except zipfile.BadZipFile:
        raise RuntimeError(f"Cannot extract `{zip_file.name}` – Bad ZIP format.")

@st.cache_resource
def process_all_zips(zips):
    summary_dict = defaultdict(dict)
    summary_table_raw = []
    zip_outputs = {}
    error_logs = io.StringIO()

    temp_dir = tempfile.mkdtemp()

    try:
        for zip_file in zips:
            zip_name = Path(zip_file.name).stem
            zip_output_path = os.path.join(temp_dir, f"{zip_name}.xlsx")

            try:
                files_in_zip = extract_zip_files(zip_file)
                if not files_in_zip:
                    error_logs.write(f"⚠️ No supported files in `{zip_file.name}`\n")
                    continue

                sheet_names = {}
                buffer = io.BytesIO()

                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    for filename, file_obj in files_in_zip:
                        base_sheet_name = clean_sheet_name(filename)
                        sheet_name = base_sheet_name
                        counter = 1
                        while sheet_name in sheet_names:
                            sheet_name = f"{base_sheet_name}_{counter}"
                            counter += 1
                        sheet_names[sheet_name] = True

                        try:
                            suffix = Path(filename).suffix.lower()
                            if suffix == ".csv":
                                df = pd.read_csv(file_obj, header=None, low_memory=False, dtype=str)
                            else:
                                df = pd.read_excel(file_obj, header=None)

                            data_cleaned, count = detect_data_and_count_rows(df)
                            if count == 0:
                                raise ValueError("No data rows found below detected header.")

                            data_cleaned.to_excel(writer, sheet_name=sheet_name, index=False)

                            summary_dict[filename][zip_file.name] = count
                            summary_table_raw.append({
                                "Zip File": zip_file.name,
                                "Unzipped File": filename,
                                "Rows": count
                            })

                        except Exception as e:
                            error_logs.write(f"❌ Error in {filename} inside {zip_file.name}: {e}\n")
                            break

                writer.close()
                with open(zip_output_path, "wb") as f:
                    f.write(buffer.getvalue())
                zip_outputs[zip_name] = zip_output_path

            except Exception as e:
                error_logs.write(f"❌ Failed to process ZIP `{zip_file.name}`: {e}\n")

        # Build summary
        summary_df = pd.DataFrame(summary_table_raw)
        pivot_summary = (
            summary_df.pivot_table(index="Unzipped File", columns="Zip File", values="Rows", fill_value=0)
            .reset_index()
        )

        summary_path = os.path.join(temp_dir, "summary.xlsx")
        with pd.ExcelWriter(summary_path, engine="xlsxwriter") as writer:
            pivot_summary.to_excel(writer, index=False, sheet_name="Summary")

        error_log_path = os.path.join(temp_dir, "error_log.txt")
        with open(error_log_path, "w", encoding="utf-8") as f:
            f.write(error_logs.getvalue())

        # Zip all outputs
        zip_bundle_path = os.path.join(temp_dir, "all_outputs.zip")
        with zipfile.ZipFile(zip_bundle_path, 'w') as zipf:
            for name, path in zip_outputs.items():
                zipf.write(path, arcname=f"{name}.xlsx")
            zipf.write(summary_path, arcname="summary.xlsx")
            zipf.write(error_log_path, arcname="error_log.txt")

        return zip_outputs, summary_df, pivot_summary, error_logs.getvalue(), zip_bundle_path

    finally:
        pass  # do not clean up temp dir so downloads remain valid

# 🔄 Main Logic
if uploaded_zips:
    if "processed_outputs" not in st.session_state:
        with st.spinner("Processing ZIPs. This may take time..."):
            st.session_state["processed_outputs"] = process_all_zips(uploaded_zips)

    zip_outputs, summary_df, pivot_summary, error_content, zip_bundle_path = st.session_state["processed_outputs"]

    # 🔢 Show Summary
    if not pivot_summary.empty:
        st.subheader("📊 Summary Table (Row Counts per File per ZIP)")
        st.dataframe(pivot_summary)

        csv_data = pivot_summary.to_csv(index=False).encode("utf-8")
        st.download_button("📥 Download Summary CSV", csv_data, "summary.csv", "text/csv")

        with open(zip_bundle_path, "rb") as f:
            st.download_button("📦 Download ALL Outputs as ZIP", f.read(), "all_outputs.zip", "application/zip")

    # 📤 Per-Workbook Download
    st.subheader("📥 Download Individual Excel Workbooks")
    for zip_name, path in zip_outputs.items():
        with open(path, "rb") as f:
            st.download_button(f"⬇️ {zip_name}.xlsx", f.read(), f"{zip_name}.xlsx")

    # 📋 Error Logs
    if error_content.strip():
        st.subheader("🚨 Error Log")
        st.text_area("Errors:", error_content, height=150)
        st.download_button("📄 Download Error Log", error_content, "error_log.txt", "text/plain")
