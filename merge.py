import streamlit as st
import pandas as pd
import io
import zipfile

def read_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            return pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.xlsb'):
            import pyxlsb
            return pd.read_excel(uploaded_file, engine='pyxlsb')
        else:
            st.warning(f"Unsupported file format: {uploaded_file.name}")
            return None
    except Exception as e:
        st.error(f"Error reading {uploaded_file.name}: {e}")
        return None

def split_dataframe(df, max_rows=10_00_000):
    chunks = [df[i:i + max_rows] for i in range(0, len(df), max_rows)]
    return chunks

def generate_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

def create_zip(files_dict):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for name, content in files_dict.items():
            zip_file.writestr(name, content.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

st.set_page_config(page_title="Excel Merger & Splitter Dashboard", layout="wide")
st.title("📊 Excel File Merger & Splitter")
st.markdown("Upload multiple Excel/CSV files and merge them into a single sheet. If the output has more than 10 lakh rows, it will split automatically into multiple files.")

uploaded_files = st.file_uploader("📂 Upload Excel or CSV files", accept_multiple_files=True, type=['csv', 'xlsx', 'xls', 'xlsb'])
output_files = {}

if uploaded_files:
    st.info(f"Total files uploaded: {len(uploaded_files)}")
    merged_df = pd.DataFrame()
    for uploaded_file in uploaded_files:
        df = read_file(uploaded_file)
        if df is not None:
            st.success(f"✅ {uploaded_file.name} - {len(df)} rows")
            merged_df = pd.concat([merged_df, df], ignore_index=True)

    if not merged_df.empty:
        st.subheader("✅ Merge Summary")
        st.write(f"Total Rows After Merge: {len(merged_df)}")

        # Splitting if necessary
        chunks = split_dataframe(merged_df)
        for idx, chunk in enumerate(chunks):
            excel_bytes = generate_excel_bytes(chunk)
            output_files[f"merged_output_part_{idx+1}.xlsx"] = excel_bytes

        if output_files:
            st.subheader("📥 Download Merged Files")
            for name, content in output_files.items():
                st.download_button(
                    label=f"Download {name}",
                    data=content,
                    file_name=name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            zip_buffer = create_zip(output_files)
            st.download_button(
                label="📦 Download All as ZIP",
                data=zip_buffer,
                file_name="merged_outputs.zip",
                mime="application/zip"
            )
        else:
            st.warning("No output files were generated. Please check the input files.")
    else:
        st.error("❌ No valid data found to merge. Please upload proper files.")
