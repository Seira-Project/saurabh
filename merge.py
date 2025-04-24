import streamlit as st
import pandas as pd
import io
import zipfile
import gc  # for memory cleanup

# --------- UTILS ---------
def read_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file), uploaded_file.name
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            return pd.read_excel(uploaded_file), uploaded_file.name
        elif uploaded_file.name.endswith('.xlsb'):
            import pyxlsb
            return pd.read_excel(uploaded_file, engine='pyxlsb'), uploaded_file.name
        else:
            st.warning(f"❗ Unsupported file format: {uploaded_file.name}")
            return None, uploaded_file.name
    except Exception as e:
        st.error(f"❌ Error reading {uploaded_file.name}: {e}")
        return None, uploaded_file.name

def split_dataframe(df, max_rows=1_000_000):
    return [df[i:i + max_rows] for i in range(0, len(df), max_rows)]

def generate_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:  # lighter on memory
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

# --------- STREAMLIT UI ---------
st.set_page_config(page_title="Excel Merger Dashboard", layout="wide")
st.title("📊 Excel/CSV File Merger")
st.caption("Upload Excel or CSV files. Automatically merges, optimizes memory, and splits large files.")

uploaded_files = st.file_uploader("📂 Upload files", type=["csv", "xlsx", "xls", "xlsb"], accept_multiple_files=True)

if uploaded_files:
    st.subheader("🧾 File Previews (first 10 rows)")
    merged_df = pd.DataFrame()
    all_columns = set()
    valid_files = 0

    for uploaded_file in uploaded_files:
        df, name = read_file(uploaded_file)
        if df is not None:
            valid_files += 1
            st.expander(f"📁 {name} ({len(df)} rows)").write(df.head(10))
            all_columns.update(df.columns)
            merged_df = pd.concat([merged_df, df], ignore_index=True)
            del df  # free memory
            gc.collect()
        else:
            st.warning(f"⚠️ Skipped {name} due to error.")

    if valid_files == 0:
        st.stop()

    st.success(f"✅ Merged {valid_files} files. {len(merged_df)} total rows.")
    st.subheader("⚙️ Merge Options")

    selected_columns = st.multiselect("🧩 Select columns to keep (optional):", list(all_columns), default=list(all_columns))
    remove_duplicates = st.checkbox("🧹 Remove duplicate rows")
    sort_column = st.selectbox("🔀 Sort by column (optional):", ["None"] + list(all_columns))

    if selected_columns:
        merged_df = merged_df[selected_columns]

    if remove_duplicates:
        merged_df = merged_df.drop_duplicates()

    if sort_column != "None":
        merged_df = merged_df.sort_values(by=sort_column)

    st.info(f"📊 Final dataset: {len(merged_df)} rows × {len(merged_df.columns)} columns")

    if not merged_df.empty:
        st.subheader("💾 Download Merged Output")

        output_files = {}
        chunks = split_dataframe(merged_df)
        progress = st.progress(0)

        for idx, chunk in enumerate(chunks):
            excel_bytes = generate_excel_bytes(chunk)
            output_files[f"merged_output_part_{idx+1}.xlsx"] = excel_bytes
            del chunk
            gc.collect()
            progress.progress((idx + 1) / len(chunks))

        for name, content in output_files.items():
            st.download_button(
                label=f"⬇️ Download {name}",
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
        st.warning("⚠️ No valid data to export.")
