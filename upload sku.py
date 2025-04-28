import streamlit as st
import pyodbc
import openpyxl
import tempfile
import os

# SQL Server connection details
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=seira-dev.cyqe19y3pvql.ap-south-1.rds.amazonaws.com;'
    r'DATABASE=SeiraOrder;'
    r'UID=SeiraDevAdmin#210#994;'
    r'PWD=SeiraDevAdmin#210#994;'
)

def upload_and_insert_sku(file):
    try:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(file.read())
            tmp_filepath = tmp_file.name

        # Load Excel file using openpyxl
        workbook = openpyxl.load_workbook(tmp_filepath)
        sheet = workbook.active  # Read the first sheet

        # Ensure the database connection
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        success_count = 0
        error_rows = []

        # Loop through each row and insert the data
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):  # Start from row 2
            try:
                sku_data = row

                query = """
                INSERT INTO [dbo].[sku]
                       ([company_id],[div_id],[depot_id],[sku_code],[sku_description],[sku_base_price],
                        [sku_case_size],[sku_preference],[microname],[brand_id],[pack],[varient],[volume],
                        [crs_code],[crs_code_description],[crs_code_case_size],[comments],[active],[created],
                        [created_by],[modified],[modified_by],[version],[compcode],[cat_gp_code],[cat_gp_desc],
                        [cat_code],[cat_desc],[sku_size],[sku_uom],[pr_clas],[prod_size],[base_code],[micro_code],
                        [available],[channel_id],[material_grade],[child],[multiple],[hsn_code],[batch],[is_combi],
                        [mgf_date],[expiry_date],[parent_code],[cgst_per],[igst_per],[sgst_per],[ean_code],
                        [stock_cover_day],[segment_id],[segment_name],[price_grp],[pro_npr_indicator],
                        [micro_brand],[is_child],[is_npd],[micro_name_cd],[is_split],[cpc],[variant],[div_name])
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """

                cursor.execute(query, tuple(sku_data))
                success_count += 1

            except Exception as row_error:
                error_rows.append(f"Row {row_idx}: {str(row_error)}")

        conn.commit()

        return success_count, error_rows

    except Exception as e:
        return None, [f"General Error: {str(e)}"]

    finally:
        try:
            cursor.close()
            conn.close()
        except:
            pass
        # Clean up the temp file
        if os.path.exists(tmp_filepath):
            os.remove(tmp_filepath)

# -------- Streamlit UI --------
st.set_page_config(page_title="SKU Upload Manager", page_icon="📦", layout="centered")
st.title("📦 SKU Upload Manager")

st.write("Upload an Excel (.xlsx) file to insert SKU data into the database.")

uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"])

if uploaded_file:
    if st.button("Upload and Insert into Database"):
        with st.spinner("Processing... Please wait."):
            success_count, errors = upload_and_insert_sku(uploaded_file)

        if success_count is None:
            st.error(f"Upload failed with error: {errors[0]}")
        else:
            st.success(f"Successfully inserted {success_count} records into database!")

            if errors:
                st.warning(f"However, {len(errors)} rows failed during upload.")
                with st.expander("See Error Details"):
                    for err in errors:
                        st.text(err)
else:
    st.info("Please upload an Excel file to continue.")
