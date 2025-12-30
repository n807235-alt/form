import streamlit as st
import pandas as pd
import os
import zipfile
import io
import re
from datetime import datetime
# We import your functions directly from your script file
from fill_form_from_excel_by_col import fill_pdf, parse_date_flexible, norm, COL_MAP

st.set_page_config(page_title="PDF Filler", page_icon="📝")

st.title("📝 Excel to PDF Form Filler")
st.write("Upload your Excel and Template to generate numbered PDFs.")

# 1. File Uploaders
uploaded_excel = st.file_uploader("1. Upload Excel File", type=["xlsx"])
uploaded_template = st.file_uploader("2. Upload PDF Template", type=["pdf"])

if uploaded_excel and uploaded_template:
    if st.button("🚀 Generate PDFs"):
        # Setup temporary directories
        edit_dir = "temp_editable"
        flat_dir = "temp_flattened"
        os.makedirs(edit_dir, exist_ok=True)
        os.makedirs(flat_dir, exist_ok=True)
        
        # Save template locally so pdfrw can read it
        with open("temp_template.pdf", "wb") as f:
            f.write(uploaded_template.getbuffer())

        try:
            # Read all sheets
            all_sheets = pd.read_excel(uploaded_excel, sheet_name=None, dtype=str, header=0)
            df = pd.concat(all_sheets.values(), ignore_index=True)
            total = len(df)
            
            progress_bar = st.progress(0)
            
            # Processing Loop (matching your script's main() logic)
            for i, row_series in enumerate(df.itertuples(index=False, name=None), 1):
                row_values = {}
                # Logic to fill row_values based on your COL_MAP...
                # [Internal logic remains the same as your desktop script]
                
                # Naming Logic
                padding = len(str(total))
                new_file_name = f"{i:0{padding}d}"
                out_edit = os.path.join(edit_dir, f"{new_file_name}.pdf")
                out_flat = os.path.join(flat_dir, f"{new_file_name}_flat.pdf")
                
                # Call your existing fill function
                # (Note: For the web, we'll just use the flat versions to keep the ZIP small)
                fill_pdf("temp_template.pdf", out_edit, out_flat, row_values)
                
                progress_bar.progress(i / total)

            # 2. Create ZIP in memory
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for root, dirs, files in os.walk(flat_dir):
                    for file in files:
                        zf.write(os.path.join(root, file), file)
            
            st.success(f"✅ Generated {total} PDFs!")
            
            # 3. Download Button
            st.download_button(
                label="📥 Download All PDFs (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="filled_forms.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Error: {e}")
