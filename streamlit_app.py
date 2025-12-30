import streamlit as st
import pandas as pd
import os
import zipfile
import io
from datetime import datetime
from fill_form_from_excel_by_col import fill_pdf, parse_date_flexible, norm, COL_MAP

st.set_page_config(page_title="PDF Filler", page_icon="📄")
st.title("📄 PDF Form Auto-Filler")

uploaded_excel = st.file_uploader("Upload Excel", type=["xlsx"])
uploaded_template = st.file_uploader("Upload PDF Template", type=["pdf"])

if uploaded_excel and uploaded_template:
    if st.button("🚀 Start Generating"):
        # Create directories
        flat_dir = "output_flat"
        # Cleanup old files if they exist from a previous failed run
        if os.path.exists(flat_dir):
            import shutil
            shutil.rmtree(flat_dir)
        os.makedirs(flat_dir, exist_ok=True)
        
        # Save template
        with open("temp_template.pdf", "wb") as f:
            f.write(uploaded_template.getbuffer())

        try:
            all_sheets = pd.read_excel(uploaded_excel, sheet_name=None, dtype=str, header=0)
            df = pd.concat(all_sheets.values(), ignore_index=True)
            total = len(df)
            
            bar = st.progress(0)
            status = st.empty()
            
            for i, row in enumerate(df.itertuples(index=False, name=None), 1):
                row_values = {}
                # Logic to fill row_values based on your COL_MAP
                for col_letter, key in COL_MAP.items():
                    # Manual letter-to-index conversion to avoid dependency issues
                    letter = col_letter.strip().upper()
                    idx = 0
                    for ch in letter:
                        idx = idx * 26 + (ord(ch) - ord('A') + 1)
                    col_idx = idx - 1
                    row_values[key] = norm(row[col_idx]) if col_idx < len(row) else ""

                # Derived fields
                name_cell = row_values.get("name_cell", "")
                parts = name_cell.split()
                row_values["surname"] = parts[0] if len(parts) > 0 else ""
                row_values["first_name"] = " ".join(parts[1:]) if len(parts) > 1 else ""
                d, m, y = parse_date_flexible(row_values.get("day_of_birth", ""))
                row_values["day_of_birth"], row_values["month_of_birth"], row_values["year_of_birth"] = d, m, y
                row_values["change"] = "Yes"
                row_values["year"] = "2026"
                row_values["declaration_date"] = datetime.today().strftime("%d/%m/%Y")

                # Generate File
                padding = len(str(total))
                fname = f"{i:0{padding}d}.pdf"
                out_path = os.path.join(flat_dir, fname)
                
                # Fill the PDF
                fill_pdf("temp_template.pdf", "temp_dummy.pdf", out_path, row_values)
                
                bar.progress(i / total)
                status.text(f"Generated {i}/{total}")

            # --- THE MEMORY FIX: Write ZIP to disk, not RAM ---
            zip_filename = "final_output.zip"
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(flat_dir):
                    for f in files:
                        zf.write(os.path.join(root, f), f)
            
            st.success("🎉 All files created successfully!")
            
            # Provide the download by reading the file back
            with open(zip_filename, "rb") as f:
                st.download_button(
                    label="📥 Download ZIP",
                    data=f,
                    file_name="filled_forms.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"Processing Error: {e}")
