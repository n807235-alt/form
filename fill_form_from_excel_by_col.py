#!/usr/bin/env python3
"""
fill_form_from_excel_by_col.py
Uses column-letter mapping to fill a PDF form from all sheets in the Excel file,
naming output files by sequential row number.
"""

import os
import re
from datetime import datetime
import pandas as pd
from pdfrw import PdfReader, PdfWriter
from dateutil import parser as dateparser

# ---------- CONFIG ----------
TEMPLATE_PATH = "fillable_form_new.pdf"
EXCEL_PATH = "google_form_responses.xlsx"
SHEET_NAME = "GOOGLE Responses BATCH 1" # This variable is ignored in the main function now, but kept for context/legacy
OUT_DIR = "output_forms"
OUT_EDITABLE = os.path.join(OUT_DIR, "editable")
OUT_FLATTENED = os.path.join(OUT_DIR, "flattened")
# os.makedirs(OUT_EDITABLE, exist_ok=True) # Removed from top, added to loop for safety
# os.makedirs(OUT_FLATTENED, exist_ok=True) # Removed from top, added to loop for safety

# Column-letter -> internal field keys (based on your mapping)
COL_MAP = {
    "AA": "declaration",
    "E": "day_of_birth", # also month/year_of_birth come from same cell
    "Z": "disabled", # disabled / not_disabled
    "J": "employer_address",
    "D": "gender", # male/female
    "R": "first_child_dob",
    "Q": "first_child_name",
    "S": "first_child_school",
    "C": "name_cell", # holds surname and first_name
    "L": "ghana_card",
    "M": "marital", # married/single
    "F": "mothers_maiden",
    "I": "name_of_employer",
    "N": "name_of_spouse",
    "AC": "names_and_dates_of_aged_dependants",
    "P": "number_of_chilfren",
    "K": "phone_number",
    "U": "second_child_dob",
    "T": "second_child_name",
    "V": "second_child_school",
    # spouse DOB & spouse derived fields
    "O": "spouse_dob",
    # spouse_year_of_birth, spouse_month_of_birth, spouse_day_of_birth derive from spouse_dob
    "G": "social_sec", # social_sec (note earlier you had spacing; this uses column G)
    "H": "staff_id",
    "X": "third_child_date",
    "W": "third_child_name",
    "Y": "third_child_school",
    
    # any other columns you want add here
}

# PDF field names list (for clarity; must match actual PDF)
PDF_FIELDS = [
    "Text1","change","day_of_birth","disabled","employer_address","female",
    "first_child_dob","first_child_name","first_child_school","first_name",
    "ghana_card","male","married","month_of_birth","mothers_maiden","name_of_employer",
    "name_of_spouse","names_and_dates_of_aged_dependants","no_change","not_disabled",
    "number_of_chilfren","phone_number","second_child_dob","second_child_name",
    "second_child_school","single","social_sec","spouse_day_of_birth","spouse_ghana_card",
    "spouse_month_of_birth","spouse_social_sec","spouse_staff_id","spouse_year_of_birth",
    "staff_id","surname","third_child_date","third_child_name","third_child_school",
    "year","year_of_birth"
]

# Checkbox groups for setting annot.AS
CHECKBOX_GROUPS = {
    "gender": ("male", "female"),
    "marital": ("married", "single"),
    "disability": ("disabled", "not_disabled"),
}

# ---------- helpers ----------
def col_letter_to_index(letter):
    letter = letter.strip().upper()
    idx = 0
    for ch in letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid column letter: {letter}")
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def get_cell_by_letter(series, letter):
    try:
        idx = col_letter_to_index(letter)
        return series.iloc[idx]
    except Exception:
        return None

def norm(v):
    if v is None:
        return ""
    return str(v).strip()

def parse_date_flexible(text):
    """Return (day, month, year) as zero-padded strings if parseable, else ('','','')."""
    t = norm(text)
    if not t:
        return "", "", ""
    # remove ordinal suffixes: 1st, 2nd, 3rd, 4th...
    t = re.sub(r'(?<=\d)(st|nd|rd|th)\b', '', t, flags=re.IGNORECASE)
    t = t.replace('.', '/').replace('-', '/')
    # Try dateutil with dayfirst
    try:
        d = dateparser.parse(t, dayfirst=True, fuzzy=True)
        return f"{d.day:02d}", f"{d.month:02d}", str(d.year)
    except Exception:
        pass
    # Fallback: extract 3 number groups (DD MM YYYY or DD MM YY)
    parts = re.findall(r'\d{1,4}', t)
    if len(parts) >= 3:
        day = parts[0]
        month = parts[1]
        year = parts[2]
        if len(year) == 2:
            # assume 1900-1999 or 2000-2099 heuristic: simple choose 1900s if >30 else 2000s
            y = int(year)
            year = str(1900 + y) if y > 30 else str(2000 + y)
        return day.zfill(2), month.zfill(2), year
    return "", "", ""

# ---------- PDF filling ----------
def fill_pdf(template_path, output_edit_path, output_flat_path, row_values):
    # Re-read the template every time for robustness
    pdf = PdfReader(template_path) 
    
    for page in pdf.pages:
        if page.Annots:
            for annot in page.Annots:
                if annot.Subtype == '/Widget' and annot.T:
                    field_name = annot.T[1:-1]
                    # Always tick 'change'
                    if field_name == "change":
                        annot.AS = '/Yes'
                        continue

                    value = norm(row_values.get(field_name, ""))

                    # checkboxes
                    if field_name in sum(CHECKBOX_GROUPS.values(), ()):
                        annot.AS = '/Yes' if value and value.lower().startswith('y') or value.lower().startswith('m') or value.lower().startswith('mar') else '/Off'
                        # above is permissive; specific logic applied when building row_values
                        continue

                    # set value and try to set appearance to allow auto-scaling (best-effort)
                    if value != "":
                        annot.V = value
                        try:
                            annot.DA = "/Helv 11 Tf 0 g" # best-effort autosize hint
                        except Exception:
                            pass
                        annot.AP = None
                    else:
                        # ensure empty fields are cleared
                        annot.V = ""
                        try:
                            annot.AP = None
                        except Exception:
                            pass

    # Write the editable PDF
    PdfWriter(output_edit_path, trailer=pdf).write()
    
    # Flatten (remove annotations/fields) and write the flattened PDF
    for page in pdf.pages:
        if "/Annots" in page:
            page.Annots = []
    PdfWriter(output_flat_path, trailer=pdf).write()

# ---------- main ----------
def main():
    # 1. Read ALL sheets from the Excel file into a dictionary of DataFrames (sheet_name=None)
    try:
        all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, dtype=str, header=0)
    except FileNotFoundError:
        print(f"ERROR: Excel file not found at '{EXCEL_PATH}'.")
        return
    except Exception as e:
        print(f"ERROR loading Excel data: {e}")
        return

    # 2. Concatenate all sheets into one single DataFrame (df)
    df = pd.concat(all_sheets.values(), ignore_index=True)
    
    # Provide feedback on what was processed
    sheet_names_processed = ", ".join(all_sheets.keys())
    total = len(df)
    print(f"Loaded {total} rows from the following sheets: {sheet_names_processed}")
    
    # prepare column-letter -> index mapping for quick access
    col_to_idx = {col: col_letter_to_index(col) for col in COL_MAP.keys()}

    for i, row_series in enumerate(df.itertuples(index=False, name=None), 1):
        def cell(letter):
            try:
                idx = col_letter_to_index(letter)
                return row_series[idx] if idx < len(row_series) else None
            except Exception:
                return None

        row_values = {}

        # Populate simple mapped fields from COL_MAP
        for col_letter, key in COL_MAP.items():
            val = norm(cell(col_letter))
            row_values[key] = val

        # --- extract only the year from the timestamp column (B) ---
        timestamp = norm(cell("B"))
        year_part = ""
        if timestamp:
            # Try to extract a 4-digit or 2-digit year
            m = re.search(r'(\d{4})', timestamp)
            if not m:
                m = re.search(r'(\d{2})', timestamp)
            if m:
                y = m.group(1)
                if len(y) == 2:
                    year_part = "20" + y
                else:
                    year_part = y
        row_values["year"] = year_part

        # split name in name_cell -> surname, first_name
        name_cell = row_values.get("name_cell", "")
        surname = ""
        first_name = ""
        if name_cell:
            parts = re.split(r'\s+', name_cell.strip())
            if len(parts) >= 1:
                surname = parts[0]
                first_name = " ".join(parts[1:]) if len(parts) > 1 else ""
        row_values["surname"] = surname
        row_values["first_name"] = first_name

        # Staff ID is no longer used for primary naming, but we keep the value extraction
        staff_id = norm(row_values.get("staff_id", ""))
        if not staff_id:
            staff_id = f"unknown_{i}"
        row_values["staff_id"] = staff_id

        # Checkbox logic
        gender_raw = norm(row_values.get("gender", ""))
        row_values["male"] = "Yes" if gender_raw.lower().startswith("m") else ""
        row_values["female"] = "Yes" if gender_raw.lower().startswith("f") else ""

        marital_raw = norm(row_values.get("marital", ""))
        row_values["married"] = "Yes" if "married" in marital_raw.lower() else ""
        row_values["single"] = "Yes" if "single" in marital_raw.lower() else ""

        dis_raw = norm(row_values.get("disabled", ""))
        if dis_raw:
            if dis_raw.lower().startswith("y"):
                row_values["disabled"] = "Yes"
                row_values["not_disabled"] = ""
            else:
                row_values["disabled"] = ""
                row_values["not_disabled"] = "Yes"
        else:
            row_values["disabled"] = ""
            row_values["not_disabled"] = ""

        # DOB Parsing
        dob_cell = row_values.get("day_of_birth", "")
        d, m, y = parse_date_flexible(dob_cell)
        row_values["day_of_birth"] = d
        row_values["month_of_birth"] = m
        row_values["year_of_birth"] = y

        # Spouse DOB Parsing
        spouse_cell = row_values.get("spouse_dob", "")
        sd, sm, sy = parse_date_flexible(spouse_cell)
        row_values["spouse_day_of_birth"] = sd
        row_values["spouse_month_of_birth"] = sm
        row_values["spouse_year_of_birth"] = sy

        # Children DOB Formatting
        fc_cell = row_values.get("first_child_dob", "")
        fcd, fcm, fcy = parse_date_flexible(fc_cell)
        row_values["first_child_dob"] = (f"{fcd}/{fcm}/{fcy}" if fcd and fcm and fcy else "").strip("/")

        sc_cell = row_values.get("second_child_dob", "")
        scd, scm, scy = parse_date_flexible(sc_cell)
        row_values["second_child_dob"] = (f"{scd}/{scm}/{scy}" if scd and scm and scy else "").strip("/")

        tc_cell = row_values.get("third_child_date", "")
        tcd, tcm, tcy = parse_date_flexible(tc_cell)
        row_values["third_child_date"] = (f"{tcd}/{tcm}/{tcy}" if tcd and tcm and tcy else "").strip("/")

        row_values["change"] = "Yes"
        row_values["phone_number"] = norm(row_values.get("phone_number", ""))
        row_values["ghana_card"] = norm(row_values.get("ghana_card", ""))
        row_values["social_sec"] = norm(row_values.get("social_sec", ""))
        
        # -----------------------------------------------------------------
        # *** APPLIED FIXES: SEQUENTIAL NAMING & DIRECTORY SAFETY ***
        
        # 1. SEQUENTIAL FILENAME LOGIC (001, 002, ...)
        padding = len(str(total)) 
        new_file_name = f"{i:0{padding}d}"

        out_edit = os.path.join(OUT_EDITABLE, f"{new_file_name}.pdf")
        out_flat = os.path.join(OUT_FLATTENED, f"{new_file_name}_flat.pdf")
        
        # 2. DIRECTORY SAFETY CHECK (prevents [Errno 2])
        os.makedirs(OUT_EDITABLE, exist_ok=True)
        os.makedirs(OUT_FLATTENED, exist_ok=True)
        # -----------------------------------------------------------------
        
        row_values["year"] = "2026"
        row_values["declaration_date"] = datetime.today().strftime("%d/%m/%Y")

        print(f"[{i}/{total}] Generating {new_file_name}")
        try:
            fill_pdf(TEMPLATE_PATH, out_edit, out_flat, row_values)
        except Exception as ex:
            # We now use the sequential name in the error message
            print(f"ERROR processing row {i} (File: {new_file_name}): {ex}")


if __name__ == "__main__":
    main()
