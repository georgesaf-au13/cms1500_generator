import pymupdf  # PyMuPDF
import pandas as pd
import os
import re
from datetime import datetime

# ==========================================
# 1. Configuration & File Paths
# ==========================================
CMS_PROJECT_PATH = r"D:\Work\cms_filling"
CMS_1500_PATH = CMS_PROJECT_PATH+r"\form-cms1500.pdf"

WEBPT_DIR = CMS_PROJECT_PATH+r"\webpt"
UTILS_DIR = CMS_PROJECT_PATH+r"\utils"

INPUT_DIR = WEBPT_DIR+r"\input" 
OUTPUT_DIR = WEBPT_DIR+r"\output"

#webpt reports
SV_FILE = INPUT_DIR+r"\scheduled_visits.csv"
UNITS_FILE = INPUT_DIR+r"\units.csv"
PATIENT_FILE = INPUT_DIR+r"\patients_list.csv"

#sub reports
CHARGES_FILE = UTILS_DIR+r"\charges.xlsx"
CREDENTIALS_FILE = UTILS_DIR+r"\provider_credentials.csv"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Payer map retained purely for attaching the [Payer ID] to the top field
PAYER_MAP = {
    "Aetna": {
        "insurance_name": "AETNA",
        "payer_id": "60054"
    },
    "UHC": {
        "insurance_name": "UNITED HEALTH CARE",
        "payer_id": "87726"
    },
    "UnitedHealthCare": {
        "insurance_name": "UNITED HEALTH CARE",
        "payer_id": "87726"
    },
}

# ==========================================
# 2. Helper Functions
# ==========================================
def parse_date(date_val):
    if pd.isna(date_val) or str(date_val).strip() == '':
        return "", "", ""
    try:
        if isinstance(date_val, str):
            date_val = pd.to_datetime(date_val)
        return date_val.strftime("%m"), date_val.strftime("%d"), date_val.strftime("%Y")
    except:
        return "", "", ""

def clean_number_string(val):
    val = str(val).strip()
    if val.lower() == 'nan' or val == '':
        return ''
    if 'e+' in val.lower():
        try:
            return str(int(float(val)))
        except:
            return val
    if val.endswith('.0'):
        return val[:-2]
    return val

def fill_pdf(output_filename, field_data):
    doc = pymupdf.open(CMS_1500_PATH)
    page = doc.load_page(0)
    
    # 1. Fill the form fields
    for widget in page.widgets():
        field_name = widget.field_name
        
        if field_name in field_data:
            insert_value = str(field_data[field_name]).strip()
            
            if insert_value != "" or field_name == "id_physician":
                widget.field_value = insert_value
                widget.update()
            
    # 2. Render Page 1 to high-resolution image bytes (300 DPI)
    pix = page.get_pixmap(dpi=300)
    img_bytes = pix.tobytes("png")
    
    # 3. Create a static PDF from the image bytes
    img_doc = pymupdf.open("png", img_bytes)
    pdf_bytes = img_doc.convert_to_pdf()
    img_doc.close()
    
    # 4. Save as a non-fillable image PDF
    out_doc = pymupdf.open("pdf", pdf_bytes)
    out_doc.save(output_filename)
    out_doc.close()
    doc.close()

# ==========================================
# 3. Main Processing Logic
# ==========================================
def process_claims(visits_file, units_file, charges_file, patient_file, creds_file):
    df_visits = pd.read_csv(visits_file, encoding='utf-8-sig', dtype=str).fillna('')
    df_units = pd.read_csv(units_file, encoding='utf-8-sig', dtype=str).fillna('')
    df_patients = pd.read_csv(patient_file, encoding='utf-8-sig', dtype=str).fillna('')
    df_creds = pd.read_csv(creds_file, encoding='utf-8-sig', dtype=str).fillna('')
    df_charges = pd.read_excel(charges_file) 

    df_visits.columns = df_visits.columns.str.strip()
    df_units.columns = df_units.columns.str.strip()
    df_patients.columns = df_patients.columns.str.strip()
    df_creds.columns = df_creds.columns.str.strip()
    df_charges.columns = df_charges.columns.str.strip()

    df_visits['EMR Patient ID'] = df_visits['EMR Patient ID'].str.strip()
    df_units['EMR Patient ID'] = df_units['EMR Patient ID'].str.strip()
    df_patients['Patient ID'] = df_patients['Patient ID'].str.strip()
    
    df_visits['Appointment Date'] = pd.to_datetime(df_visits['Appointment Date'], errors='coerce').dt.date
    df_units['Date of Service'] = pd.to_datetime(df_units['Date of Service'], errors='coerce').dt.date

    # Aggressive CPT cleaning to fix the $0 charge matching bug
    df_charges['cpt'] = df_charges['cpt'].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
    df_charges['charge'] = pd.to_numeric(df_charges['charge'], errors='coerce').fillna(0.0)
    charge_dict = dict(zip(df_charges['cpt'], df_charges['charge']))

    # ==========================================
    # GENERATE INCOMPLETE NOTES REPORT
    # ==========================================
    # Extract unique combinations of Patient + Date from the units file
    unique_units = df_units[['EMR Patient ID', 'Date of Service']].drop_duplicates()
    
    # Merge visits with units to find what's missing
    unmatched_visits = df_visits.merge(
        unique_units,
        left_on=['EMR Patient ID', 'Appointment Date'],
        right_on=['EMR Patient ID', 'Date of Service'],
        how='left',
        indicator=True
    )
    
    # Filter only the rows that exist in visits but NOT in units
    incomplete_notes_df = unmatched_visits[unmatched_visits['_merge'] == 'left_only'].copy()
    
    # Clean up the extra columns generated by the merge
    if 'Date of Service' in incomplete_notes_df.columns:
        incomplete_notes_df = incomplete_notes_df.drop(columns=['Date of Service', '_merge'])
    else:
        incomplete_notes_df = incomplete_notes_df.drop(columns=['_merge'])
        
    # Export to Excel
    report_date = datetime.today().strftime('%Y%m%d')
    report_path = os.path.join(OUTPUT_DIR, f"incomplete_notes_{report_date}.xlsx")
    incomplete_notes_df.to_excel(report_path, index=False)
    print(f"\n---> Report Generated: {len(incomplete_notes_df)} incomplete notes saved to {report_path}\n")

    # ==========================================
    # START PDF CLAIM GENERATION
    # ==========================================
    grouped_units = df_units.groupby(['EMR Patient ID', 'Date of Service'])

    for (patient_id, dos), lines in grouped_units:
        visit_info = df_visits[(df_visits['EMR Patient ID'] == patient_id) & 
                               (df_visits['Appointment Date'] == dos)]
        
        if visit_info.empty:
            continue
            
        visit_info = visit_info.iloc[0]
        claim_data = {}
        
        # -- Map Patient Header Info --
        claim_data['pt_name'] = visit_info['Patient Name']
        b_mm, b_dd, b_yy = parse_date(visit_info['Patient DOB'])
        claim_data['birth_mm'], claim_data['birth_dd'], claim_data['birth_yy'] = b_mm, b_dd, b_yy
        
        patient_demo = df_patients[df_patients['Patient ID'] == patient_id]
        
        if not patient_demo.empty:
            demo_row = patient_demo.iloc[0]
            claim_data['pt_street'] = demo_row.get('Address', '')
            claim_data['pt_city'] = demo_row.get('City', '')
            claim_data['pt_state'] = demo_row.get('State', '')
            claim_data['pt_zip'] = clean_number_string(demo_row.get('Zip', ''))
            
            raw_phone = str(demo_row.get('Phone', ''))
            clean_phone = re.sub(r'\D', '', raw_phone) 
            
            if len(clean_phone) >= 10:
                claim_data['pt_AreaCode'] = clean_phone[:3]
                claim_data['pt_phone'] = clean_phone[3:10]
            else:
                claim_data['pt_phone'] = clean_phone

        claim_data['ins_name'] = claim_data.get('pt_name', '')
        claim_data['ins_street'] = claim_data.get('pt_street', '')
        claim_data['ins_city'] = claim_data.get('pt_city', '')
        claim_data['ins_state'] = claim_data.get('pt_state', '')
        claim_data['ins_zip'] = claim_data.get('pt_zip', '')
        
        primary_ins = str(visit_info.get('Primary Insurance', ''))
        claim_data['insurance_id'] = clean_number_string(visit_info.get('Primary Insurance Policy Number', ''))
        
        # Set Patient Account Number to Insurance ID (Box 26)
        claim_data['pt_account'] = claim_data['insurance_id']
        
        matched_payer = next((k for k in PAYER_MAP.keys() if k.lower() in primary_ins.lower()), None)
        
        if matched_payer:
            payer_data = PAYER_MAP[matched_payer]
            base_name = payer_data.get('insurance_name', matched_payer)
            payer_id = payer_data.get('payer_id', '')
            claim_data['insurance_name'] = f"{base_name} [{payer_id}]" if payer_id else base_name
        else:
            claim_data['insurance_name'] = primary_ins

        # ==========================================
        # NEW: DYNAMIC CREDENTIAL LOOKUP
        # ==========================================
        clinic_name = str(visit_info.get('Clinic Name', '')).strip()
        
        # 1. Try to find exact Clinic + Insurance match (regex=False prevents string warnings)
        cred_match = df_creds[
            (df_creds['Clinic Name'].str.contains(clinic_name, case=False, na=False, regex=False)) & 
            (df_creds['Insurance Name'].str.contains(primary_ins, case=False, na=False, regex=False))
        ]
        
        # 2. Fall back to the "DEFAULT" row for that clinic if no exact insurance match
        if cred_match.empty:
            cred_match = df_creds[
                (df_creds['Clinic Name'].str.contains(clinic_name, case=False, na=False, regex=False)) & 
                (df_creds['Insurance Name'].str.upper() == 'DEFAULT')
            ]
            
        prov_data = {}
        if not cred_match.empty:
            prov_data = cred_match.iloc[0].to_dict()
            claim_data.update(prov_data)
            
        # Hardcode Box 24J bottom and Box 31 Date
        claim_data['id_physician'] = "" 
        claim_data['physician_date'] = datetime.today().strftime('%m/%d/%Y')

        # -- ICD-10 Parsing --
        all_diags = []
        for d_str in lines['Treating Diagnosis'].astype(str):
            diags = [d.strip() for d in d_str.replace(';', ',').split(',')]
            for d in diags:
                if d and d.lower() != 'nan' and d not in all_diags:
                    all_diags.append(d)
        
        diag_letters = "ABCDEFGHIJKL"
        diag_map = {}
        for idx, d in enumerate(all_diags[:12]):
            claim_data[f'diagnosis{idx+1}'] = d  
            diag_map[d] = diag_letters[idx]      

        # -- Service Lines (Box 24) --
        total_charge = 0.0
        
        for i, (_, line) in enumerate(lines.head(6).iterrows(), start=1):
            sv_mm, sv_dd, sv_yy = parse_date(line['Date of Service'])
            claim_data[f'sv{i}_mm_from'] = sv_mm
            claim_data[f'sv{i}_dd_from'] = sv_dd
            claim_data[f'sv{i}_yy_from'] = sv_yy
            
            pos = line.get('Place of Service', '').strip()
            if pos and pos.lower() != 'nan':
                claim_data[f'place{i}'] = "11"
            
            cpt_code = clean_number_string(line['CPT Code']).upper()
            claim_data[f'cpt{i}'] = cpt_code
            claim_data[f'day{i}'] = clean_number_string(line['Units'])
            
            line_diags = [d.strip() for d in str(line.get('Treating Diagnosis', '')).replace(';', ',').split(',')]
            pointers = [diag_map[d] for d in line_diags if d in diag_map]
            claim_data[f'diag{i}'] = "".join(pointers[:4])
            
            mods = [m.strip() for m in str(line.get('Modifier', '')).replace(',', ' ').split() if m.strip() and m.strip().lower() != 'nan']
            if len(mods) > 0: claim_data[f'mod{i}'] = mods[0]
            if len(mods) > 1: claim_data[f'mod{i}a'] = mods[1]
            if len(mods) > 2: claim_data[f'mod{i}b'] = mods[2]
            if len(mods) > 3: claim_data[f'mod{i}c'] = mods[3]
            
            # Specific line-item fields
            claim_data[f'emg{i}'] = 'ZZ'
            claim_data[f'local{i}'] = prov_data.get('rendering_npi', '')
            claim_data[f'local{i}a'] = prov_data.get('taxonomy', '')
            
            # Charges calculation
            unit_charge = charge_dict.get(cpt_code, 0.0)
            
            if unit_charge == 0.0 and cpt_code != '':
                print(f"  -> WARNING: No charge found in Excel for CPT: {cpt_code}")
                
            try:
                units_count = float(line['Units']) if line['Units'] else 0.0
            except ValueError:
                units_count = 0.0
                
            line_charge = unit_charge * units_count
            claim_data[f'ch{i}'] = f"{line_charge:.2f}"
            total_charge += line_charge

        claim_data['t_charge'] = f"{total_charge:.2f}"
        
        safe_pt_name = str(claim_data.get('pt_name', 'Unknown')).replace(' ', '_').replace('/', '')
        safe_dos = f"{dos.strftime('%m%d%y') if pd.notna(dos) else 'NoDate'}"
        output_filename = os.path.join(OUTPUT_DIR, f"CMS1500_{safe_pt_name}_{safe_dos}.pdf")
        
        fill_pdf(output_filename, claim_data)
        print(f"Generated: {output_filename}")


if __name__ == "__main__":
    process_claims(SV_FILE, UNITS_FILE, CHARGES_FILE, PATIENT_FILE, CREDENTIALS_FILE)