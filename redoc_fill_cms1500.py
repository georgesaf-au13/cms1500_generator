import pymupdf  # PyMuPDF
import pandas as pd
import os
import re
from datetime import datetime

# ==========================================
# 1. Configuration & File Paths
# ==========================================
CMS_PROJECT_PATH = r"D:\Work\cms_filling"
CMS_1500_PATH = r"D:\Work\cms_filling\form-cms1500.pdf"

REDOC_DIR = CMS_PROJECT_PATH + r"\redoc"
UTILS_DIR = CMS_PROJECT_PATH + r"\utils"

INPUT_DIR = REDOC_DIR + r"\input"
OUTPUT_DIR = REDOC_DIR + r"\output"

VISITS_FILE = INPUT_DIR + r"\VisitTypeReport_v2.csv"
UNITS_FILE = INPUT_DIR + r"\DailyRecordOfInterventions.csv"
MAILING_FILE = INPUT_DIR + r"\PatientListMailing.csv"

CHARGES_FILE = UTILS_DIR + r"\charges.xlsx"
CREDENTIALS_FILE = UTILS_DIR + r"\provider_credentials.csv"

os.makedirs(OUTPUT_DIR, exist_ok=True)

PAYER_MAP = {
    "Aetna": {"insurance_name": "Aetna Zaya", "payer_id": "60054"},
    "UHC": {"insurance_name": "UHC Zaya", "payer_id": "87726"}
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
    if val.lower() == 'nan' or val == '': return ''
    if 'e+' in val.lower():
        try: return str(int(float(val)))
        except: return val
    if val.endswith('.0'): return val[:-2]
    return val

def standardize_name(name):
    name = str(name).upper().strip()
    name = re.sub(r'\s+', ' ', name)
    if ',' in name:
        parts = name.split(',', 1)
        return f"{parts[0].strip()}, {parts[1].strip()}"
    return name

def extract_insurance_id(account_string):
    val = str(account_string).strip().upper()
    if val.lower() == 'nan' or not val: return ''
    if ' DOB' in val:
        val = val.split(' DOB')[0].strip()
    chunks = val.split()
    return chunks[-1] if chunks else val

def extract_insurance_name(account_string):
    """Extracts the carrier string from combined fields (e.g., 'Healthfirst 13177769' -> 'Healthfirst')."""
    val = str(account_string).strip()
    if val.lower() == 'nan' or not val: return ''
    if ' DOB' in val:
        val = val.split(' DOB')[0].strip()
    chunks = val.split()
    return chunks[0] if chunks else ''

def clean_clinic_name(name_str):
    return re.sub(r'(?i)^location:\s*', '', str(name_str)).strip()

def fill_pdf(output_filename, field_data):
    doc = pymupdf.open(CMS_1500_PATH)
    page = doc.load_page(0)
    
    for widget in page.widgets():
        field_name = widget.field_name
        if field_name in field_data:
            insert_value = str(field_data[field_name]).strip()
            if insert_value != "" or field_name == "id_physician":
                widget.field_value = insert_value
                widget.update()
            
    out_doc = pymupdf.open()
    out_doc.insert_pdf(doc, from_page=0, to_page=0)
    out_doc.bake()
    out_doc.save(output_filename)
    out_doc.close()
    doc.close()

# ==========================================
# 3. Main Processing Logic
# ==========================================
def process_redoc_claims():
    df_visits = pd.read_csv(VISITS_FILE, encoding='utf-8-sig', dtype=str, skiprows=26).fillna('')
    df_units = pd.read_csv(UNITS_FILE, encoding='utf-8-sig', dtype=str, skiprows=3).fillna('')
    df_mail = pd.read_csv(MAILING_FILE, encoding='utf-8-sig', dtype=str, skiprows=3).fillna('')
    
    df_creds = pd.read_csv(CREDENTIALS_FILE, encoding='utf-8-sig', dtype=str).fillna('')
    df_charges = pd.read_excel(CHARGES_FILE) 

    for df in [df_visits, df_units, df_mail, df_creds, df_charges]:
        df.columns = df.columns.str.strip()

    # NORMALIZATION BEFORE FILTERING
    df_visits['CleanFullName'] = df_visits['ScheduledPatientFullName'].apply(standardize_name)
    df_units['CleanFullName'] = df_units['PatientName'].apply(standardize_name)
    df_mail['CleanFullName'] = df_mail['PatientName'].apply(standardize_name)

    # ==========================================
    # APPLY ZAYA MASK TO VISITS & FILTER UNITS
    # ==========================================
    mask_zaya = (
        df_visits["PrimaryInsurance"].str.contains("Zaya", case=False, na=False) 
        & ~df_visits["PrimaryInsurance"].str.contains("Non Zaya", case=False, na=False) 
        & ~df_visits["VisitStatus"].str.contains("Missed", case=False, na=False)
    )
    df_visits = df_visits.loc[mask_zaya, :].copy()

    # FIX #1: Strictly restrict df_units to ONLY patients in the filtered Zaya visits list
    zaya_patients = df_visits['CleanFullName'].unique()
    df_units = df_units[df_units['CleanFullName'].isin(zaya_patients)].copy()

    # Filter out ReDoc totals rows
    df_units = df_units[~df_units['PatientName'].str.contains('Totals', case=False, na=False)]
    df_units = df_units[df_units['BillingCode'].str.strip() != '']

    # Standardize InsID and Dates
    df_visits['InsID'] = df_visits['PatientNumber'].apply(extract_insurance_id)
    df_units['InsID'] = df_units['AccountNumber'].apply(extract_insurance_id)
    df_mail['InsID'] = df_mail['PatientNumber'].apply(extract_insurance_id)

    df_visits['MatchDate'] = pd.to_datetime(df_visits['visitDate'], errors='coerce').dt.date
    df_units['MatchDate'] = pd.to_datetime(df_units['VisitDate'], errors='coerce').dt.date

    # Setup Charges Dictionary
    df_charges['cpt'] = df_charges['cpt'].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
    df_charges['charge'] = pd.to_numeric(df_charges['charge'], errors='coerce').fillna(0.0)
    charge_dict = dict(zip(df_charges['cpt'], df_charges['charge']))

    # ==========================================
    # GENERATE INCOMPLETE NOTES REPORT
    # ==========================================
    unique_units = df_units[['CleanFullName', 'MatchDate']].drop_duplicates()
    
    unmatched_visits = df_visits.merge(
        unique_units,
        on=['CleanFullName', 'MatchDate'],
        how='left',
        indicator=True
    )
    
    incomplete_notes_df = unmatched_visits[unmatched_visits['_merge'] == 'left_only'].copy()
    incomplete_notes_df = incomplete_notes_df.drop(columns=['_merge'])
        
    report_date = datetime.today().strftime('%Y%m%d')
    report_path = os.path.join(OUTPUT_DIR, f"incomplete_notes_{report_date}.xlsx")
    incomplete_notes_df.to_excel(report_path, index=False)
    print(f"\n---> Report Generated: {len(incomplete_notes_df)} incomplete notes saved to {report_path}\n")

    # ==========================================
    # START PDF CLAIM GENERATION
    # ==========================================
    grouped_units = df_units.groupby(['CleanFullName', 'InsID', 'MatchDate'])

    for (patient_name, ins_id, dos), lines in grouped_units:
        claim_data = {}
        first_line = lines.iloc[0]

        # ------------------------------------------
        # 1. Lookup Demographics
        # ------------------------------------------
        demo_info = df_mail[(df_mail['CleanFullName'] == patient_name) & (df_mail['InsID'] == ins_id)]
        patient_dob = pd.NaT
        
        if not demo_info.empty:
            demo = demo_info.iloc[0]
            claim_data['pt_street'] = demo.get('textbox98', '')
            claim_data['pt_city'] = demo.get('textbox82', '')
            claim_data['pt_state'] = demo.get('textbox74', '')
            claim_data['pt_zip'] = clean_number_string(demo.get('textbox50', ''))
            
            patient_dob_str = demo.get('PatientDOB', '')
            if pd.notna(patient_dob_str) and str(patient_dob_str).strip():
                patient_dob = pd.to_datetime(patient_dob_str, errors='coerce')
                b_mm, b_dd, b_yy = parse_date(patient_dob_str)
                claim_data['birth_mm'], claim_data['birth_dd'], claim_data['birth_yy'] = b_mm, b_dd, b_yy

        # ------------------------------------------
        # 2. Robust Visit & Insurance Lookup
        # ------------------------------------------
        visit_info = df_visits[(df_visits['CleanFullName'] == patient_name) & 
                               (df_visits['MatchDate'] == dos)]
        
        # Fallback to name-only match in df_visits if date matching was off
        if visit_info.empty:
            visit_info = df_visits[df_visits['CleanFullName'] == patient_name]
        
        primary_ins = ""
        clinic_name = ""

        if not visit_info.empty:
            visit = visit_info.iloc[0]
            primary_ins = str(visit.get('PrimaryInsurance', '')).strip()
            clinic_name = clean_clinic_name(visit.get('ComponentOfCare', ''))
            
            raw_phone = str(visit.get('PatientPhone', ''))
            clean_phone = re.sub(r'\D', '', raw_phone) 
            if len(clean_phone) >= 10:
                claim_data['pt_AreaCode'] = clean_phone[:3]
                claim_data['pt_phone'] = clean_phone[3:10]
            else:
                claim_data['pt_phone'] = clean_phone

        # FIX #2: If primary_ins is still missing or UNKNOWN, parse it directly from AccountNumber/MRN
        if not primary_ins or primary_ins.upper() == 'UNKNOWN':
            primary_ins = extract_insurance_name(first_line.get('AccountNumber', ''))
            if not primary_ins:
                primary_ins = extract_insurance_name(first_line.get('MRN', ''))

        if not clinic_name:
            clinic_name = clean_clinic_name(first_line.get('Location', ''))

        # ------------------------------------------
        # 3. Patient & Insurance Header
        # ------------------------------------------
        claim_data['pt_name'] = first_line['PatientName']
        claim_data['ins_name'] = claim_data.get('pt_name', '')
        claim_data['ins_street'] = claim_data.get('pt_street', '')
        claim_data['ins_city'] = claim_data.get('pt_city', '')
        claim_data['ins_state'] = claim_data.get('pt_state', '')
        claim_data['ins_zip'] = claim_data.get('pt_zip', '')

        claim_data['insurance_id'] = ins_id
        claim_data['pt_account'] = ins_id

        # ------------------------------------------
        # 4. Dynamic Credentials Lookup
        # ------------------------------------------
        cred_match = df_creds[
            (df_creds['Clinic Name'].str.contains(clinic_name, case=False, na=False, regex=False)) & 
            (df_creds['Insurance Name'].str.contains(primary_ins, case=False, na=False, regex=False))
        ]
        if cred_match.empty:
            cred_match = df_creds[
                (df_creds['Clinic Name'].str.contains(clinic_name, case=False, na=False, regex=False)) & 
                (df_creds['Insurance Name'].str.upper() == 'DEFAULT')
            ]
            
        prov_data = {}
        if not cred_match.empty:
            prov_data = cred_match.iloc[0].to_dict()
            claim_data.update(prov_data)
            
        claim_data['id_physician'] = "" 
        claim_data['physician_date'] = datetime.today().strftime('%m/%d/%Y')
        
        # Format Box 1 Insurance Name
        matched_payer = next((k for k in PAYER_MAP.keys() if k.lower() in primary_ins.lower()), None)
        if matched_payer:
            payer_data = PAYER_MAP[matched_payer]
            base_name = payer_data.get('insurance_name', matched_payer)
            payer_id = payer_data.get('payer_id', '')
            claim_data['insurance_name'] = f"{base_name} [{payer_id}]" if payer_id else base_name
        else:
            claim_data['insurance_name'] = primary_ins

        # ------------------------------------------
        # 5. ICD-10 Parsing (Box 21)
        # ------------------------------------------
        all_diags = []
        for _, line in lines.iterrows():
            icd_primary = str(line.get('ICD', '')).strip()
            if icd_primary and icd_primary.lower() != 'nan' and icd_primary not in all_diags:
                all_diags.append(icd_primary)
                
            icd_add = str(line.get('AdditionalICD', '')).replace('Additional Diagnoses:', '').strip()
            if icd_add and icd_add.lower() != 'nan':
                for d in [x.strip() for x in icd_add.split(',')]:
                    if d and d not in all_diags:
                        all_diags.append(d)

        diag_letters = "ABCDEFGHIJKL"
        diag_map = {}
        for idx, d in enumerate(all_diags[:12]):
            claim_data[f'diagnosis{idx+1}'] = d  
            diag_map[d] = diag_letters[idx]      

        # ------------------------------------------
        # 6. Service Lines (Box 24)
        # ------------------------------------------
        total_charge = 0.0
        
        for i, (_, line) in enumerate(lines.head(6).iterrows(), start=1):
            sv_mm, sv_dd, sv_yy = parse_date(dos.strftime('%m/%d/%Y'))
            claim_data[f'sv{i}_mm_from'] = sv_mm
            claim_data[f'sv{i}_dd_from'] = sv_dd
            claim_data[f'sv{i}_yy_from'] = sv_yy
            claim_data[f'place{i}'] = '11'
            
            cpt_code = clean_number_string(line['BillingCode']).upper()
            claim_data[f'cpt{i}'] = cpt_code
            claim_data[f'day{i}'] = clean_number_string(line['Units'])
            
            line_diags = [str(line.get('ICD', '')).strip()]
            add_diags = str(line.get('AdditionalICD', '')).replace('Additional Diagnoses:', '').strip()
            if add_diags: line_diags.extend([x.strip() for x in add_diags.split(',')])
            
            pointers = [diag_map[d] for d in line_diags if d in diag_map]
            claim_data[f'diag{i}'] = "".join(pointers[:4])
            
            mod = str(line.get('InterventionModifier', '')).strip()
            if mod and mod.lower() != 'nan':
                claim_data[f'mod{i}'] = mod
                
            claim_data[f'emg{i}'] = 'ZZ'
            claim_data[f'local{i}'] = prov_data.get('rendering_npi', '')
            claim_data[f'local{i}a'] = prov_data.get('taxonomy', '')
            
            unit_charge = charge_dict.get(cpt_code, 0.0)
            if unit_charge == 0.0:
                redoc_fee = pd.to_numeric(line.get('InterventionFee', 0), errors='coerce')
                unit_charge = float(redoc_fee) if pd.notna(redoc_fee) else 0.0
                
            try:
                units_count = float(line['Units']) if line['Units'] else 0.0
            except ValueError:
                units_count = 0.0
                
            line_charge = unit_charge * units_count
            claim_data[f'ch{i}'] = f"{line_charge:.2f}"
            total_charge += line_charge

        claim_data['t_charge'] = f"{total_charge:.2f}"
        
        # ------------------------------------------
        # 7. Output Generation
        # ------------------------------------------
        safe_name = patient_name.replace(', ', '_').replace(',', '_').replace(' ', '_')
        safe_dob = patient_dob.strftime('%m%d%Y') if pd.notna(patient_dob) else 'NoDOB'
        safe_ins = ins_id if ins_id else 'NoID'
        safe_dos = dos.strftime('%m%d%Y') if pd.notna(dos) else 'NoDOS'
        
        filename_str = f"{safe_name}_{safe_dob}_{safe_ins}_{safe_dos}.pdf"
        filename_str = re.sub(r'[^A-Za-z0-9_.]', '', filename_str)
        
        output_filename = os.path.join(OUTPUT_DIR, filename_str)
        
        fill_pdf(output_filename, claim_data)
        print(f"Generated: {output_filename}")

if __name__ == "__main__":
    process_redoc_claims()