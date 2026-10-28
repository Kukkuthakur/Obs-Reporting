import os
import json
import pydicom
from pynetdicom import AE, evt, StoragePresentationContexts

# --- CONFIGURATION ---
# You can change these, but make sure they match your Streamlit app and USG machine settings
AE_TITLE = 'OB_REPORT_APP'
PORT = 11112
STORAGE_DIR = 'received_sr'
JSON_FILE = 'latest_sr.json'

# Create storage directory if it doesn't exist
if not os.path.exists(STORAGE_DIR):
    os.makedirs(STORAGE_DIR)

# --- PARSER ---
def extract_sr_data(ds):
    from datetime import datetime, timedelta
    extracted = {}
    
    def walk_tree(seq, current_site=""):
        for item in seq:
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            new_site = current_site
            if concept in ["Umbilical Artery", "Uterine Artery"]:
                new_site = concept

            if item.ValueType == 'TEXT':
                val = item.get('TextValue', '')
                if val: 
                    if concept == 'EDD from average ultrasound age': extracted['edd_cua'] = val
                    if concept == 'EDD from LMP': extracted['edd_lmp'] = val
                    if concept == 'LMP': extracted['lmp'] = val

            if item.ValueType == 'NUM' and 'MeasuredValueSequence' in item:
                val_seq = item.MeasuredValueSequence[0]
                num_val = float(val_seq.get('NumericValue', 0))
                unit = val_seq.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')
                
                if unit == 'centimeter':
                    if concept == 'Biparietal Diameter': extracted['bpd'] = f"{int(num_val * 10)}"
                    elif concept == 'Head Circumference': extracted['hc'] = f"{int(num_val * 10)}"
                    elif concept == 'Abdominal Circumference': extracted['ac'] = f"{int(num_val * 10)}"
                    elif concept == 'Femur Length': extracted['fl'] = f"{int(num_val * 10)}"
                    elif concept == 'Humerus length': extracted['hl'] = f"{int(num_val * 10)}"
                    elif concept == 'Gestational Sac Diameter': extracted['msd'] = f"{int(num_val * 10)}"
                    elif concept == 'Crown Rump Length': extracted['crl'] = f"{int(num_val * 10)}"
                
                if new_site == "Umbilical Artery" and concept == 'Pulsatility Index':
                    extracted['umbilical_pi'] = str(num_val)
                if new_site == "Uterine Artery" and concept == 'Pulsatility Index':
                    extracted['uterine_pi'] = str(num_val)
                
                if concept == 'Estimated Weight': extracted['fetal_weight'] = str(int(num_val))
                if concept == 'Fetal Heart Rate': extracted['fhr'] = str(int(num_val))
                
                if concept == 'Composite Ultrasound Age': 
                    extracted['ga_usg_days'] = int(num_val)
                    extracted['mean_age'] = f"{int(num_val)//7}weeks {int(num_val)%7}days ± 2w"
                
                if concept == 'Gestational Age by LMP':
                    extracted['ga_lmp_days'] = int(num_val)

                if concept == 'Growth Percentile Rank':
                    if 'current_param' in extracted:
                        param = extracted['current_param']
                        extracted[f"{param}_pct"] = str(num_val)

                if concept == 'Gestational Age':
                    if 'current_param' in extracted:
                        param = extracted['current_param']
                        extracted[f"{param}_wks"] = str(int(num_val) // 7)
                        extracted[f"{param}_dys"] = str(int(num_val) % 7)

            if concept in ['Biparietal Diameter', 'Head Circumference', 'Abdominal Circumference', 'Femur Length', 'Humerus length']:
                param_map = {
                    'Biparietal Diameter': 'bpd', 'Head Circumference': 'hc',
                    'Abdominal Circumference': 'ac', 'Femur Length': 'fl', 'Humerus length': 'hl'
                }
                extracted['current_param'] = param_map[concept]
            
            if 'ContentSequence' in item:
                walk_tree(item.ContentSequence, new_site)
            
            if concept == 'Biometry Group':
                extracted.pop('current_param', None)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)
        
    extracted.pop('current_param', None)
    
    today = datetime.now()
    if 'lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['lmp'] = (today - timedelta(days=extracted['ga_lmp_days'])).strftime("%d-%m-%Y")
    if 'edd_lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['edd_lmp'] = (today + timedelta(days=(280 - extracted['ga_lmp_days']))).strftime("%d-%m-%Y")
    if 'edd_cua' not in extracted and 'ga_usg_days' in extracted:
        extracted['edd_cua'] = (today + timedelta(days=(280 - extracted['ga_usg_days']))).strftime("%d-%m-%Y")

    return extracted

# --- C-STORE HANDLER ---
def handle_store(event):
    ds = event.dataset
    ds.file_meta = event.file_meta
    
    # 1. Save the raw .dcm file
    filename = f"{STORAGE_DIR}/{ds.SOPInstanceUID}.dcm"
    ds.save_as(filename, write_like_original=False)
    print(f"\n[+] Received DICOM file: {ds.PatientName}")
    
    # 2. Parse the data and save to JSON for Streamlit
    try:
        parsed_data = extract_sr_data(ds)
        with open(JSON_FILE, 'w') as f:
            json.dump(parsed_data, f)
        print(f"[+] Successfully parsed data and saved to {JSON_FILE}")
    except Exception as e:
        print(f"[-] Error parsing data: {e}")

    return 0x0000

# --- START SERVER ---
def main():
    ae = AE(ae_title=AE_TITLE)
    ae.supported_contexts = StoragePresentationContexts
    
    handlers = [(evt.EVT_C_STORE, handle_store)]
    
    print(f"Starting DICOM Listener on Port {PORT} with AE Title '{AE_TITLE}'...")
    print("Waiting for the ultrasound machine to send a report...")
    print("Leave this window open while you are working.")
    
    ae.start_server(('0.0.0.0', PORT), evt_handlers=handlers)

if __name__ == '__main__':
    main()
