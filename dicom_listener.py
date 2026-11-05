import os
import re
import json
import pydicom
from pynetdicom import AE, evt
from pynetdicom.sop_class import Verification
from pynetdicom.presentation import AllStoragePresentationContexts

# --- CONFIGURATION ---
AE_TITLE = 'PCARE'
PORT = 5103
STORAGE_DIR = 'received_sr'
JSON_FILE = 'latest_sr.json'

if not os.path.exists(STORAGE_DIR):
    os.makedirs(STORAGE_DIR)

# --- DOPPLER MAPPING ---
# SR vessel container name -> key prefix used in latest_sr.json
SITE_KEYS = {
    "Umbilical Artery": "umbilical",
    "Uterine Artery": "uterine",
    "Middle Cerebral Artery": "mca",
    "MCA": "mca",
}
# SR measurement name -> key suffix
DOPPLER_METRICS = {
    "Systolic to Diastolic Velocity Ratio": "sd",
    "Resistivity Index": "ri",
    "Pulsatility Index": "pi",
}

# --- PARSER ---
def extract_sr_data(ds):
    from datetime import datetime, timedelta
    extracted = {}
    
    def walk_tree(seq, current_site="", current_side=""):
        for item in seq:
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            new_site = current_site
            new_side = current_side
            if concept in SITE_KEYS:
                new_site = concept
                new_side = ""
                # Laterality is a CODE child of the vessel container (Left / Right / Unilateral...)
                for sub in item.get('ContentSequence', []):
                    if sub.get('ValueType', '') == 'CODE' and 'ConceptNameCodeSequence' in sub:
                        if sub.ConceptNameCodeSequence[0].get('CodeMeaning', '') == 'Laterality':
                            new_side = sub.ConceptCodeSequence[0].get('CodeMeaning', '').strip().lower()
                base = SITE_KEYS[concept]
                if new_side in ('left', 'right'):
                    extracted.setdefault(f"{base}_sides", []).append(new_side)
                elif new_side:
                    extracted[f"{base}_laterality"] = new_side
            elif 'Artery' in concept or 'MCA' in concept:
                print(f"[?] Unmapped vessel concept in SR: '{concept}'")

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
                    elif concept == 'Nuchal Translucency': extracted['nt'] = f"{int(num_val * 10)}"

                if concept == 'Nasal Bone':
                    extracted['nasal_bone'] = f"{num_val}mm"
                
                if 'Choroid Plexus' in concept or 'Falx' in concept:
                    extracted['falx'] = "Appreciated, appear Normal"

                # Doppler indices (S/D, RI, PI) for Umbilical / Uterine / MCA
                if new_site in SITE_KEYS and concept in DOPPLER_METRICS:
                    base = SITE_KEYS[new_site]
                    metric = DOPPLER_METRICS[concept]
                    extracted[f"{base}_{metric}"] = str(num_val)
                    if new_side in ('left', 'right'):
                        extracted[f"{base}_{new_side}_{metric}"] = str(num_val)
                
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
                walk_tree(item.ContentSequence, new_site, new_side)
            
            if concept == 'Biometry Group':
                extracted.pop('current_param', None)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)
        
    extracted.pop('current_param', None)

    # Bilateral vessels: store a clean laterality + mean values (unsided key = mean of L/R)
    for base in ("uterine", "mca", "umbilical"):
        sides = extracted.pop(f"{base}_sides", [])
        if sides:
            extracted[f"{base}_laterality"] = "bilateral" if len(set(sides)) == 2 else sides[0]
            for metric in ("sd", "ri", "pi"):
                vals = [float(extracted[f"{base}_{s}_{metric}"]) for s in set(sides)
                        if f"{base}_{s}_{metric}" in extracted]
                if len(vals) == 2:
                    extracted[f"{base}_{metric}"] = f"{sum(vals)/2:.2f}"
    
    today = datetime.now()
    if 'lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['lmp'] = (today - timedelta(days=extracted['ga_lmp_days'])).strftime("%d-%m-%Y")
    if 'edd_lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['edd_lmp'] = (today + timedelta(days=(280 - extracted['ga_lmp_days']))).strftime("%d-%m-%Y")
    if 'edd_cua' not in extracted and 'ga_usg_days' in extracted:
        extracted['edd_cua'] = (today + timedelta(days=(280 - extracted['ga_usg_days']))).strftime("%d-%m-%Y")

    # --- PATIENT NAME / AGE / SEX ---
    if hasattr(ds, 'PatientName'):
        # Convert DICOM string to standard string and replace ^ with spaces
        raw_name = str(ds.PatientName).replace('^', ' ').strip()
        
        # Debug print: We can see exactly what the machine sends
        print(f"\n[DEBUG] RAW PATIENT NAME FROM MACHINE: '{raw_name}'")
        
        extracted['patient_name'] = raw_name
        
        # 1. Try to find the Age (any 1-3 digit number)
        # This looks for "28", "28Y", "28 Y", "28Y/F", "28 Y / F", "28 Years"
        age_match = re.search(r'(\d{1,3})', raw_name)
        if age_match:
            extracted['patient_age'] = age_match.group(1)
            # Everything BEFORE the age is the name
            extracted['patient_name'] = raw_name[:age_match.start()].strip()
        
        # 2. Try to find the Sex (F or M)
        sex_match = re.search(r'\d\s*Y\w*\s*/?\s*([FM])\b', raw_name, re.I)
        if sex_match:
            extracted['p_sex'] = sex_match.group(1).upper()
        elif hasattr(ds, 'PatientSex') and ds.PatientSex in ('F', 'M'):
            extracted['p_sex'] = ds.PatientSex

    # Fallback to standard DICOM tags if the string parsing failed
    if 'patient_age' not in extracted and getattr(ds, 'PatientAge', ''):
        extracted['patient_age'] = str(ds.PatientAge).lstrip('0').replace('Y', '').strip()
        
    if 'p_sex' not in extracted and getattr(ds, 'PatientSex', '') in ('F', 'M'):
        extracted['p_sex'] = ds.PatientSex

    return extracted

# --- C-STORE HANDLER ---
def handle_store(event):
    try:
        ds = event.dataset
        ds.file_meta = event.file_meta
        
        # Check if it's an SR or an Image
        sop_class = str(ds.get('SOPClassUID', ''))
        is_sr = '1.2.840.10008.5.1.4.1.1.88' in sop_class
        
        if is_sr:
            filename = f"{STORAGE_DIR}/{ds.SOPInstanceUID}.dcm"
            ds.save_as(filename, write_like_original=False)
            print(f"\n[+] Received SR for Patient: {ds.PatientName}")
            try:
                parsed_data = extract_sr_data(ds)
                with open(JSON_FILE, 'w') as f:
                    json.dump(parsed_data, f)
                print(f"[+] Successfully parsed SR data and saved to {JSON_FILE}")
            except Exception as e:
                print(f"[-] Error parsing SR data: {e}")
        else:
            # If it's an image, we accept it silently so the machine doesn't fail
            print(f"\n[*] Received Image ({ds.get('Modality', 'US')}) - Ignored for report.")
            
        return 0x0000
    except Exception as e:
        print(f"\n[!] Error handling file: {e}")
        return 0x0000 

# --- C-ECHO HANDLER ---
def handle_echo(event):
    print("\n[+] Verification (C-ECHO) request received from machine!")
    return 0x0000

# --- START SERVER ---
def main():
    ae = AE(ae_title=AE_TITLE)
    
    # Accept ALL possible formats (Images, SRs, Privates)
    ae.supported_contexts = AllStoragePresentationContexts
    ae.add_supported_context(Verification)
    
    handlers = [
        (evt.EVT_C_STORE, handle_store),
        (evt.EVT_C_ECHO, handle_echo)
    ]
    
    print(f"Starting DICOM Listener on Port {PORT} with AE Title '{AE_TITLE}'...")
    print("Waiting for the ultrasound machine...")
    print("Leave this window open while you are working.")
    
    ae.start_server(('0.0.0.0', PORT), evt_handlers=handlers)

if __name__ == '__main__':
    main()
