import os
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

# --- PARSER ---
def extract_sr_data(ds):
    from datetime import datetime, timedelta
    extracted = {}
    
    # We will use a flat list to grab every concept in the tree first, 
    # then parse them in order.
    flat_items = []
    
    def walk_tree(seq, parent_site=""):
        for item in seq:
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            new_site = parent_site
            # Catch any site identifier used by the machine
            if concept in ["Umbilical Artery", "Uterine Artery", "Fetal Doppler", "Maternal Doppler", "MCA"]:
                new_site = concept
            
            # Store the concept, the value sequence, and the parent site
            flat_items.append((concept, item, new_site))
            
            if 'ContentSequence' in item:
                walk_tree(item.ContentSequence, new_site)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)

    # Now we parse the flat list
    current_param = None
    for concept, item, site in flat_items:
        
        # Keep track of which biometry parameter we are under
        if concept in ['Biparietal Diameter', 'Head Circumference', 'Abdominal Circumference', 'Femur Length', 'Humerus length']:
            param_map = {
                'Biparietal Diameter': 'bpd', 'Head Circumference': 'hc',
                'Abdominal Circumference': 'ac', 'Femur Length': 'fl', 'Humerus length': 'hl'
            }
            current_param = param_map[concept]

        # --- TEXT VALUES ---
        if item.ValueType == 'TEXT':
            val = item.get('TextValue', '')
            if val: 
                if concept == 'EDD from average ultrasound age': extracted['edd_cua'] = val
                if concept == 'EDD from LMP': extracted['edd_lmp'] = val
                if concept == 'LMP': extracted['lmp'] = val

        # --- NUMERIC VALUES ---
        if item.ValueType == 'NUM' and 'MeasuredValueSequence' in item:
            val_seq = item.MeasuredValueSequence[0]
            num_val = float(val_seq.get('NumericValue', 0))
            unit = val_seq.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')
            
            # Biometry Sizes (in cm, convert to mm)
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
                extracted['nasal_bone'] = "Present" if num_val > 0 else "Absent"
            
            # Doppler (Handles Umbilical, Uterine, MCA properly)
            if concept == 'Pulsatility Index':
                if "Umbilical" in site or "Fetal" in site:
                    extracted['umbilical_pi'] = str(round(num_val, 2))
                elif "Uterine" in site or "Maternal" in site:
                    extracted['uterine_pi'] = str(round(num_val, 2))
                elif "MCA" in site:
                    extracted['mca_pi'] = str(round(num_val, 2))

            # Summary & Ratios
            if concept == 'Estimated Weight': 
                extracted['fetal_weight'] = str(int(num_val))
            if concept == 'EFW percentile rank':
                extracted['fetal_weight_pct'] = str(round(num_val, 1))
            
            if concept == 'HC/AC': extracted['hc_ac'] = str(round(num_val, 2))
            if concept == 'FL/AC': extracted['fl_ac'] = str(round(num_val, 1))
            if concept == 'FL/BPD': extracted['fl_bpd'] = str(round(num_val, 1))
            if concept == 'FL/HC': extracted['fl_hc'] = str(round(num_val, 1))

            if concept == 'Fetal Heart Rate': 
                extracted['fhr'] = str(int(num_val))
            
            if concept == 'Composite Ultrasound Age': 
                extracted['ga_usg_days'] = int(num_val)
                extracted['mean_age'] = f"{int(num_val)//7}weeks {int(num_val)%7}days ± 2w"
            
            if concept == 'Gestational Age by LMP':
                extracted['ga_lmp_days'] = int(num_val)

            # Biometry Percentiles & GA
            if concept == 'Growth Percentile Rank':
                if current_param:
                    extracted[f"{current_param}_pct"] = str(round(num_val, 1))

            if concept == 'Gestational Age':
                if current_param:
                    extracted[f"{current_param}_wks"] = str(int(num_val) // 7)
                    extracted[f"{current_param}_dys"] = str(int(num_val) % 7)

        # Reset current_param when we hit a new group
        if concept in ['Biometry Group', 'Fetal Long Bones']:
            current_param = None
    
    # Date calculations
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
