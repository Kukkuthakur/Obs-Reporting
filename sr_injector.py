import streamlit as st
import pydicom
import io

st.set_page_config(page_title="SR Injector Test", layout="wide")
st.title("🔬 DICOM SR Data Extractor")
st.write("Upload a Samsung `.dcm` SR file to extract measurements.")

uploaded_file = st.file_uploader("Choose a DICOM file", type=["dcm", "dicom", "ima", ""])

def extract_sr_data(ds):
    """Walks the SR Content Tree and extracts known measurements."""
    extracted = {}
    
    # We will walk the tree and look for specific Concept Names
    def walk_tree(seq, parent_concept=""):
        for item in seq:
            # Get the concept name (what the measurement is)
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            # Check for NUM (Numeric) values
            if item.ValueType == 'NUM' and 'MeasuredValueSequence' in item:
                val_seq = item.MeasuredValueSequence[0]
                num_val = float(val_seq.get('NumericValue', 0))
                unit = val_seq.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')
                
                # 1. BIOMETRY (Convert cm to mm)
                if unit == 'centimeter':
                    if concept == 'Biparietal Diameter': extracted['bpd'] = f"{int(num_val * 10)}"
                    elif concept == 'Head Circumference': extracted['hc'] = f"{int(num_val * 10)}"
                    elif concept == 'Abdominal Circumference': extracted['ac'] = f"{int(num_val * 10)}"
                    elif concept == 'Femur Length': extracted['fl'] = f"{int(num_val * 10)}"
                    elif concept == 'Humerus length': extracted['hl'] = f"{int(num_val * 10)}"
                    elif concept == 'Gestational Sac Diameter': extracted['msd'] = f"{int(num_val * 10)}"
                
                # 2. DOPPLER (Keep as PI/RI values)
                if 'Umbilical Artery' in parent_concept:
                    if concept == 'Pulsatility Index': extracted['umbilical_pi'] = str(num_val)
                    elif concept == 'Resistivity Index': extracted['umbilical_ri'] = str(num_val)
                    elif concept == 'Systolic to Diastolic Velocity Ratio': extracted['umbilical_sd'] = str(num_val)
                
                if 'Uterine Artery' in parent_concept:
                    if concept == 'Pulsatility Index': extracted['uterine_pi'] = str(num_val)
                    elif concept == 'Resistivity Index': extracted['uterine_ri'] = str(num_val)
                
                # 3. GENERAL FETAL METRICS
                if concept == 'Estimated Weight': extracted['fetal_weight'] = str(int(num_val))
                if concept == 'Fetal Heart Rate': extracted['fhr'] = str(int(num_val))
                if concept == 'Composite Ultrasound Age': extracted['ga_days'] = str(int(num_val))
                if concept == 'Gestational Age by LMP': extracted['ga_lmp_days'] = str(int(num_val))

            # Check for TEXT values (sometimes dates are stored here)
            if item.ValueType == 'TEXT':
                val_text = item.get('TextValue', '')
                if 'EDD from average ultrasound age' == concept: extracted['edd_cua'] = val_text
                if 'EDD from LMP' == concept: extracted['edd_lmp'] = val_text

            # Recursively check child sequences, updating the parent_concept
            if 'ContentSequence' in item:
                # If this item is a 'Findings' or 'Biometry Group', we don't change the parent name,
                # but if it's a vessel (e.g., Umbilical Artery), we pass it down.
                new_parent = concept if concept in ["Umbilical Artery", "Uterine Artery"] else parent_concept
                walk_tree(item.ContentSequence, new_parent)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)
    
    return extracted

if uploaded_file is not None:
    try:
        ds = pydicom.dcmread(io.BytesIO(uploaded_file.read()))
        st.success("File loaded successfully!")
        
        # Display basic info
        st.subheader("Patient Info")
        st.write(f"**Name:** {ds.get('PatientName', 'Unknown')}")
        st.write(f"**Modality:** {ds.get('Modality', 'Unknown')}")
        
        # Extract and display
        st.subheader("Extracted Measurements")
        extracted_data = extract_sr_data(ds)
        
        if extracted_data:
            st.json(extracted_data)
            
            # Show a preview of how it will look in the app
            st.subheader("Mapped to App Fields")
            col1, col2, col3 = st.columns(3)
            with col1:
                if 'bpd' in extracted_data: st.info(f"BPD: {extracted_data['bpd']} mm")
                if 'hc' in extracted_data: st.info(f"HC: {extracted_data['hc']} mm")
                if 'ac' in extracted_data: st.info(f"AC: {extracted_data['ac']} mm")
                if 'fl' in extracted_data: st.info(f"FL: {extracted_data['fl']} mm")
                if 'hl' in extracted_data: st.info(f"HL: {extracted_data['hl']} mm")
                if 'msd' in extracted_data: st.info(f"MSD: {extracted_data['msd']} mm")
            with col2:
                if 'fetal_weight' in extracted_data: st.success(f"Weight: {extracted_data['fetal_weight']} g")
                if 'fhr' in extracted_data: st.success(f"FHR: {extracted_data['fhr']} bpm")
                if 'ga_days' in extracted_data: 
                    weeks = int(extracted_data['ga_days']) // 7
                    days = int(extracted_data['ga_days']) % 7
                    st.success(f"Mean Age: {weeks} weeks {days} days")
            with col3:
                if 'umbilical_pi' in extracted_data: st.warning(f"Umbilical PI: {extracted_data['umbilical_pi']}")
                if 'uterine_pi' in extracted_data: st.warning(f"Uterine PI: {extracted_data['uterine_pi']}")
        else:
            st.warning("No measurements found. Check if this is a valid SR file.")
            
    except Exception as e:
        st.error(f"Error reading file: {e}")
