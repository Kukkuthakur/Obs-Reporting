import streamlit as st
import pydicom
import io

st.set_page_config(page_title="DICOM Inspector", layout="wide")
st.title("🔍 DICOM SR Inspector")
st.write("Upload one of the old `.dcm` files to see its structure.")

uploaded_file = st.file_uploader("Choose a DICOM file", type=["dcm", "dicom", "ima", ""])

if uploaded_file is not None:
    try:
        # Read the file from the uploaded bytes
        ds = pydicom.dcmread(io.BytesIO(uploaded_file.read()))
        
        st.success("File loaded successfully!")
        
        # 1. Basic Info
        st.subheader("Basic Info")
        st.write(f"**Patient Name:** {ds.get('PatientName', 'Unknown')}")
        st.write(f"**Modality:** {ds.get('Modality', 'Unknown')}")
        st.write(f"**SOP Class:** {ds.get('SOPClassUID', 'Unknown')}")
        
        # 2. Look for SR Content Tree
        st.subheader("SR Content Tree (Measurements)")
        if 'ContentSequence' in ds:
            st.write("Found a Structured Report content tree!")
            
            # A function to walk the tree and format it nicely for Streamlit
            def walk_tree(seq, depth=0):
                for item in seq:
                    concept = "Unknown"
                    if 'ConceptNameCodeSequence' in item:
                        concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', 'Unknown')
                    
                    val = "No Value"
                    if 'ValueType' in item:
                        if item.ValueType == 'NUM':
                            if 'MeasuredValueSequence' in item:
                                mvs = item.MeasuredValueSequence[0]
                                val = f"{mvs.get('NumericValue', '')} {mvs.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')}"
                        elif item.ValueType == 'TEXT':
                            val = item.get('TextValue', '')
                    
                    st.text("  " * depth + f"[{item.get('RelationshipType', '')}] {concept} = {val}")
                    
                    if 'ContentSequence' in item:
                        walk_tree(item.ContentSequence, depth + 1)
            
            walk_tree(ds.ContentSequence)
        else:
            st.warning("No ContentSequence found. Checking for Private Tags (measurements in images)...")
            has_tags = False
            for elem in ds:
                if elem.VR in ('SQ', 'UN') and elem.tag.group != 0x0002 and (elem.tag.is_private or 'Meas' in elem.name):
                    has_tags = True
                    st.text(f"Tag: {elem.tag} | Name: {elem.name} | VR: {elem.VR}")
                    if elem.VR == 'SQ':
                        for item in elem.value:
                            for sub_elem in item:
                                if isinstance(sub_elem.value, (str, int, float)):
                                    st.text(f"   -> {sub_elem.name}: {sub_elem.value}")
            if not has_tags:
                st.error("No measurements or private tags found.")
                
    except Exception as e:
        st.error(f"Error reading file: {e}")
