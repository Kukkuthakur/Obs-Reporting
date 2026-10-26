import streamlit as st
import html
from datetime import datetime

from shared.fields import FIELDS, TEMPLATES, FIELD_BY_ID
from shared.templates import TEMPLATE_BLOCKS, BLOCKS

# ============================================================
# TEXT GENERATOR (MIRRORS YOUR RADIOLOGY APP)
# ============================================================

def generate_ob_report(template_id, data):
    """Takes the data dictionary and returns formatted paragraphs."""
    out = []
    
    # 1. Header
    if data.get("patient_name"):
        out.append(f"NAME :- {data['patient_name']}      DATE :- {data['report_date']}")
    if data.get("lmp"):
        out.append(f"LMP :- {data['lmp']}")
    out.append("")
    
    # 2. Title based on template
    if template_id == "EARLY_OBS":
        out.append("EARLY OBS SCAN")
    elif template_id == "LEVEL2":
        out.append("USG OBS (ANOMALY SCAN)")
    else:
        out.append("ULTRASOUND OBSTETRICS REPORT")
    out.append("")

    # 3. Early OBS specific
    if template_id == "EARLY_OBS":
        gs = data.get("gsac_status", "")
        msd = data.get("msd", "")
        crl = data.get("crl", "")
        yolk = data.get("yolk_sac", "")
        fhr = data.get("cardiac_activity", "")
        
        if gs == "empty":
            out.append(f"G-Sac: EMPTY, MSD={msd}mm")
            out.append("CRL: EMBRYO NOT SEEN")
            out.append(f"Yolk sac: {yolk}")
            out.append(f"Cardiac pulsations: {fhr}")
        else:
            if crl:
                out.append(f"CRL: {crl}mm")
            if fhr and fhr != "N/A":
                out.append(f"Cardiac pulsations: {fhr} bpm")
        
        out.append("")
        return "\n".join(out)

    # 4. Intro (Normal OB templates)
    presentation = data.get("presentation", "variable")
    lie = data.get("lie", "")
    lie_str = f" and {lie} lie" if lie else ""
    
    out.append(f"Uterus shows single live fetus. Fetus is in {presentation.upper()} "
               f"PRESENTATION{lie_str} at the time of scanning. Fetal movements are present. "
               f"Fetal cardiac pulsations are appreciated.")
    out.append("")
    
    # 5. Biometry Table
    out.append("BIO PHYSICAL PARAMETERS:")
    out.append(f"BPD: {data.get('bpd', '')}mm")
    out.append(f"HC: {data.get('hc', '')}mm")
    out.append(f"AC: {data.get('ac', '')}mm")
    out.append(f"FL: {data.get('fl', '')}mm")
    out.append(f"HL: {data.get('hl', '')}mm")
    out.append("")
    
    out.append(f"MEAN AGE: {data.get('mean_age', '')}")
    out.append(f"FETAL WEIGHT: {data.get('fetal_weight', '')} GRAMS")
    out.append(f"AFI: Adequate, largest pocket measures {data.get('afi', '')} cm.")
    out.append(f"FHR is approx. {data.get('fhr', '')} /min.")
    out.append("")
    
    # 6. Doppler
    if data.get("mean_uterine_pi"):
        out.append(f"Mean Uterine Artery PI Is {data['mean_uterine_pi']} (WNL, No Notching)")
    if data.get("right_uterine_pi"):
        out.append(f"Right Uterine Artery PI is {data['right_uterine_pi']} (WNL, No Notching)")
    if data.get("umbilical_pi"):
        out.append(f"Umbilical Artery PI is {data['umbilical_pi']} (WNL)")
    out.append("")
    
    # 7. Anatomy
    out.append(f"Placenta is {data.get('placenta_position', '').upper()} and away from the Os.")
    out.append(f"CERVIX: normal ({data.get('cervix_length', '')}cm), internal Os is {data.get('internal_os', '')}.")
    
    loop = data.get("cord_loop", "none seen")
    if loop == "none seen":
        out.append("NO LOOP OF CORD SEEN AROUND THE NECK.")
    else:
        out.append("LOOP OF CORD SEEN AROUND THE NECK.")
    out.append("")

    # 8. Scar thickness (T3 only)
    if template_id == "T3" and data.get("scar_thickness"):
         out.append(f"Scar thickness measures {data['scar_thickness']}mm.")
         out.append("")

    # 9. Impression
    out.append("IMPRESSION:")
    if template_id == "T3":
        out.append(f"Ultrasound obstetric reveals single live intrauterine fetus in cephalic presentation "
                   f"of approx. {data.get('mean_age', '')} gestation age.")
    else:
        out.append(f"Ultrasound obstetric reveals single live intrauterine fetus in variable presentation "
                   f"of approx. {data.get('mean_age', '')} gestation age.")

    return "\n".join(out)

# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="OB Reporting App", layout="wide")

st.markdown("""
    <style>
    .preview-box {
        background-color: #2a2a2a !important;
        color: #ffffff !important;
        font-family: Consolas, Menlo, 'Courier New', monospace !important;
        font-size: 13px !important;
        line-height: 1.6 !important;
        padding: 16px !important;
        border-radius: 8px !important;
        white-space: pre-wrap !important;
    }
    div[data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #cfd6dd !important;
        border-radius: 6px !important;
        margin-bottom: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("OB Reporting App")

# Setup Layout: Left (Inputs) and Right (Preview)
col_input, col_preview = st.columns([1, 1], gap="large")

# 1. Template Selector (Top of left column)
with col_input:
    template_id = st.selectbox(
        "Report Template",
        options=list(TEMPLATES.keys()),
        format_func=lambda k: TEMPLATES[k],
    )
    st.markdown("---")

# Initialize data dictionary to store our values
data = {}

# 2. Build the Input Form on the Left
with col_input:
    st.subheader("Patient Details")
    for field_id in BLOCKS["HEADER"] + BLOCKS["LMP"]:
        field = FIELD_BY_ID[field_id]
        if field["type"] == "date":
            data[field_id] = st.text_input(field["label"], value=datetime.now().strftime("%d-%b-%y"))
        else:
            data[field_id] = st.text_input(field["label"])

    st.subheader("Findings")
    # Loop through the blocks defined for this specific template
    for block_name in TEMPLATE_BLOCKS[template_id]:
        if block_name in ("HEADER", "LMP"):
            continue # Already handled above
        
        with st.expander(block_name.replace("_", " ").title(), expanded=True):
            cols = st.columns(2)
            for i, field_id in enumerate(BLOCKS[block_name]):
                field = FIELD_BY_ID.get(field_id)
                if not field:
                    continue
                
                with cols[i % 2]:
                    ftype = field["type"]
                    label = field["label"]
                    
                    # Use session state keys to persist values
                    key = f"{template_id}_{field_id}"
                    
                    if ftype == "enum":
                        choices = field.get("choices", [])
                        data[field_id] = st.selectbox(label, choices, key=key)
                    else:
                        data[field_id] = st.text_input(label, key=key)

# 3. Render the Live Preview on the Right
with col_preview:
    st.subheader("📄 Live Preview")
    
    # Generate the text based on current form state
    report_text = generate_ob_report(template_id, data)
    
    # Render to styled HTML box
    safe_html = html.escape(report_text).replace("\n", "<br>")
    st.markdown(f'<div class="preview-box">{safe_html}</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    # Add download button
    st.download_button(
        label="⬇️ Download Report (TXT)",
        data=report_text,
        file_name=f"OB_Report_{data.get('patient_name', 'Unknown')}.txt",
        mime="text/plain"
    )
