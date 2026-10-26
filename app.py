import streamlit as st
import html
from datetime import datetime

from shared.fields import FIELDS, TEMPLATES, FIELD_BY_ID
from shared.templates import TEMPLATE_BLOCKS, BLOCKS

# ============================================================
# TEXT GENERATOR
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
# STREAMLIT UI & CSS
# ============================================================

st.set_page_config(page_title="डॉ. प्रणव एवं डॉ. गौरव इमेजिंग, आगरा", layout="wide")

# EXACT CSS FROM RADIOLOGY APP
st.markdown(
    """
    <style>
    :root, html, body {
        color-scheme: light !important;
        background-color: #f4f6f9 !important;
    }
    #MainMenu {visibility: hidden;}
    header[data-testid="stHeader"] {visibility: hidden; height: 0;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden;}
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > .main,
    section.main,
    .main .block-container {
        background-color: #f4f6f9 !important;
    }
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
    .stApp li, .stApp small, .stApp strong, .stApp em {
        color: #111111 !important;
    }
    .pg-title {
        text-align: center; margin: 0 0 0.15rem 0; padding: 0;
        color: #1F4E79 !important; font-weight: 700; font-size: 2rem;
        letter-spacing: 0.5px;
    }
    .pg-tagline {
        text-align: center; margin: 0 0 1.0rem 0; padding: 0;
        color: #555 !important; font-style: italic; font-size: 1.05rem;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        display: flex !important; flex-direction: column !important;
        justify-content: flex-end !important;
    }
    div[data-testid="stRadio"] label,
    div[data-testid="stCheckbox"] label { cursor: pointer !important; }
    div[data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #cfd6dd !important;
        border-radius: 6px !important;
        margin-bottom: 4px !important;
    }
    div[data-testid="stExpander"] details,
    div[data-testid="stExpander"] details > div {
        background-color: #ffffff !important;
    }
    div[data-testid="stExpander"] details > summary {
        padding: 8px 12px !important; min-height: 42px !important;
        box-sizing: border-box !important; display: flex !important;
        align-items: center !important;
        background-color: #ffffff !important;
    }
    div[data-testid="stExpander"] details[open] > summary {
        background-color: #eef2f7 !important;
        border-bottom: 1px solid #cfd6dd !important;
    }
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] summary * {
        color: #111111 !important; font-weight: 600 !important;
        font-size: 14px !important;
        white-space: nowrap !important; overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    .stTextInput input,
    .stTextArea textarea,
    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea,
    [data-baseweb="input"] input,
    [data-baseweb="base-input"] input,
    [data-baseweb="textarea"] textarea,
    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #111111 !important;
        border: 1px solid #cfd6dd !important;
        min-height: 42px !important;
        box-sizing: border-box !important;
    }
    .stButton button {
        background-color: #ffffff !important;
        color: #111111 !important;
        border: 1px solid #cfd6dd !important;
    }
    .stButton button:hover {
        background-color: #eef2f7 !important;
        border-color: #305496 !important;
    }
    .preview-box,
    .stApp .preview-box,
    .stApp .preview-box * {
        color: #ffffff !important;
    }
    .preview-box {
        background-color: #2a2a2a !important;
        font-family: Consolas, Menlo, 'Courier New', monospace !important;
        font-size: 12.5px !important;
        line-height: 1.5 !important;
        padding: 14px 16px !important;
        border-radius: 8px !important;
        border: 1px solid #333 !important;
        white-space: pre-wrap !important;
        word-wrap: break-word !important;
        overflow-wrap: anywhere !important;
        word-break: break-word !important;
        margin-bottom: 8px !important;
        display: block !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# HEADER
st.markdown(
    """
    <div class="pg-title">डॉ. प्रणव एवं डॉ. गौरव इमेजिंग, आगरा</div>
    <div class="pg-tagline">Precision Imaging. Trusted Diagnostics. Agra.</div>
    """,
    unsafe_allow_html=True,
)

# TOP PATIENT ROW
c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3], vertical_alignment="bottom")
with c1:
    p_name = st.text_input("Name", key="p_name")
with c2:
    p_age = st.text_input("Age (years)", key="p_age")
with c3:
    p_sex = st.radio("Sex", ["F", "M"], horizontal=True, key="p_sex")
with c4:
    p_date = st.text_input("Date", value=datetime.now().strftime("%d-%b-%y"), key="p_date")
with c5:
    p_ref = st.text_input("Referred by", key="p_ref")

st.markdown("---")

# Setup Layout: Left (Inputs) and Right (Preview)
col_input, col_preview = st.columns([1, 1], gap="large")

# 1. Template Selector (Top of left column)
with col_input:
    template_id = st.selectbox(
        "⚡ Load a Clinical Template",
        options=list(TEMPLATES.keys()),
        format_func=lambda k: TEMPLATES[k],
    )
    st.caption("Select a template to fill typical findings.")

# Initialize data dictionary to store our values
data = {
    "patient_name": p_name,
    "patient_age": p_age,
    "report_date": p_date,
    "referred_by": p_ref,
}

# 2. Build the Input Form on the Left
with col_input:
    st.subheader("Findings")
    
    # LMP is usually its own block
    data["lmp"] = st.text_input("LMP", key="p_lmp")
    
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
        label="⬇️ Download .docx",
        data=report_text,
        file_name=f"OB_Report_{p_name or 'Unknown'}.txt",
        mime="text/plain"
    )
