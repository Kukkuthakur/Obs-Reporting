import streamlit as st
import html
import hashlib
import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from shared.fields import FIELDS, TEMPLATES, FIELD_BY_ID
from shared.templates import TEMPLATE_BLOCKS, BLOCKS

# ============================================================
# TEXT GENERATORS
# ============================================================

def generate_findings(template_id, data):
    """Takes the data dictionary and returns formatted findings paragraphs."""
    out = []
    
    # 1. Title
    if template_id == "EARLY_OBS":
        out.append("EARLY OBS SCAN")
    elif template_id == "LEVEL2":
        out.append("USG OBS (ANOMALY SCAN)")
    else:
        out.append("ULTRASOUND OBSTETRICS REPORT")
    out.append("")

    # 2. Early OBS specific
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

    # 3. Intro (Normal OB templates)
    presentation = data.get("presentation", "variable")
    lie = data.get("lie", "")
    lie_str = f" and {lie} lie" if lie else ""
    
    out.append(f"Uterus shows single live fetus. Fetus is in {presentation.upper()} "
               f"PRESENTATION{lie_str} at the time of scanning. Fetal movements are present. "
               f"Fetal cardiac pulsations are appreciated.")
    out.append("")
    
    # 4. Biometry Table
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
    
    # 5. Doppler
    if data.get("mean_uterine_pi"):
        out.append(f"Mean Uterine Artery PI Is {data['mean_uterine_pi']} (WNL, No Notching)")
    if data.get("right_uterine_pi"):
        out.append(f"Right Uterine Artery PI is {data['right_uterine_pi']} (WNL, No Notching)")
    if data.get("umbilical_pi"):
        out.append(f"Umbilical Artery PI is {data['umbilical_pi']} (WNL)")
    out.append("")
    
    # 6. Anatomy
    out.append(f"Placenta is {data.get('placenta_position', '').upper()} and away from the Os.")
    out.append(f"CERVIX: normal ({data.get('cervix_length', '')}cm), internal Os is {data.get('internal_os', '')}.")
    
    loop = data.get("cord_loop", "none seen")
    if loop == "none seen":
        out.append("NO LOOP OF CORD SEEN AROUND THE NECK.")
    else:
        out.append("LOOP OF CORD SEEN AROUND THE NECK.")
    out.append("")

    # 7. Scar thickness (T3 only)
    if template_id == "T3" and data.get("scar_thickness"):
         out.append(f"Scar thickness measures {data['scar_thickness']}mm.")
         out.append("")

    return "\n".join(out)

def generate_impression(template_id, data):
    if template_id == "EARLY_OBS":
        return "Intra uterine live pregnancy of 10wks 01days +/-1w with good cardiac activity."
    
    age = data.get('mean_age', '')
    if template_id == "T3":
        return f"Ultrasound obstetric reveals single live intrauterine fetus in cephalic presentation of approx. {age} gestation age."
    else:
        return f"Ultrasound obstetric reveals single live intrauterine fetus in variable presentation of approx. {age} gestation age."

# ============================================================
# DOCX BUILDER (PORTED FROM RADIOLOGY APP)
# ============================================================

def _set_table_borders(table, size=6):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(size))
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)

def _set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        m = OxmlElement(f"w:{edge}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tcPr.append(mar)

def _add_run(paragraph, text, bold=False, underline=False, font="Calibri", size=11, italic=False, color=None):
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.underline = underline
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rPr.insert(0, rFonts)
    return run

def build_docx_bytes(data, findings_text, impression_text):
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(3.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    # Spacer
    for _ in range(2):
        p = doc.add_paragraph()
        p.add_run("").font.size = Pt(10)

    # 1. Patient Demographic Table
    table = doc.add_table(rows=2, cols=2)
    table.autofit = True
    _set_table_borders(table)
    
    cells = [
        (0, 0, f"NAME :- {data['patient_name']}"),
        (0, 1, f"DATE :- {data['report_date']}"),
        (1, 0, f"AGE/SEX :- {data['patient_age']}Y/{data['p_sex']}"),
        (1, 1, f"REF. BY: {data['referred_by']}"),
    ]
    for r, c, text in cells:
        cell = table.cell(r, c)
        cell.text = ""
        cp = cell.paragraphs[0]
        cp.paragraph_format.line_spacing = 1.0
        _add_run(cp, text, bold=True)

    p = doc.add_paragraph()
    p.add_run("").font.size = Pt(8)

    # 2. LMP
    if data.get("lmp"):
        p = doc.add_paragraph()
        _add_run(p, f"LMP :- {data['lmp']}", bold=True)
        p.add_run("").font.size = Pt(8)

    # 3. Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(title_para, "ULTRASOUND OBSTETRICS REPORT", bold=True, underline=True, size=13, color=RGBColor(0x1F, 0x4E, 0x79))
    doc.add_paragraph().add_run("").font.size = Pt(8)

    # 4. Findings (from text area)
    for line in findings_text.split("\n"):
        if not line.strip():
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _add_run(p, line)

    doc.add_paragraph().add_run("").font.size = Pt(10)

    # 5. Impression
    imp_head = doc.add_paragraph()
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True, size=12, color=RGBColor(0x1F, 0x4E, 0x79))
    
    for line in impression_text.split("\n"):
        if not line.strip():
            continue
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _add_run(p, line, bold=True, size=11)

    doc.add_paragraph().add_run("").font.size = Pt(10)

    # 6. Hardcoded Disclaimer Table
    disc_table = doc.add_table(rows=1, cols=1)
    _set_table_borders(disc_table)
    cell = disc_table.cell(0, 0)
    _set_cell_margins(cell, top=10, bottom=10, left=80, right=80)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = Pt(7)
    
    disc_text = (
        '"The radiologist has not conducted or caused to conduct any pre natal '
        'diagnostic technique (sonography) for the purpose of determining the sex '
        'of the foetus. No person including the person conducting the diagnostic '
        'procedure has communicated to the pregnant woman concerned or her '
        'relatives or any other person the sex of the foetus by words, signs, or '
        'in any other manner."'
    )
    _add_run(p, disc_text, bold=True, size=7)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ============================================================
# STREAMLIT UI & CSS
# ============================================================

st.set_page_config(page_title="डॉ. प्रणव एवं डॉ. गौरव इमेजिंग, आगरा", layout="wide")

# ... [CSS remains exactly the same as previous message] ...
st.markdown(
    """
    <style>
    :root, html, body { color-scheme: light !important; background-color: #f4f6f9 !important; }
    #MainMenu {visibility: hidden;} header[data-testid="stHeader"] {visibility: hidden; height: 0;}
    footer {visibility: hidden;} [data-testid="stToolbar"] {visibility: hidden;}
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > .main, section.main, .main .block-container { background-color: #f4f6f9 !important; }
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp li, .stApp small, .stApp strong, .stApp em { color: #111111 !important; }
    .pg-title { text-align: center; margin: 0 0 0.15rem 0; padding: 0; color: #1F4E79 !important; font-weight: 700; font-size: 2rem; letter-spacing: 0.5px; }
    .pg-tagline { text-align: center; margin: 0 0 1.0rem 0; padding: 0; color: #555 !important; font-style: italic; font-size: 1.05rem; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] { display: flex !important; flex-direction: column !important; justify-content: flex-end !important; }
    div[data-testid="stRadio"] label, div[data-testid="stCheckbox"] label { cursor: pointer !important; }
    div[data-testid="stExpander"] { background-color: #ffffff !important; border: 1px solid #cfd6dd !important; border-radius: 6px !important; margin-bottom: 4px !important; }
    div[data-testid="stExpander"] details, div[data-testid="stExpander"] details > div { background-color: #ffffff !important; }
    div[data-testid="stExpander"] details > summary { padding: 8px 12px !important; min-height: 42px !important; box-sizing: border-box !important; display: flex !important; align-items: center !important; background-color: #ffffff !important; }
    div[data-testid="stExpander"] details[open] > summary { background-color: #eef2f7 !important; border-bottom: 1px solid #cfd6dd !important; }
    div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary * { color: #111111 !important; font-weight: 600 !important; font-size: 14px !important; white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; }
    .stTextInput input, .stTextArea textarea, div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea, [data-baseweb="input"] input, [data-baseweb="base-input"] input, [data-baseweb="textarea"] textarea, div[data-baseweb="select"] > div { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; min-height: 42px !important; box-sizing: border-box !important; }
    .stButton button, .stDownloadButton button, div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; }
    .stButton button:hover, .stDownloadButton button:hover { background-color: #eef2f7 !important; border-color: #305496 !important; }
    .preview-box, .stApp .preview-box, .stApp .preview-box * { color: #ffffff !important; }
    .preview-box { background-color: #2a2a2a !important; font-family: Consolas, Menlo, 'Courier New', monospace !important; font-size: 12.5px !important; line-height: 1.5 !important; padding: 14px 16px !important; border-radius: 8px !important; border: 1px solid #333 !important; white-space: pre-wrap !important; word-wrap: break-word !important; overflow-wrap: anywhere !important; word-break: break-word !important; margin-bottom: 8px !important; display: block !important; }
    div[class*="st-key-impression_box"] textarea { background-color: #ffffff !important; color: #000000 !important; font-family: Consolas, Menlo, 'Courier New', monospace !important; font-size: 12.5px !important; line-height: 1.5 !important; border: 1px solid #333 !important; white-space: pre-wrap !important; word-wrap: break-word !important; overflow-wrap: anywhere !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="pg-title">डॉ. प्रणव एवं डॉ. गौरव इमेजिंग, आगरा</div><div class="pg-tagline">Precision Imaging. Trusted Diagnostics. Agra.</div>', unsafe_allow_html=True)

# TOP PATIENT ROW
c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3], vertical_alignment="bottom")
with c1: p_name = st.text_input("Name", key="p_name_input")
with c2: p_age = st.text_input("Age (years)", key="p_age_input")
with c3: p_sex = st.radio("Sex", ["F", "M"], horizontal=True, key="p_sex_input")
with c4: p_date = st.text_input("Date", value=datetime.now().strftime("%d-%b-%y"), key="p_date_input")
with c5: p_ref = st.text_input("Referred by", key="p_ref_input")

st.markdown("---")

col_input, col_preview = st.columns([1, 1], gap="large")

with col_input:
    template_id = st.selectbox("⚡ Load a Clinical Template", options=list(TEMPLATES.keys()), format_func=lambda k: TEMPLATES[k])
    st.caption("Select a template to fill typical findings.")

data = {"patient_name": p_name, "patient_age": p_age, "report_date": p_date, "referred_by": p_ref, "p_sex": p_sex}

with col_input:
    st.subheader("Findings")
    data["lmp"] = st.text_input("LMP", key="p_lmp")
    
    for block_name in TEMPLATE_BLOCKS[template_id]:
        if block_name in ("HEADER", "LMP"): continue
        with st.expander(block_name.replace("_", " ").title(), expanded=True):
            cols = st.columns(2)
            for i, field_id in enumerate(BLOCKS[block_name]):
                field = FIELD_BY_ID.get(field_id)
                if not field: continue
                with cols[i % 2]:
                    ftype = field["type"]
                    label = field["label"]
                    key = f"{template_id}_{field_id}"
                    if ftype == "enum":
                        data[field_id] = st.selectbox(label, field.get("choices", []), key=key)
                    else:
                        data[field_id] = st.text_input(label, key=key)

with col_preview:
    st.subheader("📄 Live Preview")
    findings_text = generate_findings(template_id, data)
    safe_html = html.escape(findings_text).replace("\n", "<br>")
    st.markdown(f'<div class="preview-box">{safe_html}</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Editable Impression Box
    st.subheader("✏️ Impression (editable)")
    auto_imp_str = generate_impression(template_id, data)
    new_auto_hash = hashlib.md5(auto_imp_str.encode("utf-8")).hexdigest()

    if "_imp_ver" not in st.session_state: st.session_state["_imp_ver"] = 0
    if "_imp_auto_hash" not in st.session_state: st.session_state["_imp_auto_hash"] = new_auto_hash
    if "_imp_widget_current" not in st.session_state: st.session_state["_imp_widget_current"] = auto_imp_str

    _user_text_key = f"impression_box_v{st.session_state['_imp_ver']}"
    _user_text = st.session_state.get(_user_text_key, None)

    if (new_auto_hash != st.session_state["_imp_auto_hash"] and (_user_text is None or _user_text == st.session_state["_imp_widget_current"])):
        st.session_state["_imp_ver"] += 1
        st.session_state["_imp_auto_hash"] = new_auto_hash
        st.session_state["_imp_widget_current"] = auto_imp_str
        _user_text_key = f"impression_box_v{st.session_state['_imp_ver']}"

    if _user_text_key not in st.session_state: st.session_state[_user_text_key] = auto_imp_str

    edited_imp = st.text_area("Impression lines", height=150, label_visibility="collapsed", key=_user_text_key)

    st.markdown("---")
    
    # Download Button (Generates DOCX)
    docx_bytes = build_docx_bytes(data, findings_text, edited_imp)
    
    st.download_button(
        label="⬇️ Download .docx",
        data=docx_bytes,
        file_name=f"OB_Report_{p_name or 'Unknown'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
