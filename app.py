import streamlit as st
import html
import hashlib
import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# SHARED CONFIGURATION (Simulating your shared/ files)
# ============================================================
TEMPLATES = {
    "EARLY_OBS": "Early OBS Scan (6-10 Weeks)",
    "NT_NB": "Early OBS NT/NB Scan (11-14 Weeks)",
    "T16": "OB Report (16 Weeks)",
    "T20": "OB Report (20 Weeks)",
    "T22": "OB Report (22 Weeks)",
    "T26": "OB Report (26 Weeks)",
    "T28": "OB Report (28 Weeks)",
    "T32": "OB Report (32 Weeks)",
    "T36": "OB Report (36 Weeks)",
    "DOPPLER": "OB Doppler",
    "LEVEL2": "Level II / Anomaly Scan"
}

# For this script, we will hardcode the fields to demonstrate the new logic
# In your real app, this comes from shared/fields.py
def get_template_blocks(template_id):
    if template_id in ["EARLY_OBS", "NT_NB"]:
        return ["HEADER", "LMP", "EARLY_PARAMS", "EDD", "OPINION"]
    elif template_id == "LEVEL2":
        return ["HEADER", "LMP", "EGA", "FETAL_SURVEY", "CRANIUM", "SPINE_FACE", "THORAX_ABDOMEN", "LIMBS", "INDICES", "DOPPLER", "OPINION"]
    elif template_id == "DOPPLER":
        return ["HEADER", "LMP", "EGA", "BIO_PHYSICAL", "DOPPLER_INDICES", "OPINION"]
    else:
        return ["HEADER", "LMP", "BIO_PHYSICAL", "ANATOMY", "OPINION"]

# ============================================================
# TEXT GENERATORS (Findings & Impression)
# ============================================================
def generate_findings_text(template_id, data):
    """Generates the plain text for the live preview box."""
    lines = []
    
    if template_id == "EARLY_OBS":
        lines.append(f"G-Sac: {data.get('gsac', '')}")
        lines.append(f"CRL: {data.get('crl', '')}")
        lines.append(f"Yolk sac: {data.get('yolk', '')}")
        lines.append(f"Cardiac pulsations: {data.get('fhr', '')}")
    elif template_id in ["T16", "T20", "T22", "T26", "T28", "T32", "T36"]:
        lines.append(f"Uterus shows single live fetus. Fetus is in {data.get('presentation', 'VARIABLE')} PRESENTATION at the time of scanning.")
        lines.append("")
        lines.append("BIO PHYSICAL PARAMETERS:")
        lines.append(f"BPD: {data.get('bpd_size', '')} | {data.get('bpd_wks', '')} {data.get('bpd_dys', '')}")
        lines.append(f"HC: {data.get('hc_size', '')} | {data.get('hc_wks', '')} {data.get('hc_dys', '')}")
        lines.append(f"AC: {data.get('ac_size', '')} | {data.get('ac_wks', '')} {data.get('ac_dys', '')}")
        lines.append(f"FL: {data.get('fl_size', '')} | {data.get('fl_wks', '')} {data.get('fl_dys', '')}")
        lines.append(f"HL: {data.get('hl_size', '')} | {data.get('hl_wks', '')} {data.get('hl_dys', '')}")
    
    return "\n".join(lines)

def generate_impression(template_id, data):
    if template_id == "EARLY_OBS":
        return f"Intra uterine live pregnancy of {data.get('crl', '')} +/-1w with good cardiac activity."
    elif template_id == "T16":
        return f"Ultrasound obstetric reveals single live intrauterine fetus in variable presentation of approx. {data.get('mean_age', '')} gestation age."
    else:
        return f"Ultrasound obstetric reveals single live intrauterine fetus in {data.get('presentation', 'cephalic')} presentation of approx. {data.get('mean_age', '')} gestation age."

# ============================================================
# DOCX BUILDER (The core engine)
# ============================================================

def _set_cell_border(cell, **kwargs):
    """Allows us to remove borders on tables to make them look clean."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def _add_run(paragraph, text, bold=False, size=11, italic=False, color=None, underline=False):
    run = paragraph.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    if color is not None:
        run.font.color.rgb = color
    return run

def _create_docx_table(doc, data_rows, headers=None, col_widths=None):
    """Creates a clean table, either with or without borders."""
    if not data_rows: return
    
    cols = len(data_rows[0])
    table = doc.add_table(rows=len(data_rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    # Optional header
    row_offset = 0
    if headers:
        for j, h in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = ""
            p = cell.paragraphs[0]
            _add_run(p, h, bold=True, size=10)
            # Header border
            _set_cell_border(cell, bottom={"sz": 12, "val": "single", "color": "000000"})
        row_offset = 1

    for i, row in enumerate(data_rows):
        for j, text in enumerate(row):
            cell = table.cell(i + row_offset, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            # Make the first column bold (like BPD or G-Sac)
            is_bold = (j == 0)
            _add_run(p, str(text), bold=is_bold, size=10)

    # Set widths if provided (e.g., [3cm, 2cm, 2cm, 2cm])
    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = width
                
    return table

def build_docx_bytes(data, template_id):
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # --- 1. Header Table (Demographics) ---
    header_table = doc.add_table(rows=2, cols=2)
    header_table.autofit = True
    cells = [
        (0, 0, f"NAME :- {data.get('patient_name', '')}"),
        (0, 1, f"DATE :- {data.get('report_date', '')}"),
        (1, 0, f"AGE/SEX :- {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}"),
        (1, 1, f"REF. BY:- {data.get('referred_by', '')}"),
    ]
    for r, c, text in cells:
        cell = header_table.cell(r, c)
        cell.text = ""
        _add_run(cell.paragraphs[0], text, bold=True, size=11)
        _set_cell_border(cell, bottom={"sz": 6, "val": "single", "color": "000000"})

    doc.add_paragraph()
    
    # --- 2. Title ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(title, TEMPLATES[template_id].upper(), bold=True, underline=True, size=12, color=RGBColor(0x1F, 0x4E, 0x79))
    
    if data.get("lmp"):
        p = doc.add_paragraph()
        _add_run(p, f"LMP- {data['lmp']}", bold=True, size=11)
    
    doc.add_paragraph()

    # --- 3. Early OBS Specific Tables ---
    if template_id == "EARLY_OBS":
        early_rows = [
            ["G-Sac", data.get('gsac', '')],
            ["CRL", data.get('crl', '')],
            ["Yolk sac", data.get('yolk', '')],
            ["Cardiac pulsations", data.get('fhr', '')],
            ["Decidual reaction", data.get('decidual', 'Present')],
            ["Subchorionic collection", data.get('subchorionic', 'Absent')],
            ["Corpus leuteal cyst", data.get('corpus', 'Not Appreciated')],
            ["Cervix", data.get('cervix', 'Normal')],
            ["Internal Os", data.get('os', 'Closed')],
            ["Placenta", data.get('placenta', '')],
        ]
        _create_docx_table(doc, early_rows, col_widths=[Cm(5.5), Cm(11.5)])
        
        # EDD Table for Early OBS
        doc.add_paragraph()
        edd_rows = [
            ["EDD (LMP)", data.get('edd_lmp', ''), "±2W"],
            ["EDD (CUA)", data.get('edd_cua', ''), "±2W"],
        ]
        _create_docx_table(doc, edd_rows, col_widths=[Cm(3.5), Cm(5), Cm(3)])

    # --- 4. Standard OB / Level 2 Specific Tables ---
    elif template_id in ["T16", "T20", "T22", "T26", "T28", "T32", "T36", "DOPPLER", "LEVEL2"]:
        # Intro Text
        p = doc.add_paragraph()
        _add_run(p, f"Uterus shows single live fetus. Fetus is in {data.get('presentation', 'VARIABLE')} PRESENTATION at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        
        # Biometry Table
        bio_head = doc.add_paragraph()
        _add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        
        bio_rows = [
            ["BPD", data.get('bpd_size', ''), data.get('bpd_wks', ''), data.get('bpd_dys', '')],
            ["HC", data.get('hc_size', ''), data.get('hc_wks', ''), data.get('hc_dys', '')],
            ["AC", data.get('ac_size', ''), data.get('ac_wks', ''), data.get('ac_dys', '')],
            ["FL", data.get('fl_size', ''), data.get('fl_wks', ''), data.get('fl_dys', '')],
            ["HL", data.get('hl_size', ''), data.get('hl_wks', ''), data.get('hl_dys', '')],
        ]
        _create_docx_table(doc, bio_rows, col_widths=[Cm(3.5), Cm(4), Cm(4.5), Cm(4.5)])
        
        # Summary Metrics
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, f"FL/HC= {data.get('fl_hc', '')} ; HC/AC = {data.get('hc_ac', '')} (BOTH WNL)", bold=True)
        p = doc.add_paragraph()
        _add_run(p, f"MEAN AGE : {data.get('mean_age', '')}", bold=True)
        p = doc.add_paragraph()
        _add_run(p, f"FETAL WEIGHT : {data.get('fetal_weight', '')}", bold=True)
        p = doc.add_paragraph()
        _add_run(p, f"AFI:- Adequate, Largest Pocket Measures {data.get('afi', '')}", bold=False)
        p = doc.add_paragraph()
        _add_run(p, f"Foetal heart rate is approx. {data.get('fhr', '')}", bold=False)
        
        # EDD
        doc.add_paragraph()
        edd_rows = [
            ["EDD (LMP)--", data.get('edd_lmp', ''), "±2W"],
            ["EDD (CUA)--", data.get('edd_cua', ''), "±2W"],
        ]
        _create_docx_table(doc, edd_rows, col_widths=[Cm(3.5), Cm(5), Cm(3)])
        
        # Anatomy Details
        doc.add_paragraph()
        anat_p = doc.add_paragraph()
        _add_run(anat_p, f"Placenta is {data.get('placenta', 'ANTERIOR')} and away from the Os.")
        anat_p = doc.add_paragraph()
        _add_run(anat_p, f"CERVIX- normal ({data.get('cervix', '3.5cm')}), internal Os is closed.")
        anat_p = doc.add_paragraph()
        _add_run(anat_p, f"{data.get('cord_loop', 'NO LOOP OF CORD SEEN AROUND THE NECK.')}")

    # --- 5. Impression ---
    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True, color=RGBColor(0x1F, 0x4E, 0x79))
    
    imp_text = generate_impression(template_id, data)
    p = doc.add_paragraph()
    _add_run(p, imp_text, bold=True, italic=True)

    # --- 6. PNDT Disclaimer & Note ---
    doc.add_paragraph()
    disc_table = doc.add_table(rows=1, cols=1)
    cell = disc_table.cell(0, 0)
    cell.text = ""
    _set_cell_border(cell, top={"sz": 6, "val": "single", "color": "000000"}, bottom={"sz": 6, "val": "single", "color": "000000"})
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, size=8)

    doc.add_paragraph()
    note_p = doc.add_paragraph()
    note_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _add_run(note_p, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", italic=True, size=8)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="OB Reporting Software", layout="wide")

st.markdown("""
    <style>
    .preview-box { background-color: #2a2a2a !important; color: #ffffff !important; font-family: monospace; padding: 15px; border-radius: 8px; white-space: pre-wrap; font-size: 13px; }
    div[data-testid="stExpander"] { background-color: #ffffff !important; border: 1px solid #cfd6dd !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🩺 OB Reporting Software")
st.caption("Generate perfectly aligned Word Reports with Native Tables.")

# TOP DEMOGRAPHICS
c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3])
with c1: p_name = st.text_input("Name", key="p_name")
with c2: p_age = st.text_input("Age", key="p_age")
with c3: p_sex = st.radio("Sex", ["F", "M"], horizontal=True, key="p_sex")
with c4: p_date = st.text_input("Date", value=datetime.now().strftime("%d-%b-%y"), key="p_date")
with c5: p_ref = st.text_input("Referred by", key="p_ref")

st.divider()

col_input, col_preview = st.columns([1, 1], gap="large")

data = {"patient_name": p_name, "patient_age": p_age, "report_date": p_date, "referred_by": p_ref, "p_sex": p_sex}

with col_input:
    template_id = st.selectbox("⚡ Load Clinical Template", options=list(TEMPLATES.keys()), format_func=lambda k: TEMPLATES[k])
    data["lmp"] = st.text_input("LMP", key="p_lmp")

    # Dynamic form blocks based on template
    if template_id == "EARLY_OBS":
        with st.expander("Early OBS Parameters", expanded=True):
            data["gsac"] = st.text_input("G-Sac", value="Seen")
            data["crl"] = st.text_input("CRL", value="17mm- 06wks 04days")
            data["yolk"] = st.text_input("Yolk sac", value="Not Seen")
            data["fhr"] = st.text_input("Cardiac pulsations", value="Present, 176 bpm")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value="26/03/2027")
            data["edd_cua"] = st.text_input("EDD (CUA)", value="04/04/2027")
            
    elif template_id in ["T16", "T20", "T22", "T26", "T28", "T32", "T36", "LEVEL2"]:
        with st.expander("Biometry & EGA", expanded=True):
            data["presentation"] = st.selectbox("Presentation", ["VARIABLE", "CEPHALIC"])
            
            # Biometry Grid
            st.markdown("**Biometry (Size, Weeks, Days)**")
            b1, b2, b3 = st.columns(3)
            data["bpd_size"] = b1.text_input("BPD Size", value="33mm")
            data["bpd_wks"] = b2.text_input("BPD Wks", value="16weeks")
            data["bpd_dys"] = b3.text_input("BPD Days", value="03days±2w")
            
            b1, b2, b3 = st.columns(3)
            data["hc_size"] = b1.text_input("HC Size", value="124mm")
            data["hc_wks"] = b2.text_input("HC Wks", value="16weeks")
            data["hc_dys"] = b3.text_input("HC Days", value="02days±2w")
            
            b1, b2, b3 = st.columns(3)
            data["ac_size"] = b1.text_input("AC Size", value="106mm")
            data["ac_wks"] = b2.text_input("AC Wks", value="16weeks")
            data["ac_dys"] = b3.text_input("AC Days", value="04days±2w")

            b1, b2, b3 = st.columns(3)
            data["fl_size"] = b1.text_input("FL Size", value="22mm")
            data["fl_wks"] = b2.text_input("FL Wks", value="16weeks")
            data["fl_dys"] = b3.text_input("FL Days", value="05days±2w")

            b1, b2, b3 = st.columns(3)
            data["hl_size"] = b1.text_input("HL Size", value="21mm")
            data["hl_wks"] = b2.text_input("HL Wks", value="16weeks")
            data["hl_dys"] = b3.text_input("HL Days", value="03days±2w")

            st.markdown("**Indices & Averages**")
            data["fl_hc"] = st.text_input("FL/HC", value="17.9%")
            data["hc_ac"] = st.text_input("HC/AC", value="1.17")
            data["mean_age"] = st.text_input("Mean Age", value="16WEEKS 03DAYS ± 2W")
            data["fetal_weight"] = st.text_input("Fetal Weight", value="161GRAMS ± 31GM")
            data["afi"] = st.text_input("AFI", value="4.9cm")
            data["fhr"] = st.text_input("FHR", value="155/min")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value="04/08/2027")
            data["edd_cua"] = st.text_input("EDD (CUA)", value="06/08/2027")
            data["placenta"] = st.text_input("Placenta", value="ANTERIOR")
            data["cervix"] = st.text_input("Cervix Length", value="3.1cm")

with col_preview:
    st.subheader("📄 Live Preview (Text)")
    findings_text = generate_findings_text(template_id, data)
    safe_html = html.escape(findings_text).replace("\n", "<br>")
    st.markdown(f'<div class="preview-box">{safe_html}</div>', unsafe_allow_html=True)
    
    st.divider()
    
    st.subheader("⬇️ Download Report")
    docx_bytes = build_docx_bytes(data, template_id)
    
    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=f"OB_Report_{p_name or 'Unknown'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
