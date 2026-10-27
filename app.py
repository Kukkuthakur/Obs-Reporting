import streamlit as st
import html
import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# SHARED CONFIGURATION
# ============================================================
TEMPLATES = {
    "EARLY_OBS": "Early OBS Scan",
    "T16": "OB Report (16 Weeks)",
    "T22": "OB Report (22 Weeks)",
    "T36": "OB Report (36 Weeks)",
}

# ============================================================
# TEXT GENERATOR (For Live Preview)
# ============================================================
def generate_findings_text(template_id, data):
    lines = []
    if template_id == "EARLY_OBS":
        lines.append(f"G-Sac: {data.get('gsac', '')}")
        lines.append(f"CRL: {data.get('crl', '')}")
        lines.append(f"Yolk sac: {data.get('yolk', '')}")
        lines.append(f"Cardiac pulsations: {data.get('fhr', '')}")
    else:
        lines.append(f"Uterus shows single live fetus. Fetus is in {data.get('presentation', 'VARIABLE')} PRESENTATION at the time of scanning.")
        lines.append("\nBIO PHYSICAL PARAMETERS:")
        for param in ["bpd", "hc", "ac", "fl", "hl"]:
            size = data.get(f"{param}_size", "")
            wks = data.get(f"{param}_wks", "")
            dys = data.get(f"{param}_dys", "")
            if size:
                lines.append(f"{param.upper()}: {size}mm | {wks}weeks {dys}days ±2w")
        lines.append(f"\nMEAN AGE: {data.get('mean_age', '')}")
        lines.append(f"FETAL WEIGHT: {data.get('fetal_weight', '')}")
        lines.append(f"AFI: {data.get('afi', '')}")
        lines.append(f"FHR: {data.get('fhr', '')}")
    return "\n".join(lines)

# ============================================================
# DOCX HELPERS & FORMATTING
# ============================================================

def _set_cell_border(cell, **kwargs):
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

def _set_table_borders(table):
    """Forces grid lines on the entire table."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)

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

# ============================================================
# DOCX GENERATION ENGINE
# ============================================================

def build_docx_bytes(data, template_id):
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # --- 1. Demographics Table ---
    # Exactly like the original: 2 rows, 2 columns, but with an inner divider
    header_table = doc.add_table(rows=2, cols=2)
    header_table.autofit = True
    _set_table_borders(header_table)
    
    cells = [
        (0, 0, f"NAME :- {data.get('patient_name', '')}"),
        (0, 1, f"DATE :- {data.get('report_date', '')}"),
        (1, 0, f"AGE/SEX :- {data.get('patient_age', '')}Y/ {data.get('p_sex', 'F')}"),
        (1, 1, f"REF. BY:- {data.get('referred_by', '')}"),
    ]
    for r, c, text in cells:
        cell = header_table.cell(r, c)
        cell.text = ""
        p = cell.paragraphs[0]
        _add_run(p, text, bold=True, size=11)

    doc.add_paragraph()

    # --- 2. Title ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(title, "ULTRASOUND OBSTETRICS REPORT", bold=True, underline=True, size=14, color=RGBColor(0x1F, 0x4E, 0x79))
    
    lmp_p = doc.add_paragraph()
    _add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=True, size=11)
    doc.add_paragraph()

    # --- 3. EARLY OBS SPECIFIC ---
    if template_id == "EARLY_OBS":
        early_table = doc.add_table(rows=10, cols=2)
        _set_table_borders(early_table)
        
        early_data = [
            ("G-Sac", data.get('gsac', '')),
            ("CRL", data.get('crl', '')),
            ("Yolk sac", data.get('yolk', '')),
            ("Cardiac pulsations", data.get('fhr', '')),
            ("Decidual reaction", "Present"),
            ("Subchorionic collection", "Absent"),
            ("Corpus leuteal cyst", "Appreciated"),
            ("Cervix", data.get('cervix', 'Normal')),
            ("Internal Os", "Closed"),
            ("Placenta", "-"),
        ]
        for i, (label, val) in enumerate(early_data):
            row = early_table.rows[i]
            row.cells[0].text = ""
            _add_run(row.cells[0].paragraphs[0], label, bold=True, italic=True, size=10)
            row.cells[1].text = ""
            _add_run(row.cells[1].paragraphs[0], val, size=10)
            
        doc.add_paragraph()
        # EDD Table (Early)
        edd_table = doc.add_table(rows=2, cols=3)
        _set_table_borders(edd_table)
        edd_data = [
            ("EDD (LMP)", data.get('edd_lmp', ''), "±2W"),
            ("EDD (CUA)", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            row = edd_table.rows[i]
            row.cells[0].text = ""
            _add_run(row.cells[0].paragraphs[0], label, bold=True, italic=True, size=10)
            row.cells[1].text = ""
            _add_run(row.cells[1].paragraphs[0], date, size=10)
            row.cells[2].text = ""
            _add_run(row.cells[2].paragraphs[0], err, size=10)

    # --- 4. STANDARD OB (16, 22, 36 Weeks) ---
    else:
        # Intro Paragraph
        p = doc.add_paragraph()
        _add_run(p, f"Uterus shows single live fetus. Fetus is in {data.get('presentation', 'VARIABLE')} PRESENTATION at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        
        # Biometry Header
        bio_head = doc.add_paragraph()
        _add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        
        # Biometry Table (BPD, HC, AC, FL, HL) - 4 Columns
        bio_table = doc.add_table(rows=5, cols=4)
        _set_table_borders(bio_table)
        
        params = ["bpd", "hc", "ac", "fl", "hl"]
        for i, param in enumerate(params):
            row = bio_table.rows[i]
            
            size = data.get(f"{param}_size", "")
            wks = data.get(f"{param}_wks", "")
            dys = data.get(f"{param}_dys", "")
            
            # Col 1: Name
            row.cells[0].text = ""
            _add_run(row.cells[0].paragraphs[0], param.upper(), bold=True, italic=True, size=10)
            
            # Col 2: Size
            row.cells[1].text = ""
            _add_run(row.cells[1].paragraphs[0], f"{size}mm", size=10)
            
            # Col 3: Weeks
            row.cells[2].text = ""
            _add_run(row.cells[2].paragraphs[0], f"{wks}weeks", size=10)
            
            # Col 4: Days
            row.cells[3].text = ""
            _add_run(row.cells[3].paragraphs[0], f"{dys}days±2w", size=10)

        doc.add_paragraph()
        
        # Indices (FL/HC, HC/AC)
        fl_hc = data.get('fl_hc', '')
        hc_ac = data.get('hc_ac', '')
        if template_id in ["T16", "T20"]:
            ratios_text = f"FL/HC= {fl_hc} ; HC/AC = {hc_ac} (BOTH WNL)"
        else:
            ratios_text = f"FL/AC= {data.get('fl_ac', '')} ; FL/BPD = {data.get('fl_bpd', '')} (BOTH WNL)"
        
        p = doc.add_paragraph()
        _add_run(p, ratios_text, bold=True)
        
        # Averages & Other Metrics
        metrics = [
            ("MEAN AGE :", data.get('mean_age', '')),
            ("FETAL WEIGHT :", data.get('fetal_weight', '')),
            ("AFI:-", f"Adequate, Largest Pocket Measures {data.get('afi', '')}"),
            ("Foetal heart rate", f"is approx. {data.get('fhr', '')}"),
        ]
        for label, val in metrics:
            p = doc.add_paragraph()
            _add_run(p, f"{label} {val}", bold=False)

        doc.add_paragraph()
        
        # EDD Table (Standard)
        edd_table = doc.add_table(rows=2, cols=3)
        _set_table_borders(edd_table)
        edd_data = [
            ("EDD (LMP)--", data.get('edd_lmp', ''), "±2W"),
            ("EDD (CUA)--", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            row = edd_table.rows[i]
            row.cells[0].text = ""
            _add_run(row.cells[0].paragraphs[0], label, bold=True, italic=True, size=10)
            row.cells[1].text = ""
            _add_run(row.cells[1].paragraphs[0], date, size=10)
            row.cells[2].text = ""
            _add_run(row.cells[2].paragraphs[0], err, size=10)

        # Anatomy / Placenta / Cervix
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, f"Placenta is {data.get('placenta', 'ANTERIOR')} and away from the Os.", bold=False)
        p = doc.add_paragraph()
        _add_run(p, f"CERVIX- normal ({data.get('cervix', '')}), internal Os is closed.", bold=False)
        p = doc.add_paragraph()
        _add_run(p, "NO LOOP OF CORD SEEN AROUND THE NECK.", bold=True)

    # --- 5. Impression ---
    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True, color=RGBColor(0x1F, 0x4E, 0x79))
    
    mean_age = data.get('mean_age', '')
    presentation = data.get('presentation', 'variable').lower()
    
    if template_id == "EARLY_OBS":
        imp_text = f"Intra uterine live pregnancy of {data.get('crl', '')} +/-1w with good cardiac activity."
    elif template_id == "T16":
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in variable presentation of approx. {mean_age} gestation age."
    else:
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {presentation} presentation of approx. {mean_age} gestation age."
        
    p = doc.add_paragraph()
    _add_run(p, imp_text, bold=True, italic=True)

    # --- 6. Disclaimer & Note ---
    doc.add_paragraph()
    
    # Disclaimer Table (1x1)
    disc_table = doc.add_table(rows=1, cols=1)
    _set_table_borders(disc_table)
    cell = disc_table.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, size=8)

    doc.add_paragraph()
    
    # Standard Note
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
            data["cervix"] = st.text_input("Cervix", value="Normal (4.2cm)")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value="26/03/2027")
            data["edd_cua"] = st.text_input("EDD (CUA)", value="04/04/2027")
            
    else:
        with st.expander("Biometry & EGA", expanded=True):
            data["presentation"] = st.selectbox("Presentation", ["VARIABLE", "CEPHALIC"])
            
            # Biometry Grid
            st.markdown("**Biometry (Size, Weeks, Days)**")
            for param in ["bpd", "hc", "ac", "fl", "hl"]:
                b1, b2, b3 = st.columns(3)
                data[f"{param}_size"] = b1.text_input(f"{param.upper()} Size", value="33" if param=="bpd" else "")
                data[f"{param}_wks"] = b2.text_input(f"{param.upper()} Wks", value="16")
                data[f"{param}_dys"] = b3.text_input(f"{param.upper()} Days", value="03")

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
