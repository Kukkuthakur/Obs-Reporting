import streamlit as st
import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# DOCX HELPERS (Exact Border & Alignment Control)
# ============================================================

def set_cell_border(cell, **kwargs):
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

def set_table_borders(table):
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

def add_run(paragraph, text, bold=False, italic=False, size=11, underline=False, color=None):
    """Helper to add text with specific formatting."""
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
# DOCX BUILDER (Matching the exact Word reports)
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
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    # --- 1. DEMOGRAPHIC HEADER TABLE (2x2 Grid) ---
    demo_table = doc.add_table(rows=2, cols=2)
    demo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(demo_table)
    
    for row in demo_table.rows:
        row.cells[0].width = Cm(8.5)
        row.cells[1].width = Cm(8.5)

    cell = demo_table.cell(0, 0)
    add_run(cell.paragraphs[0], f"NAME :- {data.get('patient_name', '')}", bold=True)
    
    cell = demo_table.cell(0, 1)
    add_run(cell.paragraphs[0], f"DATE :- {data.get('report_date', '')}", bold=True)

    cell = demo_table.cell(1, 0)
    add_run(cell.paragraphs[0], f"AGE/SEX :- {data.get('patient_age', '')}Y/ {data.get('p_sex', 'F')}", bold=True)
    
    cell = demo_table.cell(1, 1)
    add_run(cell.paragraphs[0], f"REF. BY:- {data.get('referred_by', '')}", bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- 2. TITLE & LMP ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_p, "ULTRASOUND OBSTETRICS REPORT", bold=True, underline=True, size=14, color=RGBColor(0x1F, 0x4E, 0x79))

    lmp_p = doc.add_paragraph()
    add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=True)
    doc.add_paragraph()

    # --- 3. EARLY OBS SPECIFIC FORMAT ---
    if template_id == "EARLY_OBS":
        early_table = doc.add_table(rows=10, cols=2)
        set_table_borders(early_table)
        
        for row in early_table.rows:
            row.cells[0].width = Cm(5.5)
            row.cells[1].width = Cm(11.5)

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
            cell_l = early_table.cell(i, 0)
            cell_r = early_table.cell(i, 1)
            
            add_run(cell_l.paragraphs[0], label, italic=True, size=10)
            add_run(cell_r.paragraphs[0], val, italic=True, size=10)

        doc.add_paragraph()

        # EDD Table
        edd_table = doc.add_table(rows=2, cols=3)
        set_table_borders(edd_table)
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5)
            row.cells[1].width = Cm(5.0)
            row.cells[2].width = Cm(3.0)

        edd_data = [
            ("EDD (LMP)", data.get('edd_lmp', ''), "±2W"),
            ("EDD (CUA)", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, size=10)

    # --- 4. STANDARD OB (16, 22, 36 Weeks) FORMAT ---
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(p, f"Uterus shows single live fetus. Fetus is in ")
        add_run(p, f"{data.get('presentation', 'VARIABLE')} PRESENTATION", bold=True, italic=True)
        add_run(p, f" at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        
        doc.add_paragraph()

        bio_head = doc.add_paragraph()
        add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        bio_head.paragraph_format.space_after = Pt(6)

        # Biometry Table (5 rows, 4 cols)
        bio_table = doc.add_table(rows=5, cols=4)
        set_table_borders(bio_table)
        
        for row in bio_table.rows:
            row.cells[0].width = Cm(3.5)
            row.cells[1].width = Cm(4.0)
            row.cells[2].width = Cm(4.5)
            row.cells[3].width = Cm(4.5)

        params = ["bpd", "hc", "ac", "fl", "hl"]
        for i, param in enumerate(params):
            size = data.get(f"{param}_size", "")
            wks = data.get(f"{param}_wks", "")
            dys = data.get(f"{param}_dys", "")
            
            add_run(bio_table.cell(i, 0).paragraphs[0], param.upper(), bold=True, italic=True, size=10)
            add_run(bio_table.cell(i, 1).paragraphs[0], f"{size}mm" if size else "", italic=True, size=10)
            add_run(bio_table.cell(i, 2).paragraphs[0], f"{wks}weeks" if wks else "", italic=True, size=10)
            add_run(bio_table.cell(i, 3).paragraphs[0], f"{dys}days±2w" if dys else "", italic=True, size=10)

        doc.add_paragraph()
        
        p = doc.add_paragraph()
        add_run(p, "FL/HC= ", bold=True)
        add_run(p, f"{data.get('fl_hc', '')} ; ", italic=True)
        add_run(p, "HC/AC = ", bold=True)
        add_run(p, f"{data.get('hc_ac', '')} (BOTH WNL)", italic=True)

        p = doc.add_paragraph()
        add_run(p, "MEAN AGE : ", bold=True)
        add_run(p, f"{data.get('mean_age', '')}", italic=True)

        p = doc.add_paragraph()
        add_run(p, "FETAL WEIGHT : ", bold=True)
        add_run(p, f"{data.get('fetal_weight', '')}", italic=True)

        p = doc.add_paragraph()
        add_run(p, "AFI:- ", bold=True)
        add_run(p, f"Adequate, Largest pocket measures {data.get('afi', '')}.", italic=True)

        p = doc.add_paragraph()
        add_run(p, "Foetal heart rate ", bold=True)
        add_run(p, f"is approx. {data.get('fhr', '')}.", italic=True)

        doc.add_paragraph()
        
        edd_table = doc.add_table(rows=2, cols=3)
        set_table_borders(edd_table)
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5)
            row.cells[1].width = Cm(5.0)
            row.cells[2].width = Cm(3.0)
            
        edd_data = [
            ("EDD (LMP)--", data.get('edd_lmp', ''), "±2W"),
            ("EDD (CUA)--", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, bold=True, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, italic=True, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, italic=True, size=10)

        doc.add_paragraph()
        
        p = doc.add_paragraph()
        add_run(p, "Placenta ", bold=True)
        add_run(p, f"is {data.get('placenta', 'ANTERIOR')} and away from the Os.")

        p = doc.add_paragraph()
        add_run(p, "CERVIX- ", bold=True)
        add_run(p, f"normal ({data.get('cervix', '')}), internal Os is closed.")

        p = doc.add_paragraph()
        add_run(p, "NO LOOP OF CORD SEEN AROUND THE NECK.", bold=True)

    # --- 5. IMPRESSION ---
    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    add_run(imp_head, "IMPRESSION:", bold=True, underline=True)
    
    mean_age = data.get('mean_age', '')
    if template_id == "EARLY_OBS":
        imp_text = f"Ultrasound early obs reveals intra uterine live pregnancy of {data.get('crl', '')} +/-1w with good cardiac activity."
    else:
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in variable presentation of approx. {mean_age} gestation age."
        
    p = doc.add_paragraph()
    add_run(p, imp_text, bold=True, italic=True)

    # --- 6. PNDT DISCLAIMER & NOTE BOX ---
    doc.add_paragraph()
    
    disc_table = doc.add_table(rows=1, cols=1)
    disc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(disc_table)
    cell = disc_table.cell(0, 0)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=8)

    doc.add_paragraph()
    
    note_p = doc.add_paragraph()
    note_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(note_p, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=8)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="OB Reporting Software", layout="wide")

# CSS for the clinic name and clean layout
st.markdown("""
    <style>
    .clinic-title { text-align: center; color: #1F4E79; font-weight: 700; font-size: 2rem; margin-bottom: 0; padding-bottom: 0; }
    .clinic-sub { text-align: center; color: #555; font-style: italic; font-size: 1.05rem; margin-top: 0; padding-top: 0; margin-bottom: 1.5rem; }
    div[data-testid="stExpander"] { background-color: #ffffff !important; border: 1px solid #cfd6dd !important; }
    </style>
""", unsafe_allow_html=True)

# CLINIC HEADER
st.markdown('<div class="clinic-title">डॉ. प्रणव एवं डॉ. गौरव इमेजिंग, आगरा</div>', unsafe_allow_html=True)
st.markdown('<div class="clinic-sub">Precision Imaging. Trusted Diagnostics. Agra.</div>', unsafe_allow_html=True)

# Patient Demographics Top Bar
c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3])
with c1: p_name = st.text_input("Name")
with c2: p_age = st.text_input("Age")
with c3: p_sex = st.radio("Sex", ["F", "M"], horizontal=True)
with c4: p_date = st.text_input("Date", value=datetime.now().strftime("%d-%b-%y"))
with c5: p_ref = st.text_input("Referred by")

data = {"patient_name": p_name, "patient_age": p_age, "report_date": p_date, "referred_by": p_ref, "p_sex": p_sex}

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    template_id = st.selectbox("Template", ["T16", "EARLY_OBS"])
    data["lmp"] = st.text_input("LMP")
    
    if template_id == "EARLY_OBS":
        data["gsac"] = st.text_input("G-Sac", value="Seen")
        data["crl"] = st.text_input("CRL", value="17mm- 06wks 04days")
        data["yolk"] = st.text_input("Yolk sac", value="Not Seen")
        data["fhr"] = st.text_input("Cardiac pulsations", value="Present, 176 bpm")
        data["cervix"] = st.text_input("Cervix", value="Normal(4.2cm)")
        data["edd_lmp"] = st.text_input("EDD (LMP)", value="26/03/2027")
        data["edd_cua"] = st.text_input("EDD (CUA)", value="04/04/2027")
    else:
        data["presentation"] = st.text_input("Presentation", value="VARIABLE")
        for param in ["bpd", "hc", "ac", "fl", "hl"]:
            b1, b2, b3 = st.columns(3)
            data[f"{param}_size"] = b1.text_input(f"{param.upper()} Size", value="33")
            data[f"{param}_wks"] = b2.text_input(f"{param.upper()} Wks", value="16")
            data[f"{param}_dys"] = b3.text_input(f"{param.upper()} Days", value="03")
        
        data["fl_hc"] = st.text_input("FL/HC", value="17.9%")
        data["hc_ac"] = st.text_input("HC/AC", value="1.17")
        data["mean_age"] = st.text_input("Mean Age", value="16WEEKS 03DAYS ± 2W")
        data["fetal_weight"] = st.text_input("Fetal Weight", value="161GRAMS ± 31GM")
        data["afi"] = st.text_input("AFI", value="4.9cm")
        data["fhr"] = st.text_input("FHR", value="155/min")
        data["edd_lmp"] = st.text_input("EDD (LMP)", value="04/08/2027")
        data["edd_cua"] = st.text_input("EDD (CUA)", value="06/08/2027")
        data["placenta"] = st.text_input("Placenta", value="ANTERIOR")
        data["cervix"] = st.text_input("Cervix", value="3.1cm")

with col2:
    st.subheader("⬇️ Download Report")
    docx_bytes = build_docx_bytes(data, template_id)
    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=f"OB_Report_{p_name or 'Unknown'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
