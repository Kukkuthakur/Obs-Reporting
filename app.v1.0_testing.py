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
# DOCX HELPERS
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
    run = paragraph.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    if color is not None:
        run.font.color.rgb = color
    return run

def generate_afi_text(afi_val):
    """Smart logic for AFI descriptive text based on clinical values."""
    try:
        val = float(afi_val)
    except (ValueError, TypeError):
        return "Adequate", False  # Default if blank or invalid

    if val == 0:
        return "Nil, Not a single pocket of even one cm seen", True
    elif 0 < val < 4:
        return "Moderate to severe oligohydramnios", True
    elif 4 <= val < 6:
        return "Mild to moderate oligohydramnios", True
    elif 6 <= val < 8:
        return "Mild oligohydramnios", True
    elif 8 <= val < 10:
        return "On the lower side", True
    elif 10 <= val <= 20:
        return "Adequate", False
    elif 20 < val < 25:
        return "On the higher side", True
    else: # >= 25
        return "Polyhydramnios", True

# ============================================================
# DOCX BUILDER
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

    # --- 1. DEMOGRAPHIC HEADER TABLE ---
    demo_table = doc.add_table(rows=2, cols=2)
    demo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(demo_table)
    
    for row in demo_table.rows:
        row.cells[0].width = Cm(8.5)
        row.cells[1].width = Cm(8.5)

    add_run(demo_table.cell(0, 0).paragraphs[0], f"NAME :- {data.get('patient_name', '')}", bold=True)
    add_run(demo_table.cell(0, 1).paragraphs[0], f"DATE :- {data.get('report_date', '')}", bold=True)
    add_run(demo_table.cell(1, 0).paragraphs[0], f"AGE/SEX :- {data.get('patient_age', '')}Y/ {data.get('p_sex', 'F')}", bold=True)
    add_run(demo_table.cell(1, 1).paragraphs[0], f"REF. BY:- {data.get('referred_by', '')}", bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- 2. TITLE & LMP ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_p, "ULTRASOUND OBSTETRICS REPORT", bold=True, underline=True, size=14, color=RGBColor(0x1F, 0x4E, 0x79))

    lmp_p = doc.add_paragraph()
    if data.get('lmp') == "Not Sure":
        add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=False, italic=True)
    else:
        add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=True)
    doc.add_paragraph()

    # --- 3. EARLY OBS SPECIFIC FORMAT ---
    if template_id == "EARLY_OBS":
        early_table = doc.add_table(rows=10, cols=2)
        set_table_borders(early_table)
        
        for row in early_table.rows:
            row.cells[0].width = Cm(5.5)
            row.cells[1].width = Cm(11.5)

        subchorionic = "Absent"
        if data.get("subchorionic_present"):
            subchorionic = f"Present; {data.get('subchorionic_grade', 'Grade-I')}"

        corpus = "Not Appreciated"
        if data.get("corpus_present"):
            side = data.get("corpus_side", "Right")
            size = data.get("corpus_size", "")
            corpus = f"Appreciated; {size}mm in {side} Ovary"

        cardiac = f"Present, {data.get('fhr_val', '')} bpm" if data.get('cardiac_status') == "Present" else data.get('cardiac_status')

        early_data = [
            ("G-Sac", data.get('gsac', 'Seen')),
            ("CRL", data.get('crl', '')),
            ("Yolk sac", data.get('yolk', 'Not Seen')),
            ("Cardiac pulsations", cardiac),
            ("Decidual reaction", "Present"),
            ("Subchorionic collection", subchorionic),
            ("Corpus leuteal cyst", corpus),
            ("Cervix", data.get('cervix', 'Normal')),
            ("Internal Os", "Closed"),
            ("Placenta", "-"),
        ]
        
        for i, (label, val) in enumerate(early_data):
            add_run(early_table.cell(i, 0).paragraphs[0], label, italic=True, size=10)
            add_run(early_table.cell(i, 1).paragraphs[0], val, italic=True, size=10)

        doc.add_paragraph()

        edd_table = doc.add_table(rows=2, cols=3)
        set_table_borders(edd_table)
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5)
            row.cells[1].width = Cm(5.0)
            row.cells[2].width = Cm(3.0)

        edd_data = [
            ("EDD (LMP)", data.get('edd_lmp', '') if data.get('lmp') != "Not Sure" else "-", "±2W"),
            ("EDD (CUA)", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, size=10)

    # --- 4. STANDARD OB FORMAT ---
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(p, f"Uterus shows single live fetus. Fetus is in ")
        # Lowercase bold italic for presentation
        add_run(p, f"{data.get('presentation', 'variable').lower()} presentation", bold=True, italic=True)
        add_run(p, f" at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        
        doc.add_paragraph()

        bio_head = doc.add_paragraph()
        add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        bio_head.paragraph_format.space_after = Pt(6)

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
        
        # --- RATIO TABLE (For Alignment) ---
        ratio_table = doc.add_table(rows=1, cols=2)
        ratio_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        
        # Set widths to match the Biometry table's first 2 columns
        ratio_table.columns[0].width = Cm(7.5)
        ratio_table.columns[1].width = Cm(8.5)

        p1 = ratio_table.cell(0, 0).paragraphs[0]
        add_run(p1, "FL/HC= ", bold=True)
        add_run(p1, f"{data.get('fl_hc', '')}", italic=True)

        p2 = ratio_table.cell(0, 1).paragraphs[0]
        add_run(p2, "HC/AC = ", bold=True)
        add_run(p2, f"{data.get('hc_ac', '')} (BOTH WNL)", italic=True)

        doc.add_paragraph()
        
        p = doc.add_paragraph()
        add_run(p, "MEAN AGE : ", bold=True)
        add_run(p, f"{data.get('mean_age', '')}", italic=True)

        p = doc.add_paragraph()
        add_run(p, "FETAL WEIGHT : ", bold=True)
        add_run(p, f"{data.get('fetal_weight', '')}", italic=True)

        # --- AFI SMART LOGIC ---
        p = doc.add_paragraph()
        add_run(p, "AFI:- ", bold=True)
        
        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', '')
            desc, is_abnormal = generate_afi_text(afi_val)
            if desc == "Nil, Not a single pocket of even one cm seen":
                afi_str = desc
            else:
                afi_str = f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', '')
            desc, is_abnormal = generate_afi_text(afi_val)
            afi_str = f"{desc}, largest pocket measures {afi_val} cm"

        if is_abnormal:
            add_run(p, afi_str, bold=True, italic=True)
        else:
            add_run(p, afi_str, italic=False)

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
            ("EDD (LMP)--", data.get('edd_lmp', '') if data.get('lmp') != "Not Sure" else "-", "±2W"),
            ("EDD (CUA)--", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, bold=True, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, italic=True, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, italic=True, size=10)

        doc.add_paragraph()
        
        # --- PLACENTA LOGIC ---
        p = doc.add_paragraph()
        add_run(p, "Placenta ", bold=True)
        
        plac_pos = data.get('placenta_pos', 'ANTERIOR')
        plac_os = data.get('placenta_os', 'away from the Os')
        
        if plac_os == "Placenta Praevia":
            plac_str = f"is {plac_pos} and Placenta Praevia ({data.get('praevia_grade', 'Grade I')})."
        elif plac_os == "Low-lying":
            dist = data.get('placenta_dist', '')
            if dist.strip():
                plac_str = f"is {plac_pos} and Low-lying, about {dist}cm away from the Internal Os."
            else:
                plac_str = f"is {plac_pos} and Low-lying, but away from the Internal Os."
        else:
            plac_str = f"is {plac_pos} and {plac_os}."
            
        add_run(p, plac_str)

        p = doc.add_paragraph()
        add_run(p, "CERVIX- ", bold=True)
        add_run(p, f"normal ({data.get('cervix', '')}), internal Os is {data.get('internal_os', 'closed').lower()}.")

        p = doc.add_paragraph()
        if data.get('cord_around') or data.get('cord_straddling'):
            if data.get('cord_around'): add_run(p, "SINGLE LOOP OF CORD AROUND THE NECK.", bold=True)
            if data.get('cord_straddling'): add_run(p, "SINGLE LOOP OF CORD STRADDLING THE NECK.", bold=True)
        else:
            add_run(p, "NO LOOP OF CORD SEEN AROUND THE NECK.", bold=True)

        if template_id == "T3" and data.get('measure_scar'):
            p = doc.add_paragraph()
            add_run(p, "Scar Thickness: ", bold=True)
            add_run(p, f"{data.get('scar_size', '')} mm", italic=True)

    # --- 5. IMPRESSION ---
    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    add_run(imp_head, "IMPRESSION:", bold=True, underline=True)
    
    mean_age = data.get('mean_age', '')
    
    if template_id == "EARLY_OBS":
        if data.get('cardiac_status') == "Blighted Ovum":
            imp_text = f"INTRA UTERINE EMPTY GESTATIONAL SAC OF {data.get('ga_crl', '')} +/-1w WITH NO APPRECIABLE FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY FAILURE (BLIGHTED OVUM)."
        elif data.get('cardiac_status') == "Not yet appreciable":
            imp_text = f"Intra uterine pregnancy of {data.get('ga_crl', '')} +/-1w with not yet appreciable cardiac activity. Adv- Review scan after 2 weeks."
        else:
            imp_text = f"Intra uterine live pregnancy of {data.get('ga_crl', '')} +/-1w with good cardiac activity."
    else:
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {data.get('presentation', 'variable').lower()} presentation of approx. {mean_age} gestation age."
        
    p = doc.add_paragraph()
    add_run(p, imp_text, bold=True, italic=True)

    # --- 6. PNDT DISCLAIMER ---
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

st.markdown("""
    <style>
    .clinic-title { text-align: center; color: #1F4E79; font-weight: 700; font-size: 2rem; margin-bottom: 0; padding-bottom: 0; }
    .clinic-sub { text-align: center; color: #555; font-style: italic; font-size: 1.05rem; margin-top: 0; padding-top: 0; margin-bottom: 1.5rem; }
    div[data-testid="stExpander"] { background-color: #ffffff !important; border: 1px solid #cfd6dd !important; }
    .dicom-box { background-color: #f0f5fa !important; border: 1px solid #b0c4de !important; border-radius: 6px; padding: 10px; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

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
    template_id = st.selectbox("Template", ["EARLY_OBS", "T16", "T3"])
    
    # LMP Logic
    lmp_not_sure = st.checkbox("LMP Not Sure / Not Known")
    data["lmp"] = "Not Sure" if lmp_not_sure else st.text_input("LMP")

    # ==========================================
    # DICOM ACCORDION (Machine Data)
    # ==========================================
    with st.expander("🔗 DICOM SR Data (Machine Generated)", expanded=False):
        st.markdown('<div class="dicom-box">Verify or override data received from the ultrasound machine.</div>', unsafe_allow_html=True)
        if template_id == "EARLY_OBS":
            data["gsac"] = st.text_input("G-Sac", value="Seen")
            data["crl"] = st.text_input("CRL", value="")
            data["ga_crl"] = st.text_input("GA by CRL", value="")
            data["yolk"] = st.text_input("Yolk sac", value="")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value="")
            data["edd_cua"] = st.text_input("EDD (CUA)", value="")
        else:
            for param in ["bpd", "hc", "ac", "fl", "hl"]:
                b1, b2, b3 = st.columns(3)
                data[f"{param}_size"] = b1.text_input(f"{param.upper()} Size")
                data[f"{param}_wks"] = b2.text_input(f"{param.upper()} Wks")
                data[f"{param}_dys"] = b3.text_input(f"{param.upper()} Days")
            
            b1, b2 = st.columns(2)
            data["fl_hc"] = b1.text_input("FL/HC")
            data["hc_ac"] = b2.text_input("HC/AC")
            
            b1, b2 = st.columns(2)
            data["mean_age"] = b1.text_input("Mean Age")
            data["fetal_weight"] = b2.text_input("Fetal Weight (g)")
            
            b1, b2 = st.columns(2)
            data["fhr"] = b1.text_input("FHR")
            data["edd_lmp"] = b2.text_input("EDD (LMP)")
            data["edd_cua"] = st.text_input("EDD (CUA)")

    # ==========================================
    # MANUAL ENTRY (Typist Data)
    # ==========================================
    if template_id == "EARLY_OBS":
        with st.expander("Early OBS Details", expanded=True):
            data["cardiac_status"] = st.selectbox("Cardiac Activity", ["Present", "Not yet appreciable", "Blighted Ovum"])
            if data["cardiac_status"] == "Present":
                data["fhr_val"] = st.text_input("FHR Value (bpm)")
            
            data["cervix"] = st.text_input("Cervix Length", value="Normal")
            
            data["subchorionic_present"] = st.checkbox("Subchorionic collection present")
            if data["subchorionic_present"]:
                data["subchorionic_grade"] = st.radio("Grade", ["Grade-I", "Grade-I/II", "Grade-II", "Grade-II/III"], horizontal=True)
            
            data["corpus_present"] = st.checkbox("Corpus Luteum Cyst present")
            if data["corpus_present"]:
                b1, b2 = st.columns(2)
                data["corpus_side"] = b1.selectbox("Side", ["Right", "Left"])
                data["corpus_size"] = b2.text_input("Size (mm)")

    else:
        with st.expander("Anatomy & Presentation", expanded=True):
            b1, b2 = st.columns(2)
            data["presentation"] = b1.selectbox("Presentation", ["VARIABLE", "CEPHALIC", "BREECH"])
            data["lie"] = b2.selectbox("Lie", ["None", "LOA", "ROA", "LOP", "ROP"])
            
            # AFI Logic
            data["afi_method"] = st.radio("AFI Method", ["Largest Pocket", "Total 4 Quadrants"], horizontal=True)
            if data["afi_method"] == "Total 4 Quadrants":
                data["afi_total"] = st.text_input("AFI (cm)")
            else:
                data["afi_pocket"] = st.text_input("Largest Pocket (cm)")
            
            b1, b2 = st.columns(2)
            data["placenta_pos"] = b1.selectbox("Placenta Position", ["ANTERIOR", "POSTERIOR", "FUNDAL", "LOW-LYING"])
            data["placenta_os"] = b2.selectbox("Placenta Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"])
            
            if data["placenta_os"] == "Low-lying":
                data["placenta_dist"] = st.text_input("Distance from Internal Os (cm) [Optional]")
            elif data["placenta_os"] == "Placenta Praevia":
                data["praevia_grade"] = st.radio("Praevia Grade", ["Grade I", "Grade II", "Grade III", "Grade IV"], horizontal=True)
            
            b1, b2 = st.columns(2)
            data["cervix"] = b1.text_input("Cervix Length")
            data["internal_os"] = b2.selectbox("Internal Os", ["Closed", "Open"])
            
            st.markdown("**Cord Loop**")
            data["cord_around"] = st.checkbox("Single loop of cord around the neck")
            data["cord_straddling"] = st.checkbox("Single loop of cord straddling the neck")
            
            if template_id == "T3":
                st.markdown("**Scar Thickness**")
                data["measure_scar"] = st.checkbox("Measure Scar Thickness")
                if data["measure_scar"]:
                    data["scar_size"] = st.text_input("Scar Size (mm)")

with col2:
    st.subheader("⬇️ Download Report")
    docx_bytes = build_docx_bytes(data, template_id)
    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=f"OB_Report_{p_name or 'Unknown'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
