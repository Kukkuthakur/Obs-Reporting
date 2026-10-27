import streamlit as st
import io
import html
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
# LIVE PREVIEW RENDERER
# ============================================================

def render_ob_preview(data, template_id):
    """Return a plain-text rendering of the report in the same order
    as the DOCX builder. Used for the live preview box only."""
    lines = []

    # --- 1. Demographics ---
    lines.append(f"NAME :- {data.get('patient_name', '')}        "
                 f"DATE :- {data.get('report_date', '')}")
    lines.append(f"AGE/SEX :- {data.get('patient_age', '')}Y/ "
                 f"{data.get('p_sex', 'F')}        "
                 f"REF. BY:- {data.get('referred_by', '')}")
    lines.append("")

    # --- 2. Title & LMP ---
    lines.append("         ULTRASOUND OBSTETRICS REPORT")
    lines.append("")
    lines.append(f"LMP- {data.get('lmp', '')}")
    lines.append("")

    # --- 3. EARLY_OBS path ---
    if template_id == "EARLY_OBS":
        subchorionic = "Absent"
        if data.get("subchorionic_present"):
            subchorionic = f"Present; {data.get('subchorionic_grade', 'Grade-I')}"

        corpus = "Not Appreciated"
        if data.get("corpus_present"):
            side = data.get("corpus_side", "Right")
            size = data.get("corpus_size", "")
            corpus = f"Appreciated; {size}mm in {side} Ovary"

        if data.get('cardiac_status') == "Present":
            cardiac = f"Present, {data.get('fhr_val', '')} bpm"
        else:
            cardiac = data.get('cardiac_status')

        early_rows = [
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
        for label, val in early_rows:
            lines.append(f"  {label:<28}: {val}")
        lines.append("")

        lines.append(
            f"  {'EDD (LMP)':<20}: "
            f"{data.get('edd_lmp', '') if data.get('lmp') != 'Not Sure' else '-'}"
            f"    ±2W"
        )
        lines.append(
            f"  {'EDD (CUA)':<20}: {data.get('edd_cua', '')}    ±2W"
        )

    # --- 4. Standard path (T16 / T3) ---
    else:
        lines.append(
            f"Uterus shows single live fetus. Fetus is in "
            f"{data.get('presentation', 'variable').lower()} presentation "
            f"at the time of scanning. Fetal movements are present. "
            f"Fetal cardiac pulsations are appreciated."
        )
        lines.append("")

        lines.append("BIO PHYSICAL PARAMETERS:")
        for param in ["bpd", "hc", "ac", "fl", "hl"]:
            size = data.get(f"{param}_size", "")
            wks = data.get(f"{param}_wks", "")
            dys = data.get(f"{param}_dys", "")
            lines.append(
                f"  {param.upper():<4}: "
                f"{(str(size) + 'mm') if size else '':<10} "
                f"{(str(wks) + 'weeks') if wks else '':<12} "
                f"{(str(dys) + 'days±2w') if dys else ''}"
            )
        lines.append("")

        lines.append(
            f"FL/HC= {data.get('fl_hc', '')}        "
            f"HC/AC = {data.get('hc_ac', '')} (BOTH WNL)"
        )
        lines.append("")

        lines.append(f"MEAN AGE : {data.get('mean_age', '')}")
        lines.append(f"FETAL WEIGHT : {data.get('fetal_weight', '')}")

        # AFI (same logic as DOCX)
        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', '')
            desc, _is_abn = generate_afi_text(afi_val)
            if desc == "Nil, Not a single pocket of even one cm seen":
                afi_str = desc
            else:
                afi_str = (f"{desc}, total of four quadrants measures "
                           f"{afi_val} cm")
        else:
            afi_val = data.get('afi_pocket', '')
            desc, _is_abn = generate_afi_text(afi_val)
            afi_str = f"{desc}, largest pocket measures {afi_val} cm"
        lines.append(f"AFI:- {afi_str}")

        lines.append(f"Foetal heart rate is approx. {data.get('fhr', '')}.")
        lines.append("")

        lines.append(
            f"  {'EDD (LMP)--':<18}: "
            f"{data.get('edd_lmp', '') if data.get('lmp') != 'Not Sure' else '-'}"
            f"    ±2W"
        )
        lines.append(
            f"  {'EDD (CUA)--':<18}: {data.get('edd_cua', '')}    ±2W"
        )
        lines.append("")

        # Placenta
        plac_pos = data.get('placenta_pos', 'ANTERIOR')
        plac_os = data.get('placenta_os', 'away from the Os')
        if plac_os == "Placenta Praevia":
            plac_str = (f"is {plac_pos} and Placenta Praevia "
                        f"({data.get('praevia_grade', 'Grade I')}).")
        elif plac_os == "Low-lying":
            dist = data.get('placenta_dist', '')
            if dist.strip():
                plac_str = (f"is {plac_pos} and Low-lying, about {dist}cm "
                            f"away from the Internal Os.")
            else:
                plac_str = (f"is {plac_pos} and Low-lying, but away from "
                            f"the Internal Os.")
        else:
            plac_str = f"is {plac_pos} and {plac_os}."
        lines.append(f"Placenta {plac_str}")

        lines.append(
            f"CERVIX- normal ({data.get('cervix', '')}), internal Os is "
            f"{data.get('internal_os', 'closed').lower()}."
        )

        if data.get('cord_around') or data.get('cord_straddling'):
            parts = []
            if data.get('cord_around'):
                parts.append("SINGLE LOOP OF CORD AROUND THE NECK.")
            if data.get('cord_straddling'):
                parts.append("SINGLE LOOP OF CORD STRADDLING THE NECK.")
            lines.append(" ".join(parts))
        else:
            lines.append("NO LOOP OF CORD SEEN AROUND THE NECK.")

        if template_id == "T3" and data.get('measure_scar'):
            lines.append(f"Scar Thickness: {data.get('scar_size', '')} mm")

    # --- 5. Impression ---
    lines.append("")
    lines.append("IMPRESSION:")
    mean_age = data.get('mean_age', '')
    if template_id == "EARLY_OBS":
        if data.get('cardiac_status') == "Blighted Ovum":
            imp_text = (f"INTRA UTERINE EMPTY GESTATIONAL SAC OF "
                        f"{data.get('ga_crl', '')} +/-1w WITH NO APPRECIABLE "
                        f"FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY "
                        f"FAILURE (BLIGHTED OVUM).")
        elif data.get('cardiac_status') == "Not yet appreciable":
            imp_text = (f"Intra uterine pregnancy of {data.get('ga_crl', '')} "
                        f"+/-1w with not yet appreciable cardiac activity. "
                        f"Adv- Review scan after 2 weeks.")
        else:
            imp_text = (f"Intra uterine live pregnancy of "
                        f"{data.get('ga_crl', '')} +/-1w with good cardiac "
                        f"activity.")
    else:
        imp_text = (f"Ultrasound obstetric reveals single live intrauterine "
                    f"fetus in {data.get('presentation', 'variable').lower()} "
                    f"presentation of approx. {mean_age} gestation age.")
    lines.append(f"  - {imp_text}")

    # --- 6. Disclaimer (short) ---
    lines.append("")
    lines.append("PNDT DISCLAIMER: The radiologist has not conducted or "
                 "caused to conduct any pre-natal diagnostic technique for "
                 "the purpose of determining fetal sex.")

    return "\n".join(lines)

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="OB Reporting Software", layout="wide")

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

    /* === Background chain (matches USG Whole Abdomen app) === */
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > .main,
    section.main,
    .main .block-container {
        background-color: #f4f6f9 !important;
    }

    /* === Global text color === */
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
    .stApp li, .stApp small, .stApp strong, .stApp em {
        color: #111111 !important;
    }

    /* === OB header === */
    .clinic-title {
        text-align: center; color: #1F4E79 !important;
        font-weight: 700; font-size: 2rem;
        margin-bottom: 0; padding-bottom: 0; letter-spacing: 0.5px;
    }
    .clinic-sub {
        text-align: center; color: #555 !important;
        font-style: italic; font-size: 1.05rem;
        margin-top: 0; padding-top: 0; margin-bottom: 1.5rem;
    }

    /* === DICOM info box === */
    .dicom-box {
        background-color: #f0f5fa !important;
        border: 1px solid #b0c4de !important;
        border-radius: 6px; padding: 10px; margin-bottom: 15px;
    }

    /* === Column layout === */
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        display: flex !important; flex-direction: column !important;
        justify-content: flex-end !important;
    }
    @media (max-width: 900px) {
        div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            min-width: 0 !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            flex-wrap: wrap !important;
            gap: 0.4rem 1rem !important;
        }
        div[data-testid="stRadio"] label,
        div[data-testid="stCheckbox"] label {
            min-height: 44px !important;
            padding: 6px 10px !important;
            display: flex !important;
            align-items: center !important;
            cursor: pointer !important;
            user-select: none !important;
        }
    }
    div[data-testid="stRadio"] label,
    div[data-testid="stCheckbox"] label { cursor: pointer !important; }

    /* === Expanders === */
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

    /* === Number input === */
    div[data-testid="stNumberInput"] input {
        background-color: #ffffff !important; color: #111111 !important;
        border: 1px solid #cfd6dd !important; border-radius: 6px !important;
        min-height: 42px !important; box-sizing: border-box !important;
        font-weight: 600 !important; text-align: center !important;
        padding: 0 6px !important;
    }
    div[data-testid="stNumberInput"] button { display: none !important; }

    /* === Text inputs / textareas === */
    .stTextInput input,
    .stTextArea textarea,
    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea,
    [data-baseweb="input"] input,
    [data-baseweb="base-input"] input,
    [data-baseweb="textarea"] textarea {
        background-color: #ffffff !important;
        color: #111111 !important;
        border: 1px solid #cfd6dd !important;
        min-height: 42px !important;
        box-sizing: border-box !important;
    }

    /* === Selectbox popover / menu === */
    div[data-baseweb="select"] > div,
    div[data-baseweb="popover"] > div,
    ul[data-baseweb="menu"] {
        background-color: #ffffff !important;
        color: #111111 !important;
    }

    /* === Buttons === */
    .stButton button,
    .stDownloadButton button,
    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button {
        background-color: #ffffff !important;
        color: #111111 !important;
        border: 1px
