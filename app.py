import streamlit as st
import io
import html
import math
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pydicom
import matplotlib.pyplot as plt
import numpy as np

def generate_percentile_bar(percentile):
    """Generates a small PNG image of a percentile bar."""
    if not percentile:
        return None
    
    try:
        pct = float(percentile)
    except:
        return None

    fig, ax = plt.subplots(figsize=(3, 0.4))
    fig.patch.set_alpha(0)
    
    ax.plot([0, 100], [0, 0], color='lightgray', linewidth=4, zorder=1)
    ax.plot([10, 90], [0, 0], color='#4CAF50', linewidth=4, zorder=2)
    ax.scatter(pct, 0, color='#D32F2F', s=80, zorder=3, edgecolors='black')
    
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.5, 0.5)
    ax.axis('off')
    
    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches='tight', transparent=True, dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

# ============================================================
# BARCELONA / FMF Z-SCORE CALCULATOR
# ============================================================

def _ga_days(ga_weeks: float) -> float:
    return ga_weeks * 7.0

def _z_to_percentile(z: float) -> float:
    sign = 1 if z >= 0 else -1
    x = abs(z) / math.sqrt(2.0)
    t = 1.0 / (1.0 + 0.3275911 * x)
    a1, a2, a3, a4, a5 = 0.254829592, -0.284496736, 1.421413741, -1.453152027, 1.061405429
    erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)
    cdf = 0.5 * (1.0 + sign * erf)
    return cdf * 100.0

def calculate_ua_pi(pi: float, ga_weeks: float):
    ga = _ga_days(ga_weeks)
    median = -0.0016 * ga + 1.2876
    sd = 10 ** (-0.0012 * ga + 0.1147)
    z = (pi - median) / sd if sd else 0
    return {"z_score": round(z, 2), "percentile": round(_z_to_percentile(z), 1)}

def calculate_uta_pi(pi: float, ga_weeks: int):
    UTA_PI_REF = {
        11: (1.18, 1.79, 2.70), 12: (1.11, 1.68, 2.53), 13: (1.05, 1.58, 2.38),
        14: (0.99, 1.49, 2.24), 15: (0.94, 1.41, 2.11), 16: (0.89, 1.33, 1.99),
        17: (0.85, 1.27, 1.88), 18: (0.81, 1.20, 1.79), 19: (0.78, 1.15, 1.70),
        20: (0.74, 1.10, 1.61), 21: (0.71, 1.05, 1.54), 22: (0.69, 1.00, 1.47),
        23: (0.66, 0.96, 1.41), 24: (0.64, 0.93, 1.35), 25: (0.62, 0.89, 1.30),
        26: (0.60, 0.86, 1.25), 27: (0.58, 0.84, 1.21), 28: (0.56, 0.81, 1.17),
        29: (0.55, 0.79, 1.13), 30: (0.54, 0.77, 1.10), 31: (0.52, 0.75, 1.06),
        32: (0.51, 0.73, 1.04), 33: (0.50, 0.71, 1.01), 34: (0.50, 0.70, 0.99),
        35: (0.49, 0.69, 0.97), 36: (0.48, 0.68, 0.95), 37: (0.48, 0.67, 0.94),
        38: (0.47, 0.66, 0.92), 39: (0.47, 0.65, 0.91), 40: (0.47, 0.65, 0.90),
        41: (0.47, 0.65, 0.89)
    }
    if ga_weeks not in UTA_PI_REF: return None
    p5, p50, p95 = UTA_PI_REF[ga_weeks]
    if pi <= p5: pct = 5.0
    elif pi >= p95: pct = 95.0
    elif pi <= p50: pct = 5.0 + 45.0 * (pi - p5) / (p50 - p5)
    else: pct = 50.0 + 45.0 * (pi - p50) / (p95 - p50)
    z = math.sqrt(2) * math.erfinv(2 * (pct / 100.0) - 1)
    return {"z_score": round(z, 2), "percentile": round(pct, 1)}

# ============================================================
# DICOM SR PARSER
# ============================================================
def extract_sr_data(ds):
    from datetime import datetime, timedelta
    extracted = {}
    
    def walk_tree(seq, current_site=""):
        for item in seq:
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            new_site = current_site
            if concept in ["Umbilical Artery", "Uterine Artery"]:
                new_site = concept

            if item.ValueType == 'TEXT':
                val = item.get('TextValue', '')
                if val: 
                    if concept == 'EDD from average ultrasound age': extracted['edd_cua'] = val
                    if concept == 'EDD from LMP': extracted['edd_lmp'] = val
                    if concept == 'LMP': extracted['lmp'] = val

            if item.ValueType == 'NUM' and 'MeasuredValueSequence' in item:
                val_seq = item.MeasuredValueSequence[0]
                num_val = float(val_seq.get('NumericValue', 0))
                unit = val_seq.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')
                
                if unit == 'centimeter':
                    if concept == 'Biparietal Diameter': extracted['bpd'] = f"{int(num_val * 10)}"
                    elif concept == 'Head Circumference': extracted['hc'] = f"{int(num_val * 10)}"
                    elif concept == 'Abdominal Circumference': extracted['ac'] = f"{int(num_val * 10)}"
                    elif concept == 'Femur Length': extracted['fl'] = f"{int(num_val * 10)}"
                    elif concept == 'Humerus length': extracted['hl'] = f"{int(num_val * 10)}"
                    elif concept == 'Gestational Sac Diameter': extracted['msd'] = f"{int(num_val * 10)}"
                    elif concept == 'Crown Rump Length': extracted['crl'] = f"{int(num_val * 10)}"
                    elif concept == 'Nuchal Translucency': extracted['nt'] = f"{int(num_val * 10)}"

                if concept == 'Nasal Bone':
                    extracted['nasal_bone'] = "Present" if num_val > 0 else "Absent"
                
                if new_site == "Umbilical Artery" and concept == 'Pulsatility Index':
                    extracted['umbilical_pi'] = str(num_val)
                if new_site == "Uterine Artery" and concept == 'Pulsatility Index':
                    extracted['uterine_pi'] = str(num_val)
                
                if concept == 'Estimated Weight': extracted['fetal_weight'] = str(int(num_val))
                if concept == 'Fetal Heart Rate': extracted['fhr'] = str(int(num_val))
                
                if concept == 'Composite Ultrasound Age': 
                    extracted['ga_usg_days'] = int(num_val)
                    extracted['mean_age'] = f"{int(num_val)//7}weeks {int(num_val)%7}days ± 2w"
                
                if concept == 'Gestational Age by LMP':
                    extracted['ga_lmp_days'] = int(num_val)

                if concept == 'Growth Percentile Rank':
                    if 'current_param' in extracted:
                        param = extracted['current_param']
                        extracted[f"{param}_pct"] = str(num_val)

                if concept == 'Gestational Age':
                    if 'current_param' in extracted:
                        param = extracted['current_param']
                        extracted[f"{param}_wks"] = str(int(num_val) // 7)
                        extracted[f"{param}_dys"] = str(int(num_val) % 7)

            if concept in ['Biparietal Diameter', 'Head Circumference', 'Abdominal Circumference', 'Femur Length', 'Humerus length']:
                param_map = {
                    'Biparietal Diameter': 'bpd', 'Head Circumference': 'hc',
                    'Abdominal Circumference': 'ac', 'Femur Length': 'fl', 'Humerus length': 'hl'
                }
                extracted['current_param'] = param_map[concept]
            
            if 'ContentSequence' in item:
                walk_tree(item.ContentSequence, new_site)
            
            if concept == 'Biometry Group':
                extracted.pop('current_param', None)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)
        
    extracted.pop('current_param', None)
    
    today = datetime.now()
    if 'lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['lmp'] = (today - timedelta(days=extracted['ga_lmp_days'])).strftime("%d-%m-%Y")
    if 'edd_lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['edd_lmp'] = (today + timedelta(days=(280 - extracted['ga_lmp_days']))).strftime("%d-%m-%Y")
    if 'edd_cua' not in extracted and 'ga_usg_days' in extracted:
        extracted['edd_cua'] = (today + timedelta(days=(280 - extracted['ga_usg_days']))).strftime("%d-%m-%Y")

    return extracted

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

def generate_afi_text(afi_val, is_total=True):
    try:
        val = float(afi_val)
    except (ValueError, TypeError):
        return "Adequate", False
    if not is_total: return "", False 
    if val == 0: return "Nil, Not a single pocket of even one cm seen", True
    elif 0 < val < 4: return "Moderate to severe oligohydramnios", True
    elif 4 <= val < 6: return "Mild to moderate oligohydramnios", True
    elif 6 <= val < 8: return "Mild oligohydramnios", True
    elif 8 <= val < 10: return "On the lower side", True
    elif 10 <= val <= 20: return "Adequate", False
    elif 20 < val < 25: return "On the higher side", True
    else: return "Polyhydramnios", True

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

    # 1. HEADER TABLE
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

    # 2. TITLE
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_p, "ULTRASOUND OBSTETRICS REPORT", bold=True, underline=True, size=14, color=RGBColor(0x1F, 0x4E, 0x79))
    lmp_p = doc.add_paragraph()
    if data.get('lmp') == "Not Sure":
        add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=False, italic=True)
    else:
        add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=True)
    doc.add_paragraph()

    # 3. EARLY OBS
    if template_id == "EARLY_OBS":
        early_table = doc.add_table(rows=10, cols=2)
        set_table_borders(early_table)
        for row in early_table.rows:
            row.cells[0].width = Cm(5.5)
            row.cells[1].width = Cm(11.5)
        subchorionic = "Absent" if not data.get("subchorionic_present") else f"Present; {data.get('subchorionic_grade', 'Grade-I')}"
        corpus = "Not Appreciated" if not data.get("corpus_present") else f"Appreciated; {data.get('corpus_size', '')}mm in {data.get('corpus_side', 'Right')} Ovary"
        cardiac = f"Present, {data.get('fhr_val', '')} bpm" if data.get('cardiac_status') == "Present" else data.get('cardiac_status')
        
        early_data = [
            ("G-Sac", data.get('gsac', 'Seen')), ("CRL", data.get('crl', '')),
            ("Yolk sac", data.get('yolk', 'Not Seen')), ("Cardiac pulsations", cardiac),
            ("Decidual reaction", "Present"), ("Subchorionic collection", subchorionic),
            ("Corpus leuteal cyst", corpus), ("Cervix", data.get('cervix', 'Normal')),
            ("Internal Os", "Closed"), ("Placenta", "-"),
        ]
        for i, (label, val) in enumerate(early_data):
            add_run(early_table.cell(i, 0).paragraphs[0], label, italic=True, size=10)
            add_run(early_table.cell(i, 1).paragraphs[0], val, italic=True, size=10)
        doc.add_paragraph()

        edd_table = doc.add_table(rows=2, cols=3)
        set_table_borders(edd_table)
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5); row.cells[1].width = Cm(5.0); row.cells[2].width = Cm(3.0)
        edd_data = [
            ("EDD (LMP)", data.get('edd_lmp', '') if data.get('lmp') != "Not Sure" else "-", "±2W"),
            ("EDD (CUA)", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, size=10)

    # 3.5 NT_NB SCAN
    elif template_id == "NT_NB":
        nt_table = doc.add_table(rows=4, cols=4)
        set_table_borders(nt_table)
        for row in nt_table.rows:
            row.cells[0].width = Cm(4.0); row.cells[1].width = Cm(4.0); row.cells[2].width = Cm(4.0); row.cells[3].width = Cm(4.0)
            
        nt_data = [
            ("CRL", f"{data.get('crl', '')}mm", "NT", f"{data.get('nt', '')}mm"),
            ("Nasal Bone", data.get('nasal_bone', ''), "FHR", f"{data.get('fhr_val', '')} bpm"),
            ("DV PI", data.get('dv_pi', ''), "DV A-Wave", data.get('dv_awave', '')),
            ("Placenta", data.get('placenta_pos', ''), "Cervix", f"{data.get('cervix', '')}cm"),
        ]
        
        for i, (l1, v1, l2, v2) in enumerate(nt_data):
            add_run(nt_table.cell(i, 0).paragraphs[0], l1, bold=True, size=10)
            add_run(nt_table.cell(i, 1).paragraphs[0], v1, italic=True, size=10)
            add_run(nt_table.cell(i, 2).paragraphs[0], l2, bold=True, size=10)
            add_run(nt_table.cell(i, 3).paragraphs[0], v2, italic=True, size=10)
            
        doc.add_paragraph()
        
        edd_table = doc.add_table(rows=2, cols=3)
        set_table_borders(edd_table)
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5); row.cells[1].width = Cm(5.0); row.cells[2].width = Cm(3.0)
        edd_data = [
            ("EDD (LMP)--", data.get('edd_lmp', '') if data.get('lmp') != "Not Sure" else "-", "±2W"),
            ("EDD (CUA)--", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, bold=True, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, italic=True, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, italic=True, size=10)

    # 4. STANDARD OB
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(p, f"Uterus shows single live fetus. Fetus is in ")
        add_run(p, f"{data.get('presentation', 'variable').lower()} presentation", bold=True, italic=True)
        add_run(p, f" at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        doc.add_paragraph()

        bio_head = doc.add_paragraph()
        add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        bio_head.paragraph_format.space_after = Pt(6)

        bio_table = doc.add_table(rows=5, cols=5)
        set_table_borders(bio_table)
        for row in bio_table.rows:
            row.cells[0].width = Cm(3.5); row.cells[1].width = Cm(2.5); row.cells[2].width = Cm(2.5); row.cells[3].width = Cm(2.5); row.cells[4].width = Cm(5.0)
        params = ["bpd", "hc", "ac", "fl", "hl"]
        for i, param in enumerate(params):
            size = data.get(f"{param}_size", ""); wks = data.get(f"{param}_wks", ""); dys = data.get(f"{param}_dys", "")
            add_run(bio_table.cell(i, 0).paragraphs[0], param.upper(), bold=True, italic=True, size=10)
            add_run(bio_table.cell(i, 1).paragraphs[0], f"{size}mm" if size else "", italic=True, size=10)
            add_run(bio_table.cell(i, 2).paragraphs[0], f"{wks}weeks" if wks else "", italic=True, size=10)
            add_run(bio_table.cell(i, 3).paragraphs[0], f"{dys}days±2w" if dys else "", italic=True, size=10)
            
            pct_val = data.get(f"{param}_pct", "")
            if pct_val:
                bar_img = generate_percentile_bar(pct_val)
                if bar_img:
                    cell = bio_table.cell(i, 4)
                    p_img = cell.paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p_img.add_run()
                    run.add_picture(bar_img, width=Cm(3.5))
                    
                    p_txt = cell.add_paragraph(f"{pct_val}%")
                    p_txt.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if p_txt.runs:
                        p_txt.runs[0].font.size = Pt(8)
        doc.add_paragraph()

        ratio_table = doc.add_table(rows=1, cols=2)
        ratio_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        ratio_table.columns[0].width = Cm(7.5); ratio_table.columns[1].width = Cm(8.5)
        add_run(ratio_table.cell(0, 0).paragraphs[0], "FL/HC= ", bold=True)
        add_run(ratio_table.cell(0, 0).paragraphs[0], f"{data.get('fl_hc', '')}", italic=True)
        add_run(ratio_table.cell(0, 1).paragraphs[0], "HC/AC = ", bold=True)
        add_run(ratio_table.cell(0, 1).paragraphs[0], f"{data.get('hc_ac', '')} (BOTH WNL)", italic=True)
        doc.add_paragraph()

        p = doc.add_paragraph()
        add_run(p, "MEAN AGE : ", bold=True); add_run(p, f"{data.get('mean_age', '')}", italic=True)
        p = doc.add_paragraph()
        add_run(p, "FETAL WEIGHT : ", bold=True); add_run(p, f"{data.get('fetal_weight', '')}", italic=True)

        p = doc.add_paragraph()
        add_run(p, "AFI:- ", bold=True)
        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', ''); desc, is_abn = generate_afi_text(afi_val, is_total=True)
            afi_str = desc if desc == "Nil, Not a single pocket of even one cm seen" else f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', ''); desc, is_abn = generate_afi_text(afi_val, is_total=False)
            afi_str = f"largest pocket measures {afi_val} cm" if not desc else f"{desc}, largest pocket measures {afi_val} cm"
        add_run(p, afi_str, bold=is_abn, italic=is_abn)

        p = doc.add_paragraph()
        add_run(p, "Foetal heart rate ", bold=True); add_run(p, f"is approx. {data.get('fhr', '')}.", italic=True)
        doc.add_paragraph()

        edd_table = doc.add_table(rows=2, cols=3)
        set_table_borders(edd_table)
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5); row.cells[1].width = Cm(5.0); row.cells[2].width = Cm(3.0)
        edd_data = [
            ("EDD (LMP)--", data.get('edd_lmp', '') if data.get('lmp') != "Not Sure" else "-", "±2W"),
            ("EDD (CUA)--", data.get('edd_cua', ''), "±2W"),
        ]
        for i, (label, date, err) in enumerate(edd_data):
            add_run(edd_table.cell(i, 0).paragraphs[0], label, bold=True, italic=True, size=10)
            add_run(edd_table.cell(i, 1).paragraphs[0], date, italic=True, size=10)
            add_run(edd_table.cell(i, 2).paragraphs[0], err, italic=True, size=10)
        doc.add_paragraph()

        p = doc.add_paragraph()
        add_run(p, "Placenta ", bold=True)
        plac_pos = data.get('placenta_pos', 'ANTERIOR'); plac_os = data.get('placenta_os', 'away from the Os')
        if plac_os == "Placenta Praevia":
            plac_str = f"is {plac_pos} and Placenta Praevia ({data.get('praevia_grade', 'Grade I')})."
        elif plac_os == "Low-lying":
            dist = data.get('placenta_dist', '')
            plac_str = f"is {plac_pos} and Low-lying, about {dist}cm away from the Internal Os." if dist.strip() else f"is {plac_pos} and Low-lying, but away from the Internal Os."
        else:
            plac_str = f"is {plac_pos} and {plac_os}."
        add_run(p, plac_str)

        p = doc.add_paragraph()
        add_run(p, "CERVIX- ", bold=True); add_run(p, f"normal ({data.get('cervix', '')}), internal Os is {data.get('internal_os', 'closed').lower()}.")

        p = doc.add_paragraph()
        if data.get('cord_around') or data.get('cord_straddling'):
            if data.get('cord_around'): add_run(p, "SINGLE LOOP OF CORD AROUND THE NECK.", bold=True)
            if data.get('cord_straddling'): add_run(p, "SINGLE LOOP OF CORD STRADDLING THE NECK.", bold=True)
        else:
            add_run(p, "NO LOOP OF CORD SEEN AROUND THE NECK.", bold=True)

        if template_id == "T3" and data.get('measure_scar'):
            p = doc.add_paragraph()
            add_run(p, "Scar Thickness: ", bold=True); add_run(p, f"{data.get('scar_size', '')} mm", italic=True)

    # 5. IMPRESSION
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
    elif template_id == "NT_NB":
        imp_text = f"Single live fetus of approx. {mean_age} gestation age. NT measures {data.get('nt', '')}mm. Nasal bone is {data.get('nasal_bone', '').lower()}."
    else:
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {data.get('presentation', 'variable').lower()} presentation of approx. {mean_age} gestation age."
    p = doc.add_paragraph(); add_run(p, imp_text, bold=True, italic=True)

    # 6. PNDT
    doc.add_paragraph()
    disc_table = doc.add_table(rows=1, cols=1)
    disc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(disc_table)
    cell = disc_table.cell(0, 0); p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=8)
    
    doc.add_paragraph()
    note_p = doc.add_paragraph(); note_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(note_p, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=8)

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ============================================================
# LIVE PREVIEW RENDERER
# ============================================================
def render_ob_preview(data, template_id):
    lines = []
    lines.append(f"NAME :- {data.get('patient_name', '')}        DATE :- {data.get('report_date', '')}")
    lines.append(f"AGE/SEX :- {data.get('patient_age', '')}Y/ {data.get('p_sex', 'F')}        REF. BY:- {data.get('referred_by', '')}")
    lines.append("")
    lines.append("         ULTRASOUND OBSTETRICS REPORT")
    lines.append("")
    lines.append(f"LMP- {data.get('lmp', '')}")
    lines.append("")

    if template_id == "EARLY_OBS":
        subchorionic = "Absent" if not data.get("subchorionic_present") else f"Present; {data.get('subchorionic_grade', 'Grade-I')}"
        corpus = "Not Appreciated" if not data.get("corpus_present") else f"Appreciated; {data.get('corpus_size', '')}mm in {data.get('corpus_side', 'Right')} Ovary"
        cardiac = f"Present, {data.get('fhr_val', '')} bpm" if data.get('cardiac_status') == "Present" else data.get('cardiac_status')
        early_rows = [
            ("G-Sac", data.get('gsac', 'Seen')), ("CRL", data.get('crl', '')),
            ("Yolk sac", data.get('yolk', 'Not Seen')), ("Cardiac pulsations", cardiac),
            ("Decidual reaction", "Present"), ("Subchorionic collection", subchorionic),
            ("Corpus leuteal cyst", corpus), ("Cervix", data.get('cervix', 'Normal')),
            ("Internal Os", "Closed"), ("Placenta", "-"),
        ]
        for label, val in early_rows: lines.append(f"  {label:<28}: {val}")
        lines.append("")
        lines.append(f"  {'EDD (LMP)':<20}: {data.get('edd_lmp', '') if data.get('lmp') != 'Not Sure' else '-'}    ±2W")
        lines.append(f"  {'EDD (CUA)':<20}: {data.get('edd_cua', '')}    ±2W")

    elif template_id == "NT_NB":
        lines.append("NT / NB SCAN DETAILS:")
        lines.append(f"  CRL: {data.get('crl', '')}mm")
        lines.append(f"  NT: {data.get('nt', '')}mm")
        lines.append(f"  Nasal Bone: {data.get('nasal_bone', '')}")
        lines.append(f"  FHR: {data.get('fhr_val', '')} bpm")
        lines.append(f"  DV PI: {data.get('dv_pi', '')} (A-wave: {data.get('dv_awave', '')})")
        lines.append("")
        lines.append(f"  Placenta: {data.get('placenta_pos', '')}")
        lines.append(f"  Cervix: {data.get('cervix', '')}cm")
        lines.append("")
        lines.append(f"  EDD (LMP): {data.get('edd_lmp', '')}")
        lines.append(f"  EDD (CUA): {data.get('edd_cua', '')}")

    else:
        lines.append(f"Uterus shows single live fetus. Fetus is in {data.get('presentation', 'variable').lower()} presentation at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        lines.append("")
        lines.append("BIO PHYSICAL PARAMETERS:")
        for param in ["bpd", "hc", "ac", "fl", "hl"]:
            size = data.get(f"{param}_size", ""); wks = data.get(f"{param}_wks", ""); dys = data.get(f"{param}_dys", "")
            lines.append(f"  {param.upper():<4}: {(str(size) + 'mm') if size else '':<10} {(str(wks) + 'weeks') if wks else '':<12} {(str(dys) + 'days±2w') if dys else ''}")
        lines.append("")
        lines.append(f"FL/HC= {data.get('fl_hc', '')}        HC/AC = {data.get('hc_ac', '')} (BOTH WNL)")
        lines.append("")
        lines.append(f"MEAN AGE : {data.get('mean_age', '')}")
        lines.append(f"FETAL WEIGHT : {data.get('fetal_weight', '')}")

        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', ''); desc, _ = generate_afi_text(afi_val, is_total=True)
            afi_str = desc if desc == "Nil, Not a single pocket of even one cm seen" else f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', ''); desc, _ = generate_afi_text(afi_val, is_total=False)
            afi_str = f"largest pocket measures {afi_val} cm" if not desc else f"{desc}, largest pocket measures {afi_val} cm"
        lines.append(f"AFI:- {afi_str}")
        lines.append(f"Foetal heart rate is approx. {data.get('fhr', '')}.")
        lines.append("")

        lines.append(f"  {'EDD (LMP)--':<18}: {data.get('edd_lmp', '') if data.get('lmp') != 'Not Sure' else '-'}    ±2W")
        lines.append(f"  {'EDD (CUA)--':<18}: {data.get('edd_cua', '')}    ±2W")
        lines.append("")

        plac_pos = data.get('placenta_pos', 'ANTERIOR'); plac_os = data.get('placenta_os', 'away from the Os')
        if plac_os == "Placenta Praevia": plac_str = f"is {plac_pos} and Placenta Praevia ({data.get('praevia_grade', 'Grade I')})."
        elif plac_os == "Low-lying":
            dist = data.get('placenta_dist', '')
            plac_str = f"is {plac_pos} and Low-lying, about {dist}cm away from the Internal Os." if dist.strip() else f"is {plac_pos} and Low-lying, but away from the Internal Os."
        else: plac_str = f"is {plac_pos} and {plac_os}."
        lines.append(f"Placenta {plac_str}")

        lines.append(f"CERVIX- normal ({data.get('cervix', '')}), internal Os is {data.get('internal_os', 'closed').lower()}.")
        if data.get('cord_around') or data.get('cord_straddling'):
            parts = []
            if data.get('cord_around'): parts.append("SINGLE LOOP OF CORD AROUND THE NECK.")
            if data.get('cord_straddling'): parts.append("SINGLE LOOP OF CORD STRADDLING THE NECK.")
            lines.append(" ".join(parts))
        else: lines.append("NO LOOP OF CORD SEEN AROUND THE NECK.")

        if template_id == "T3" and data.get('measure_scar'): lines.append(f"Scar Thickness: {data.get('scar_size', '')} mm")

    lines.append(""); lines.append("IMPRESSION:")
    mean_age = data.get('mean_age', '')
    if template_id == "EARLY_OBS":
        if data.get('cardiac_status') == "Blighted Ovum":
            imp_text = f"INTRA UTERINE EMPTY GESTATIONAL SAC OF {data.get('ga_crl', '')} +/-1w WITH NO APPRECIABLE FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY FAILURE (BLIGHTED OVUM)."
        elif data.get('cardiac_status') == "Not yet appreciable":
            imp_text = f"Intra uterine pregnancy of {data.get('ga_crl', '')} +/-1w with not yet appreciable cardiac activity. Adv- Review scan after 2 weeks."
        else:
            imp_text = f"Intra uterine live pregnancy of {data.get('ga_crl', '')} +/-1w with good cardiac activity."
    elif template_id == "NT_NB":
        imp_text = f"Single live fetus of approx. {mean_age} gestation age. NT measures {data.get('nt', '')}mm. Nasal bone is {data.get('nasal_bone', '').lower()}."
    else:
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {data.get('presentation', 'variable').lower()} presentation of approx. {mean_age} gestation age."
    lines.append(f"  - {imp_text}")

    lines.append("")
    lines.append("PNDT DISCLAIMER: The radiologist has not conducted or caused to conduct any pre-natal diagnostic technique for the purpose of determining fetal sex.")
    return "\n".join(lines)

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="OB Reporting Software", layout="wide")

st.markdown("""
    <style>
    :root, html, body { color-scheme: light !important; background-color: #f4f6f9 !important; }
    #MainMenu {visibility: hidden;} header[data-testid="stHeader"] {visibility: hidden; height: 0;}
    footer {visibility: hidden;} [data-testid="stToolbar"] {visibility: hidden;}
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > .main, section.main, .main .block-container { background-color: #f4f6f9 !important; }
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp li, .stApp small, .stApp strong, .stApp em { color: #111111 !important; }
    .clinic-title { text-align: center; color: #1F4E79 !important; font-weight: 700; font-size: 2rem; margin-bottom: 0; padding-bottom: 0; letter-spacing: 0.5px; }
    .clinic-sub { text-align: center; color: #555 !important; font-style: italic; font-size: 1.05rem; margin-top: 0; padding-top: 0; margin-bottom: 1.5rem; }
    .dicom-box { background-color: #f0f5fa !important; border: 1px solid #b0c4de !important; border-radius: 6px; padding: 10px; margin-bottom: 15px; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] { display: flex !important; flex-direction: column !important; justify-content: flex-end !important; }
    div[data-testid="stRadio"] label, div[data-testid="stCheckbox"] label { cursor: pointer !important; }
    div[data-testid="stExpander"] { background-color: #ffffff !important; border: 1px solid #cfd6dd !important; border-radius: 6px !important; margin-bottom: 4px !important; }
    div[data-testid="stExpander"] details, div[data-testid="stExpander"] details > div { background-color: #ffffff !important; }
    div[data-testid="stExpander"] details > summary { padding: 8px 12px !important; min-height: 42px !important; box-sizing: border-box !important; display: flex !important; align-items: center !important; background-color: #ffffff !important; }
    div[data-testid="stExpander"] details[open] > summary { background-color: #eef2f7 !important; border-bottom: 1px solid #cfd6dd !important; }
    div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary * { color: #111111 !important; font-weight: 600 !important; font-size: 14px !important; white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; }
    div[data-testid="stNumberInput"] input { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; border-radius: 6px !important; min-height: 42px !important; box-sizing: border-box !important; font-weight: 600 !important; text-align: center !important; padding: 0 6px !important; }
    div[data-testid="stNumberInput"] button { display: none !important; }
    .stTextInput input, .stTextArea textarea, div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea, [data-baseweb="input"] input, [data-baseweb="base-input"] input, [data-baseweb="textarea"] textarea { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; min-height: 42px !important; box-sizing: border-box !important; }
    div[data-baseweb="select"] > div, div[data-baseweb="popover"] > div, ul[data-baseweb="menu"] { background-color: #ffffff !important; color: #111111 !important; }
    .stButton button, .stDownloadButton button, div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; }
    .stButton button:hover, .stDownloadButton button:hover { background-color: #eef2f7 !important; border-color: #305496 !important; }
    div[data-testid="stAlert"] { background-color: #ffffff !important; color: #111111 !important; }
    div[data-testid="stCaptionContainer"], div[data-testid="stCaptionContainer"] * { color: #555555 !important; }
    hr { border-color: #cfd6dd !important; }
    div[data-testid="stMarkdownContainer"] * { color: #111111; }
    .preview-box, .stApp .preview-box, .stApp .preview-box * { color: #ffffff !important; }
    .preview-box { background-color: #2a2a2a !important; font-family: Consolas, Menlo, 'Courier New', monospace !important; font-size: 12.5px !important; line-height: 1.5 !important; padding: 14px 16px !important; border-radius: 8px !important; border: 1px solid #333 !important; white-space: pre-wrap !important; word-wrap: break-word !important; overflow-wrap: anywhere !important; word-break: break-word !important; margin-bottom: 8px !important; display: block !important; }
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
    template_id = st.selectbox("Template", ["EARLY_OBS", "NT_NB", "T16", "T3"])
    
    # --- DUAL DICOM INPUT (Manual Upload + Listener) ---
    import json
    import os
    
    st.markdown("### 🔗 DICOM Input")
    col_up, col_ref = st.columns([2, 1])
    
    with col_up:
        uploaded_sr = st.file_uploader("Auto-Fill from File", type=["dcm", "dicom", "ima", ""], label_visibility="collapsed")
        if uploaded_sr is not None:
            try:
                ds = pydicom.dcmread(io.BytesIO(uploaded_sr.read()))
                st.session_state["sr_data"] = extract_sr_data(ds)
                st.success("DICOM SR File Loaded!")
            except Exception as e:
                st.error(f"Error reading SR: {e}")
                
    with col_ref:
        if st.button("🔄 Refresh from Listener", use_container_width=True):
            st.rerun()
            
    # Check for automated listener file (if no manual file was uploaded)
    if uploaded_sr is None and os.path.exists('latest_sr.json'):
        try:
            with open('latest_sr.json', 'r') as f:
                st.session_state["sr_data"] = json.load(f)
            st.info("✅ Auto-loaded from DICOM Listener")
        except:
            pass
            
    # Push values into session_state widgets
    sr_data = st.session_state.get("sr_data", {})
    
    if 'bpd' in sr_data: st.session_state["sr_bpd_size"] = sr_data['bpd']
    if 'hc' in sr_data: st.session_state["sr_hc_size"] = sr_data['hc']
    if 'ac' in sr_data: st.session_state["sr_ac_size"] = sr_data['ac']
    if 'fl' in sr_data: st.session_state["sr_fl_size"] = sr_data['fl']
    if 'hl' in sr_data: st.session_state["sr_hl_size"] = sr_data['hl']
    
    for p in ["bpd", "hc", "ac", "fl", "hl"]:
        if f"{p}_wks" in sr_data: st.session_state[f"sr_{p}_wks"] = sr_data[f"{p}_wks"]
        if f"{p}_dys" in sr_data: st.session_state[f"sr_{p}_dys"] = sr_data[f"{p}_dys"]
        if f"{p}_pct" in sr_data: st.session_state[f"sr_{p}_pct"] = sr_data[f"{p}_pct"]

    if 'msd' in sr_data: st.session_state["sr_msd"] = sr_data['msd']
    if 'crl' in sr_data: st.session_state["sr_crl"] = sr_data['crl']
    if 'nt' in sr_data: st.session_state["sr_nt"] = sr_data['nt']
    if 'nasal_bone' in sr_data: st.session_state["sr_nasal_bone"] = sr_data['nasal_bone']
    if 'fetal_weight' in sr_data: st.session_state["sr_fetal_weight"] = sr_data['fetal_weight']
    if 'fhr' in sr_data: st.session_state["sr_fhr"] = sr_data['fhr']
    if 'mean_age' in sr_data: st.session_state["sr_mean_age"] = sr_data['mean_age']
    
    if 'edd_lmp' in sr_data: st.session_state["sr_edd_lmp"] = sr_data['edd_lmp']
    if 'edd_cua' in sr_data: st.session_state["sr_edd_cua"] = sr_data['edd_cua']
    if 'lmp' in sr_data: st.session_state["sr_lmp"] = sr_data['lmp']
    
    if 'ga_lmp_days' in sr_data: 
        st.session_state["sr_ga_lmp"] = f"{sr_data['ga_lmp_days']//7}w{sr_data['ga_lmp_days']%7}d"

    if 'umbilical_pi' in sr_data: st.session_state["sr_umbilical_pi"] = sr_data['umbilical_pi']
    if 'uterine_pi' in sr_data: st.session_state["sr_uterine_pi"] = sr_data['uterine_pi']
    # --------------------------
    lmp_not_sure = st.checkbox("LMP Not Sure / Not Known")
    if lmp_not_sure:
        data["lmp"] = "Not Sure"
    else:
        data["lmp"] = st.text_input("LMP", value=st.session_state.get("sr_lmp", ""))

    with st.expander("🔗 DICOM SR Data (Machine Generated)", expanded=False):
        st.markdown('<div class="dicom-box">Verify or override data received from the ultrasound machine.</div>', unsafe_allow_html=True)
        
        if template_id != "EARLY_OBS":
            # Timeline Header
            st.markdown(f"**LMP:** {st.session_state.get('sr_lmp', '')} &nbsp;|&nbsp; "
                        f"**GA(LMP):** {st.session_state.get('sr_ga_lmp', '')} &nbsp;|&nbsp; "
                        f"**AUA:** {st.session_state.get('sr_mean_age', '')} &nbsp;|&nbsp; "
                        f"**EDD(LMP):** {st.session_state.get('sr_edd_lmp', '')}")
            st.markdown("---")

        if template_id == "EARLY_OBS":
            data["gsac"] = st.text_input("G-Sac", value=st.session_state.get("sr_gsac", "Seen"), key="sr_gsac")
            data["msd"] = st.text_input("MSD", value=st.session_state.get("sr_msd", ""), key="sr_msd")
            data["crl"] = st.text_input("CRL", value=st.session_state.get("sr_crl", ""), key="sr_crl")
            data["ga_crl"] = st.text_input("GA by CRL", value="")
            data["yolk"] = st.text_input("Yolk sac", value="")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get("sr_edd_lmp", ""))
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get("sr_edd_cua", ""))
        elif template_id == "NT_NB":
            b1, b2 = st.columns(2)
            data["crl"] = b1.text_input("CRL (mm)", value=st.session_state.get("sr_crl", ""))
            data["nt"] = b2.text_input("NT (mm)", value=st.session_state.get("sr_nt", ""))
            
            b1, b2 = st.columns(2)
            data["nasal_bone"] = b1.text_input("Nasal Bone", value=st.session_state.get("sr_nasal_bone", ""))
            data["fhr_val"] = b2.text_input("FHR (bpm)", value=st.session_state.get("sr_fhr", ""))
            
            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get("sr_edd_lmp", ""))
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get("sr_edd_cua", ""))
        else:
            st.markdown("**Fetal Growth Parameters**")
            h1, h2, h3, h4, h5 = st.columns([2, 1, 1, 1, 3])
            h1.caption("Parameter")
            h2.caption("Size (mm)")
            h3.caption("Wks")
            h4.caption("Days")
            h5.caption("Percentile")

            params = [
                ("bpd", "Biparietal Diameter"),
                ("hc", "Head Circumference"),
                ("ac", "Abdominal Circumference"),
                ("fl", "Femur Length"),
                ("hl", "Humerus Length")
            ]

            for p_key, p_name in params:
                c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 3])
                c1.markdown(f"**{p_name}**")
                
                data[f"{p_key}_size"] = c2.text_input(f"{p_key} size", value=st.session_state.get(f"sr_{p_key}_size", ""), key=f"sr_{p_key}_size", label_visibility="collapsed")
                data[f"{p_key}_wks"] = c3.text_input(f"{p_key} wks", value=st.session_state.get(f"sr_{p_key}_wks", ""), key=f"sr_{p_key}_wks", label_visibility="collapsed")
                data[f"{p_key}_dys"] = c4.text_input(f"{p_key} dys", value=st.session_state.get(f"sr_{p_key}_dys", ""), key=f"sr_{p_key}_dys", label_visibility="collapsed")
                
                pct_val = st.session_state.get(f"sr_{p_key}_pct", "")
                if pct_val:
                    bar_img = generate_percentile_bar(pct_val)
                    if bar_img:
                        c5.image(bar_img, use_column_width=True)
                        c5.caption(f"{pct_val}%")
                else:
                    c5.text_input(f"{p_key} pct", value="", key=f"sr_{p_key}_pct_display", label_visibility="collapsed")

            st.markdown("---")
            st.markdown("**Ratios & Averages**")
            b1, b2 = st.columns(2)
            
            # Calculate ratios on the fly if we have the values
            try:
                if data.get("fl_size") and data.get("hc_size"):
                    fl_hc_val = (float(data["fl_size"]) / float(data["hc_size"])) * 100
                    st.session_state["sr_fl_hc"] = f"{fl_hc_val:.1f}%"
                if data.get("hc_size") and data.get("ac_size"):
                    hc_ac_val = float(data["hc_size"]) / float(data["ac_size"])
                    st.session_state["sr_hc_ac"] = f"{hc_ac_val:.2f}"
                if data.get("fl_size") and data.get("bpd_size"):
                    fl_bpd_val = (float(data["fl_size"]) / float(data["bpd_size"])) * 100
                    st.session_state["sr_fl_bpd"] = f"{fl_bpd_val:.1f}%"
                if data.get("fl_size") and data.get("ac_size"):
                    fl_ac_val = (float(data["fl_size"]) / float(data["ac_size"])) * 100
                    st.session_state["sr_fl_ac"] = f"{fl_ac_val:.1f}%"
            except:
                pass

            data["fl_hc"] = b1.text_input("FL/HC", value=st.session_state.get("sr_fl_hc", ""))
            data["hc_ac"] = b2.text_input("HC/AC", value=st.session_state.get("sr_hc_ac", ""))
            
            b1, b2 = st.columns(2)
            data["fl_bpd"] = b1.text_input("FL/BPD", value=st.session_state.get("sr_fl_bpd", ""))
            data["fl_ac"] = b2.text_input("FL/AC", value=st.session_state.get("sr_fl_ac", ""))

            b1, b2 = st.columns(2)
            data["mean_age"] = b1.text_input("Mean Age", value=st.session_state.get("sr_mean_age", ""), key="sr_mean_age")
            data["fetal_weight"] = b2.text_input("Fetal Weight (g)", value=st.session_state.get("sr_fetal_weight", ""), key="sr_fetal_weight")

            b1, b2 = st.columns(2)
            data["fhr"] = b1.text_input("FHR", value=st.session_state.get("sr_fhr", ""), key="sr_fhr")
            data["edd_lmp"] = b2.text_input("EDD (LMP)", value=st.session_state.get("sr_edd_lmp", ""))
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get("sr_edd_cua", ""))
            
            st.markdown("**Doppler Indices & Percentiles**")
            b1, b2, b3 = st.columns(3)
            data["umbilical_pi"] = b1.text_input("Umbilical PI", value=st.session_state.get("sr_umbilical_pi", ""), key="sr_umbilical_pi")
            data["mca_pi"] = b2.text_input("MCA PI", value="")
            data["uterine_pi"] = b3.text_input("Mean Uterine PI", value=st.session_state.get("sr_uterine_pi", ""), key="sr_uterine_pi")

            data["ga_weeks_dop"] = st.number_input("GA (weeks) for Percentiles", min_value=10.0, max_value=42.0, value=28.0, step=0.1)

            if data.get("umbilical_pi"):
                try:
                    res = calculate_ua_pi(float(data["umbilical_pi"]), data["ga_weeks_dop"])
                    st.caption(f"Umbilical PI: Z-Score = {res['z_score']}, Percentile = {res['percentile']}th")
                except: pass
            if data.get("uterine_pi"):
                try:
                    res = calculate_uta_pi(float(data["uterine_pi"]), int(round(data["ga_weeks_dop"])))
                    if res: st.caption(f"Uterine PI: Z-Score = {res['z_score']}, Percentile = {res['percentile']}th")
                except: pass

    if template_id == "EARLY_OBS":
        with st.expander("Early OBS Details", expanded=True):
            data["cardiac_status"] = st.selectbox("Cardiac Activity", ["Present", "Not yet appreciable", "Blighted Ovum"])
            if data["cardiac_status"] == "Present":
                data["fhr_val"] = st.text_input("FHR Value (bpm)", value=st.session_state.get("sr_fhr", ""))
            data["cervix"] = st.text_input("Cervix Length", value="Normal")
            data["subchorionic_present"] = st.checkbox("Subchorionic collection present")
            if data["subchorionic_present"]:
                data["subchorionic_grade"] = st.radio("Grade", ["Grade-I", "Grade-I/II", "Grade-II", "Grade-II/III"], horizontal=True)
            data["corpus_present"] = st.checkbox("Corpus Luteum Cyst present")
            if data["corpus_present"]:
                b1, b2 = st.columns(2)
                data["corpus_side"] = b1.selectbox("Side", ["Right", "Left"])
                data["corpus_size"] = b2.text_input("Size (mm)")

    elif template_id == "NT_NB":
        with st.expander("Typist Inputs (NT / NB)", expanded=True):
            data["dv_pi"] = st.text_input("Ductus Venosus PI", value="")
            data["dv_awave"] = st.selectbox("DV A-Wave", ["Positive", "Absent", "Reversed"])
            
            b1, b2 = st.columns(2)
            data["placenta_pos"] = b1.selectbox("Placenta Position", ["ANTERIOR", "POSTERIOR", "FUNDAL", "LOW-LYING"])
            data["cervix"] = b2.text_input("Cervix Length", value="")

    else:
        with st.expander("Anatomy & Presentation", expanded=True):
            b1, b2 = st.columns(2)
            data["presentation"] = b1.selectbox("Presentation", ["VARIABLE", "CEPHALIC", "BREECH"])
            data["lie"] = b2.selectbox("Lie", ["None", "LOA", "ROA", "LOP", "ROP"])

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
    st.subheader("📄 Live Preview")
    _preview_text = render_ob_preview(data, template_id)
    _safe = html.escape(_preview_text).replace("\n", "<br>")
    st.markdown(f'<div class="preview-box">{_safe}</div>', unsafe_allow_html=True)

    st.subheader("⬇️ Download Report")
    docx_bytes = build_docx_bytes(data, template_id)
    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=f"OB_Report_{p_name or 'Unknown'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
