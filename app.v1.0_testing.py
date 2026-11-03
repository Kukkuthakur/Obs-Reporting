import streamlit as st
import io
import html
import math
import re
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
    """Generates a color-coded PNG image of a percentile bar."""
    if not percentile:
        return None
    
    try:
        pct = float(percentile)
    except:
        return None

    # Make the figure slightly taller to fit the color zones better
    fig, ax = plt.subplots(figsize=(3, 0.5))
    fig.patch.set_alpha(0)
    
    # 1. Draw the color zones
    # Red: < 5th
    ax.plot([0, 5], [0, 0], color='#D32F2F', linewidth=6, zorder=1, solid_capstyle='butt')
    # Yellow: 5th - 10th
    ax.plot([5, 10], [0, 0], color='#FFC107', linewidth=6, zorder=1, solid_capstyle='butt')
    # Green: 10th - 90th
    ax.plot([10, 90], [0, 0], color='#4CAF50', linewidth=6, zorder=1, solid_capstyle='butt')
    # Yellow: 90th - 95th
    ax.plot([90, 95], [0, 0], color='#FFC107', linewidth=6, zorder=1, solid_capstyle='butt')
    # Red: > 95th
    ax.plot([95, 100], [0, 0], color='#D32F2F', linewidth=6, zorder=1, solid_capstyle='butt')
    
    # 2. Draw the marker dot
    # If the dot is in a red or yellow zone, make it stand out more
    dot_color = '#000000' if 10 <= pct <= 90 else '#D32F2F'
    ax.scatter(pct, 0, color=dot_color, s=60, zorder=3, edgecolors='white', linewidths=1.5)
    
    # 3. Add subtle tick marks at 10 and 90
    ax.plot([10, 10], [-0.15, 0.15], color='black', linewidth=1, zorder=2)
    ax.plot([90, 90], [-0.15, 0.15], color='black', linewidth=1, zorder=2)
    
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.3, 0.3)
    ax.axis('off')
    
    buf = io.BytesIO()
    # Increased DPI slightly for a crisper look in Word
    plt.savefig(buf, format="png", bbox_inches='tight', transparent=True, dpi=250)
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
    return {"z_score": 0.0, "percentile": round(pct, 1)}

# ============================================================
# DICOM SR PARSER
# ============================================================
def extract_sr_data(ds):
    from datetime import datetime, timedelta
    extracted = {}
    
    # We will use a flat list to grab every concept in the tree first, 
    # then parse them in order.
    flat_items = []
    
    def walk_tree(seq, parent_site=""):
        for item in seq:
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            new_site = parent_site
            # Catch any site identifier used by the machine
            if concept in ["Umbilical Artery", "Uterine Artery", "Fetal Doppler", "Maternal Doppler", "MCA"]:
                new_site = concept
            
            # Store the concept, the value sequence, and the parent site
            flat_items.append((concept, item, new_site))
            
            if 'ContentSequence' in item:
                walk_tree(item.ContentSequence, new_site)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)

    # Now we parse the flat list
    current_param = None
    for concept, item, site in flat_items:
        
        # Keep track of which biometry parameter we are under
        if concept in ['Biparietal Diameter', 'Head Circumference', 'Abdominal Circumference', 'Femur Length', 'Humerus length']:
            param_map = {
                'Biparietal Diameter': 'bpd', 'Head Circumference': 'hc',
                'Abdominal Circumference': 'ac', 'Femur Length': 'fl', 'Humerus length': 'hl'
            }
            current_param = param_map[concept]

        # --- TEXT VALUES ---
        if item.ValueType == 'TEXT':
            val = item.get('TextValue', '')
            if val: 
                if concept == 'EDD from average ultrasound age': extracted['edd_cua'] = val
                if concept == 'EDD from LMP': extracted['edd_lmp'] = val
                if concept == 'LMP': extracted['lmp'] = val

        # --- NUMERIC VALUES ---
        if item.ValueType == 'NUM' and 'MeasuredValueSequence' in item:
            val_seq = item.MeasuredValueSequence[0]
            num_val = float(val_seq.get('NumericValue', 0))
            unit = val_seq.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')
            
            # Biometry Sizes (in cm, convert to mm)
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
            
            # --- BULLETPROOF DOPPLER INDICES PARSER ---
            concept_lower = concept.lower()
            
            # 1. Determine the Vessel
            vessel = None
            if "umbilical" in site.lower() or "umbilical" in concept_lower:
                vessel = "umbilical"
            elif "uterine" in site.lower() or "uterine" in concept_lower:
                vessel = "uterine"
            elif "mca" in site.lower() or "cerebral" in site.lower() or "mca" in concept_lower:
                vessel = "mca"

            # 2. Determine the Measurement Type and Save
            if vessel:
                val_str = str(round(num_val, 2))
                
                # PI (Pulsatility)
                if "pulsatility" in concept_lower or " pi" in concept_lower or concept_lower == "pi":
                    if vessel == "uterine":
                        if 'uterine_pi' in extracted:
                            extracted['uterine_pi_right'] = val_str
                        else:
                            extracted['uterine_pi'] = val_str
                    else:
                        extracted[f'{vessel}_pi'] = val_str
                        
                # RI (Resistivity)
                elif "resistivity" in concept_lower or " ri" in concept_lower or concept_lower == "ri":
                    if vessel == "uterine":
                        if 'uterine_ri' in extracted:
                            extracted['uterine_right_ri'] = val_str
                        else:
                            extracted['uterine_ri'] = val_str
                    else:
                        extracted[f'{vessel}_ri'] = val_str
                        
                # S/D (Systolic / Diastolic)
                elif "diastolic" in concept_lower or "s/d" in concept_lower or "systolic" in concept_lower:
                    if vessel == "uterine":
                        if 'uterine_sd' in extracted:
                            extracted['uterine_right_sd'] = val_str
                        else:
                            extracted['uterine_sd'] = val_str
                    else:
                        extracted[f'{vessel}_sd'] = val_str

            # Summary & Ratios
            if concept == 'Estimated Weight': 
                extracted['fetal_weight'] = str(int(num_val))
            if concept == 'EFW percentile rank':
                extracted['fetal_weight_pct'] = str(round(num_val, 1))
            
            if concept == 'HC/AC': extracted['hc_ac'] = str(round(num_val, 2))
            if concept == 'FL/AC': extracted['fl_ac'] = str(round(num_val, 1))
            if concept == 'FL/BPD': extracted['fl_bpd'] = str(round(num_val, 1))
            if concept == 'FL/HC': extracted['fl_hc'] = str(round(num_val, 1))

            if concept == 'Fetal Heart Rate': 
                extracted['fhr'] = str(int(num_val))
            
            if concept == 'Composite Ultrasound Age': 
                extracted['ga_usg_days'] = int(num_val)
                extracted['mean_age'] = f"{int(num_val)//7}weeks {int(num_val)%7}days ± 2w"
            
            if concept == 'Gestational Age by LMP':
                extracted['ga_lmp_days'] = int(num_val)

            # Biometry Percentiles & GA
            if concept == 'Growth Percentile Rank':
                if current_param:
                    extracted[f"{current_param}_pct"] = str(round(num_val, 1))

            if concept == 'Gestational Age':
                if current_param:
                    extracted[f"{current_param}_wks"] = str(int(num_val) // 7)
                    extracted[f"{current_param}_dys"] = str(int(num_val) % 7)

        # Reset current_param when we hit a new group
        if concept in ['Biometry Group', 'Fetal Long Bones']:
            current_param = None
    
    # Date calculations
    today = datetime.now()
    if 'lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['lmp'] = (today - timedelta(days=extracted['ga_lmp_days'])).strftime("%d-%m-%Y")
    if 'edd_lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['edd_lmp'] = (today + timedelta(days=(280 - extracted['ga_lmp_days']))).strftime("%d-%m-%Y")
    if 'edd_cua' not in extracted and 'ga_usg_days' in extracted:
        extracted['edd_cua'] = (today + timedelta(days=(280 - extracted['ga_usg_days']))).strftime("%d-%m-%Y")

    # --- GRAB AND SPLIT PATIENT NAME/AGE ---
    if hasattr(ds, 'PatientName'):
        raw_name = str(ds.PatientName).replace('^', ' ').strip()
        extracted['patient_name'] = raw_name
        # Try to split the name and age (e.g., "SHIVANI 22Y/F")
        if 'Y/F' in raw_name.upper():
            parts = raw_name.upper().split('Y/F')
            age_part = parts[0].strip().split(' ')[-1] # Get the last word before Y/F (the age)
            name_part = parts[0].strip().replace(age_part, '').strip() # Remove the age from the name
            extracted['patient_age'] = age_part
            extracted['patient_name'] = name_part
    # ---------------------------------------

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

def set_cell_background(cell, hex_color):
    """Sets the background shading of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def get_gradient_color(row_index, total_rows=6):
    """Returns a light green alternating gradient hex code."""
    # Base light green is E8F5E9. Darker is C8E6C9.
    # We alternate to keep it readable like the reference image.
    if row_index == 0:
        return "A5D6A7" # Header: Darker Green
    elif row_index % 2 == 1:
        return "E8F5E9" # Row 1: Lighter
    else:
        return "D3EAD5" # Row 2: Slightly darker

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
def generate_doppler_text(data):
    doppler_lines = []
    if data.get("umbilical_pi"):
        doppler_lines.append(f"Umbilical Artery: PI = {data.get('umbilical_pi')}")
    if data.get("uterine_pi"):
        notching = data.get("uterine_notching", "No notching")
        if notching == "No notching":
            doppler_lines.append(f"Uterine Artery: PI = {data.get('uterine_pi')} (No notching seen)")
        else:
            doppler_lines.append(f"Uterine Artery: PI = {data.get('uterine_pi')} ({notching})")
    if data.get("mca_pi"):
        doppler_lines.append(f"MCA: PI = {data.get('mca_pi')}")
    return "\n".join(doppler_lines)

def build_docx_bytes(data, template_id, sr_data=None):
    if sr_data is None:
        sr_data = {}
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

    # === FLEXIBLE TOP SPACER ===
    # Adds exactly 2 blank lines (11pt) above the patient demographics box
    # to leave room for pre-printed stationery/letterhead.
    for _ in range(2):
        _p = doc.add_paragraph()
        _p.paragraph_format.space_after = Pt(0)
        _p.paragraph_format.space_before = Pt(0)
        _p.paragraph_format.line_spacing = 1.0
        _run = _p.add_run("")
        _run.font.name = "Calibri"
        _run.font.size = Pt(11)
        
    # Adds 2 more blank lines (10pt) to push the report down further
    for _ in range(2):
        _p = doc.add_paragraph()
        _p.paragraph_format.space_after = Pt(0)
        _p.paragraph_format.space_before = Pt(0)
        _p.paragraph_format.line_spacing = 1.0
        _run = _p.add_run("")
        _run.font.name = "Calibri"
        _run.font.size = Pt(10)
    # === END SPACER ===

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
    
    # 8pt gap
    p_gap1 = doc.add_paragraph()
    p_gap1.paragraph_format.space_after = Pt(0)
    p_gap1.paragraph_format.line_spacing = Pt(8)

    # 2. TITLE
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_p, "ULTRASOUND OBSTETRICS REPORT", bold=True, underline=True, size=14, color=RGBColor(0x1F, 0x4E, 0x79))
    
    # 5pt gap
    p_gap2 = doc.add_paragraph()
    p_gap2.paragraph_format.space_after = Pt(0)
    p_gap2.paragraph_format.line_spacing = Pt(5)
    
    lmp_p = doc.add_paragraph()
    lmp_p.paragraph_format.space_after = Pt(0)
    if data.get('lmp') == "Not Sure":
        add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=False, italic=True)
    else:
        add_run(lmp_p, f"LMP- {data.get('lmp', '')}", bold=True)
        
    # 5pt gap
    p_gap3 = doc.add_paragraph()
    p_gap3.paragraph_format.space_after = Pt(0)
    p_gap3.paragraph_format.line_spacing = Pt(5)

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
        
        # Combine Presentation and Lie
        pres_text = data.get('presentation', 'variable').lower()
        lie_text = data.get('lie', 'None')
        
        lie_map = {
            "LOA": "Left Occiputo-Anterior",
            "ROA": "Right Occiputo-Anterior",
            "LOP": "Left Occiputo-Posterior",
            "ROP": "Right Occiputo-Posterior"
        }
        
        if lie_text != "None":
            pres_text = f"{pres_text.capitalize()} presentation & {lie_map.get(lie_text, lie_text)} lie"
        else:
            pres_text = f"{pres_text.capitalize()} presentation"
            
        add_run(p, f"Uterus shows single live fetus. Fetus is in ")
        add_run(p, f"{pres_text}", bold=True, italic=True)
        add_run(p, f" at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        
        # 8pt gap
        p_gap4 = doc.add_paragraph()
        p_gap4.paragraph_format.space_after = Pt(0)
        p_gap4.paragraph_format.line_spacing = Pt(8)

        bio_head = doc.add_paragraph()
        add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        bio_head.paragraph_format.space_after = Pt(6)

        # --- NEW COMPACT COLORED BIOMETRY TABLE ---
        bio_table = doc.add_table(rows=6, cols=5)
        set_table_borders(bio_table)
        bio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # 1. Header Row
        headers = ["Fetal growth parameters", "mm", "Weeks", "Days", "Percentile"]
        widths = [Cm(4.5), Cm(1.5), Cm(1.5), Cm(1.5), Cm(5.0)] # Condensed widths
        for i, text in enumerate(headers):
            cell = bio_table.cell(0, i)
            cell.width = widths[i]
            set_cell_background(cell, get_gradient_color(0))
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
            add_run(p, text, bold=True, size=10)

        # 2. Data Rows
        params = ["bpd", "hc", "ac", "fl", "hl"]
        param_names = ["Biparietal Diameter", "Head circumference", "Abdominal Circumference", "Femoral Length", "Humerus Length"]
        
        for row_idx, param in enumerate(params):
            row_i = row_idx + 1 # Offset by 1 for header
            size = data.get(f"{param}_size", ""); wks = data.get(f"{param}_wks", ""); dys = data.get(f"{param}_dys", "")
            
            # Row background color
            bg_color = get_gradient_color(row_i)
            
            # Col 0: Parameter Name
            cell0 = bio_table.cell(row_i, 0)
            cell0.width = widths[0]
            set_cell_background(cell0, bg_color)
            add_run(cell0.paragraphs[0], param_names[row_idx], size=10)
            
            # Col 1: Size (mm)
            cell1 = bio_table.cell(row_i, 1)
            cell1.width = widths[1]
            set_cell_background(cell1, bg_color)
            cell1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cell1.paragraphs[0], size if size else "", size=10)
            
            # Col 2: Weeks
            cell2 = bio_table.cell(row_i, 2)
            cell2.width = widths[2]
            set_cell_background(cell2, bg_color)
            cell2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cell2.paragraphs[0], wks if wks else "", size=10)
            
            # Col 3: Days
            cell3 = bio_table.cell(row_i, 3)
            cell3.width = widths[3]
            set_cell_background(cell3, bg_color)
            cell3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cell3.paragraphs[0], dys if dys else "", size=10)
            
            # Col 4: Percentile Chart (Text next to the bar to save space)
            cell4 = bio_table.cell(row_i, 4)
            cell4.width = widths[4]
            set_cell_background(cell4, bg_color)
            cell4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            pct_val = data.get(f"{param}_pct", "")
            if pct_val and str(pct_val).strip() != "":
                try:
                    bar_img = generate_percentile_bar(pct_val)
                except:
                    bar_img = None
                if bar_img:
                    p = cell4.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.0
                    
                    # 1. Add the percentage text
                    add_run(p, f"{pct_val}% ", bold=True, size=9)
                    
                    # 2. Add the image (Reduced to 2.5cm to fit perfectly)
                    run = p.add_run()
                    run.add_picture(bar_img, width=Cm(2.5))
        # --- END BIOMETRY TABLE ---
                        
        # (Removed doc.add_paragraph() here to make them flush)

        # --- DYNAMIC RATIO TABLE BASED ON GESTATIONAL AGE ---
        import re
        mean_age_str = str(data.get('mean_age', '')).lower()
        
        # Find the digits immediately before "weeks"
        match = re.search(r'(\d+)\s*weeks', mean_age_str)
        
        if match:
            ga_weeks = int(match.group(1))
        else:
            ga_weeks = 0
        
        ratio_table = doc.add_table(rows=1, cols=2)
        ratio_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        ratio_table.columns[0].width = Cm(7.5); ratio_table.columns[1].width = Cm(8.5)
        
        if ga_weeks <= 22 and ga_weeks > 0:
            # Before 22 weeks
            add_run(ratio_table.cell(0, 0).paragraphs[0], "FL/HC= ", bold=True)
            add_run(ratio_table.cell(0, 0).paragraphs[0], f"{data.get('fl_hc', '')}", italic=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], "HC/AC = ", bold=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], f"{data.get('hc_ac', '')} (BOTH WNL)", italic=True)
        else:
            # 23 weeks and beyond
            add_run(ratio_table.cell(0, 0).paragraphs[0], "FL/AC= ", bold=True)
            add_run(ratio_table.cell(0, 0).paragraphs[0], f"{data.get('fl_ac', '')}", italic=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], "FL/BPD = ", bold=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], f"{data.get('fl_bpd', '')} (BOTH WNL)", italic=True)
        
        # (Removed doc.add_paragraph() here to make them flush)

        p = doc.add_paragraph()
        add_run(p, "MEAN AGE : ", bold=True); add_run(p, f"{data.get('mean_age', '')}", italic=True)
        # --- FETAL WEIGHT WITH BAR ---
        p = doc.add_paragraph()
        add_run(p, "FETAL WEIGHT : ", bold=True)
        
        # Calculate the ± 15% error margin
        fw_str = f"{data.get('fetal_weight', '')}g"
        try:
            fw_val = float(data.get('fetal_weight', ''))
            error_margin = int(round(fw_val * 0.15))
            fw_str = f"{int(fw_val)} GRAMS ± {error_margin}GMS"
        except:
            pass # If no weight is entered, just leave it blank
            
        add_run(p, fw_str, italic=True)
        if data.get("fetal_weight_pct"):
            p.add_run("  (Percentile: ")
            p.add_run(f"{data.get('fetal_weight_pct')}%)")
            bar_img_efw = generate_percentile_bar(data.get("fetal_weight_pct"))
            if bar_img_efw:
                p.add_run().add_picture(bar_img_efw, width=Cm(3.0))

        p = doc.add_paragraph()
        add_run(p, "AFI:- ", bold=True)
        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        
        # We must define is_abn here so the bolding works!
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', '')
            desc, is_abn = generate_afi_text(afi_val, is_total=True)
            if desc == "Nil, Not a single pocket of even one cm seen":
                afi_str = desc
            else:
                afi_str = f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', '')
            try:
                afi_float = float(afi_val)
                if 2.0 <= afi_float <= 8.0:
                    afi_str = f"Adequate, largest pocket measures {afi_val} cm"
                    is_abn = False
                elif afi_float < 2.0:
                    afi_str = f"Reduced, largest pocket measures {afi_val} cm"
                    is_abn = True
                else:
                    afi_str = f"Increased, largest pocket measures {afi_val} cm"
                    is_abn = True
            except:
                afi_str = f"largest pocket measures {afi_val} cm"
                is_abn = False
                
        add_run(p, afi_str, bold=is_abn, italic=is_abn)
        p = doc.add_paragraph()
        add_run(p, "Foetal heart rate ", bold=True); add_run(p, f"is approx. {data.get('fhr', '')}.", italic=True)
        doc.add_paragraph()

        # --- EDD TABLE ---
        # 1. Safely get the EDDs. If they are empty, use a dash.
        _lmp_val = data.get('lmp', '')
        _edd_lmp = data.get('edd_lmp', '')
        _edd_cua = data.get('edd_cua', '')
        
        # 2. Determine if we need 1 or 2 rows based on LMP status
        if _lmp_val == "Not Sure":
            edd_data = [("EDD (CUA)--", _edd_cua, "±2W")]
            num_rows = 1
        else:
            edd_data = [
                ("EDD (LMP)--", _edd_lmp, "±2W"),
                ("EDD (CUA)--", _edd_cua, "±2W"),
            ]
            num_rows = 2

        # 3. Build the table
        edd_table = doc.add_table(rows=num_rows, cols=3)
        set_table_borders(edd_table)
        
        # 4. Set widths
        for row in edd_table.rows:
            row.cells[0].width = Cm(3.5)
            row.cells[1].width = Cm(5.0)
            row.cells[2].width = Cm(3.0)
            
        # 5. Fill the cells with formatting
        for i, (label, date, err) in enumerate(edd_data):
            cell_0 = edd_table.cell(i, 0)
            cell_1 = edd_table.cell(i, 1)
            cell_2 = edd_table.cell(i, 2)
            
            # Clear default paragraph
            cell_0.text = ""
            cell_1.text = ""
            cell_2.text = ""
            
            add_run(cell_0.paragraphs[0], label, bold=True, italic=True, size=10)
            add_run(cell_1.paragraphs[0], str(date), italic=True, size=10)
            add_run(cell_2.paragraphs[0], err, italic=True, size=10)
            
        # 5. Add a gap after the table so it doesn't collide with the Doppler section
        doc.add_paragraph()
        # --- END EDD TABLE ---

        # --- DOPPLER SECTION (GREEN & COMPACT) ---
        if data.get("umbilical_pi") or data.get("uterine_pi") or data.get("mca_pi"):
            
            # Values are now correctly passed via data dict, thanks to the UI boxes
            pass 
            # -----------------------------------------------------------
            doc.add_paragraph()
            doppler_head = doc.add_paragraph()
            add_run(doppler_head, "COLOR DOPPLER INDICES:", bold=True, underline=True)
            
            # Create a 7-column table (Added CPR)
            dop_table = doc.add_table(rows=1, cols=7)
            set_table_borders(dop_table)
            dop_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Column widths (Compact)
            dop_widths = [Cm(3.8), Cm(1.2), Cm(1.2), Cm(1.2), Cm(4.5), Cm(3.0), Cm(2.5)]
            headers = ["Vessels", "S/D", "RI", "PI", "PI Percentile", "CPR", "Remarks"]
            
            # 1. Header Row (Green Gradient)
            for i, text in enumerate(headers):
                cell = dop_table.cell(0, i)
                cell.width = dop_widths[i]
                set_cell_background(cell, get_gradient_color(0)) # Green Header
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
                add_run(cell.paragraphs[0], text, bold=True, size=10)
            
            row_idx = 1
            
            def add_dop_row(vessel, sd, ri, pi, pct, cpr, remark):
                nonlocal row_idx
                row_cells = dop_table.add_row().cells
                bg_color = get_gradient_color(row_idx) # Green Gradient Body
                
                vals = [vessel, sd, ri, pi, pct, cpr, remark]
                for i, val in enumerate(vals):
                    cell = row_cells[i]
                    cell.width = dop_widths[i]
                    set_cell_background(cell, bg_color)
                    
                    if i == 0 or i == 6: # Left align Vessel and Remarks
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                    if i == 4 and pct: # PI Percentile Bar
                        try:
                            bar_img = generate_percentile_bar(pct)
                            if bar_img:
                                p = cell.paragraphs[0]
                                p.paragraph_format.space_after = Pt(0)
                                p.paragraph_format.line_spacing = 1.0
                                add_run(p, f"{pct}% ", bold=True, size=9)
                                p.add_run().add_picture(bar_img, width=Cm(2.5))
                        except: pass
                    else:
                        # FORCE it to print the exact value from the data dict
                        _val_to_print = str(val) if val is not None else "-"
                        if _val_to_print.strip() == "":
                            _val_to_print = "-"
                        add_run(cell.paragraphs[0], _val_to_print, size=10)
                        
                row_idx += 1
            
            # --- Populate Rows ---
            # 1. Umbilical Artery
            if data.get("umbilical_pi"):
                _umb_sd = sr_data.get("umbilical_sd", data.get("umbilical_sd", ""))
                _umb_ri = sr_data.get("umbilical_ri", data.get("umbilical_ri", ""))
                
                add_dop_row(
                    "Umbilical Artery", 
                    _umb_sd, 
                    _umb_ri, 
                    data.get("umbilical_pi", ""), 
                    data.get("umbilical_pct", ""), 
                    "", # CPR left blank here, calculated on MCA row
                    "Within Normal Limit"
                )
            
            # 2. Right Uterine Artery
            if data.get("uterine_right_pi"):
                add_dop_row(
                    "Right Uterine Artery", 
                    data.get("uterine_right_sd", ""), 
                    data.get("uterine_right_ri", ""), 
                    data.get("uterine_right_pi", ""), 
                    data.get("uterine_right_pct", ""), 
                    "", 
                    "Within Normal Limit"
                )
                
            # 3. Left Uterine Artery
            if data.get("uterine_left_pi"):
                add_dop_row(
                    "Left Uterine Artery", 
                    data.get("uterine_left_sd", ""), 
                    data.get("uterine_left_ri", ""), 
                    data.get("uterine_left_pi", ""), 
                    data.get("uterine_left_pct", ""), 
                    "", 
                    "Within Normal Limit"
                )
            
            # 4. Uterine Artery (Check if we have both sides to call it "Mean")
            if data.get("uterine_pi"):
                if data.get("uterine_left_pi") and data.get("uterine_right_pi"):
                    _ut_label = "Uterine Arteries\nMean PI = "
                else:
                    _ut_label = "Uterine Artery" # Fallback if only one was sent
                    
                add_dop_row(
                    _ut_label, 
                    data.get("uterine_left_sd", ""), # Will be blank if not sent
                    data.get("uterine_left_ri", ""), # Will be blank if not sent
                    data.get("uterine_pi", ""), 
                    data.get("uterine_pct", ""), 
                    "", 
                    "Normal"
                )
                
            # 5. MCA & CPR
            if data.get("mca_pi"):
                # Calculate CPR
                cpr_val = ""
                cpr_pct = ""
                if data.get("umbilical_pi"):
                    try:
                        umb = float(data.get("umbilical_pi"))
                        mca = float(data.get("mca_pi"))
                        if umb > 0:
                            cpr_float = mca / umb
                            cpr_val = f"{cpr_float:.2f}"
                            # Simple CPR threshold for remark
                            if cpr_float < 1.0:
                                cpr_remark = "Low CPR"
                            else:
                                cpr_remark = "Normal CPR"
                    except:
                        pass

                add_dop_row(
                    "Middle Cerebral Artery", 
                    data.get("mca_sd", ""), 
                    data.get("mca_ri", ""), 
                    data.get("mca_pi", ""), 
                    data.get("mca_pct", ""), 
                    cpr_val, 
                    "Within Normal Limit"
                )

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
        pres_text = data.get('presentation', 'variable').lower()
        lie_text = data.get('lie', 'None')
        if lie_text != "None":
            pres_text += f", {lie_text}"
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {pres_text} presentation of approx. {mean_age} gestation age."
    p = doc.add_paragraph(); add_run(p, imp_text, bold=True, italic=True)

    # 6. PNDT & NOTE COMBINED BOX
    doc.add_paragraph()
    disc_table = doc.add_table(rows=1, cols=1)
    disc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(disc_table)
    cell = disc_table.cell(0, 0)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # First Line: PNDT Disclaimer
    add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=8)
    
    # Add a line break, then the Note
    p.add_run().add_break()
    add_run(p, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=8)

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
        # Combine Presentation and Lie
        pres_text = data.get('presentation', 'variable').lower()
        lie_text = data.get('lie', 'None')
        
        lie_map = {
            "LOA": "Left Occiputo-Anterior",
            "ROA": "Right Occiputo-Anterior",
            "LOP": "Left Occiputo-Posterior",
            "ROP": "Right Occiputo-Posterior"
        }
        
        if lie_text != "None":
            pres_text = f"{pres_text.capitalize()} presentation & {lie_map.get(lie_text, lie_text)} lie"
        else:
            pres_text = f"{pres_text.capitalize()} presentation"
            
        lines.append(f"Uterus shows single live fetus. Fetus is in {pres_text} at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        lines.append("")
        lines.append("BIO PHYSICAL PARAMETERS:")
        for param in ["bpd", "hc", "ac", "fl", "hl"]:
            size = data.get(f"{param}_size", ""); wks = data.get(f"{param}_wks", ""); dys = data.get(f"{param}_dys", "")
            lines.append(f"  {param.upper():<4}: {(str(size) + 'mm') if size else '':<10} {(str(wks) + 'weeks') if wks else '':<12} {(str(dys) + 'days±2w') if dys else ''}")
        lines.append("")
        # --- DYNAMIC RATIO PREVIEW BASED ON GESTATIONAL AGE ---
        import re
        _mean_age_str = str(data.get('mean_age', '')).lower()
        
        # Find the digits immediately before "weeks"
        _match = re.search(r'(\d+)\s*weeks', _mean_age_str)
        
        if _match:
            _ga_weeks = int(_match.group(1))
        else:
            _ga_weeks = 0
            
        # If GA is 22 or less, use FL/HC & HC/AC. Otherwise (or if unknown), use FL/AC & FL/BPD
        if 0 < _ga_weeks <= 22:
            lines.append(f"FL/HC= {data.get('fl_hc', '')}        HC/AC = {data.get('hc_ac', '')} (BOTH WNL)")
        else:
            lines.append(f"FL/AC= {data.get('fl_ac', '')}        FL/BPD = {data.get('fl_bpd', '')} (BOTH WNL)")
        lines.append("")
        lines.append(f"MEAN AGE : {data.get('mean_age', '')}")
        
        # FETAL WEIGHT PREVIEW
        fw_str_prev = f"{data.get('fetal_weight', '')}g"
        try:
            fw_val_prev = float(data.get('fetal_weight', ''))
            err_prev = int(round(fw_val_prev * 0.15))
            fw_str_prev = f"{int(fw_val_prev)} GRAMS ± {err_prev}GMS"
        except:
            pass
            
        fw_line = f"FETAL WEIGHT : {fw_str_prev}"
        if data.get("fetal_weight_pct"):
            fw_line += f" (Percentile: {data.get('fetal_weight_pct')}%)"
        lines.append(fw_line)

        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', '')
            desc, _ = generate_afi_text(afi_val, is_total=True)
            if desc == "Nil, Not a single pocket of even one cm seen":
                afi_str = desc
            else:
                afi_str = f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', '')
            try:
                afi_float = float(afi_val)
                if 2.0 <= afi_float <= 8.0:
                    afi_str = f"Adequate, largest pocket measures {afi_val} cm"
                elif afi_float < 2.0:
                    afi_str = f"Reduced, largest pocket measures {afi_val} cm"
                else:
                    afi_str = f"Increased, largest pocket measures {afi_val} cm"
            except:
                afi_str = f"largest pocket measures {afi_val} cm"
        lines.append(f"AFI:- {afi_str}")
        lines.append(f"Foetal heart rate is approx. {data.get('fhr', '')}.")
        lines.append("")

        # --- EDD PREVIEW FIX ---
        _lmp_val = data.get('lmp', '')
        _edd_lmp = data.get('edd_lmp', '')
        _edd_cua = data.get('edd_cua', '')
        
        if _lmp_val != 'Not Sure':
            lines.append(f"  {'EDD (LMP)--':<18}: {_edd_lmp}    ±2W")
            
        lines.append(f"  {'EDD (CUA)--':<18}: {_edd_cua}    ±2W")
        lines.append("")
        # -----------------------
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

        if template_id == "T3" and data.get('measure_scar'): 
            lines.append(f"Scar Thickness: {data.get('scar_size', '')} mm")

        # --- DOPPLER SECTION PREVIEW ---
        if data.get("umbilical_pi") or data.get("uterine_pi") or data.get("mca_pi"):
            lines.append("")
            lines.append("DOPPLER INDICES:")
            
            if data.get("umbilical_pi"):
                ua_str = f"  Umbilical Artery: PI = {data.get('umbilical_pi')}"
                if data.get("umbilical_pct"): ua_str += f" (Percentile: {data.get('umbilical_pct')}%)"
                lines.append(ua_str)
                
            if data.get("uterine_pi"):
                notching = data.get("uterine_notching", "No notching")
                uta_str = f"  Uterine Artery: PI = {data.get('uterine_pi')} ({notching})"
                if data.get("uterine_pct"): uta_str += f" (Percentile: {data.get('uterine_pct')}%)"
                lines.append(uta_str)
                
            if data.get("mca_pi"):
                lines.append(f"  MCA: PI = {data.get('mca_pi')}")

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
        pres_text = data.get('presentation', 'variable').lower()
        lie_text = data.get('lie', 'None')
        if lie_text != "None":
            pres_text += f", {lie_text}"
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {pres_text} presentation of approx. {mean_age} gestation age."
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

# --- NEW: CHECK LISTENER JSON AT THE VERY TOP ---
import json
import os
json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'latest_sr.json')

if os.path.exists(json_path):
    try:
        mtime = os.path.getmtime(json_path)
        if st.session_state.get("last_json_mtime") != mtime:
            with open(json_path, 'r') as f:
                new_data = json.load(f)
            st.session_state["sr_data"] = new_data
            st.session_state["last_json_mtime"] = mtime
            st.session_state["dicom_version"] = st.session_state.get("dicom_version", 0) + 1
            st.toast("✅ Loaded fresh data from DICOM Listener!")
    except Exception as e:
        pass
# -----------------------------------------------

v = st.session_state.get("dicom_version", 0)
sr_data = st.session_state.get("sr_data", {})

# FORCE NAME/AGE INTO SESSION STATE IMMEDIATELY
if 'patient_name' in sr_data: st.session_state[f"sr_patient_name_{v}"] = sr_data['patient_name']
if 'patient_age' in sr_data: st.session_state[f"sr_patient_age_{v}"] = sr_data['patient_age']
if 'fetal_weight_pct' in sr_data: st.session_state[f"sr_fetal_weight_pct_edit_{v}"] = sr_data['fetal_weight_pct']

c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3])
with c1: p_name = st.text_input("Name", value=st.session_state.get(f"sr_patient_name_{v}", ""), key=f"p_name_input_{v}")
with c2: p_age = st.text_input("Age", value=st.session_state.get(f"sr_patient_age_{v}", ""), key=f"p_age_input_{v}")
with c3: p_sex = st.radio("Sex", ["F", "M"], horizontal=True)
with c4: p_date = st.text_input("Date", value=datetime.now().strftime("%d-%b-%y"))
with c5: p_ref = st.text_input("Referred by")

data = {"patient_name": p_name, "patient_age": p_age, "report_date": p_date, "referred_by": p_ref, "p_sex": p_sex}

# --- NEW PATIENT BUTTON ---
if st.button("🚀 Start New Patient / Clear Form", type="primary"):
    # Clear the DICOM data
    st.session_state["sr_data"] = {}
    st.session_state["dicom_version"] += 1
    # Clear the input fields by deleting their session state keys
    for key in list(st.session_state.keys()):
        if key.startswith("sr_") or key in ["p_name_input", "p_age_input"]:
            del st.session_state[key]
    st.rerun()
# ---------------------------

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    template_id = st.selectbox("Template", ["EARLY_OBS", "NT_NB", "T16", "T3"])
    
    # --- DUAL DICOM INPUT (Manual Upload + Listener) ---
    import json
    import os
    
    st.markdown("### 🔗 DICOM Input")
    if "sr_data" in st.session_state:
        st.write("DEBUG: Parsed SR Data:", st.session_state["sr_data"])
    # Initialize the version bumper
    if "dicom_version" not in st.session_state:
        st.session_state["dicom_version"] = 0
        
    col_up, col_ref = st.columns([2, 1])
    
    with col_up:
        uploaded_sr = st.file_uploader("Auto-Fill from File", type=["dcm", "dicom", "ima", "json"], label_visibility="collapsed")
        if uploaded_sr is not None:
            try:
                if uploaded_sr.name.endswith('.json'):
                    st.session_state["sr_data"] = json.load(uploaded_sr)
                else:
                    ds = pydicom.dcmread(io.BytesIO(uploaded_sr.read()))
                    st.session_state["sr_data"] = extract_sr_data(ds)
                
                st.session_state["dicom_version"] += 1  # Force refresh
                st.session_state["last_json_mtime"] = None # Ignore old listener file
                st.success("DICOM SR File Loaded (Manual)")
            except Exception as e:
                st.error(f"Error reading SR: {e}")
                     
    with col_ref:
        if st.button("🔄 Refresh from Listener", use_container_width=True, key="refresh_btn"):
            st.session_state["last_json_mtime"] = None  # Forces the app to look at the file again
            st.rerun()
            
    # Push values into session_state widgets with versioning
    sr_data = st.session_state.get("sr_data", {})
    v = st.session_state["dicom_version"]
    
    if "last_applied_version" not in st.session_state:
        st.session_state["last_applied_version"] = -1
        
    apply_new_data = st.session_state["last_applied_version"] != v
    
    if apply_new_data:
        # --- NEW: Auto-fill Patient Info ---
        if 'patient_name' in sr_data: st.session_state[f"sr_patient_name_{v}"] = sr_data['patient_name']
        if 'patient_age' in sr_data: st.session_state[f"sr_patient_age_{v}"] = sr_data['patient_age']
        
        if 'bpd' in sr_data: st.session_state[f"sr_bpd_size_{v}"] = sr_data['bpd']
        if 'hc' in sr_data: st.session_state[f"sr_hc_size_{v}"] = sr_data['hc']
        if 'ac' in sr_data: st.session_state[f"sr_ac_size_{v}"] = sr_data['ac']
        if 'fl' in sr_data: st.session_state[f"sr_fl_size_{v}"] = sr_data['fl']
        if 'hl' in sr_data: st.session_state[f"sr_hl_size_{v}"] = sr_data['hl']
        
        for p in ["bpd", "hc", "ac", "fl", "hl"]:
            if f"{p}_wks" in sr_data: st.session_state[f"sr_{p}_wks_{v}"] = sr_data[f"{p}_wks"]
            if f"{p}_dys" in sr_data: st.session_state[f"sr_{p}_dys_{v}"] = sr_data[f"{p}_dys"]
            if f"{p}_pct" in sr_data: st.session_state[f"sr_{p}_pct_{v}"] = sr_data[f"{p}_pct"]

        if 'msd' in sr_data: st.session_state[f"sr_msd_{v}"] = sr_data['msd']
        if 'crl' in sr_data: st.session_state[f"sr_crl_{v}"] = sr_data['crl']
        if 'nt' in sr_data: st.session_state[f"sr_nt_{v}"] = sr_data['nt']
        if 'nasal_bone' in sr_data: st.session_state[f"sr_nasal_bone_{v}"] = sr_data['nasal_bone']
        if 'fetal_weight' in sr_data: st.session_state[f"sr_fetal_weight_{v}"] = sr_data['fetal_weight']
        if 'fetal_weight_pct' in sr_data: st.session_state[f"sr_fetal_weight_pct_edit_{v}"] = sr_data['fetal_weight_pct']
        if 'fhr' in sr_data: st.session_state[f"sr_fhr_{v}"] = sr_data['fhr']
        if 'mean_age' in sr_data: st.session_state[f"sr_mean_age_{v}"] = sr_data['mean_age']
        
        if 'edd_lmp' in sr_data: st.session_state[f"sr_edd_lmp_{v}"] = sr_data['edd_lmp']
        if 'edd_cua' in sr_data: st.session_state[f"sr_edd_cua_{v}"] = sr_data['edd_cua']
        if 'lmp' in sr_data: st.session_state[f"sr_lmp_{v}"] = sr_data['lmp']
        
        if 'ga_lmp_days' in sr_data: 
            st.session_state[f"sr_ga_lmp_{v}"] = f"{sr_data['ga_lmp_days']//7}w{sr_data['ga_lmp_days']%7}d"

        if 'umbilical_pi' in sr_data: st.session_state[f"sr_umbilical_pi_{v}"] = sr_data['umbilical_pi']
        if 'uterine_pi' in sr_data: st.session_state[f"sr_uterine_pi_{v}"] = sr_data['uterine_pi']
        
        # --- NEW: Pass S/D and RI values to the UI ---
        if 'umbilical_sd' in sr_data: st.session_state[f"sr_umbilical_sd_{v}"] = sr_data['umbilical_sd']
        if 'umbilical_ri' in sr_data: st.session_state[f"sr_umbilical_ri_{v}"] = sr_data['umbilical_ri']
        
        if 'uterine_sd' in sr_data: st.session_state[f"sr_uterine_sd_{v}"] = sr_data['uterine_sd']
        if 'uterine_ri' in sr_data: st.session_state[f"sr_uterine_ri_{v}"] = sr_data['uterine_ri']
        
        if 'uterine_left_pi' in sr_data: st.session_state[f"sr_uterine_left_pi_{v}"] = sr_data['uterine_left_pi']
        if 'uterine_left_sd' in sr_data: st.session_state[f"sr_uterine_left_sd_{v}"] = sr_data['uterine_left_sd']
        if 'uterine_left_ri' in sr_data: st.session_state[f"sr_uterine_left_ri_{v}"] = sr_data['uterine_left_ri']

        if 'uterine_right_pi' in sr_data: st.session_state[f"sr_uterine_right_pi_{v}"] = sr_data['uterine_right_pi']
        if 'uterine_right_sd' in sr_data: st.session_state[f"sr_uterine_right_sd_{v}"] = sr_data['uterine_right_sd']
        if 'uterine_right_ri' in sr_data: st.session_state[f"sr_uterine_right_ri_{v}"] = sr_data['uterine_right_ri']

        if 'mca_pi' in sr_data: st.session_state[f"sr_mca_pi_{v}"] = sr_data['mca_pi']
        if 'mca_sd' in sr_data: st.session_state[f"sr_mca_sd_{v}"] = sr_data['mca_sd']
        if 'mca_ri' in sr_data: st.session_state[f"sr_mca_ri_{v}"] = sr_data['mca_ri']
        # -----------------------------------------------

        st.session_state["last_applied_version"] = v
    # --------------------------
    lmp_not_sure = st.checkbox("LMP Not Sure / Not Known")
    if lmp_not_sure:
        data["lmp"] = "Not Sure"
    else:
        data["lmp"] = st.text_input("LMP", value=st.session_state.get(f"sr_lmp_{v}", ""))

    with st.expander("🔗 DICOM SR Data (Machine Generated)", expanded=False):
        st.markdown('<div class="dicom-box">Verify or override data received from the ultrasound machine.</div>', unsafe_allow_html=True)
        
        if template_id != "EARLY_OBS":
            # Timeline Header
            st.markdown(f"**LMP:** {st.session_state.get(f'sr_lmp_{v}', '')} &nbsp;|&nbsp; "
                        f"**GA(LMP):** {st.session_state.get(f'sr_ga_lmp_{v}', '')} &nbsp;|&nbsp; "
                        f"**AUA:** {st.session_state.get(f'sr_mean_age_{v}', '')} &nbsp;|&nbsp; "
                        f"**EDD(LMP):** {st.session_state.get(f'sr_edd_lmp_{v}', '')}")
            st.markdown("---")

        if template_id == "EARLY_OBS":
            data["gsac"] = st.text_input("G-Sac", value=st.session_state.get(f"sr_gsac_{v}", "Seen"), key=f"sr_gsac_{v}")
            data["msd"] = st.text_input("MSD", value=st.session_state.get(f"sr_msd_{v}", ""), key=f"sr_msd_{v}")
            data["crl"] = st.text_input("CRL", value=st.session_state.get(f"sr_crl_{v}", ""), key=f"sr_crl_{v}")
            data["ga_crl"] = st.text_input("GA by CRL", value="")
            data["yolk"] = st.text_input("Yolk sac", value="")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
        elif template_id == "NT_NB":
            b1, b2 = st.columns(2)
            data["crl"] = b1.text_input("CRL (mm)", value=st.session_state.get(f"sr_crl_{v}", ""), key=f"sr_crl_{v}")
            data["nt"] = b2.text_input("NT (mm)", value=st.session_state.get(f"sr_nt_{v}", ""), key=f"sr_nt_{v}")
            
            b1, b2 = st.columns(2)
            data["nasal_bone"] = b1.text_input("Nasal Bone", value=st.session_state.get(f"sr_nasal_bone_{v}", ""), key=f"sr_nasal_bone_{v}")
            data["fhr_val"] = b2.text_input("FHR (bpm)", value=st.session_state.get(f"sr_fhr_{v}", ""), key=f"sr_fhr_{v}")
            
            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
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
                
                data[f"{p_key}_size"] = c2.text_input(f"{p_key} size", value=st.session_state.get(f"sr_{p_key}_size_{v}", ""), key=f"sr_{p_key}_size_{v}", label_visibility="collapsed")
                data[f"{p_key}_wks"] = c3.text_input(f"{p_key} wks", value=st.session_state.get(f"sr_{p_key}_wks_{v}", ""), key=f"sr_{p_key}_wks_{v}", label_visibility="collapsed")
                data[f"{p_key}_dys"] = c4.text_input(f"{p_key} dys", value=st.session_state.get(f"sr_{p_key}_dys_{v}", ""), key=f"sr_{p_key}_dys_{v}", label_visibility="collapsed")
                
                # Make the percentile an editable text box, pre-filled with the DICOM value
                pct_val = c5.text_input(f"{p_key} pct", value=st.session_state.get(f"sr_{p_key}_pct_{v}", ""), key=f"sr_{p_key}_pct_edit_{v}", label_visibility="collapsed")
                
                if pct_val:
                    bar_img = generate_percentile_bar(pct_val)
                    if bar_img:
                        c5.image(bar_img, use_container_width=True)
            st.markdown("---")
            st.markdown("**Ratios & Averages**")
            b1, b2 = st.columns(2)
            
            try:
                if data.get("fl_size") and data.get("hc_size"):
                    fl_hc_val = (float(data["fl_size"]) / float(data["hc_size"])) * 100
                    st.session_state[f"sr_fl_hc_{v}"] = f"{fl_hc_val:.1f}%"
                if data.get("hc_size") and data.get("ac_size"):
                    hc_ac_val = float(data["hc_size"]) / float(data["ac_size"])
                    st.session_state[f"sr_hc_ac_{v}"] = f"{hc_ac_val:.2f}"
                if data.get("fl_size") and data.get("bpd_size"):
                    fl_bpd_val = (float(data["fl_size"]) / float(data["bpd_size"])) * 100
                    st.session_state[f"sr_fl_bpd_{v}"] = f"{fl_bpd_val:.1f}%"
                if data.get("fl_size") and data.get("ac_size"):
                    fl_ac_val = (float(data["fl_size"]) / float(data["ac_size"])) * 100
                    st.session_state[f"sr_fl_ac_{v}"] = f"{fl_ac_val:.1f}%"
            except:
                pass

            # --- CRITICAL FIX: Pull calculated percentiles from session_state into data dict ---
            for p in ["bpd", "hc", "ac", "fl", "hl"]:
                pct_box_val = st.session_state.get(f"sr_{p}_pct_edit_{v}", "")
                
                # If LMP is "Not Sure", we send empty data to the document/preview
                if lmp_not_sure:
                    data[f"{p}_pct"] = ""
                elif pct_box_val:
                    data[f"{p}_pct"] = pct_box_val

            data["fl_hc"] = b1.text_input("FL/HC", value=st.session_state.get(f"sr_fl_hc_{v}", ""), key=f"sr_fl_hc_{v}")
            data["hc_ac"] = b2.text_input("HC/AC", value=st.session_state.get(f"sr_hc_ac_{v}", ""), key=f"sr_hc_ac_{v}")
            
            b1, b2 = st.columns(2)
            data["fl_bpd"] = b1.text_input("FL/BPD", value=st.session_state.get(f"sr_fl_bpd_{v}", ""), key=f"sr_fl_bpd_{v}")
            data["fl_ac"] = b2.text_input("FL/AC", value=st.session_state.get(f"sr_fl_ac_{v}", ""), key=f"sr_fl_ac_{v}")

            b1, b2 = st.columns(2)
            data["mean_age"] = b1.text_input("Mean Age", value=st.session_state.get(f"sr_mean_age_{v}", ""), key=f"sr_mean_age_{v}")
            data["fetal_weight"] = b2.text_input("Fetal Weight (g)", value=st.session_state.get(f"sr_fetal_weight_{v}", ""), key=f"sr_fetal_weight_{v}")
            
            # EFW Percentile Fix
            # Streamlit doesn't update the data dictionary automatically when you change a widget with a key. 
            # We have to pull it directly from session_state right before building the doc.
            efw_pct_val = b2.text_input("EFW Percentile (%)", value=st.session_state.get(f"sr_fetal_weight_pct_{v}", ""), key=f"sr_fetal_weight_pct_edit_{v}")
            
            if lmp_not_sure:
                data["fetal_weight_pct"] = ""
            else:
                data["fetal_weight_pct"] = efw_pct_val
            if efw_pct_val:
                bar_efw = generate_percentile_bar(efw_pct_val)
                if bar_efw: b2.image(bar_efw, use_container_width=True)
            b1, b2 = st.columns(2)
            data["fhr"] = b1.text_input("FHR", value=st.session_state.get(f"sr_fhr_{v}", ""), key=f"sr_fhr_{v}")
            data["edd_lmp"] = b2.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
            
            # --- DOPPLER CALCULATIONS ---
            # 1. Calculate Mean Uterine PI
            ut_pi_left = st.session_state.get(f"sr_uterine_pi_{v}", "")
            ut_pi_right = st.session_state.get(f"sr_uterine_pi_right_{v}", "")
            
            mean_ut_pi = ut_pi_left
            if ut_pi_left and ut_pi_right:
                try:
                    mean_val = (float(ut_pi_left) + float(ut_pi_right)) / 2.0
                    mean_ut_pi = f"{mean_val:.2f}"
                except:
                    pass
                    
            # 2. Check if MCA is present in DICOM
            mca_present = 'mca_pi' in sr_data
            mca_val = sr_data.get('mca_pi', "")

            st.markdown("**Doppler Indices & Percentiles**")
            
            if mca_present:
                b1, b2, b3 = st.columns(3)
                data["umbilical_pi"] = b1.text_input("Umbilical PI", value=st.session_state.get(f"sr_umbilical_pi_{v}", ""), key=f"sr_umbilical_pi_{v}")
                data["uterine_pi"] = b2.text_input("Mean Uterine PI", value=mean_ut_pi, key=f"sr_uterine_pi_mean_{v}")
                data["mca_pi"] = b3.text_input("MCA PI", value=mca_val, key=f"sr_mca_pi_{v}")
            else:
                b1, b2 = st.columns(2)
                data["umbilical_pi"] = b1.text_input("Umbilical PI", value=st.session_state.get(f"sr_umbilical_pi_{v}", ""), key=f"sr_umbilical_pi_{v}")
                data["uterine_pi"] = b2.text_input("Mean Uterine PI", value=mean_ut_pi, key=f"sr_uterine_pi_mean_{v}")
                data["mca_pi"] = "" # Ensure it stays blank if not present

            # --- THE ACTUAL FIX: REGISTER S/D AND RI SO STREAMLIT DOESNT LOSE THEM ---
            st.markdown("*Doppler Waveform Indices (auto-filled)*")
            w1, w2, w3, w4 = st.columns(4)
            data["umbilical_sd"] = w1.text_input("Umb S/D", value=st.session_state.get(f"sr_umbilical_sd_{v}", ""), key=f"sr_umbilical_sd_{v}")
            data["umbilical_ri"] = w2.text_input("Umb RI", value=st.session_state.get(f"sr_umbilical_ri_{v}", ""), key=f"sr_umbilical_ri_{v}")
            data["uterine_sd"] = w3.text_input("Ut S/D", value=st.session_state.get(f"sr_uterine_sd_{v}", ""), key=f"sr_uterine_sd_{v}")
            data["uterine_ri"] = w4.text_input("Ut RI", value=st.session_state.get(f"sr_uterine_ri_{v}", ""), key=f"sr_uterine_ri_{v}")
            
            # Same for MCA if present
            if mca_present:
                data["mca_sd"] = st.text_input("MCA S/D", value=st.session_state.get(f"sr_mca_sd_{v}", ""), key=f"sr_mca_sd_{v}")
                data["mca_ri"] = st.text_input("MCA RI", value=st.session_state.get(f"sr_mca_ri_{v}", ""), key=f"sr_mca_ri_{v}")
            else:
                data["mca_sd"] = ""
                data["mca_ri"] = ""
            
            # --- NEW CRITICAL FIX: PUSH S/D AND RI DIRECTLY FROM SR_DATA ---
            data["umbilical_sd"] = sr_data.get('umbilical_sd', st.session_state.get(f"sr_umbilical_sd_{v}", ""))
            data["umbilical_ri"] = sr_data.get('umbilical_ri', st.session_state.get(f"sr_umbilical_ri_{v}", ""))
            
            data["uterine_left_pi"] = sr_data.get('uterine_left_pi', "")
            data["uterine_left_sd"] = sr_data.get('uterine_left_sd', "")
            data["uterine_left_ri"] = sr_data.get('uterine_left_ri', "")
            
            data["uterine_right_pi"] = sr_data.get('uterine_right_pi', "")
            data["uterine_right_sd"] = sr_data.get('uterine_right_sd', "")
            data["uterine_right_ri"] = sr_data.get('uterine_right_ri', "")
            
            data["uterine_sd"] = sr_data.get('uterine_sd', st.session_state.get(f"sr_uterine_sd_{v}", ""))
            data["uterine_ri"] = sr_data.get('uterine_ri', st.session_state.get(f"sr_uterine_ri_{v}", ""))
            
            data["mca_sd"] = sr_data.get('mca_sd', st.session_state.get(f"sr_mca_sd_{v}", ""))
            data["mca_ri"] = sr_data.get('mca_ri', st.session_state.get(f"sr_mca_ri_{v}", ""))
            # ---------------------------------------------------------------
                        
            data["uterine_notching"] = st.selectbox("Uterine Artery Notching", ["No notching", "Unilateral Notching", "Bilateral Notching"])

            data["ga_weeks_dop"] = st.number_input("GA (weeks) for Percentiles", min_value=10.0, max_value=42.0, value=28.0, step=0.1)

            b1, b2, b3 = st.columns(3)
            if data.get("umbilical_pi"):
                try:
                    res = calculate_ua_pi(float(data["umbilical_pi"]), data["ga_weeks_dop"])
                    data["umbilical_pct"] = str(res['percentile'])
                    b1.caption(f"Umbilical: {res['percentile']}th")
                    bar_ua = generate_percentile_bar(res['percentile'])
                    if bar_ua: b1.image(bar_ua, use_container_width=True)
                except: pass
                
            # Calculate Left Uterine PI Percentile
            if data.get("uterine_left_pi"):
                try:
                    res_left = calculate_uta_pi(float(data["uterine_left_pi"]), int(round(data["ga_weeks_dop"])))
                    if res_left:
                        data["uterine_left_pct"] = str(res_left['percentile'])
                except: pass

            # Calculate Right Uterine PI Percentile
            if data.get("uterine_right_pi"):
                try:
                    res_right = calculate_uta_pi(float(data["uterine_right_pi"]), int(round(data["ga_weeks_dop"])))
                    if res_right:
                        data["uterine_right_pct"] = str(res_right['percentile'])
                except: pass

            # Calculate Mean Uterine PI Percentile
            if data.get("uterine_pi"):
                try:
                    res = calculate_uta_pi(float(data["uterine_pi"]), int(round(data["ga_weeks_dop"])))
                    if res: 
                        data["uterine_pct"] = str(res['percentile'])
                        b2.caption(f"Uterine: {res['percentile']}th")
                        bar_uta = generate_percentile_bar(res['percentile'])
                        if bar_uta: b2.image(bar_uta, use_container_width=True)
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
    
    # --- THE ULTIMATE FIX ---
    # We re-fetch the session state right here to guarantee it is perfectly fresh.
    final_sr_data = st.session_state.get("sr_data", {})
    
    docx_bytes = build_docx_bytes(data, template_id, final_sr_data) 
    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=f"OB_Report_{p_name or 'Unknown'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
